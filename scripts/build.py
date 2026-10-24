#!/usr/bin/env python3
"""Build the repository content from the SLR extraction spreadsheet.

Usage:
    pip install openpyxl
    python scripts/build.py path/to/Final_new_all_394.xlsx

Generates:
    README.md                 overview, statistics, taxonomy, full paper list
    papers/<citekey>.md       one page per paper with every extracted field
    data/papers.csv           all extracted fields, one row per paper
    data/references.bib       BibTeX for all papers
    docs/datasets.md          datasets ranked by number of papers using them
    docs/metrics.md           evaluation metrics ranked by usage
    docs/applications.md      positive applications of detection
"""
import csv
import html
import re
import sys
import urllib.parse
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent

# Column index in the "All" sheet -> (csv name, human label)
FIELDS = [
    (0, "id", "No."),
    (1, "category", "Category"),
    (2, "title", "Title"),
    (3, "venue_type", "Venue type"),
    (4, "publisher", "Publisher"),
    (5, "year", "Year"),
    (6, "country", "First-author country"),
    (7, "generation_method", "Deepfake generation method"),
    (8, "modalities_raw", "Modalities"),
    (9, "positive_applications", "Positive applications"),
    (10, "ethical_concerns", "Ethical concerns (generation)"),
    (11, "detection_limitations", "Limitations in detection"),
    (12, "generation_limitations", "Limitations in generation"),
    (13, "robustness", "Robustness / generalization"),
    (14, "data_collection", "Data collection techniques"),
    (15, "metrics", "Evaluation metrics"),
    (16, "detection_models", "Detection models"),
    (17, "datasets", "Datasets"),
    (18, "feature_extraction", "Feature extraction"),
    (19, "sota_baselines", "SOTA / baseline techniques"),
    (20, "summary", "Methodology summary"),
    (22, "bibtex", "BibTeX"),
    (23, "citekey", "Citation key"),
]
DETAIL_FIELDS = [f for f in FIELDS if f[1] in {
    "positive_applications", "ethical_concerns", "detection_limitations",
    "generation_limitations", "robustness", "data_collection", "metrics",
    "detection_models", "datasets", "feature_extraction", "sota_baselines",
    "generation_method"}]

EMPTY = {"", "none", "n/a", "na", "nan", "not mentioned", "not mentioned.",
         "not explicitly mentioned in the paper.", "not mentioned in the paper"}

PUBLISHER = {"science direct": "Elsevier", "springer nature": "Springer"}
COUNTRY = {"china": "China", "uk": "United Kingdom", "united states": "USA",
           "south korea": "Republic of Korea"}

# Canonical dataset name -> regex matched against the "Datasets" field
DATASETS = {
    "Weibo": r"weibo(?!-?21|\s*21)", "Weibo21": r"weibo\s*-?\s*21",
    "Twitter (MediaEval)": r"twitter|mediaeval", "Fakeddit": r"fakeddit",
    "PolitiFact": r"politifact", "GossipCop": r"gossipcop",
    "FakeNewsNet": r"fakenewsnet", "PHEME": r"pheme",
    "FakeSV": r"fakesv", "FakeTT": r"fakett", "FVC": r"\bfvc\b",
    "NewsCLIPpings": r"newsclippings", "DGM4": r"dgm4|dgm\^4",
    "FakeAVCeleb": r"fakeavceleb", "LAV-DF": r"lav-?df",
    "AV-Deepfake1M": r"av-?deepfake1m", "DFDC": r"\bdfdc\b|deepfake detection challenge",
    "FaceForensics++": r"faceforensics|ff\+\+", "Celeb-DF": r"celeb-?df",
    "DeeperForensics": r"deeperforensics", "DF-TIMIT": r"df-?timit",
    "KoDF": r"\bkodf\b", "ASVspoof": r"asvspoof", "WildDeepfake": r"wild-?deepfake",
    "ReCOVery": r"recovery", "CoAID": r"coaid", "MM-COVID": r"mm-?covid",
}
METRICS = {
    "Accuracy": r"accura", "F1-score": r"\bf1|f-score|f-measure",
    "Precision": r"precision", "Recall": r"recall", "AUC / ROC-AUC": r"\bauc\b|roc",
    "EER": r"\beer\b|equal error", "Average Precision (AP/mAP)": r"\bm?ap\b|average precision",
    "IoU": r"\biou\b", "Macro-F1": r"macro", "MCC": r"\bmcc\b|matthews",
}
# Positive-application categories (detection), keyword-assigned
APPLICATIONS = {
    "Social-media fake news & misinformation mitigation": r"fake news|misinformation|disinformation|rumou?r",
    "Short-video / video platform misinformation": r"short.?video|tiktok|youtube|kuaishou|video platform",
    "Fact-checking & content moderation support": r"fact.?check|moderat|human.?review|human.?in.?the.?loop|assessor",
    "Elections & political integrity": r"election|politic",
    "Public health & crisis response": r"covid|health|vaccine|pandemic|crisis|crises|disaster|panic",
    "Social stability, public trust & national security": r"stabilit|public trust|social order|\bnational|harmony|unrest",
    "Identity, biometric authentication & fraud prevention": r"biometric|authentication system|identity|impersonat|fraud|spoof",
    "Privacy & reputation protection": r"privacy|reputation|blackmail|cyberbully|cyberstalk|scam",
    "Digital forensics & legal evidence": r"forensic|legal|law enforcement|evidence|locali[sz]",
    "Media integrity & journalism": r"media integrity|journalis|authenticit|copyright|traceab",
    "Cybersecurity & information security": r"cyber|information security|network security",
    "Real-time communication security": r"real.?time|\blive\b|live-?stream|conferenc|streaming|edge device|\bedge\b|mobile|surveillance",
    "E-commerce & economic protection": r"fake review|consumer|marketplace|e-commerce|financ|econom",
    "Low-resource & multilingual detection": r"hindi|bangla|amharic|indian|multilingual|low.?resource",
    "Explainability & user awareness": r"explainab|interpretab|awareness|user trust|education",
}


def clean(v):
    if v is None:
        return ""
    s = html.unescape(str(v)).replace("\xa0", " ").replace("\r", "").strip()
    return re.sub(r"[ \t]+\n", "\n", s)


def is_empty(s):
    return s.strip().lower().rstrip(".") in {e.rstrip(".") for e in EMPTY}


def norm_modalities(raw):
    s = raw.lower()
    return [m for m in ("Text", "Image", "Audio", "Video") if m.lower() in s]


def norm_venue(raw):
    s = raw.lower()
    if "workshop" in s:
        return "Workshop"
    if "journal" in s:
        return "Journal"
    return "Conference" if "conf" in s or "comf" in s else raw.title()


def load(xlsx):
    ws = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)["All"]
    papers = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row[2]:
            continue
        p = {name: clean(row[i]) if i < len(row) else "" for i, name, _ in FIELDS}
        p["id"] = str(int(float(p["id"]))) if p["id"] else ""
        p["year"] = str(int(float(p["year"]))) if p["year"] else ""
        p["title"] = " ".join(p["title"].split())
        p["citekey"] = p["citekey"].strip()
        p["category"] = "Deepfake" if "deep" in p["category"].lower() else "Fake News"
        p["venue_type"] = norm_venue(p["venue_type"])
        p["publisher"] = PUBLISHER.get(p["publisher"].lower(), p["publisher"])
        c = " ".join(p["country"].split())
        p["country"] = COUNTRY.get(c.lower(), c)
        p["modalities"] = norm_modalities(p["modalities_raw"])
        bib = p["bibtex"]
        if bib and not bib.startswith("@"):
            bib = "@" + bib  # fix entries missing the leading @
        p["bibtex"] = bib
        papers.append(p)
    return papers


def scholar(title):
    return "https://scholar.google.com/scholar?q=" + urllib.parse.quote(title)


def md_escape(s):
    return s.replace("|", "\\|")


def count_patterns(papers, field, patterns):
    hits = defaultdict(list)
    for p in papers:
        text = p[field].lower()
        for name, pat in patterns.items():
            if re.search(pat, text):
                hits[name].append(p)
    return sorted(hits.items(), key=lambda kv: -len(kv[1]))


def table(headers, rows):
    out = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
    out += ["| " + " | ".join(r) + " |" for r in rows]
    return "\n".join(out)


def paper_row(p, link_prefix="papers/"):
    return [p["year"],
            f"[{md_escape(p['title'])}]({link_prefix}{p['citekey']}.md)",
            ", ".join(p["modalities"]) or "-",
            f"{p['venue_type']} · {p['publisher']}",
            f"[Scholar]({scholar(p['title'])})"]


def keep_lines(s):
    """Keep the spreadsheet's line breaks when rendered as Markdown."""
    return re.sub(r"(?<!\n)\n(?!\n)", "  \n", s)


def write_paper_pages(papers):
    d = ROOT / "papers"
    d.mkdir(exist_ok=True)
    for p in papers:
        lines = [f"# {p['title']}", "",
                 f"**Category:** {p['category']} · **Modalities:** {', '.join(p['modalities']) or '-'}  ",
                 f"**Venue:** {p['venue_type']} ({p['publisher']}, {p['year']}) · "
                 f"**First-author country:** {p['country'] or '-'}  ",
                 f"**Citation key:** `{p['citekey']}` · [Google Scholar]({scholar(p['title'])})", ""]
        if p["summary"] and not is_empty(p["summary"]):
            lines += ["## Methodology summary", "", keep_lines(p["summary"]), ""]
        for _, name, label in DETAIL_FIELDS:
            if p[name] and not is_empty(p[name]):
                lines += [f"## {label}", "", keep_lines(p[name]), ""]
        if p["bibtex"]:
            lines += ["## BibTeX", "", "```bibtex", p["bibtex"], "```", ""]
        lines += ["[← Back to the paper list](../README.md#paper-list)", ""]
        (d / f"{p['citekey']}.md").write_text("\n".join(lines), encoding="utf-8")


def write_data(papers):
    (ROOT / "data").mkdir(exist_ok=True)
    cols = [n for _, n, _ in FIELDS] + ["modalities"]
    with open(ROOT / "data/papers.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for p in papers:
            w.writerow({**p, "modalities": ";".join(p["modalities"])})
    bib = "\n\n".join(p["bibtex"] for p in papers if p["bibtex"])
    (ROOT / "data/references.bib").write_text(bib + "\n", encoding="utf-8")


def ranked_doc(title, intro, ranked, total):
    lines = [f"# {title}", "", intro, "",
             table(["Name", "# Papers", "Share"],
                   [[n, str(len(ps)), f"{100 * len(ps) / total:.1f}%"] for n, ps in ranked]), ""]
    for n, ps in ranked:
        lines += [f"## {n}", ""]
        lines += [f"- [{md_escape(p['title'])}](../papers/{p['citekey']}.md) ({p['year']})"
                  for p in sorted(ps, key=lambda p: p["year"], reverse=True)]
        lines.append("")
    lines += ["[← Back to README](../README.md)", ""]
    return "\n".join(lines)


def write_docs(papers):
    d = ROOT / "docs"
    d.mkdir(exist_ok=True)
    n = len(papers)
    note = ("Counts come from keyword matching on the extracted fields, so they are "
            "approximate; open a paper page for the exact wording.")
    (d / "datasets.md").write_text(ranked_doc(
        "Datasets", "Benchmark datasets used across the reviewed papers. " + note,
        count_patterns(papers, "datasets", DATASETS), n), encoding="utf-8")
    (d / "metrics.md").write_text(ranked_doc(
        "Evaluation metrics", "Evaluation metrics reported across the reviewed papers. " + note,
        count_patterns(papers, "metrics", METRICS), n), encoding="utf-8")
    (d / "applications.md").write_text(ranked_doc(
        "Positive applications of detection",
        "Societal benefits of detection reported by the reviewed papers. "
        "A paper can fall under several categories. " + note,
        count_patterns(papers, "positive_applications", APPLICATIONS), n), encoding="utf-8")


def write_readme(papers):
    n = len(papers)
    cat = Counter(p["category"] for p in papers)
    years = Counter(p["year"] for p in papers)
    pubs = Counter(p["publisher"] for p in papers)
    venues = Counter(p["venue_type"] for p in papers)
    countries = Counter(p["country"] for p in papers if p["country"])
    mods = Counter("+".join(p["modalities"]) or "Unspecified" for p in papers)
    datasets = count_patterns(papers, "datasets", DATASETS)[:10]

    def counts(c, k=None):
        return ", ".join(f"{a} ({b})" for a, b in c.most_common(k))

    lines = [
        "# Awesome Multimodal Deepfake & Fake News Detection",
        "",
        "![Papers](https://img.shields.io/badge/papers-%d-blue) "
        "![Years](https://img.shields.io/badge/years-%s--%s-green) "
        "![License](https://img.shields.io/badge/license-CC%%20BY%%204.0-lightgrey)"
        % (n, min(years), max(years)),
        "",
        "A curated, structured collection of **%d papers** on multimodal deepfake and "
        "fake news detection (text, image, audio, video), compiled for a systematic "
        "literature review. Every paper has its own page with the extracted methodology, "
        "datasets, metrics, features, limitations, robustness findings and BibTeX." % n,
        "",
        "> **Paper:** _Add your SLR title, authors and link here once available._",
        "",
        "## Contents",
        "- [Research questions](#research-questions)",
        "- [Statistics](#statistics)",
        "- [Resources](#resources)",
        "- [Paper list](#paper-list)",
        "- [Citation](#citation)",
        "- [Contributing](#contributing)",
        "",
        "## Research questions",
        "1. **RQ1** What is a deepfake, and which kinds exist in audio, text, image and video?",
        "2. **RQ2** What are the benefits, risks and challenges in generation and detection?",
        "3. **RQ3** Which data collection techniques, benchmarks (metrics, datasets, baselines) "
        "and feature extraction methods are used in detection?",
        "4. **RQ4** What are the state-of-the-art detection techniques in each modality?",
        "",
        "## Statistics",
        "",
        table(["Dimension", "Breakdown"], [
            ["Category", counts(cat)],
            ["Venue type", counts(venues)],
            ["Publisher", counts(pubs)],
            ["Year", ", ".join(f"{y} ({years[y]})" for y in sorted(years))],
            ["Modalities", counts(mods, 8)],
            ["Top countries (first author)", counts(countries, 8)],
            ["Top datasets", ", ".join(f"{k} ({len(v)})" for k, v in datasets)],
        ]),
        "",
        "",
        "![Papers per year](figures/fig_papers_per_year.png)",
        "",
        "<p float=\"left\"><img src=\"figures/fig_venue_type.png\" width=\"49%\"/> "
        "<img src=\"figures/fig_publisher.png\" width=\"49%\"/></p>",
        "<p float=\"left\"><img src=\"figures/fig_countries.png\" width=\"49%\"/> "
        "<img src=\"figures/fig_modalities.png\" width=\"49%\"/></p>",
        "",
        "Figures are generated by `scripts/plots.py`.",
        "",
        "## Resources",
        "",
        "| Resource | Description |",
        "|---|---|",
        "| [`data/papers.csv`](data/papers.csv) | All extracted fields for every paper (spreadsheet-friendly) |",
        "| [`data/references.bib`](data/references.bib) | BibTeX for all papers |",
        "| [`docs/datasets.md`](docs/datasets.md) | Benchmark datasets ranked by usage, with papers per dataset |",
        "| [`docs/metrics.md`](docs/metrics.md) | Evaluation metrics ranked by usage |",
        "| [`docs/applications.md`](docs/applications.md) | Positive applications of detection, by category |",
        "| [`papers/`](papers/) | One page per paper with all extracted information |",
        "",
        "## Paper list",
        "",
    ]
    headers = ["Year", "Title", "Modalities", "Venue", "Link"]
    for c in ("Deepfake", "Fake News"):
        lines += [f"### {c} detection", ""]
        groups = defaultdict(list)
        for p in papers:
            if p["category"] == c:
                groups[", ".join(p["modalities"]) or "Unspecified"].append(p)
        for g, ps in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            ps.sort(key=lambda p: (p["year"], p["title"]), reverse=True)
            lines += [f"<details><summary><b>{g}</b> ({len(ps)} papers)</summary>", "",
                      table(headers, [paper_row(p) for p in ps]), "", "</details>", ""]
    lines += [
        "## Citation",
        "",
        "If this collection helps your research, please cite:",
        "",
        "```bibtex",
        "@misc{awesome_multimodal_deepfake_detection,",
        "  title  = {Awesome Multimodal Deepfake and Fake News Detection},",
        "  author = {Your Name and Co-authors},",
        "  year   = {2026},",
        "  url    = {https://github.com/<user>/<repo>}",
        "}",
        "```",
        "",
        "## Contributing",
        "",
        "Missing a paper or spotted an error? Open an issue or a pull request, see "
        "[CONTRIBUTING.md](CONTRIBUTING.md).",
        "",
        "## License",
        "",
        "Content and data: [CC BY 4.0](LICENSE). Scripts: MIT.",
        "",
    ]
    (ROOT / "README.md").write_text("\n".join(lines), encoding="utf-8")


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    papers = load(sys.argv[1])
    keys = Counter(p["citekey"] for p in papers)
    dupes = [k for k, v in keys.items() if v > 1 or not k]
    if dupes:
        sys.exit(f"Duplicate or empty citation keys: {dupes}")
    write_paper_pages(papers)
    write_data(papers)
    write_docs(papers)
    write_readme(papers)
    print(f"Built repository content for {len(papers)} papers in {ROOT}")


if __name__ == "__main__":
    main()
