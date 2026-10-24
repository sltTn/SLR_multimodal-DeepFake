# Contributing

Thanks for helping keep this list accurate and up to date.

## Suggest a paper
Open an issue with:
- Title, venue, year and a link (DOI or arXiv)
- Category: Deepfake or Fake News
- Modalities: text, image, audio, video
- Datasets and evaluation metrics used

## Fix an error
Open an issue or a pull request that edits the relevant page in `papers/` and say what was wrong.

## Maintainers: regenerating the repository
All Markdown, CSV and BibTeX files are generated from the extraction spreadsheet:

```bash
pip install -r requirements.txt
python scripts/build.py path/to/Final_new_all_394.xlsx
pip install matplotlib
python scripts/plots.py path/to/Final_new_all_394.xlsx
```

Edit the spreadsheet, rerun the script and commit the result so pages never drift from the data.
