# Google Sheet to 800x480 PNG

This repository downloads a public Google Spreadsheet CSV, renders its first rows as a strictly black-and-white 800x480 PNG, and publishes it with GitHub Pages.

No npm or Node.js is used.

## Replace the old repository files

Delete the old project files, then upload everything from this project. The workflow file must remain at exactly:

```text
.github/workflows/pages.yml
```

Because `.github` begins with a dot, the safest way to create the workflow on github.com is:

1. Open the repository's **Code** tab.
2. Select **Add file -> Create new file**.
3. Enter `.github/workflows/pages.yml` as the filename.
4. Paste the contents of this project's `pages.yml` file and commit it to `main`.

Upload the other files normally:

```text
src/generate.py
requirements.txt
.gitignore
README.md
```

## Enable GitHub Pages

1. Open **Settings -> Pages**.
2. Under **Build and deployment**, set **Source** to **GitHub Actions**.
3. Open the **Actions** tab.
4. The workflow should run automatically after the commit. It can also be run manually using **Run workflow**.

After a successful run, the site will be available at:

```text
https://YOUR-USERNAME.github.io/YOUR-REPOSITORY/
```

The direct PNG URL will be:

```text
https://YOUR-USERNAME.github.io/YOUR-REPOSITORY/sheet.png
```

## Updating

The workflow runs:

- whenever code is pushed to `main`;
- when manually started from the Actions tab; and
- once per hour, at minute 17.

GitHub may delay scheduled workflows during periods of heavy demand.

## What the image contains

The renderer uses:

- the first non-empty CSV row as the header;
- up to five columns;
- up to nine data rows;
- black and white only, using a 1-bit PNG;
- automatic truncation for text that does not fit.

To change the layout, edit the constants near the top of `src/generate.py`.

## Use a different sheet

Either replace `SHEET_CSV_URL` in `src/generate.py`, or set a `SHEET_CSV_URL` environment variable in the workflow.
