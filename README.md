# Wholesale / Retail Name Scraper

A small desktop program that takes the text you copy from a wholesale
marketplace page (e.g. Faire) and pulls out **only the store names** — the first
line of each listing — into an Excel (`.xlsx`) worksheet.

It automatically removes the noise from each listing:

- the minimum-order line (`$150 min`, `$0 min`, …)
- star ratings on their own line (`4.8`)
- shipping / discount lines (`Free shipping`, `Up to 5% off`, `Up to $150 off + free shipping`, …)

## What you get

Given pasted text like:

```
Chic Couture Inc

$150 min

Free shipping


Bailey Rose

4.8

$50 min

Free shipping
```

You get a spreadsheet with one column:

| Store Name        |
|-------------------|
| Chic Couture Inc  |
| Bailey Rose       |

## Setup (one time)

1. Install [Python 3](https://www.python.org/downloads/) (on Windows, check
   "Add Python to PATH" during install).
2. Open a terminal / Command Prompt in this folder and run:

   ```
   pip install -r requirements.txt
   ```

## Using it — the easy way (window/GUI)

Double-click `namescraper.py`, or run:

```
python namescraper.py
```

Then:

1. Copy the page text from the marketplace.
2. Click **Paste from clipboard** (or paste with Ctrl+V into the box).
3. The extracted store names appear in the preview list.
4. Click **Save to Excel** and choose where to save the file.

Options:

- **Append to existing file** — adds new names to the bottom of an existing
  spreadsheet instead of starting over.
- **Skip duplicates** — won't add a store name that's already in the file.

## Using it — command line

Read straight from your clipboard and append to `store_names.xlsx`:

```
python namescraper.py --clipboard
```

Read from a text file and choose the output name:

```
python namescraper.py listings.txt -o my_stores.xlsx
```

Other flags:

- `--overwrite` — replace the Excel file instead of appending.
- `--allow-duplicates` — keep duplicate names.
- `--gui` — force the window even when other arguments are given.
