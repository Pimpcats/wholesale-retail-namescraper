# Wholesale / Retail Name Scraper

Takes the text you copy from a wholesale marketplace page (e.g. Faire, Temu /
TikTok Shop wholesale) and pulls out **only the brand / store names** into an
Excel (`.xlsx`) worksheet.

It finds the brand by anchoring on each listing's **minimum-order line**
(`$X min`) — the brand name is the line right above it. Because of that, it
ignores product titles, `MSRP` prices, star ratings, badges ("New",
"Bestseller", "5 colors"), and shipping / discount lines, and it works even when
different listings have a different number of lines. (Underlines and links don't
survive copy-paste — copied text is always plain — so the `$X min` anchor is
what makes this reliable.)

There are two versions in this repo:

- **`index.html` — web version (no install).** Recommended for Windows. Runs
  entirely in your browser; nothing is uploaded. See **"Web version"** below.
- **`namescraper.py` — Python version.** Requires installing Python. See
  **"Python version"** below.

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

## Web version (no install — recommended for Windows)

Nothing to install, no pip, no Python. Pick whichever is easier:

### Option A — open a clickable URL (GitHub Pages)

Host it on your own GitHub for free so you get a link you can bookmark:

1. Push this repo to your GitHub account (already done if you're reading this there).
2. On GitHub, go to **Settings → Pages**.
3. Under **Build and deployment → Source**, choose **Deploy from a branch**.
4. Pick the branch that has `index.html` and the **`/ (root)`** folder, then **Save**.
5. Wait ~1 minute, then open the URL GitHub shows you
   (`https://<your-username>.github.io/wholesale-retail-namescraper/`).

### Option B — just open the file (zero setup)

1. Download `index.html` from GitHub (open the file → **Download raw file**).
2. Double-click it. It opens in your default browser and works offline
   (the CSV download needs no internet; the `.xlsx` download loads a small
   script from the web).

### Using the web version

1. Copy the page text from the marketplace.
2. Paste it into the box (or click **Paste from clipboard**).
3. The store names appear in the preview as you paste.
4. Click **Download Excel (.xlsx)** or **Download CSV** (CSV also opens in Excel).

To add more stores later, open the saved spreadsheet and paste the new batch
underneath — the tool itself always produces a fresh download.

---

## Python version

### Setup (one time)

1. Install [Python 3](https://www.python.org/downloads/) (on Windows, check
   "Add Python to PATH" during install).
2. Open a terminal / Command Prompt in this folder and run:

   ```
   pip install -r requirements.txt
   ```

### Using it — the easy way (window/GUI)

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

### Using it — command line

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
