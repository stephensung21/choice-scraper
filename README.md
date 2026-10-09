# CHOICE product scraper

Collects CHOICE (choice.com.au) reviews for the products you're shopping for
(air fryers, robot vacuums, kettles, TVs by default) into one spreadsheet,
sorted by CHOICE score. It needs your own CHOICE membership and is meant for
personal use. Check CHOICE's terms of use before running it.

## Setup (once)

Requires Python 3.9+.

```bash
git clone <this repo> && cd choice-scraper
python -m venv .venv
# macOS/Linux:
source .venv/bin/activate
# Windows:
.venv\Scripts\activate

pip install -r requirements.txt
playwright install chromium
```

## Run

```bash
python scrape.py login              # a browser opens: log in to CHOICE, then press Enter in the terminal
python scrape.py scrape --limit 2   # quick test: 2 products per category
python scrape.py scrape             # everything
```

Your password is never stored. The script keeps only the logged-in browser
session in `.browser-profile/`, which git ignores.

## Output

- `output/choice_products.xlsx`
  - **All products**: name, brand, price, CHOICE score, recommended / Best Buy, pros, cons, verdict, link
  - one sheet per category with every spec and review section it found
- `output/<category>.json`: everything captured, including the full page text
- `output/products/...`: one file per product. A re-run skips products already saved, so an interrupted run carries on where it stopped. Use `--refresh` to re-scrape.

Recommended / Best Buy / score are picked out of the page text by pattern
matching, so check them against the site for anything you're about to buy.

## Options

```bash
python scrape.py scrape kettles tvs   # only some categories (keys from categories.json)
python scrape.py scrape --delay 5     # slower (default ~3-6 s between pages)
python scrape.py scrape --headless    # no browser window
python scrape.py export               # rebuild the spreadsheet from saved data
```

## If something doesn't work

- **"no product links found"**: the category URL in `categories.json` is
  probably wrong. The script pauses and asks you to browse to the right
  listing page in the open browser window, then press Enter. It prints the
  URL to put in `categories.json` for next time. If links still aren't found,
  add a `"link_pattern"` regex for product URLs to that category.
- **"looks locked - are you logged in?"**: run `python scrape.py login` again.
- **Fields missing or wrong**: run `python scrape.py debug <product-url>` and
  share `debug/page-extract.json` (or `debug/page.html`) so the extraction can
  be tuned to CHOICE's real page layout.
