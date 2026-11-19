# Book Catalogue Scraper

Scrapes the **entire catalogue** of [books.toscrape.com](https://books.toscrape.com) (1,000 books across 50 pages) and delivers a clean, ready-to-use **Excel report** and **CSV**.

The site is a public sandbox made for practising web scraping, so this project shows the full workflow a real client job needs: crawling every page, parsing, cleaning, de-duplicating, and handing over a file a non-technical person can open.

## What it does

- Follows pagination automatically until the last page
- Extracts title, price, star rating, availability and link for every book
- Optional `--details` mode visits each book page for **category, UPC, exact stock count and description**
- Cleans the data: numeric prices, ratings as 1–5, duplicates removed
- Exports `books.xlsx` with formatted headers, frozen first row, sized columns and summary sheets (**by rating**, **by category**), plus `books.csv` (UTF-8, opens correctly in Excel)
- Retries failed requests and waits between requests so it doesn't overload the site

## Run it

```bash
pip install -r requirements.txt

python books_scraper.py              # all 50 pages, listing data (~1 minute)
python books_scraper.py --details    # + category, UPC, stock, description (~1,000 extra requests)
python books_scraper.py --pages 2    # quick test
```

Files are written to `output/`.

| Option | Default | Meaning |
|---|---|---|
| `--pages N` | all | stop after N catalogue pages |
| `--details` | off | visit every book page for extra fields |
| `--delay S` | 0.3 | seconds between requests |
| `--out DIR` | `output` | output folder |

## Output

One row per book (sample):

| title | price_gbp | rating | in_stock | category | stock_count | upc |
|---|---|---|---|---|---|---|
| A Light in the Attic | 51.77 | 3 | True | Poetry | 22 | a897fe39b1053632 |

<!-- Add a screenshot of output/books.xlsx here: ![Excel output](docs/excel-output.png) -->

## Tests

```bash
python -m pytest -q
```

The tests run offline against saved HTML pages and cover parsing, pagination, detail pages, de-duplication and the Excel/CSV export.

## Project structure

```
books_scraper.py        # the scraper (crawl → parse → clean → export)
tests/                  # offline tests with HTML fixtures
notebooks/              # earlier practice notebooks (single-page scraping, currency rate, weather API)
```

## Tech

Python · requests · BeautifulSoup · pandas · openpyxl · pytest
