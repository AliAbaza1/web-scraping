"""Scrape the full catalogue of books.toscrape.com into a clean Excel report and CSV.

books.toscrape.com is a public sandbox built for practising web scraping.

Usage:
    python books_scraper.py                 # all 50 pages, listing data only
    python books_scraper.py --details       # also visit every book page (UPC, category, stock, description)
    python books_scraper.py --pages 2       # quick run on the first 2 pages
"""

from __future__ import annotations

import argparse
import logging
import re
import time
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from typing import Iterator, Optional
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Font
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/"
START_URL = urljoin(BASE_URL, "catalogue/page-1.html")
RATING_WORDS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("books_scraper")


@dataclass
class Book:
    title: str
    price_gbp: float
    rating: int
    in_stock: bool
    url: str
    category: Optional[str] = None
    stock_count: Optional[int] = None
    upc: Optional[str] = None
    description: Optional[str] = None


# ---------- HTTP ----------

def make_session() -> requests.Session:
    """Session with automatic retries on network errors and 429/5xx responses."""
    session = requests.Session()
    retry = Retry(total=4, backoff_factor=1, status_forcelist=(429, 500, 502, 503, 504))
    session.mount("http://", HTTPAdapter(max_retries=retry))
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.headers["User-Agent"] = "Mozilla/5.0 (portfolio scraper; +https://github.com/AliAbaza1)"
    return session


def get_soup(session: requests.Session, url: str) -> BeautifulSoup:
    response = session.get(url, timeout=20)
    response.raise_for_status()
    response.encoding = "utf-8"
    return BeautifulSoup(response.text, "html.parser")


# ---------- Parsing ----------

def parse_price(text: str) -> float:
    """'£51.77' or 'Â£51.77' -> 51.77"""
    match = re.search(r"\d+(?:\.\d+)?", text)
    if not match:
        raise ValueError(f"No price in {text!r}")
    return float(match.group())


def parse_rating(tag) -> int:
    classes = tag.get("class", []) if tag else []
    for name in classes:
        if name in RATING_WORDS:
            return RATING_WORDS[name]
    return 0


def parse_listing_page(soup: BeautifulSoup, page_url: str) -> tuple[list[Book], Optional[str]]:
    """Return the books on one catalogue page and the URL of the next page (or None)."""
    books = []
    for pod in soup.select("article.product_pod"):
        link = pod.select_one("h3 a")
        availability = pod.select_one("p.instock.availability, p.availability")
        books.append(
            Book(
                title=link["title"].strip(),
                price_gbp=parse_price(pod.select_one("p.price_color").get_text()),
                rating=parse_rating(pod.select_one("p.star-rating")),
                in_stock="in stock" in (availability.get_text().lower() if availability else ""),
                url=urljoin(page_url, link["href"]),
            )
        )
    next_link = soup.select_one("li.next a")
    next_url = urljoin(page_url, next_link["href"]) if next_link else None
    return books, next_url


def parse_detail_page(soup: BeautifulSoup) -> dict:
    """Extra fields from a single book page."""
    table = {
        row.th.get_text(strip=True): row.td.get_text(strip=True)
        for row in soup.select("table.table-striped tr")
        if row.th and row.td
    }
    crumbs = soup.select("ul.breadcrumb li a")
    category = crumbs[2].get_text(strip=True) if len(crumbs) >= 3 else None
    stock_match = re.search(r"(\d+)\s+available", table.get("Availability", ""))
    description_tag = soup.select_one("#product_description + p")
    return {
        "upc": table.get("UPC"),
        "category": category,
        "stock_count": int(stock_match.group(1)) if stock_match else 0,
        "description": description_tag.get_text(strip=True) if description_tag else None,
    }


# ---------- Crawling ----------

def crawl(session: requests.Session, max_pages: Optional[int], details: bool, delay: float) -> Iterator[Book]:
    url, page = START_URL, 0
    while url and (max_pages is None or page < max_pages):
        page += 1
        log.info("Page %d: %s", page, url)
        books, url = parse_listing_page(get_soup(session, url), url)
        for book in books:
            if details:
                try:
                    for key, value in parse_detail_page(get_soup(session, book.url)).items():
                        setattr(book, key, value)
                except requests.RequestException as exc:
                    log.warning("Skipped details for %s: %s", book.url, exc)
                time.sleep(delay)
            yield book
        time.sleep(delay)


# ---------- Output ----------

def clean(books: list[Book]) -> pd.DataFrame:
    df = pd.DataFrame([asdict(b) for b in books], columns=[f.name for f in fields(Book)])
    df = df.drop_duplicates(subset="url").reset_index(drop=True)
    df["title"] = df["title"].str.strip()
    return df


def summary_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    tables = {
        "By rating": df.groupby("rating")
        .agg(books=("title", "count"), avg_price_gbp=("price_gbp", "mean"))
        .round(2)
        .reset_index(),
    }
    if df["category"].notna().any():
        tables["By category"] = (
            df.groupby("category")
            .agg(books=("title", "count"), avg_price_gbp=("price_gbp", "mean"), avg_rating=("rating", "mean"))
            .round(2)
            .sort_values("books", ascending=False)
            .reset_index()
        )
    return tables


def export(df: pd.DataFrame, out_dir: Path) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path, xlsx_path = out_dir / "books.csv", out_dir / "books.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.dropna(axis=1, how="all").to_excel(writer, sheet_name="Books", index=False)
        for name, table in summary_tables(df).items():
            table.to_excel(writer, sheet_name=name, index=False)
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            for cell in sheet[1]:
                cell.font = Font(bold=True)
            for column in sheet.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in column)
                sheet.column_dimensions[column[0].column_letter].width = min(max(10, width + 2), 60)
    return csv_path, xlsx_path


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--pages", type=int, default=None, help="Stop after N catalogue pages (default: all)")
    parser.add_argument("--details", action="store_true", help="Visit each book page for UPC, category, stock, description")
    parser.add_argument("--delay", type=float, default=0.3, help="Seconds to wait between requests (be polite)")
    parser.add_argument("--out", type=Path, default=Path("output"), help="Output folder")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    started = time.time()
    df = clean(list(crawl(make_session(), args.pages, args.details, args.delay)))
    csv_path, xlsx_path = export(df, args.out)
    log.info("Saved %d books to %s and %s in %.0fs", len(df), xlsx_path, csv_path, time.time() - started)


if __name__ == "__main__":
    main()
