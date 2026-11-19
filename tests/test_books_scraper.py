from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup

import books_scraper as bs

FIXTURES = Path(__file__).parent / "fixtures"


def soup(name):
    return BeautifulSoup((FIXTURES / name).read_text(encoding="utf-8"), "html.parser")


def test_parse_price_handles_mojibake():
    assert bs.parse_price("Â£51.77") == 51.77
    assert bs.parse_price("£9") == 9.0


def test_listing_page_fields_and_next_link():
    books, next_url = bs.parse_listing_page(soup("page-1.html"), bs.START_URL)
    assert [b.title for b in books] == ["A Light in the Attic", "Tipping the Velvet"]
    assert books[0].price_gbp == 51.77 and books[0].rating == 3 and books[0].in_stock
    assert books[0].url == "https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html"
    assert next_url == "https://books.toscrape.com/catalogue/page-2.html"


def test_last_page_has_no_next():
    _, next_url = bs.parse_listing_page(soup("page-2.html"), "https://books.toscrape.com/catalogue/page-2.html")
    assert next_url is None


def test_detail_page():
    d = bs.parse_detail_page(soup("detail.html"))
    assert d == {
        "upc": "a897fe39b1053632",
        "category": "Poetry",
        "stock_count": 22,
        "description": "It's hard to imagine a world without A Light in the Attic.",
    }


class FakeSession:
    """Serves fixture pages instead of hitting the network."""

    def get(self, url, timeout=None):
        name = "detail.html" if url.endswith("/index.html") else url.rsplit("/", 1)[-1]
        resp = type("R", (), {})()
        resp.text = (FIXTURES / name).read_text(encoding="utf-8")
        resp.encoding = "utf-8"
        resp.raise_for_status = lambda: None
        return resp


def test_crawl_follows_pagination_and_exports(tmp_path):
    books = list(bs.crawl(FakeSession(), max_pages=None, details=True, delay=0))
    assert len(books) == 3 and all(b.category == "Poetry" for b in books)

    df = bs.clean(books + books)  # duplicates are removed
    assert len(df) == 3

    csv_path, xlsx_path = bs.export(df, tmp_path)
    assert pd.read_csv(csv_path).shape[0] == 3
    sheets = pd.read_excel(xlsx_path, sheet_name=None)
    assert set(sheets) == {"Books", "By rating", "By category"}
    assert sheets["By category"].loc[0, "books"] == 3


def test_max_pages_limit():
    books = list(bs.crawl(FakeSession(), max_pages=1, details=False, delay=0))
    assert len(books) == 2
