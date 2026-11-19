#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Royal Enfield price scraper (source: 91Wheels brand page only).
Page: https://www.91wheels.com/bikes/royalenfield

Page structure found: every current model is a card with
  <a href="/bikes/royalenfield/<slug>">Royal Enfield <Model></a>
  "₹1.89 - ₹2.26 Lakh*"  +  "Ex-Showroom Price"
Upcoming models say "Expected Price" and discontinued models have no
"Ex-Showroom Price" label, so both are excluded automatically.
"""

import re
import sys
import time
import logging
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
BASE_URL = "https://www.91wheels.com"
PAGE_URL = "https://www.91wheels.com/bikes/royalenfield"
OUT_DIR = Path(__file__).resolve().parent
XLSX_FILE = OUT_DIR / "Royal_Enfield_Prices.xlsx"
CSV_FILE = OUT_DIR / "Royal_Enfield_Prices.csv"

# True  -> only the 10 models from the specification project
# False -> every current Royal Enfield model listed on the page
ONLY_TARGET_MODELS = True

TARGET_MODELS = {
    "classic-350": "Classic 350",
    "bullet-350": "Bullet 350",
    "hunter-350": "Hunter 350",
    "meteor-350": "Meteor 350",
    "himalayan-450": "Himalayan 450",
    "interceptor-650": "Interceptor 650",
    "continental-gt-650": "Continental GT 650",
    "super-meteor-650": "Super Meteor 650",
    "guerrilla-450": "Guerrilla 450",
    "shotgun-650": "Shotgun 650",
}

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
RETRIES = 3
NAV_TIMEOUT_MS = 60000

COL_MODEL = "Model Name"
COL_MIN = "Base Price - Minimum (INR)"
COL_MAX = "Top-End Price - Maximum (INR)"

RUPEE = "\u20b9"
PRICE_RE = re.compile(
    RUPEE + r"\s*([\d,]+(?:\.\d+)?)\s*(?:-|\u2013|to)?\s*(?:" + RUPEE
    + r"\s*([\d,]+(?:\.\d+)?))?\s*(Lakh|Lac|Crore|Cr)\b",
    re.I,
)
SLUG_RE = re.compile(r"/bikes/royalenfield/([a-z0-9\-]+)")

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("re_prices")


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def clean(text):
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).replace("\xa0", " ")).strip()


def slug_of(href):
    path = urlparse(urljoin(BASE_URL, href.split("#")[0].split("?")[0])).path.rstrip("/")
    m = SLUG_RE.fullmatch(path)
    return m.group(1) if m else None


def parse_price_text(text):
    """'₹1.89 - ₹2.26 Lakh*' -> (189000, 226000); '₹3.65 Lakh*' -> (365000, 365000)."""
    m = PRICE_RE.search(text)
    if not m:
        return None
    unit = m.group(3).lower()
    mult = 10_000_000 if unit.startswith("cr") else 100_000
    low = float(m.group(1).replace(",", ""))
    high = float(m.group(2).replace(",", "")) if m.group(2) else low
    low_i, high_i = int(round(low * mult)), int(round(high * mult))
    return min(low_i, high_i), max(low_i, high_i)


def display_name(raw_name, slug):
    if slug in TARGET_MODELS:
        return TARGET_MODELS[slug]
    return re.sub(r"^Royal Enfield\s+", "", clean(raw_name), flags=re.I)


# --------------------------------------------------------------------------- #
# Browser
# --------------------------------------------------------------------------- #
def scroll_page(page):
    try:
        for _ in range(12):
            page.evaluate("window.scrollBy(0, Math.max(document.body.scrollHeight / 8, 700))")
            page.wait_for_timeout(400)
        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(500)
    except Exception:
        pass


def fetch_html(page, url):
    for attempt in range(1, RETRIES + 1):
        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except PWTimeout:
                pass
            scroll_page(page)
            html = page.content()
            if resp is not None and resp.status >= 400:
                raise RuntimeError(f"HTTP {resp.status}")
            if len(html) < 5000:
                raise RuntimeError("page content too small")
            return html
        except Exception as exc:
            log.warning("Attempt %d/%d failed: %s", attempt, RETRIES, exc)
            time.sleep(2 * attempt)
    return None


# --------------------------------------------------------------------------- #
# Extraction
# --------------------------------------------------------------------------- #
def parse_cards_dom(html):
    """Find each model link, climb to its card, read the 'Ex-Showroom Price' range."""
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    results = {}
    for a in soup.find_all("a", href=True):
        slug = slug_of(a["href"])
        if not slug or slug in results:
            continue
        raw_name = clean(a.get("title") or a.get_text(" "))
        if not raw_name:
            continue

        node = a.parent
        while node is not None and node.name not in ("html", "body"):
            inner = {slug_of(x["href"]) for x in node.find_all("a", href=True)}
            inner.discard(None)
            if len(inner) > 1:
                break
            text = clean(node.get_text(" "))
            if "ex-showroom" in text.lower():
                price = parse_price_text(text)
                if price:
                    results[slug] = (display_name(raw_name, slug), price[0], price[1])
                    break
            node = node.parent
    return results


def parse_cards_text(html):
    """Fallback: line-based parsing (model name line followed by price + 'Ex-Showroom Price')."""
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    lines = [clean(x) for x in soup.get_text("\n").split("\n")]
    lines = [x for x in lines if x]

    slug_by_name = {"royal enfield " + v.lower(): k for k, v in TARGET_MODELS.items()}
    results = {}
    for i, line in enumerate(lines):
        slug = slug_by_name.get(line.lower())
        if not slug or slug in results:
            continue
        window = lines[i + 1:i + 4]
        if any("ex-showroom" in w.lower() for w in window):
            for w in window:
                price = parse_price_text(w)
                if price:
                    results[slug] = (TARGET_MODELS[slug], price[0], price[1])
                    break
    return results


# --------------------------------------------------------------------------- #
# Export
# --------------------------------------------------------------------------- #
def export_files(df):
    df.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(XLSX_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Prices", index=False)
        ws = writer.sheets["Prices"]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2, min_col=2, max_col=3):
            for cell in row:
                cell.number_format = "#,##0"
        ws.freeze_panes = "A2"
        for idx, col in enumerate(df.columns, start=1):
            longest = max([len(str(col))] + [len(str(v)) for v in df[col]])
            ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 4, 16), 40)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(user_agent=USER_AGENT, locale="en-IN",
                                      viewport={"width": 1366, "height": 900})
        context.route(
            "**/*",
            lambda route: route.abort()
            if route.request.resource_type in {"image", "media", "font"}
            else route.continue_(),
        )
        page = context.new_page()
        html = fetch_html(page, PAGE_URL)
        browser.close()

    if not html:
        log.error("Could not load %s", PAGE_URL)
        sys.exit(1)

    results = parse_cards_dom(html)
    if not results:
        log.warning("DOM card parsing found nothing; trying text fallback")
        results = parse_cards_text(html)
    if not results:
        log.error("No prices found on the page.")
        sys.exit(1)

    log.info("Models with prices found on page: %d", len(results))

    if ONLY_TARGET_MODELS:
        ordered = [(slug, results[slug]) for slug in TARGET_MODELS if slug in results]
        for slug, name in TARGET_MODELS.items():
            if slug not in results:
                log.warning("Price not found for: %s", name)
    else:
        ordered = list(results.items())

    rows = [{COL_MODEL: name, COL_MIN: low, COL_MAX: high} for _, (name, low, high) in ordered]
    df = pd.DataFrame(rows, columns=[COL_MODEL, COL_MIN, COL_MAX])

    export_files(df)

    print()
    print(df.to_string(index=False, formatters={COL_MIN: "{:,}".format, COL_MAX: "{:,}".format}))
    print()
    log.info("Saved: %s", XLSX_FILE)
    log.info("Saved: %s", CSV_FILE)


if __name__ == "__main__":
    main()