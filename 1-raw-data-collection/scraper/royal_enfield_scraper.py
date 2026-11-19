#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Royal Enfield specification scraper (source: 91Wheels only) - BASE VARIANT ONLY.

Per model (3 page loads):
  1. /specifications page -> names the "Base Variant" + base-variant spec tables.
  2. Model page           -> variant links + ABS type shown in the variants table.
  3. Base variant page    -> full variant-specific specs (JSON meta + tables + overview text).
"""

import json
import random
import re
import sys
import time
import logging
from datetime import date
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
OUT_DIR = Path(__file__).resolve().parent
XLSX_FILE = OUT_DIR / "Royal_Enfield_Bike_Specifications.xlsx"
CSV_FILE = OUT_DIR / "Royal_Enfield_Bike_Specifications.csv"

MODELS = [
    ("Classic 350", "https://www.91wheels.com/bikes/royalenfield/classic-350/specifications"),
    ("Bullet 350", "https://www.91wheels.com/bikes/royalenfield/bullet-350/specifications"),
    ("Hunter 350", "https://www.91wheels.com/bikes/royalenfield/hunter-350/specifications"),
    ("Meteor 350", "https://www.91wheels.com/bikes/royalenfield/meteor-350/specifications"),
    ("Himalayan 450", "https://www.91wheels.com/bikes/royalenfield/himalayan-450/specifications"),
    ("Interceptor 650", "https://www.91wheels.com/bikes/royalenfield/interceptor-650/specifications"),
    ("Continental GT 650", "https://www.91wheels.com/bikes/royalenfield/continental-gt-650/specifications"),
    ("Super Meteor 650", "https://www.91wheels.com/bikes/royalenfield/super-meteor-650/specifications"),
    ("Guerrilla 450", "https://www.91wheels.com/bikes/royalenfield/guerrilla-450/specifications"),
    ("Shotgun 650", "https://www.91wheels.com/bikes/royalenfield/shotgun-650/specifications"),
]

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
RETRIES = 3
NAV_TIMEOUT_MS = 60000

COLUMNS = [
    "Model Name",
    "Variant",
    "Engine Displacement (cc)",
    "Engine Type",
    "Cooling System",
    "Maximum Power (PS)",
    "Maximum Torque (Nm)",
    "Transmission (Gears)",
    "Mileage (kmpl)",
    "Top Speed (kmph)",
    "Kerb Weight (kg)",
    "Fuel Tank Capacity (Litres)",
    "Front Suspension",
    "Rear Suspension",
    "Front Brake",
    "Rear Brake",
    "ABS (Single/Dual Channel)",
    "Source_URL",
    "Scraping_Date",
]

RESERVED_SLUGS = {
    "specifications", "images", "colours", "colors", "mileage", "reviews", "qna",
    "offers", "expert-review", "news", "videos", "dealers", "compare", "emi",
    "user-reviews", "variants", "price", "on-road-price", "service-centers",
    "service-cost", "features", "brochure", "spare-parts", "accessories",
}

FIELD_KEYS = {
    "displacement": ["displacement", "enginecapacity", "enginedisplacement"],
    "engine_type": ["enginetype"],
    "cooling": ["coolingsystem", "cooling", "coolingtype"],
    "power": ["maximumpower", "maxpower", "power"],
    "torque": ["maximumtorque", "maxtorque", "torque"],
    "gears": ["noofgears", "numberofgears", "gearbox", "transmission", "gears"],
    "mileage": ["araimileage", "claimedmileage", "mileagearai", "mileage", "fuelefficiency"],
    "top_speed": ["maxspeed", "topspeed"],
    "kerb": ["unladdenkerbweight", "unladenkerbweight", "kerbweight"],
    "tank": ["fueltanksize", "fueltankcapacity"],
    "front_susp": ["frontsuspension", "frontsuspensiontype"],
    "rear_susp": ["rearsuspension", "rearsuspensiontype"],
    "front_brake": ["frontbraketype", "frontbrake", "fronttype"],
    "rear_brake": ["rearbraketype", "rearbrake", "reartype"],
    "abs": ["brakingsystem", "abs", "absmode", "typeofabs", "antilockbrakingsystem"],
}
ALL_KEYS = {k for keys in FIELD_KEYS.values() for k in keys}

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("re_scraper")


# --------------------------------------------------------------------------- #
# Text helpers
# --------------------------------------------------------------------------- #
def clean(text):
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text).replace("\xa0", " ")).strip()


def norm(text):
    return re.sub(r"[^a-z0-9]", "", clean(text).lower())


RANGE_RE = re.compile(r"\d\s*(?:-|\u2013|\bto\b)\s*\d", re.I)
NUM_RE = re.compile(r"\d+(?:\.\d+)?")
ABS_ROW_RE = re.compile(r"(single|dual)[\s-]*channel\s*abs", re.I)
ARAI_RE = re.compile(r"ARAI\s+Mileage\s+of\s+(\d+(?:\.\d+)?)\s*kmpl", re.I)


def to_number(value):
    value = round(float(value), 2)
    return int(value) if value == int(value) else value


def make_num_parser(lo, hi, integer=False):
    def parser(raw):
        if raw is None:
            return None
        s = clean(raw).split("@")[0].replace(",", "")
        if not s or RANGE_RE.search(s):
            return None
        m = NUM_RE.search(s)
        if not m:
            return None
        val = float(m.group())
        if not (lo <= val <= hi):
            return None
        return int(val) if integer else to_number(val)
    return parser


def parse_power(raw):
    s = clean(raw).split("@")[0]
    m = re.search(r"(\d+(?:\.\d+)?)\s*(ps|bhp|kw|hp)\b", s, re.I)
    if not m:
        return None
    val, unit = float(m.group(1)), m.group(2).lower()
    if unit in ("bhp", "hp"):
        val *= 1.01387
    elif unit == "kw":
        val *= 1.35962
    return to_number(val) if 3 <= val <= 150 else None


def parse_torque(raw):
    s = clean(raw).split("@")[0]
    m = re.search(r"(\d+(?:\.\d+)?)\s*(nm|lb[\s-]?ft)", s, re.I)
    if not m:
        return None
    val, unit = float(m.group(1)), m.group(2).lower()
    if unit.startswith("lb"):
        val *= 1.35582
    return to_number(val) if 3 <= val <= 200 else None


def parse_abs(raw):
    t = clean(raw).lower()
    if "abs" not in t:
        return None
    if "dual" in t:
        return "Dual Channel"
    if "single" in t:
        return "Single Channel"
    return None


def parse_susp(raw):
    s = clean(raw)
    if not s or len(s) > 200 or s.lower() in ("-", "n/a", "na"):
        return None
    if norm(s) in ALL_KEYS or (len(s) < 30 and s.lower().endswith("travel")):
        return None
    return re.sub(r"\s*&\s*", ", ", s)


def parse_brake(raw):
    s = clean(raw)
    if not s or len(s) > 80 or norm(s) in ALL_KEYS:
        return None
    if not re.search(r"disc|drum|caliper|brake", s, re.I):
        return None
    return s


def parse_engine_raw(raw):
    s = clean(raw)
    if not s or len(s) > 150 or norm(s) in ALL_KEYS:
        return None
    return s


def parse_cooling_raw(raw):
    s = clean(raw)
    if s and len(s) < 80 and "cool" in s.lower():
        return re.sub(r"\s*&\s*", ", ", s)
    return None


def split_engine(raw):
    """'Single Cylinder& 4 Stroke& Air-Oil Cooled' -> ('Single Cylinder, 4 Stroke', 'Air-Oil Cooled')."""
    if not raw:
        return None, None
    parts = [clean(p) for p in re.split(r"&|,", raw) if clean(p)]
    cooling = [p for p in parts if "cool" in p.lower()]
    others = [p for p in parts if "cool" not in p.lower()]
    return (", ".join(others) if others else None), (", ".join(cooling) if cooling else None)


def norm_url(href):
    path = urlparse(urljoin(BASE_URL, href.split("#")[0].split("?")[0])).path.rstrip("/")
    return BASE_URL + path


# --------------------------------------------------------------------------- #
# Browser helpers
# --------------------------------------------------------------------------- #
def scroll_page(page):
    try:
        for _ in range(6):
            page.evaluate("window.scrollBy(0, Math.max(document.body.scrollHeight / 5, 600))")
            page.wait_for_timeout(250)
        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(300)
    except Exception:
        pass


def fetch_html(page, url):
    for attempt in range(1, RETRIES + 1):
        try:
            resp = page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
            try:
                page.wait_for_load_state("networkidle", timeout=12000)
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
            log.warning("Attempt %d/%d failed for %s: %s", attempt, RETRIES, url, exc)
            time.sleep(2 * attempt)
    return None


# --------------------------------------------------------------------------- #
# Extraction
# --------------------------------------------------------------------------- #
def find_base_variant_name(html):
    """On the /specifications page the base variant name sits right above the 'Base Variant' label."""
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    lines = [clean(x) for x in soup.get_text("\n").split("\n")]
    lines = [x for x in lines if x]
    for i, line in enumerate(lines):
        if line.lower().startswith("base variant") and i > 0:
            return lines[i - 1]
    return ""

def extract_highlight_mileage(html):
    """Reads the highlight tile where the value comes BEFORE the label, e.g. '25.35 Kmpl' ... 'Mileage'."""
    if not html:
        return None
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    lines = [clean(x) for x in soup.get_text("\n").split("\n")]
    lines = [x for x in lines if x]
    for i, line in enumerate(lines):
        m = re.fullmatch(r"(?:up to\s*)?(\d+(?:\.\d+)?)\s*kmpl", line, re.I)
        if not m:
            continue
        if any(norm(x) in ("mileage", "araimileage") for x in lines[i + 1:i + 4]):
            return m.group(1) + " kmpl"
    return None

def get_variant_links(html, model_url):
    """Return (ordered [(variant_url, name)], {variant_url: ABS type from the variants table})."""
    parts = urlparse(model_url).path.strip("/").split("/")
    prefix = f"/{parts[0]}/{parts[1]}/{parts[2]}/"
    soup = BeautifulSoup(html, "lxml")

    found = {}
    for a in soup.find_all("a", href=True):
        full = norm_url(a["href"])
        path = urlparse(full).path
        if not path.startswith(prefix):
            continue
        rest = path[len(prefix):].strip("/")
        if not rest or "/" in rest:
            continue
        if rest in RESERVED_SLUGS or rest.startswith(("price-in", "on-road", "emi")):
            continue
        name = clean(a.get("title") or a.get_text(" "))
        name = re.sub(r"\s*Compare$", "", name).strip()
        found.setdefault(full, name)

    urls = set(found)
    abs_map = {}
    for a in soup.find_all("a", href=True):
        full = norm_url(a["href"])
        if full not in urls or full in abs_map:
            continue
        node = a.parent
        while node is not None and node.name not in ("html", "body"):
            inner = {norm_url(x["href"]) for x in node.find_all("a", href=True)} & urls
            if len(inner) > 1:
                break
            m = ABS_ROW_RE.search(node.get_text(" "))
            if m:
                abs_map[full] = m.group(1).capitalize() + " Channel"
                break
            node = node.parent
    return list(found.items()), abs_map


def parse_json_safe(text):
    if not text:
        return {}
    try:
        data = json.loads(text)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def flatten_dict(obj, out, depth=0):
    if not isinstance(obj, dict) or depth > 3:
        return
    for k, v in obj.items():
        if "color" in str(k).lower():
            continue
        if isinstance(v, dict):
            flatten_dict(v, out, depth + 1)
        elif isinstance(v, (str, int, float)) and str(v).strip() != "":
            out.append((norm(k), str(v)))


def build_sources(html):
    """Return (sources, variant_json, gtm_json, h1_text)."""
    soup = BeautifulSoup(html, "lxml")

    metas = {}
    for m in soup.find_all("meta"):
        key = m.get("name") or m.get("property")
        if key and key not in metas:
            metas[key] = m.get("content", "")
    variant_json = parse_json_safe(metas.get("variant-data"))
    gtm_json = parse_json_safe(metas.get("gtm-tag"))

    json_pairs = []
    flatten_dict(variant_json, json_pairs)

    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    table_pairs = []
    for tr in soup.find_all("tr"):
        cells = [clean(c.get_text(" ")) for c in tr.find_all(["th", "td"])]
        if len(cells) >= 2 and cells[0] and cells[1]:
            table_pairs.append((norm(cells[0]), cells[1]))

    page_text = soup.get_text("\n")
    lines = [clean(x) for x in page_text.split("\n")]
    lines = [x for x in lines if x]
    text_pairs = []
    for i in range(len(lines) - 1):
        k = norm(lines[i])
        if k in ALL_KEYS and len(lines[i]) < 40:
            if k in FIELD_KEYS["mileage"] and "kmpl" not in lines[i + 1].lower():
                continue
            text_pairs.append((k, lines[i + 1]))

    extra = []
    m = ARAI_RE.search(clean(metas.get("atf-content", "")) + " " + clean(page_text))
    if m:
        extra.append(("mileage", m.group(1) + " kmpl"))

    h1 = soup.find("h1")
    h1_text = clean(h1.get_text(" ")) if h1 else ""
    return json_pairs + table_pairs + text_pairs + extra, variant_json, gtm_json, h1_text


def find_value(sources, field, parser):
    for key in FIELD_KEYS[field]:
        for k, v in sources:
            if k == key:
                result = parser(v)
                if result not in (None, ""):
                    return result
    return None


def extract_row(html, model_name, url, link_name, scraping_date, base_sources, abs_hint,
                mileage_hint=None):
    """base_sources = tables from the /specifications page (base variant) used as fallback."""
    sources, variant_json, gtm_json, h1_text = build_sources(html)

    def fv(field, parser):
        value = find_value(sources, field, parser)
        if value in (None, ""):
            value = find_value(base_sources, field, parser)
        return value

    variant = (clean(variant_json.get("v_variant_name"))
               or clean(link_name)
               or clean(gtm_json.get("Variant")))
    if not variant and h1_text:
        variant = re.sub(r"^Royal Enfield\s+" + re.escape(model_name) + r"\s*", "", h1_text,
                         flags=re.I).strip()

    engine_type, cooling = split_engine(fv("engine_type", parse_engine_raw))
    explicit_cooling = fv("cooling", parse_cooling_raw)
    if explicit_cooling:
        cooling = explicit_cooling

    return {
        "Model Name": model_name,
        "Variant": variant or None,
        "Engine Displacement (cc)": fv("displacement", make_num_parser(50, 1500)),
        "Engine Type": engine_type,
        "Cooling System": cooling,
        "Maximum Power (PS)": fv("power", parse_power),
        "Maximum Torque (Nm)": fv("torque", parse_torque),
        "Transmission (Gears)": fv("gears", make_num_parser(1, 8, integer=True)),
        "Mileage (kmpl)": (fv("mileage", make_num_parser(5, 120))
                           or make_num_parser(5, 120)(mileage_hint)),
        "Top Speed (kmph)": fv("top_speed", make_num_parser(60, 250)),
        "Kerb Weight (kg)": fv("kerb", make_num_parser(80, 400)),
        "Fuel Tank Capacity (Litres)": fv("tank", make_num_parser(3, 40)),
        "Front Suspension": fv("front_susp", parse_susp),
        "Rear Suspension": fv("rear_susp", parse_susp),
        "Front Brake": fv("front_brake", parse_brake),
        "Rear Brake": fv("rear_brake", parse_brake),
        "ABS (Single/Dual Channel)": fv("abs", parse_abs) or abs_hint,
        "Source_URL": url,
        "Scraping_Date": scraping_date,
    }


def row_has_data(row):
    skip = {"Model Name", "Variant", "Source_URL", "Scraping_Date"}
    return any(row[c] not in (None, "") for c in COLUMNS if c not in skip)


# --------------------------------------------------------------------------- #
# Export
# --------------------------------------------------------------------------- #
def export_files(rows, source_rows, scraping_date):
    df = pd.DataFrame(rows, columns=COLUMNS, dtype=object)
    df = df.drop_duplicates(subset=["Model Name", "Variant", "Source_URL"]).reset_index(drop=True)
    df = df.astype(object).where(pd.notnull(df), None)

    df.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")

    src_df = pd.DataFrame(source_rows, columns=["Model Name", "Source_URL", "Base Variant",
                                                "Variant Page URL", "Scraping_Date"])
    with pd.ExcelWriter(XLSX_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Bike_Specifications", index=False)
        src_df.to_excel(writer, sheet_name="Data_Source", index=False)

        header_fill = PatternFill("solid", fgColor="1F3864")
        header_font = Font(bold=True, color="FFFFFF")

        ws = writer.sheets["Bike_Specifications"]
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.freeze_panes = "C2"
        ws.auto_filter.ref = ws.dimensions
        for idx, col in enumerate(COLUMNS, start=1):
            longest = max([len(str(col))] + [len(str(v)) for v in df[col] if v is not None])
            ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 2, 12), 55)

        ws2 = writer.sheets["Data_Source"]
        for cell in ws2[1]:
            cell.fill = header_fill
            cell.font = header_font
        notes = [
            [],
            ["Notes"],
            ["Data source: 91Wheels (https://www.91wheels.com) only."],
            [f"Scraping date: {scraping_date}"],
            ["Scope: one row per model - the base variant named on the /specifications page."],
            ["Method: Playwright (Chromium) rendering + BeautifulSoup parsing."],
            ["Unavailable specifications are left blank; no values were estimated or fabricated."],
            ["Units: power in PS, torque in Nm, speed in kmph, mileage in kmpl, weight in kg, "
             "tank in litres, displacement in cc."],
            ["Ranges (e.g. '35-41.55') are not stored as single numbers; such cells are blank."],
        ]
        for n in notes:
            ws2.append(n)
        ws2.cell(row=len(source_rows) + 3, column=1).font = Font(bold=True)
        for col_letter, width in (("A", 28), ("B", 80), ("C", 32), ("D", 90), ("E", 16)):
            ws2.column_dimensions[col_letter].width = width

    return df


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    scraping_date = date.today().isoformat()
    rows = []
    source_rows = []

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

        for model_name, specs_url in MODELS:
            log.info("=== %s ===", model_name)
            model_url = specs_url.rsplit("/specifications", 1)[0] + "/"

            # 1) Specifications page -> base variant name + base-variant tables
            specs_html = fetch_html(page, specs_url)
            base_sources, base_name = [], ""
            if specs_html:
                base_sources = build_sources(specs_html)[0]
                base_name = find_base_variant_name(specs_html)
            else:
                log.warning("Could not load specifications page: %s", specs_url)
            log.info("%s: base variant on specs page = %r", model_name, base_name)

            # 2) Model page -> variant links + ABS from variants table
            time.sleep(random.uniform(0.5, 1.0))
            model_html = fetch_html(page, model_url)
            links, abs_map = ([], {})
            if model_html:
                links, abs_map = get_variant_links(model_html, model_url)
            mileage_hint = (extract_highlight_mileage(model_html)
                or extract_highlight_mileage(specs_html))
            log.info("%s: highlight mileage = %s", model_name, mileage_hint)
            chosen_url, chosen_name = None, base_name
            for v_url, v_name in links:
                if base_name and norm(v_name) == norm(base_name):
                    chosen_url, chosen_name = v_url, v_name
                    break
            if not chosen_url and links:
                chosen_url, chosen_name = links[0]
                log.warning("%s: base variant name not matched; using first listed variant %r",
                            model_name, chosen_name)

            # 3) Base variant page (or fall back to specs page tables)
            row = None
            if chosen_url:
                time.sleep(random.uniform(0.5, 1.0))
                v_html = fetch_html(page, chosen_url)
                if v_html:
                    row = extract_row(v_html, model_name, chosen_url, chosen_name, scraping_date,
                                      base_sources, abs_map.get(chosen_url), mileage_hint)
            if (row is None or not row_has_data(row)) and specs_html:
                log.warning("%s: using specifications page data only", model_name)
                row = extract_row(specs_html, model_name, specs_url, base_name, scraping_date,
                                  base_sources, abs_map.get(chosen_url) if chosen_url else None,
                                  mileage_hint)

            if row is not None and row_has_data(row):
                rows.append(row)
                log.info("  + %s | %s", model_name, row["Variant"])
                source_rows.append([model_name, specs_url, row["Variant"], row["Source_URL"],
                                    scraping_date])
            else:
                log.error("%s: no data extracted", model_name)
                source_rows.append([model_name, specs_url, None, None, scraping_date])

            time.sleep(random.uniform(0.5, 1.0))

        browser.close()

    if not rows:
        log.error("No data scraped. Nothing to export.")
        sys.exit(1)

    df = export_files(rows, source_rows, scraping_date)

    log.info("-" * 60)
    log.info("Total rows (one base variant per model): %d", len(df))
    blanks = {c: int(df[c].isna().sum()) for c in COLUMNS if df[c].isna().any()}
    if blanks:
        log.info("Blank (unavailable) cells per column: %s", blanks)
    log.info("Saved: %s", XLSX_FILE)
    log.info("Saved: %s", CSV_FILE)


if __name__ == "__main__":
    main()