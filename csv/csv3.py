# coding: utf-8
import re
import csv
import os
import sys
import io
import json
import time
import argparse
from urllib.parse import urlparse, parse_qs
from bs4 import BeautifulSoup, NavigableString, Tag
import requests

try:
    import pasteboard
except Exception:
    pasteboard = None

# =====================
# Settings
# =====================
ENABLE_CSV_EXPORT = True
COPY_SUBJECT_PK_TO_CLIPBOARD = True
CSV_FILE_ENCODING = "utf-8-sig"

BLOOD_CSV_NAME = "blood.csv"
STAKES_CSV_NAME = "stakes_horses.csv"
OUTPUT_DIR_NAME = "output"

PEDIGREEQUERY_BASE = "https://www.pedigreequery.com"
DEFAULT_TIMEOUT = 30
DEFAULT_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"

CSV_HEADER = [
    "PrimaryKey",
    "Sire",
    "Dam",
    "Sex",
    "Color",
    "Year",
    "Country",
    "Family",
    "DP",
    "DI",
    "CD",
    "Starts",
    "Wins",
    "Places",
    "Shows",
    "CareerEarnings",
    "Owner",
    "Breeder",
    "StateBred",
    "WinningsText",
    "Details",
    "URL",
    "Horse Name",
    "SubjectInfoText",
]

STAKES_HEADER = ["PrimaryKey", "URL", "RaceDataJSON"]

YEAR_4DIGIT_RE = re.compile(r"(\d{4})")
COUNTRY_IN_TEXT_RE = re.compile(r"\(\s*([A-Za-z]{2,4})\s*\)")
PLACEHOLDER_YEAR_RE = re.compile(r"[~〜]\s*\d{4}")
YEAR_RE = re.compile(r"(?<![~〜])\b(\d{4}c?)\b")

PQ_COLOR_RE = re.compile(
    r"\b("
    r"blk\.?|black|"
    r"b\.?|bay|"
    r"br\.?|brown|"
    r"ch\.?|chestnut|"
    r"gr\.?|grey|gray|"
    r"dkb/br\.?|dkb/br|"
    r"roan|rn\.?"
    r")\b",
    re.I
)

PQ_HEADER_INFO_RE = re.compile(
    r"\)\s*"
    r"(?P<color>blk\.?|black|b\.?|bay|br\.?|brown|ch\.?|chestnut|gr\.?|grey|gray|dkb/br\.?|dkb/br|roan|rn\.?)"
    r"\s*,?\s*(?P<sex>[A-Z])\s*,?\s*(?P<year>(18|19|20)\d{2})\b",
    re.I
)

# =====================
# Common utils
# =====================
def norm_space(s: str) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def clean_horse_name(name: str) -> str:
    if not name:
        return ""
    return (
        str(name)
        .replace("　", " ")
        .replace("\u00A0", " ")
        .replace("\n", " ")
        .replace("\r", " ")
        .replace("\t", " ")
        .strip()
    )

def normalize_horse_key_name(name: str) -> str:
    return clean_horse_name(name).upper() if name else ""

def _year_digits(y: str) -> str:
    if not y:
        return ""
    m = YEAR_4DIGIT_RE.search(str(y))
    return m.group(1) if m else ""

def _abs_pedigreequery_url(href: str) -> str:
    if not href:
        return ""
    href = href.strip()
    if not href:
        return ""
    low = href.lower()
    if low.startswith("javascript:") or href == "/":
        return ""
    if low.startswith("http://") or low.startswith("https://"):
        return href
    if href.startswith("/"):
        return PEDIGREEQUERY_BASE + href
    return PEDIGREEQUERY_BASE + "/" + href.lstrip("/")

# =====================
# PK utils
# =====================
def _normalize_href_for_pk(href: str) -> str:
    if not href:
        return ""
    s = href.strip()
    if s.upper().startswith("HREF:"):
        s = s.split(":", 1)[1].strip()
    s = re.sub(r"^https?://[^/]+", "", s, flags=re.I).strip().lower()
    if s.startswith("javascript:") or s == "/":
        return ""
    return s.lstrip("/")

def _pk_from_key_tuple(key_tuple):
    if not key_tuple:
        return ""
    t = key_tuple[0]
    if t == "HREF":
        href = (key_tuple[1] or "")
        return _normalize_href_for_pk(href)
    if t == "NAMEYEAR":
        nm, yy = key_tuple[1]
        nm = normalize_horse_key_name(nm)
        yy = _year_digits((yy or "").strip())
        return f"{nm}:{yy}" if nm and yy else (f"N:{nm}" if nm else "")
    if t == "NAME":
        nm = normalize_horse_key_name(key_tuple[1])
        return f"{nm}" if nm else ""
    return str(key_tuple)

def _pk_from_horse_href(href: str) -> str:
    if not href:
        return ""
    s = href.strip()
    if not s or s == "/":
        return ""
    m = re.search(r"[?&]h=([^&#]+)", s, flags=re.I)
    if m:
        return m.group(1).strip().lower()

    normalized = _normalize_url(s)
    try:
        u = urlparse(normalized)
    except Exception:
        return ""

    qs = parse_qs(u.query or "")
    if "h" in qs and qs["h"]:
        return (qs["h"][0] or "").strip().lower()

    path = (u.path or "").strip("/")
    if not path:
        return ""
    if "/" in path:
        return ""
    low = path.lower()
    if low in ("index.php", "query.php"):
        return ""
    return low

def _is_valid_horse_url(url: str) -> bool:
    if not url:
        return False
    normalized = _normalize_url(url)
    if not normalized:
        return False
    return bool(_pk_from_horse_href(normalized))

def _horse_url_from_href(href: str) -> str:
    pk = _pk_from_horse_href(href)
    if pk:
        return f"{PEDIGREEQUERY_BASE}/{pk}"
    return ""

def _normalize_url(url: str) -> str:
    if not url:
        return ""
    url = url.strip()
    if not url:
        return ""
    if url.startswith("//"):
        return "https:" + url
    if url.startswith("/"):
        return PEDIGREEQUERY_BASE + url
    if not re.match(r"^https?://", url, flags=re.I):
        return PEDIGREEQUERY_BASE + "/" + url.lstrip("/")
    return url

def _normalize_race_record_for_key(obj: dict):
    return (
        str(obj.get("race_page_id", "") or ""),
        str(obj.get("year", "") or ""),
        str(obj.get("placing", "") or ""),
    )

def _ensure_parent_dir(filepath: str):
    parent = os.path.dirname(os.path.abspath(filepath))
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)

def _blood_row_value(row, field_name: str):
    if isinstance(row, dict):
        return row.get(field_name, "")
    if isinstance(row, (list, tuple)):
        try:
            idx = CSV_HEADER.index(field_name)
        except ValueError:
            return ""
        return row[idx] if idx < len(row) else ""
    return ""

def _blood_row_sort_key(row):
    pk = str(_blood_row_value(row, "PrimaryKey") or "")
    year_text = _year_digits(_blood_row_value(row, "Year") or "")
    year_num = int(year_text) if year_text.isdigit() else 999999
    horse_name = norm_space(_blood_row_value(row, "Horse Name") or "").upper()
    is_founder = 0 if pk == "" else 1
    return (is_founder, year_num, horse_name, pk.upper())

# =====================
# CSV helpers
# =====================
def upsert_row(rows_by_pk, row):
    pk = row.get("PrimaryKey", "")
    if pk is None or pk == "":
        return
    if pk not in rows_by_pk:
        rows_by_pk[pk] = row
        return
    cur = rows_by_pk[pk]
    for k, v in row.items():
        if (cur.get(k) in (None, "", "不明")) and v not in (None, "", "不明"):
            cur[k] = v

def _ensure_founder(rows_by_pk):
    if "" not in rows_by_pk:
        rows_by_pk[""] = {h: "" for h in CSV_HEADER}

def dump_rows_as_csv(rows_by_pk):
    out = io.StringIO()
    w = csv.writer(out, lineterminator="\n")
    w.writerow(CSV_HEADER)
    for r in sorted(rows_by_pk.values(), key=_blood_row_sort_key):
        w.writerow([_blood_row_value(r, h) for h in CSV_HEADER])
    return out.getvalue()

def dump_stakes_rows_csv(stakes_rows: list):
    out = io.StringIO()
    w = csv.writer(out, lineterminator="\n")
    w.writerow(STAKES_HEADER)
    for row in stakes_rows:
        race_data = row.get("RaceDataJSON", [])
        if isinstance(race_data, dict):
            race_data = [race_data]
        elif not isinstance(race_data, list):
            race_data = []
        w.writerow([
            row.get("PrimaryKey", ""),
            row.get("URL", ""),
            json.dumps(race_data, ensure_ascii=False, separators=(",", ":"), sort_keys=True),
        ])
    return out.getvalue()

def sort_blood_csv_file(filepath: str, encoding: str = CSV_FILE_ENCODING):
    if not os.path.exists(filepath):
        return 0

    with open(filepath, "r", encoding=encoding, errors="replace", newline="") as f:
        rr = csv.reader(f)
        rows = list(rr)

    if not rows:
        return 0

    header = rows[0]
    data_rows = [row for row in rows[1:] if row]

    with open(filepath, "w", encoding=encoding, errors="replace", newline="") as f:
        ww = csv.writer(f, lineterminator="\n")
        ww.writerow(header)
        for row in sorted(data_rows, key=_blood_row_sort_key):
            ww.writerow(row)

    return len(data_rows)

def append_unique_csv(csv_text: str, filepath: str, subject_pk: str):
    if not csv_text.strip():
        return
    _ensure_parent_dir(filepath)

    src = io.StringIO(csv_text)
    r = csv.reader(src)
    rows = list(r)
    if not rows:
        return

    header = rows[0]
    new_rows = rows[1:]

    rows_by_pk = {}
    if os.path.exists(filepath):
        with open(filepath, "r", encoding=CSV_FILE_ENCODING, errors="replace", newline="") as f:
            rr = csv.reader(f)
            for i, row in enumerate(rr):
                if i == 0:
                    continue
                if not row:
                    continue
                pk = (row[0] or "").strip()
                if pk:
                    rows_by_pk[pk] = row

    for row in new_rows:
        if not row:
            continue
        pk = (row[0] or "").strip()
        if not pk:
            continue

        if subject_pk and pk == subject_pk:
            rows_by_pk[pk] = row
        elif pk not in rows_by_pk:
            rows_by_pk[pk] = row

    with open(filepath, "w", encoding=CSV_FILE_ENCODING, errors="replace", newline="") as f:
        ww = csv.writer(f, lineterminator="\n")
        ww.writerow(header)
        for row in rows_by_pk.values():
            ww.writerow(row)

def append_unique_stakes_rows_csv(csv_text: str, filepath: str, encoding: str = CSV_FILE_ENCODING):
    if not csv_text.strip():
        return
    _ensure_parent_dir(filepath)

    src = io.StringIO(csv_text)
    r = csv.reader(src)
    rows = list(r)
    if not rows:
        return

    header = rows[0]
    data_rows = rows[1:]
    existing = {}
    file_exists = os.path.exists(filepath)

    if file_exists:
        with open(filepath, "r", encoding=encoding, errors="replace", newline="") as f:
            rr = csv.reader(f)
            for i, row in enumerate(rr):
                if i == 0 or len(row) < 3:
                    continue
                pk = (row[0] or "").strip()
                url = (row[1] or "").strip()
                js = (row[2] or "").strip()
                if not pk:
                    continue
                try:
                    arr = json.loads(js) if js else []
                    if isinstance(arr, dict):
                        arr = [arr]
                    elif not isinstance(arr, list):
                        arr = []
                except Exception:
                    arr = []
                existing[pk] = {"URL": url, "RaceDataJSON": arr}

    for row in data_rows:
        if len(row) < 3:
            continue

        pk = (row[0] or "").strip()
        url = (row[1] or "").strip()
        js = (row[2] or "").strip()
        if not pk:
            continue

        try:
            incoming = json.loads(js) if js else []
            if isinstance(incoming, dict):
                incoming = [incoming]
            elif not isinstance(incoming, list):
                incoming = []
        except Exception:
            incoming = []

        if pk not in existing:
            existing[pk] = {"URL": url, "RaceDataJSON": []}

        if url and not existing[pk]["URL"]:
            existing[pk]["URL"] = url

        seen = set(_normalize_race_record_for_key(x) for x in existing[pk]["RaceDataJSON"])
        for rec in incoming:
            k = _normalize_race_record_for_key(rec)
            if k in seen:
                continue
            seen.add(k)
            existing[pk]["RaceDataJSON"].append(rec)

        existing[pk]["RaceDataJSON"].sort(
            key=lambda x: (
                str(x.get("year", "") or ""),
                int(x.get("placing", 999) or 999),
                str(x.get("race_page_id", "") or ""),
            )
        )

    with open(filepath, "w", encoding=encoding, errors="replace", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(header)
        for pk in sorted(existing.keys()):
            w.writerow([
                pk,
                existing[pk]["URL"],
                json.dumps(existing[pk]["RaceDataJSON"], ensure_ascii=False, separators=(",", ":"), sort_keys=True),
            ])

def load_existing_blood_pks(filepath: str, encoding: str = CSV_FILE_ENCODING):
    out = set()
    if not os.path.exists(filepath):
        return out
    with open(filepath, "r", encoding=encoding, errors="replace", newline="") as f:
        rr = csv.reader(f)
        for i, row in enumerate(rr):
            if i == 0 or not row:
                continue
            pk = (row[0] or "").strip()
            if pk:
                out.add(pk)
    return out

def load_horse_urls_from_stakes_csv(filepath: str, encoding: str = CSV_FILE_ENCODING):
    urls = []
    seen = set()
    if not os.path.exists(filepath):
        raise FileNotFoundError(filepath)

    with open(filepath, "r", encoding=encoding, errors="replace", newline="") as f:
        rr = csv.reader(f)
        for i, row in enumerate(rr):
            if i == 0 or len(row) < 2:
                continue
            url = _normalize_url(row[1] or "")
            if not url or url in seen:
                continue
            seen.add(url)
            urls.append(url)
    return urls

def make_session():
    s = requests.Session()
    s.headers.update(
        {
            "User-Agent": DEFAULT_USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        }
    )
    return s

def fetch_html(url: str, session=None, timeout: int = DEFAULT_TIMEOUT):
    session = session or make_session()
    resp = session.get(_normalize_url(url), timeout=timeout)
    resp.raise_for_status()
    resp.encoding = resp.encoding or resp.apparent_encoding or "utf-8"
    return resp.text

def fetch_soup(url: str, session=None, timeout: int = DEFAULT_TIMEOUT):
    html = fetch_html(url, session=session, timeout=timeout)
    return BeautifulSoup(html, "html.parser")

def collect_race_urls_from_list_page(soup_obj):
    race_dict = {}
    for td in soup_obj.find_all("td", class_="w2"):
        a_tag = td.find("a", href=True)
        if not a_tag:
            continue
        href = _normalize_url(a_tag.get("href", "")).split("#")[0]
        title = clean_horse_name(a_tag.get_text(" ", strip=True)) or "(no title)"
        if href and href not in race_dict:
            race_dict[href] = title
    return race_dict

def extract_frontier_horse_urls(soup_obj, parse_depth=5):
    pedigree_table = soup_obj.find("table", class_="pedigreetable")
    if not pedigree_table:
        return []

    urls = []
    seen = set()
    selector = f'td[data-g="{int(parse_depth)}"] a.horseName[href]'
    for a in pedigree_table.select(selector):
        url = _normalize_url(a.get("href", ""))
        if not _is_valid_horse_url(url) or url in seen:
            continue
        seen.add(url)
        urls.append(url)
    return urls

def process_race_page_url(url: str, session, stakes_out_path: str):
    soup = fetch_soup(url, session=session)
    csv_text, _ = build_csv_from_pedigreequery_stakes_per_horse(soup)
    if not csv_text.strip():
        return 0
    append_unique_stakes_rows_csv(csv_text, stakes_out_path, encoding=CSV_FILE_ENCODING)
    rows = list(csv.reader(io.StringIO(csv_text)))
    return max(len(rows) - 1, 0)

def process_race_targets(urls, session, stakes_out_path: str, sleep_sec: float = 0.0):
    total = 0
    for i, url in enumerate(urls, 1):
        print(f"[RACE] {i}/{len(urls)} {url}")
        try:
            count = process_race_page_url(url, session, stakes_out_path)
            total += count
            print(f"[OK] horses added/merged: {count}")
        except Exception as e:
            print(f"[WARN] race scrape failed: {url} ({e})")
        if sleep_sec > 0 and i < len(urls):
            time.sleep(sleep_sec)
    return total

def process_horse_targets(urls, session, blood_out_path: str, register_depth=4, parse_depth=5, sleep_sec: float = 0.0):
    queue = []
    for x in urls:
        nx = _normalize_url(x)
        if _is_valid_horse_url(nx) and nx not in queue:
            queue.append(nx)
    queued = set(queue)
    visited_urls = set()
    existing_pks = load_existing_blood_pks(blood_out_path, encoding=CSV_FILE_ENCODING)
    processed_count = 0
    failed_count = 0
    skipped_count = 0

    while queue:
        url = queue.pop()
        queued.discard(url)
        if not url or url in visited_urls:
            skipped_count += 1
            continue
        visited_urls.add(url)
        print(f"[HORSE] {processed_count + 1} url={url} pending={len(queue)}")

        try:
            soup = fetch_soup(url, session=session)
            family_map = extract_family_data_map(soup)
            csv_text, subject_pk = build_csv_from_pedigreequery(
                soup,
                family_map,
                register_depth=register_depth,
                parse_depth=parse_depth,
            )
            written_rows = max(len(list(csv.reader(io.StringIO(csv_text)))) - 1, 0) if csv_text.strip() else 0
            if csv_text.strip():
                append_unique_csv(csv_text, blood_out_path, subject_pk)
                if subject_pk:
                    existing_pks.add(subject_pk)
            else:
                print(f"[WARN] pedigree parse returned empty csv: {url}")

            added_frontier = 0
            for next_url in extract_frontier_horse_urls(soup, parse_depth=parse_depth):
                pk = _pk_from_horse_href(next_url)
                if not pk:
                    continue
                if next_url in visited_urls or next_url in queued:
                    continue
                if pk and pk in existing_pks:
                    continue
                queued.add(next_url)
                queue.append(next_url)
                added_frontier += 1

            processed_count += 1
            print(
                f"[OK] horse={subject_pk or '(unknown)'} rows={written_rows} "
                f"next={added_frontier} pending={len(queue)} out={blood_out_path}"
            )
        except Exception as e:
            failed_count += 1
            print(f"[WARN] horse scrape failed: {url} ({e})")

        if sleep_sec > 0 and queue:
            time.sleep(sleep_sec)

    print(
        f"[DONE] horses_processed={processed_count} "
        f"failed={failed_count} skipped={skipped_count} out={blood_out_path}"
    )
    sorted_rows = sort_blood_csv_file(blood_out_path, encoding=CSV_FILE_ENCODING)
    print(f"[DONE] blood.csv sorted rows={sorted_rows} out={blood_out_path}")
    return processed_count

# =====================
# pedigreequery pedigree helpers
# =====================
def _extract_pq_color_token(text: str) -> str:
    if not text:
        return ""
    m = PQ_COLOR_RE.search(text)
    if not m:
        return ""
    tok = m.group(1).lower().strip().replace(".", "")
    if tok in ("black", "blk"):
        return "blk"
    if tok in ("bay", "b"):
        return "b"
    if tok in ("brown", "br"):
        return "br"
    if tok in ("chestnut", "ch"):
        return "ch"
    if tok in ("grey", "gray", "gr"):
        return "gr"
    if tok in ("dkb/br", "dkbbr"):
        return "dkb/br"
    if tok in ("roan", "rn"):
        return "rn"
    return tok

def _country_token_after(tag):
    if not tag:
        return ""
    start = tag
    if isinstance(tag.parent, Tag) and tag.parent.name == "b":
        start = tag.parent

    for el in start.next_elements:
        if el is start:
            continue
        if isinstance(el, Tag):
            continue
        if isinstance(el, NavigableString):
            s = str(el).strip()
            if not s:
                continue
            m = COUNTRY_IN_TEXT_RE.search(s)
            if m:
                return m.group(1).strip().upper()
    return ""

def _display_name_with_country(a_tag):
    base = clean_horse_name(a_tag.get_text(strip=True)) if a_tag else ""
    c = _country_token_after(a_tag) if a_tag else ""
    return f"{base} ({c})" if (base and c) else base

def _sex_from_pq_td(td: Tag) -> str:
    if not td:
        return ""
    classes = td.get("class") or []
    classes = [c.lower() for c in classes]
    if "m" in classes:
        return "H"
    if "f" in classes:
        return "M"
    return ""

def _cell_text_with_neighbors(td):
    def ok(n):
        if not n or n.has_attr("data-g"):
            return False
        classes = n.get("class") or []
        return not any(cls in ("w", "w2") for cls in classes)

    parts = [td.get_text(" ", strip=True)]
    nxt = td.find_next_sibling("td")
    if ok(nxt):
        parts.append(nxt.get_text(" ", strip=True))
        nxt2 = nxt.find_next_sibling("td")
        if ok(nxt2):
            parts.append(nxt2.get_text(" ", strip=True))

    text = " ".join(parts)
    text = PLACEHOLDER_YEAR_RE.sub("", text)
    return text

def _family_alias_key_variants(name: str):
    base = normalize_horse_key_name(name)
    if not base:
        return []

    variants = [base]
    m = re.match(r"^(.*?)(\d+)$", base)
    if m and m.group(1):
        variants.append(m.group(1))

    out = []
    seen = set()
    for v in variants:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out

def extract_family_data_map(soup_obj):
    horse_to_family_map = {}
    summary = next((t for t in soup_obj.find_all("font") if "Family Summary:" in t.get_text()), None)
    if not summary:
        return horse_to_family_map

    for link in summary.find_all("a"):
        raw_label = link.get_text(strip=True)
        fam_label = re.sub(r"\s*\(\d+\)$", "", raw_label).strip() or ""

        onmouseover = link.get("onmouseover")
        if not onmouseover:
            continue

        m = re.search(r"changeFloat\('(.*?)',\s*event\)", onmouseover)
        if not m:
            continue

        horse_list_raw = m.group(1)
        for horse_chunk in [x.strip() for x in horse_list_raw.split(",")]:
            m2 = re.match(r"(.+?)\s*\(([^)]*)\)\s*$", horse_chunk)
            if m2:
                base_name = m2.group(1).strip()
                yearpart = m2.group(2).strip()
            else:
                base_name = horse_chunk.strip()
                yearpart = ""

            for cname_key in _family_alias_key_variants(base_name):
                if cname_key and cname_key != "#":
                    horse_to_family_map[(cname_key, yearpart)] = fam_label
                    horse_to_family_map[(cname_key, "")] = fam_label

    return horse_to_family_map

def _canonical_ancestor_id(cell):
    tag = cell.find("a", class_="horseName")

    country = _country_token_after(tag) if tag else ""
    name_disp = _display_name_with_country(tag) if tag else ""
    name_key = normalize_horse_key_name(re.sub(r"\s*\([A-Z]{2,4}\)\s*$", "", name_disp or ""))

    combo = _cell_text_with_neighbors(cell)

    m_year = YEAR_RE.search(combo)
    raw_year_token = m_year.group(1) if m_year else ""
    raw_year = _year_digits(raw_year_token)

    color = _extract_pq_color_token(combo)
    sex = _sex_from_pq_td(cell)

    href = ""
    if tag and tag.has_attr("href"):
        href = (tag["href"] or "").strip()
        href_low = href.lower()
        if href_low.startswith("javascript:") or href == "/":
            href = ""

    if href:
        key = ("HREF", href.lower())
    elif name_key:
        key = ("NAMEYEAR", (name_key, raw_year)) if raw_year else ("NAME", name_key)
    else:
        key = None

    g = int(cell.get("data-g", "0"))
    return {
        "name_disp": name_disp,
        "name_key": name_key,
        "year_token": raw_year_token,
        "year": raw_year,
        "color": color,
        "sex": sex,
        "country": country,
        "href": href,
        "key": key,
        "gen": g,
    }

def _clean_subject_info_text(text: str) -> str:
    if not text:
        return ""

    s = str(text).replace("\r", "\n")
    s = re.sub(r'(?mi)^\s*(Owner|Breeder|State Bred|Winnings)\s*\n\s*:\s*', r'\1: ', s)
    s = re.sub(r'\(\s*CLOSE\s*\)\s*$', '', s, flags=re.I | re.S)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()

def _extract_pq_subject_extra_info(soup_obj):
    result = {
        "Country": "",
        "Family": "",
        "DP": "",
        "DI": "",
        "CD": "",
        "Starts": "",
        "Wins": "",
        "Places": "",
        "Shows": "",
        "CareerEarnings": "",
        "Owner": "",
        "Breeder": "",
        "StateBred": "",
        "WinningsText": "",
        "SubjectInfoText": "",
    }

    header_font = soup_obj.select_one("center > font.normal")
    header_text = header_font.get_text(" ", strip=True) if header_font else ""

    m = re.search(r"\(([A-Z]{2,4})\)", header_text)
    if m:
        result["Country"] = m.group(1)

    m = re.search(r"\{([^}]+)\}", header_text)
    if m:
        result["Family"] = m.group(1).strip()

    m = re.search(r"DP\s*=\s*([^D]+?)\s+DI\s*=", header_text)
    if m:
        result["DP"] = m.group(1).strip()

    m = re.search(r"DI\s*=\s*([0-9.]+)", header_text)
    if m:
        result["DI"] = m.group(1).strip()

    m = re.search(r"CD\s*=\s*([0-9.\-]+)", header_text)
    if m:
        result["CD"] = m.group(1).strip()

    m = re.search(r"-\s*(\d+)\s+Starts,\s*(\d+)\s+Wins,\s*(\d+)\s+Places,\s*(\d+)\s+Shows", header_text)
    if m:
        result["Starts"] = m.group(1)
        result["Wins"] = m.group(2)
        result["Places"] = m.group(3)
        result["Shows"] = m.group(4)

    m = re.search(r"Career Earnings:\s*([¥$€£0-9,.\sA-Za-z]+)$", header_text)
    if m:
        result["CareerEarnings"] = m.group(1).strip()

    info_div = soup_obj.find("div", id="subjectinfo")
    if info_div:
        info_text = info_div.get_text("\n", strip=True)
        info_text = _clean_subject_info_text(info_text)
        result["SubjectInfoText"] = info_text

        m = re.search(r"(?mi)^\s*Owner\s*:\s*(.+)$", info_text)
        if m:
            result["Owner"] = m.group(1).strip()

        m = re.search(r"(?mi)^\s*Breeder\s*:\s*(.+)$", info_text)
        if m:
            result["Breeder"] = m.group(1).strip()

        m = re.search(r"(?mi)^\s*State\s*Bred\s*:\s*(.+)$", info_text)
        if m:
            result["StateBred"] = m.group(1).strip()

        m = re.search(r"(?mi)^\s*Winnings\s*:\s*(.+)$", info_text)
        if m:
            result["WinningsText"] = m.group(1).strip()

    return result

def _extract_subject_from_header(soup_obj, by_bits=None):
    by_bits = by_bits or {}

    subj_href = ""
    for a in soup_obj.select("div#menu_queries a[href^='/']"):
        if a.get_text(strip=True).lower() == "pedigree":
            subj_href = a.get("href", "")
            break
    subject_pk = subj_href.lstrip("/").strip() if subj_href else ""

    subject_display = ""
    top_a = soup_obj.select_one("center b a.nounderline")
    if top_a:
        base = clean_horse_name(top_a.get_text(strip=True))
        c = _country_token_after(top_a)
        subject_display = f"{base} ({c})" if (base and c) else base

    if not subject_display:
        title = soup_obj.title.get_text(strip=True) if soup_obj.title else ""
        horse_name = title.replace(" Horse Pedigree", "").strip() if title else ""
        subject_display = clean_horse_name(horse_name)

    if not subject_pk:
        subject_pk = f"PQ:MAIN:{normalize_horse_key_name(subject_display)}" if subject_display else "PQ:MAIN:UNKNOWN"

    header_font = soup_obj.select_one("center > font.normal")
    header_text = header_font.get_text(" ", strip=True) if header_font else ""

    year = ""
    sex = ""
    color = ""

    m = PQ_HEADER_INFO_RE.search(header_text or "")
    if m:
        color = _extract_pq_color_token(m.group("color") or "")
        sex = (m.group("sex") or "").strip().upper()
        year = m.group("year") or ""
    else:
        mm = re.search(r"\b(18\d{2}|19\d{2}|20\d{2})\b", header_text)
        if mm:
            year = mm.group(1)
        color = _extract_pq_color_token(header_text)

    extra = _extract_pq_subject_extra_info(soup_obj)
    details = "pedigreequery"
    url = _abs_pedigreequery_url(subj_href) if subj_href else ""

    return {
        "PrimaryKey": subject_pk,
        "Sire": by_bits.get("0", "") or "",
        "Dam": by_bits.get("1", "") or "",
        "Horse Name": subject_display,
        "Year": _year_digits(year),
        "Sex": sex,
        "Color": color,
        "Country": extra.get("Country", ""),
        "Family": extra.get("Family", ""),
        "DP": extra.get("DP", ""),
        "DI": extra.get("DI", ""),
        "CD": extra.get("CD", ""),
        "Starts": extra.get("Starts", ""),
        "Wins": extra.get("Wins", ""),
        "Places": extra.get("Places", ""),
        "Shows": extra.get("Shows", ""),
        "CareerEarnings": extra.get("CareerEarnings", ""),
        "Owner": extra.get("Owner", ""),
        "Breeder": extra.get("Breeder", ""),
        "StateBred": extra.get("StateBred", ""),
        "WinningsText": extra.get("WinningsText", ""),
        "Details": details,
        "URL": url,
        "SubjectInfoText": extra.get("SubjectInfoText", ""),
    }

def build_csv_from_pedigreequery(soup_obj, horse_family_map=None, register_depth=4, parse_depth=5):
    pedigree_table = soup_obj.find("table", class_="pedigreetable")
    if not pedigree_table:
        return "", ""

    rows_by_pk = {}
    _ensure_founder(rows_by_pk)
    by_bits = {}

    for g in range(1, parse_depth + 1):
        cells = pedigree_table.find_all("td", attrs={"data-g": str(g)})
        for i, cell in enumerate(cells):
            ident = _canonical_ancestor_id(cell)

            pk = _pk_from_key_tuple(ident.get("key"))
            if not pk:
                nm = ident.get("name_key", "")
                pk = f"N:{nm}" if nm else ""
            if not pk:
                continue

            bits = bin(i)[2:].zfill(g)
            by_bits[bits] = pk

            if g <= register_depth:
                horse_name_disp = ident.get("name_disp", "")
                raw_year_token = ident.get("year_token", "")
                raw_year = ident.get("year", "")

                fam_value = ""
                if horse_family_map:
                    key_name = ident.get("name_key", "")
                    fam_value = (
                        horse_family_map.get((key_name, raw_year_token), "")
                        or horse_family_map.get((key_name, raw_year), "")
                        or horse_family_map.get((key_name, ""), "")
                    )

                details = "pedigreequery"
                abs_url = _abs_pedigreequery_url(ident.get("href", ""))

                upsert_row(
                    rows_by_pk,
                    {
                        "PrimaryKey": pk,
                        "Sire": "",
                        "Dam": "",
                        "Sex": ident.get("sex", "") or "",
                        "Color": ident.get("color", "") or "",
                        "Year": raw_year or "",
                        "Country": ident.get("country", "") or "",
                        "Family": fam_value,
                        "DP": "",
                        "DI": "",
                        "CD": "",
                        "Starts": "",
                        "Wins": "",
                        "Places": "",
                        "Shows": "",
                        "CareerEarnings": "",
                        "Owner": "",
                        "Breeder": "",
                        "StateBred": "",
                        "WinningsText": "",
                        "Details": details,
                        "URL": abs_url,
                        "Horse Name": horse_name_disp,
                        "SubjectInfoText": "",
                    },
                )

    def get_pk(bits):
        return by_bits.get(bits, "")

    for bits, pk in list(by_bits.items()):
        g = len(bits)
        if g > register_depth:
            continue

        sire_pk = get_pk(bits + "0") or ""
        dam_pk = get_pk(bits + "1") or ""

        if pk in rows_by_pk:
            rows_by_pk[pk]["Sire"] = sire_pk
            rows_by_pk[pk]["Dam"] = dam_pk

    subject = _extract_subject_from_header(soup_obj, by_bits=by_bits)
    subject_pk = subject.get("PrimaryKey", "") or ""
    if subject_pk:
        upsert_row(
            rows_by_pk,
            {
                "PrimaryKey": subject_pk,
                "Sire": subject.get("Sire", ""),
                "Dam": subject.get("Dam", ""),
                "Sex": subject.get("Sex", ""),
                "Color": subject.get("Color", ""),
                "Year": _year_digits(subject.get("Year", "")),
                "Country": subject.get("Country", ""),
                "Family": subject.get("Family", ""),
                "DP": subject.get("DP", ""),
                "DI": subject.get("DI", ""),
                "CD": subject.get("CD", ""),
                "Starts": subject.get("Starts", ""),
                "Wins": subject.get("Wins", ""),
                "Places": subject.get("Places", ""),
                "Shows": subject.get("Shows", ""),
                "CareerEarnings": subject.get("CareerEarnings", ""),
                "Owner": subject.get("Owner", ""),
                "Breeder": subject.get("Breeder", ""),
                "StateBred": subject.get("StateBred", ""),
                "WinningsText": subject.get("WinningsText", ""),
                "Details": subject.get("Details", "pedigreequery"),
                "URL": subject.get("URL", ""),
                "Horse Name": subject.get("Horse Name", ""),
                "SubjectInfoText": subject.get("SubjectInfoText", ""),
            },
        )

    return dump_rows_as_csv(rows_by_pk), subject_pk

# =====================
# pedigreequery stakes -> per horse CSV
# =====================
def _find_pq_stakes_table(soup_obj):
    for table in soup_obj.find_all("table", attrs={"border": "1"}):
        header_text = norm_space(table.get_text(" ", strip=True)).lower()
        if all(x in header_text for x in ["year", "winner", "2nd", "3rd", "time"]):
            return table
    return None

def is_pedigreequery_stakes_page(soup_obj):
    if soup_obj.find("table", class_="pedigreetable"):
        return False
    if _find_pq_stakes_table(soup_obj) is None:
        return False
    txt = soup_obj.get_text(" ", strip=True).lower()
    if "search for races where" in txt:
        return True
    return "query_type=stakes" in str(soup_obj)

def _extract_race_page_meta(table):
    meta = {
        "race_page_id": "",
        "race_name": "",
        "country": "",
        "series_grade": "",
    }

    first_tr = table.find("tr")
    if first_tr:
        header_cells = first_tr.find_all("td")
        if header_cells:
            text_divs = header_cells[0].find_all("div")
            if len(text_divs) >= 1:
                meta["race_name"] = norm_space(text_divs[0].get_text(" ", strip=True))
            if len(text_divs) >= 2:
                meta["country"] = norm_space(text_divs[1].get_text(" ", strip=True))
            if len(text_divs) >= 3:
                meta["series_grade"] = norm_space(text_divs[2].get_text(" ", strip=True))

            a = first_tr.find("a", href=True)
            if a:
                href = a.get("href", "")
                m = re.search(r"[?&]id=(\d+)", href)
                if m:
                    meta["race_page_id"] = m.group(1)

    return meta

def _safe_text(td):
    return norm_space(td.get_text(" ", strip=True)) if td else ""

def _horse_td_info(td):
    if not td:
        return {"name": "", "pk": "", "url": ""}

    a = td.find("a", href=True)
    if not a:
        txt = _safe_text(td).replace("\xa0", " ").strip()
        return {"name": txt, "pk": "", "url": ""}

    href = a.get("href", "").strip()
    name = norm_space(a.get_text(" ", strip=True)).replace("\xa0", " ").strip()
    pk = _pk_from_horse_href(href)
    url = _horse_url_from_href(href)
    return {"name": name, "pk": pk, "url": url}

def build_csv_from_pedigreequery_stakes_per_horse(soup_obj):
    table = _find_pq_stakes_table(soup_obj)
    if table is None:
        return "", ""

    page_meta = _extract_race_page_meta(table)
    rows = table.find_all("tr")

    horse_map = {}
    copied_example_pk = ""

    for tr in rows[2:]:
        tds = tr.find_all("td")
        if len(tds) < 14:
            continue

        year = _year_digits(_safe_text(tds[0]))
        if not year:
            continue

        if "race not run" in _safe_text(tds[1]).lower():
            continue

        winner = _horse_td_info(tds[1])
        sire = _horse_td_info(tds[2])
        dam = _horse_td_info(tds[3])
        trainer = _safe_text(tds[4])
        family = _safe_text(tds[5])
        track = _safe_text(tds[6])
        distance = _safe_text(tds[7])
        grade = _safe_text(tds[8])
        surface = _safe_text(tds[9])
        second = _horse_td_info(tds[10])
        third = _horse_td_info(tds[11])
        time_text = _safe_text(tds[12])
        comment = _safe_text(tds[13])

        placing_map = [(1, winner), (2, second), (3, third)]

        for placing, horse in placing_map:
            horse_pk = horse.get("pk", "")
            horse_url = horse.get("url", "")
            horse_name = horse.get("name", "")

            if not horse_pk or not horse_url:
                continue

            race_record = {
                "race_page_id": page_meta.get("race_page_id", ""),
                "race_name": page_meta.get("race_name", ""),
                "country": page_meta.get("country", ""),
                "series_grade": page_meta.get("series_grade", ""),
                "year": year,
                "placing": placing,
                "horse_name": horse_name,
                "horse_pk": horse_pk,
                "track": track,
                "distance": distance,
                "grade": grade,
                "surface": surface,
                "time": time_text,
                "comment": comment,
                "trainer": trainer if placing == 1 else "",
                "family": family if placing == 1 else "",
                "winner": winner,
                "sire": sire if placing == 1 else {"name": "", "pk": "", "url": ""},
                "dam": dam if placing == 1 else {"name": "", "pk": "", "url": ""},
                "second": second,
                "third": third,
            }

            if horse_pk not in horse_map:
                horse_map[horse_pk] = {
                    "PrimaryKey": horse_pk,
                    "URL": horse_url,
                    "RaceDataJSON": [],
                }

            horse_map[horse_pk]["RaceDataJSON"].append(race_record)

            if not copied_example_pk:
                copied_example_pk = horse_pk

    for pk in horse_map:
        uniq = {}
        for rec in horse_map[pk]["RaceDataJSON"]:
            k = _normalize_race_record_for_key(rec)
            uniq[k] = rec
        horse_map[pk]["RaceDataJSON"] = sorted(
            uniq.values(),
            key=lambda x: (
                str(x.get("year", "") or ""),
                int(x.get("placing", 999) or 999),
                str(x.get("race_page_id", "") or ""),
            )
        )

    return dump_stakes_rows_csv(list(horse_map.values())), copied_example_pk

# =====================
# Main
# =====================
def main():
    try:
        html_code = pasteboard.string()
    except Exception as e:
        print(f"エラー: クリップボード参照で例外: {e}")
        sys.exit(1)

    if not html_code:
        print("エラー: クリップボードにHTMLがありません。")
        sys.exit(1)

    try:
        soup = BeautifulSoup(html_code, "html.parser")
    except Exception as e:
        print(f"エラー: HTMLパース失敗: {e}")
        sys.exit(1)

    csv_text = ""
    subject_pk = ""
    out_path = ""
    mode = ""

    if soup.find("table", class_="pedigreetable"):
        mode = "pedigreequery_pedigree"
        family_map = extract_family_data_map(soup)
        csv_text, subject_pk = build_csv_from_pedigreequery(
            soup,
            family_map,
            register_depth=4,
            parse_depth=5,
        )
        out_path = os.path.join(os.getcwd(), BLOOD_CSV_NAME)

    elif is_pedigreequery_stakes_page(soup):
        mode = "pedigreequery_stakes_per_horse"
        csv_text, subject_pk = build_csv_from_pedigreequery_stakes_per_horse(soup)
        out_path = os.path.join(os.getcwd(), STAKES_CSV_NAME)

    else:
        print("エラー: 対応していないHTMLです（pedigreequery Pedigree / pedigreequery Stakes Results のみ対応）。")
        sys.exit(1)

    if not ENABLE_CSV_EXPORT:
        print("ENABLE_CSV_EXPORT=False です。")
        sys.exit(0)

    print(csv_text)

    try:
        if mode == "pedigreequery_pedigree":
            append_unique_csv(csv_text, out_path, subject_pk)
        elif mode == "pedigreequery_stakes_per_horse":
            append_unique_stakes_rows_csv(csv_text, out_path, encoding=CSV_FILE_ENCODING)
    except Exception as e:
        print(f"\n[WARN] CSV保存失敗: {e}")

    try:
        to_clip = subject_pk if (COPY_SUBJECT_PK_TO_CLIPBOARD and subject_pk) else csv_text
        pasteboard.set_string(to_clip)
        if COPY_SUBJECT_PK_TO_CLIPBOARD and subject_pk:
            print(f"\n[OK] PrimaryKeyをクリップボードへコピー: {subject_pk}")
        else:
            print("\n[OK] CSVをクリップボードへコピー")

        if mode == "pedigreequery_stakes_per_horse":
            print(f"[OK] {STAKES_CSV_NAME} へ保存（1頭1行・年別JSON配列をマージ）: {out_path} (encoding={CSV_FILE_ENCODING})")
        else:
            print(f"[OK] {BLOOD_CSV_NAME} を更新（主役PKは上書き / その他PKは重複除外）: {out_path} (encoding={CSV_FILE_ENCODING})")
    except Exception as e:
        print(f"\n[WARN] クリップボードコピー失敗: {e}")

def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="PedigreeQuery scraper: race results and horse pedigree in one script."
    )
    parser.add_argument("--race-url", action="append", default=[], help="Specific race result URL. Repeatable.")
    parser.add_argument("--race-list-url", action="append", default=[], help="Stakes search result/list URL. Repeatable.")
    parser.add_argument("--horse-url", action="append", default=[], help="Specific horse pedigree URL. Repeatable.")
    parser.add_argument("--stakes-csv", action="append", default=[], help="Existing stakes_horses.csv path to seed horse URLs from its URL column.")
    parser.add_argument("--blood-out", default=os.path.join(os.getcwd(), OUTPUT_DIR_NAME, BLOOD_CSV_NAME), help="Output path for horse pedigree CSV.")
    parser.add_argument("--stakes-out", default=os.path.join(os.getcwd(), OUTPUT_DIR_NAME, STAKES_CSV_NAME), help="Output path for race result CSV.")
    parser.add_argument("--register-depth", type=int, default=4, help="Depth to register into blood CSV per fetched horse page.")
    parser.add_argument("--parse-depth", type=int, default=5, help="Depth to parse and use as next frontier horse URLs.")
    parser.add_argument("--sleep", type=float, default=0.0, help="Sleep seconds between requests.")
    parser.add_argument("--html-file", help="Local HTML file for one-shot conversion.")
    parser.add_argument("--from-clipboard", action="store_true", help="Read one-shot HTML from clipboard.")
    parser.add_argument("--no-clipboard-copy", action="store_true", help="Disable clipboard copy of result when using one-shot mode.")
    return parser.parse_args(argv)

def run_one_shot_html(html_code: str, blood_out_path: str, stakes_out_path: str, allow_clipboard_copy: bool):
    if not html_code:
        print("ERROR: empty HTML input")
        return 1

    try:
        soup = BeautifulSoup(html_code, "html.parser")
    except Exception as e:
        print(f"ERROR: html parse failed: {e}")
        return 1

    csv_text = ""
    subject_pk = ""
    out_path = ""
    mode = ""

    if soup.find("table", class_="pedigreetable"):
        mode = "pedigreequery_pedigree"
        family_map = extract_family_data_map(soup)
        csv_text, subject_pk = build_csv_from_pedigreequery(
            soup,
            family_map,
            register_depth=4,
            parse_depth=5,
        )
        out_path = blood_out_path
    elif is_pedigreequery_stakes_page(soup):
        mode = "pedigreequery_stakes_per_horse"
        csv_text, subject_pk = build_csv_from_pedigreequery_stakes_per_horse(soup)
        out_path = stakes_out_path
    else:
        print("ERROR: unsupported HTML. Expected pedigree page or stakes result page.")
        return 1

    if not ENABLE_CSV_EXPORT:
        print("ENABLE_CSV_EXPORT=False")
        return 0

    print(csv_text)

    try:
        if mode == "pedigreequery_pedigree":
            append_unique_csv(csv_text, out_path, subject_pk)
        else:
            append_unique_stakes_rows_csv(csv_text, out_path, encoding=CSV_FILE_ENCODING)
    except Exception as e:
        print(f"[WARN] CSV write failed: {e}")

    if allow_clipboard_copy and pasteboard is not None:
        try:
            to_clip = subject_pk if (COPY_SUBJECT_PK_TO_CLIPBOARD and subject_pk) else csv_text
            pasteboard.set_string(to_clip)
            if COPY_SUBJECT_PK_TO_CLIPBOARD and subject_pk:
                print(f"[OK] copied PrimaryKey: {subject_pk}")
            else:
                print("[OK] copied CSV to clipboard")
        except Exception as e:
            print(f"[WARN] clipboard copy failed: {e}")

    print(f"[OK] wrote {out_path}")
    return 0

def cli_main(argv=None):
    args = parse_args(argv)
    blood_out_path = args.blood_out
    stakes_out_path = args.stakes_out

    horse_seed_urls = []
    for x in args.horse_url:
        x = _normalize_url(x)
        if x and x not in horse_seed_urls:
            horse_seed_urls.append(x)

    race_urls = []
    for x in args.race_url:
        x = _normalize_url(x)
        if x and x not in race_urls:
            race_urls.append(x)

    race_list_urls = []
    for x in args.race_list_url:
        x = _normalize_url(x)
        if x and x not in race_list_urls:
            race_list_urls.append(x)

    if args.html_file:
        with open(args.html_file, "r", encoding="utf-8", errors="replace") as f:
            return run_one_shot_html(
                f.read(),
                blood_out_path=blood_out_path,
                stakes_out_path=stakes_out_path,
                allow_clipboard_copy=(not args.no_clipboard_copy),
            )

    if args.from_clipboard:
        if pasteboard is None:
            print("ERROR: pasteboard module is not available in this environment")
            return 1
        try:
            html_code = pasteboard.string()
        except Exception as e:
            print(f"ERROR: clipboard read failed: {e}")
            return 1
        return run_one_shot_html(
            html_code,
            blood_out_path=blood_out_path,
            stakes_out_path=stakes_out_path,
            allow_clipboard_copy=(not args.no_clipboard_copy),
        )

    session = make_session()

    for list_url in race_list_urls:
        print(f"[LIST] collecting race URLs from {list_url}")
        try:
            soup = fetch_soup(list_url, session=session)
            race_map = collect_race_urls_from_list_page(soup)
            for race_url in race_map.keys():
                if race_url not in race_urls:
                    race_urls.append(race_url)
            print(f"[OK] collected {len(race_map)} race URLs")
        except Exception as e:
            print(f"[WARN] race list fetch failed: {list_url} ({e})")

    if race_urls:
        process_race_targets(race_urls, session=session, stakes_out_path=stakes_out_path, sleep_sec=args.sleep)

    for stakes_csv_path in args.stakes_csv:
        print(f"[CSV] loading horse URLs from {stakes_csv_path}")
        try:
            for url in load_horse_urls_from_stakes_csv(stakes_csv_path, encoding=CSV_FILE_ENCODING):
                if url not in horse_seed_urls:
                    horse_seed_urls.append(url)
        except Exception as e:
            print(f"[WARN] stakes csv load failed: {stakes_csv_path} ({e})")

    if horse_seed_urls:
        process_horse_targets(
            horse_seed_urls,
            session=session,
            blood_out_path=blood_out_path,
            register_depth=args.register_depth,
            parse_depth=args.parse_depth,
            sleep_sec=args.sleep,
        )

    if not race_urls and not horse_seed_urls and not args.stakes_csv:
        print("ERROR: no input. Use --race-url / --race-list-url / --horse-url / --stakes-csv / --html-file / --from-clipboard")
        return 1

    return 0

if __name__ == "__main__":
    sys.exit(cli_main())
