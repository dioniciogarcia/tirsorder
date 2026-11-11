# -*- coding: utf-8 -*-
"""
NuORDER Client
==============
Turns brand linesheets / order forms (.xlsx / .xlsm / .xlsb / .csv) into NuORDER-ready product data.

  - One file, or several files combined into one master file (each linesheet in its own light color).
  - Data Identifier: recognises JOOR exports, Ferragamo, Gucci, Velvet and NuORDER Default Data files
    (brand profiles); anything else uses the general rules with columns chosen in the window.
  - Images: pictures pasted in the sheets are extracted, image links are downloaded, and a folder of
    loose images can be matched to the products and copied renamed - all as BRAND_ID_01, _02 ...

Needs openpyxl (Excel files) and Pillow (image resizing - optional: without it images are copied as they are).
"""

import csv
import datetime
import hashlib
import io
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import traceback
import urllib.error
import urllib.request
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
except Exception:  # pragma: no cover - lets the core logic be tested without a screen
    tk = None

APP_NAME = "NuORDER Client"
APP_VERSION = "3.5"
GITHUB_REPO = "dioniciogarcia/tirsorder"          # where new versions are published (Releases)
RELEASE_EXE = "NuORDER-Client.exe"                # the program GitHub compiles for every version
RELEASE_PY = "nuorder_client.py"

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".tif", ".tiff", ".bmp"}

FONT_HEADER = Font(name="Calibri", size=11, bold=True)
FONT_DATA = Font(name="Calibri", size=11)
FILL_UNMATCHED = PatternFill("solid", fgColor="FFC7CE")


class UserError(Exception):
    """An expected problem that we can explain to the user in plain words."""


# ======================================================================
#  SMALL HELPERS
# ======================================================================

def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def clean_text(v):
    """Any cell value -> tidy single-line text ('' if empty)."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).replace("\u00a0", " ").strip()
    return re.sub(r"\s+", " ", s)


def clean_code(v):
    """Codes (style numbers, color codes ...): remove float artefacts such as 3501.0."""
    s = clean_text(v)
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]
    return s


def norm_header(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def safe_filename(name):
    s = re.sub(r"[/\\]", "-", str(name))                    # RED/PRT -> RED-PRT
    s = re.sub(r'[<>:"|?*\x00-\x1f]', "_", s).strip(" .")
    return s[:150] if s else "image"


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", str(s))]


def parse_number(v):
    if is_blank(v) or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    s = re.sub(r"[^\d.,\-]", "", str(v))
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s:
        # 1,234 -> 1234   but   12,50 -> 12.50
        s = s.replace(",", "") if re.fullmatch(r"-?\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def strip_style_from_name(name, style):
    """'A72D03 - OFF SHOULDER DRESS' -> 'OFF SHOULDER DRESS' (only if style is at the start)."""
    if not name or not style:
        return name
    n = name.strip()
    if n.upper().startswith(style.upper()):
        rest = n[len(style):]
        m = re.match(r"^[\s\-_\u2013\u2014:|/]+", rest)
        if m:
            rest2 = rest[m.end():].strip()
            if rest2:
                return rest2
    return n


# ----------------------------------------------------------------------
#  Sizes
# ----------------------------------------------------------------------

ONE_SIZE_TOKENS = {
    "ONE SIZE", "ONESIZE", "ONE-SIZE", "ONE SIZE FITS ALL", "ONE SIZE FITS MOST",
    "O/S", "OS", "OSFA", "OSFM", "OSFM.", "ONE", "UNI", "UNISIZE", "UNIVERSAL", "NO SIZE",
}

ALPHA_RANK = {
    "XXXS": 0, "3XS": 0, "XXS": 1, "2XS": 1, "XS": 2, "XS/S": 3, "S": 4, "S/M": 5,
    "M": 6, "M/L": 7, "L": 8, "L/XL": 9, "XL": 10, "XL/XXL": 11, "XXL": 12, "2XL": 12,
    "XXXL": 13, "3XL": 13, "XXXXL": 14, "4XL": 14, "5XL": 15,
}


def normalize_size(raw):
    s = clean_code(raw).upper().replace("*", "").strip()          # '*00' -> '00'
    s = re.sub(r"^SIZE\s+", "", s)
    if s in ONE_SIZE_TOKENS:
        return "OS"
    if re.fullmatch(r"\d+\.\d*0+", s):          # 7.0 / 8.50 -> 7 / 8.5
        s = s.rstrip("0").rstrip(".")
    return s


def size_sort_key(s):
    if s == "00":
        return (1, -1.0, "")
    if s == "000":
        return (1, -2.0, "")
    if s in ALPHA_RANK:
        return (0, ALPHA_RANK[s], "")
    try:
        return (1, float(s), "")
    except ValueError:
        return (2, 0, s)


ALPHA_SCALE = ["XXXS", "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXXXL"]
ALPHA_ALIAS = {"3XS": "XXXS", "2XS": "XXS", "2XL": "XXL", "3XL": "XXXL", "4XL": "XXXXL"}
RANGE_RE = re.compile(r"^\s*([A-Z0-9.]+)\s*(?:-|–|—)\s*([A-Z0-9.]+)\s*$")


def expand_size_range(tok, even_all=False):
    """'23-34' -> 23, 24 ... 34   |   'XXS-XL' -> XXS, XS, S, M, L, XL   |   '00-14' -> 00, 0, 2 ... 14
    Both ends even and the end 20 or under (women's sizes 0-14, 2-12) -> steps of 2.
    even_all=True (Italian / EU clothing, e.g. Armani 46-60) -> steps of 2 whenever both ends are even."""
    m = RANGE_RE.match(tok.upper())
    if not m:
        return None
    a, b = m.group(1), m.group(2)
    a_, b_ = ALPHA_ALIAS.get(a, a), ALPHA_ALIAS.get(b, b)
    if a_ in ALPHA_SCALE and b_ in ALPHA_SCALE:
        i, j = ALPHA_SCALE.index(a_), ALPHA_SCALE.index(b_)
        return ALPHA_SCALE[i:j + 1] if i < j else None
    try:
        x, y = float(a), float(b)
    except ValueError:
        return None
    lead = []
    if re.fullmatch(r"0{2,}", a):                    # '00-14' -> 00, 0, 2 ...
        lead = [a]
    if x % 1 or y % 1:
        step = 0.5                                  # 5.5-11 -> half sizes
    elif x % 2 == 0 and y % 2 == 0 and (y <= 20 or even_all):
        step = 2.0
    else:
        step = 1.0
    if not (x < y) or (y - x) / step > 60:
        return None
    out, n = list(lead), x
    while n <= y + 1e-9:
        out.append(str(int(n)) if n == int(n) else str(n))
        n += step
    return out


def split_sizes(v, even_all=False):
    """One cell -> list of size tokens ('XS, S, M' -> ['XS','S','M']; ranges like '23-34' are expanded,
    size notes such as '5-16 (HALF SIZES OFFERED BUT NO 5.5)' are followed)."""
    if is_blank(v):
        return []
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return [normalize_size(v)]
    text = str(v)
    if "(" in text and re.search(r"\d\s*(?:-|TO)\s*\d", text, re.I):
        return parse_size_scale(text, even_all=even_all)
    if "/" in text and re.search(r"\d\s*-\s*\d.*/", text):
        return parse_size_scale(text, even_all=even_all)
    out = []
    for tok in re.split(r"[,;\n]+", text):
        tok = normalize_size(tok)
        if not tok:
            continue
        rng = expand_size_range(tok, even_all)
        out.extend(rng if rng else [tok])
    return out


def sort_sizes(sizes):
    uniq = list(dict.fromkeys(sizes))
    return sorted(uniq, key=size_sort_key)


# ----------------------------------------------------------------------
#  Dates
# ----------------------------------------------------------------------

MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}

DATE_RE = re.compile(
    r"(?P<d4>\d{1,2})[\s\-]+(?P<mon4>[A-Za-z]{3,9})\.?,?[\s\-]+(?P<y4>\d{4})"
    r"|(?P<y1>\d{4})[-/.](?P<m1>\d{1,2})[-/.](?P<d1>\d{1,2})"
    r"|(?P<m2>\d{1,2})(?P<sep>[/.-])(?P<d2>\d{1,2})(?:(?P=sep)(?P<y2>\d{2,4}))?"
    r"|(?P<mon>[A-Za-z]{3,9})\.?\s+(?P<d3>\d{1,2})(?!\d)(?:,?\s+(?P<y3>\d{4}))?"
)


def parse_dates(v, default_year):
    """Anything that looks like a date / delivery window -> list of datetime.date (0, 1 or 2+)."""
    if is_blank(v):
        return []
    if isinstance(v, datetime.datetime):
        return [v.date()]
    if isinstance(v, datetime.date):
        return [v]
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 < v < 80000:
            return [(datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(v))).date()]
        return []
    out = []
    for m in DATE_RE.finditer(str(v)):
        try:
            if m.group("d4"):
                mon = MONTHS.get(m.group("mon4")[:3].lower())
                if not mon:
                    continue
                out.append(datetime.date(int(m.group("y4")), mon, int(m.group("d4"))))
            elif m.group("y1"):
                out.append(datetime.date(int(m.group("y1")), int(m.group("m1")), int(m.group("d1"))))
            elif m.group("m2"):
                y = m.group("y2")
                y = default_year if not y else (2000 + int(y) if len(y) == 2 else int(y))
                out.append(datetime.date(y, int(m.group("m2")), int(m.group("d2"))))
            else:
                mon = MONTHS.get(m.group("mon")[:3].lower())
                if not mon:
                    continue
                y = int(m.group("y3")) if m.group("y3") else default_year
                out.append(datetime.date(y, mon, int(m.group("d3"))))
        except ValueError:
            continue
    return out


def year_from_season(season):
    s = str(season or "")
    m = re.search(r"(20\d{2})", s)
    if m:
        return int(m.group(1))
    m = re.search(r"[A-Za-z](\d{2})\b", s)      # F27, SU27, SS27 ...
    if m:
        return 2000 + int(m.group(1))
    return datetime.date.today().year


SEASON_WORDS = {
    "SU": "SUMMER", "SUM": "SUMMER", "SUMMER": "SUMMER",
    "SP": "SPRING", "SPR": "SPRING", "SPRING": "SPRING",
    "F": "FALL", "FA": "FALL", "FALL": "FALL", "AUTUMN": "FALL",
    "PF": "PRE-FALL", "PREFALL": "PRE-FALL", "PRE-FALL": "PRE-FALL", "PRE FALL": "PRE-FALL",
    "WI": "WINTER", "WINTER": "WINTER",
    "RESORT": "RESORT", "CRUISE": "CRUISE", "HOLIDAY": "HOLIDAY",
}


def season_full(text, default_year=None):
    """'SU27' / 'Pre-Fall 2027' / 'summer 27' -> 'SUMMER 2027'.  Unknown codes are kept as typed (upper case)."""
    s_ = clean_text(text).upper()
    if not s_:
        return ""
    m = re.fullmatch(r"([A-Z][A-Z \-]*?)\s*[-_]?\s*(\d{4}|\d{2})?", s_)
    if not m:
        return s_
    word = re.sub(r"\s+", " ", m.group(1)).strip(" -")
    year = m.group(2)
    full = SEASON_WORDS.get(word) or SEASON_WORDS.get(word.replace(" ", ""))
    if not full:
        return s_
    if year and len(year) == 2:
        year = str(2000 + int(year))
    if not year and default_year:
        year = str(default_year)
    return f"{full} {year}".strip()


def detect_season_from_name(name):
    """Finds the season written in a file name and returns it AS WRITTEN (SU27, SS27, FALL 2027 ...)."""
    n = os.path.splitext(os.path.basename(name))[0].upper()
    m = re.search(r"(?<![A-Z0-9])(SUMMER|SPRING|FALL|PRE-?FALL|WINTER|RESORT|CRUISE|HOLIDAY)[\s_\-]*(\d{4}|\d{2})(?![0-9])", n)
    if not m:
        m = re.search(r"(?<![A-Z0-9])(SU|SP|PF|FA|F|WI|SS|FW|AW)[\s_\-]?(\d{2})(?![0-9])", n)
    if not m:
        return ""
    return re.sub(r"[\s_\-]+", " ", m.group(0)).strip() if len(m.group(1)) > 2 else m.group(1) + m.group(2)


def sheet_linesheet_name(path, cfg):
    """File name + ' RP'; when one file holds several sheets: 'File - Sheet RP'."""
    base = default_linesheet_name(path)
    if cfg.get("multi_sheet") and cfg.get("sheet"):
        base = base[:-3] + " - " + clean_text(cfg["sheet"]) + " RP"
    return base


def linesheet_label(v):
    s_ = clean_text(v)
    if not s_:
        return ""
    return s_ if s_.upper().endswith(" RP") else s_ + " RP"


def default_linesheet_name(path):
    """'TH_ACC_WOMENS_SU27_DAVID_JONES_ORDER_FORM__1_.xlsx' -> 'TH ACC WOMENS SU27 DAVID JONES ORDER FORM RP'."""
    b = os.path.splitext(os.path.basename(path))[0]
    for _ in range(2):      # download tags such as "(1)", "(version 1)", "__1_", "_Processed"
        b = re.sub(r"(__\d+_?|\s*\((?:version\s*)?\d+\)|_version_\d+|_processed(_v\d+)?)\s*$", "", b, flags=re.I)
    if " " not in b:
        b = re.sub(r"\s+", " ", b.replace("_", " "))
    b = re.sub(r"\s+RP$", "", b.strip())
    return b + " RP"


def clean_country(v):
    s_ = clean_text(v)
    s_ = re.sub(r"(?i)^\s*(made\s+in|product\s+of|country\s+of\s+origin\s*:?)\s*", "", s_)
    s_ = s_.replace(";", " ").strip(" ,;.")
    return re.sub(r"\s+", " ", s_).upper()


def clean_fabric(v):
    s_ = clean_text(v)
    s_ = re.sub(r"\s*;(?:\s*;)*\s*", "; ", s_).strip(" ;,")
    return s_


CURRENCY_CODES = {"USD", "AUD", "NZD", "EUR", "GBP", "CAD", "JPY", "CHF", "SEK", "DKK", "NOK",
                  "HKD", "SGD", "MXN", "CNY", "ZAR", "AED", "KRW", "INR", "BRL"}


def currency_code(v):
    c = clean_text(v).upper()
    return c if c in CURRENCY_CODES else ""


def infer_file_currency(mapping):
    """Currency of an un-labelled wholesale column, guessed from columns such as 'RRP AUD'."""
    for cur in ("AUD", "EUR", "GBP", "CAD", "USD", "NZD"):
        if mapping.get("WHOLESALE " + cur) is not None or mapping.get("RETAIL " + cur) is not None:
            return cur
    return "USD"


DATE_FORMAT_CHOICES = OrderedDict([
    ("Real Excel date, shown as 1/15/2027   (NuORDER)", "EXCEL"),
    ("Text 1/15/2027", "M/D/YYYY"),
    ("Text MM-DD-YY   (old Gemini rule)", "MM-DD-YY"),
])


def format_date(d, fmt):
    if d is None:
        return None
    if fmt == "MM-DD-YY":
        return d.strftime("%m-%d-%y")
    if fmt == "MM/DD/YYYY":
        return d.strftime("%m/%d/%Y")
    if fmt == "M/D/YYYY":
        return f"{d.month}/{d.day}/{d.year}"
    return d


# ======================================================================
#  READING SOURCE FILES (xlsx / xlsm / csv)
# ======================================================================

class _CsvCell:
    __slots__ = ("value", "fill", "hyperlink", "row", "column")

    def __init__(self, value, row, column):
        self.value = value
        self.fill = None
        self.hyperlink = None
        self.row = row
        self.column = column


class CsvSheet:
    """Looks just enough like an openpyxl worksheet for our needs (used for CSV and .xlsb files)."""

    def __init__(self, path=None, rows=None, title="CSV"):
        self.rows = rows if rows is not None else read_csv_rows(path)
        self.max_row = len(self.rows)
        self.max_column = max((len(r) for r in self.rows), default=0)
        self.title = title

        class _M:
            ranges = []
        self.merged_cells = _M()

    def cell(self, row, column):
        try:
            v = self.rows[row - 1][column - 1]
        except IndexError:
            v = None
        return _CsvCell(v if v != "" else None, row, column)

    def iter_rows(self, min_row=1, max_row=None, max_col=None):
        max_row = max_row or self.max_row
        max_col = max_col or self.max_column
        for r in range(min_row, max_row + 1):
            yield [self.cell(r, c) for c in range(1, max_col + 1)]


def read_csv_rows(path):
    raw = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                raw = f.read()
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raise UserError("Could not read this CSV file (unknown text encoding).")
    try:
        dialect = csv.Sniffer().sniff(raw[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    return [list(r) for r in csv.reader(raw.splitlines(), dialect)]


# ----------------------------------------------------------------------
#  Excel Binary Workbook (.xlsb) - small built-in reader (values only)
# ----------------------------------------------------------------------

_XLSB_ERRORS = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!", 0x1D: "#NAME?",
                0x24: "#NUM!", 0x2A: "#N/A", 0x2B: "#GETTING_DATA"}


def _xlsb_records(data):
    """Yields (record_type, payload) from an .xlsb part."""
    i, n = 0, len(data)
    while i < n:
        rt = data[i]
        i += 1
        if rt & 0x80:
            rt = (rt & 0x7F) | ((data[i] & 0x7F) << 7)
            i += 1
        size, shift = 0, 0
        for _ in range(4):
            b = data[i]
            i += 1
            size |= (b & 0x7F) << shift
            shift += 7
            if not b & 0x80:
                break
        yield rt, data[i:i + size]
        i += size


def _xlsb_wstr(buf, pos):
    import struct
    cch = struct.unpack_from("<I", buf, pos)[0]
    pos += 4
    if cch == 0xFFFFFFFF:
        return None, pos
    s = buf[pos:pos + cch * 2].decode("utf-16-le", errors="replace")
    return s, pos + cch * 2


def _xlsb_rk(v):
    import struct
    if v & 2:
        num = v >> 2
        if num & (1 << 29):
            num -= 1 << 30
        num = float(num)
    else:
        num = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    if v & 1:
        num /= 100.0
    return int(num) if float(num).is_integer() else num


def xlsb_sheet_parts(z):
    """[(sheet name, part path inside the zip, visible?)] in workbook order."""
    import struct
    import xml.etree.ElementTree as ET
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.bin.rels"))}
    out = []
    for rt, p in _xlsb_records(z.read("xl/workbook.bin")):
        if rt == 156:                                     # BrtBundleSh
            state = struct.unpack_from("<I", p, 0)[0]
            rid, pos = _xlsb_wstr(p, 8)
            name, _ = _xlsb_wstr(p, pos)
            out.append((name, _zip_join("xl", rels.get(rid, "")), state == 0))
    return out


def read_xlsb(path, sheet=None, max_rows=None):
    """Returns (sheet names, rows of the chosen sheet as lists of values)."""
    import struct
    import zipfile
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise UserError("This .xlsb file could not be opened (it may be damaged or password protected).")
    with z:
        parts = xlsb_sheet_parts(z)
        names = [p[0] for p in parts]
        if sheet is None:
            return names, []
        part = next((p for p in parts if p[0] == sheet), parts[0])[1]
        sst = []
        if "xl/sharedStrings.bin" in z.namelist():
            for rt, p in _xlsb_records(z.read("xl/sharedStrings.bin")):
                if rt == 19:                              # BrtSSTItem
                    s_, _ = _xlsb_wstr(p, 1)
                    sst.append(s_ or "")
        rows, row, cur = {}, None, -1
        for rt, p in _xlsb_records(z.read(part)):
            if rt == 0:                                   # BrtRowHdr
                cur = struct.unpack_from("<I", p, 0)[0]
                if max_rows and cur >= max_rows:
                    break
                row = rows.setdefault(cur, {})
                continue
            if row is None or not 1 <= rt <= 11 or len(p) < 8:
                continue
            col = struct.unpack_from("<I", p, 0)[0]
            v = None
            if rt == 2:
                v = _xlsb_rk(struct.unpack_from("<I", p, 8)[0])
            elif rt in (3, 11):
                v = _XLSB_ERRORS.get(p[8], "#ERROR")
            elif rt in (4, 10):
                v = bool(p[8])
            elif rt in (5, 9):
                v = struct.unpack_from("<d", p, 8)[0]
                v = int(v) if v.is_integer() and abs(v) < 1e15 else v
            elif rt in (6, 8):
                v, _ = _xlsb_wstr(p, 8)
            elif rt == 7:
                k = struct.unpack_from("<I", p, 8)[0]
                v = sst[k] if k < len(sst) else None
            if v not in (None, ""):
                row[col] = v
    used = [r for r, cells in rows.items() if cells]
    if not used:
        return names, []
    width = max(max(c) for c in rows.values() if c) + 1
    out = []
    for r in range(0, max(used) + 1):
        cells = rows.get(r, {})
        out.append([cells.get(c) for c in range(width)])
    return names, out


def check_extension(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        raise UserError("Old .xls files are not supported. Open the file in Excel and use "
                        "File > Save As > Excel Workbook (.xlsx), then try again.")
    if ext not in (".xlsx", ".xlsm", ".xlsb", ".csv"):
        raise UserError("Please choose an .xlsx, .xlsm, .xlsb or .csv file.")
    return ext


def list_sheets(path, visible_only=False):
    ext = check_extension(path)
    if ext == ".csv":
        return ["CSV"]
    if ext == ".xlsb":
        import zipfile
        with zipfile.ZipFile(path) as z:
            return [n for n, _p, vis in xlsb_sheet_parts(z) if vis or not visible_only]
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        return [ws.title for ws in wb.worksheets if not visible_only or ws.sheet_state == "visible"]
    finally:
        wb.close()


def read_preview(path, sheet, max_rows=60):
    ext = check_extension(path)
    if ext == ".csv":
        return read_csv_rows(path)[:max_rows]
    if ext == ".xlsb":
        return read_xlsb(path, sheet, max_rows)[1][:max_rows]
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
        return [list(r) for r in ws.iter_rows(min_row=1, max_row=max_rows, values_only=True)]
    finally:
        wb.close()


def open_sheet(path, sheet, data_only=True):
    ext = check_extension(path)
    if ext == ".csv":
        return CsvSheet(path)
    if ext == ".xlsb":
        names, rows = read_xlsb(path, sheet if sheet else None)
        if sheet not in names:
            names, rows = read_xlsb(path, names[0])
        return CsvSheet(rows=rows, title=sheet or names[0])
    wb = load_workbook(path, data_only=data_only)
    return wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]


def merged_value_map(ws):
    """Values for the hidden cells of merged ranges, so merged style numbers apply to every row."""
    m = {}
    try:
        for rng in ws.merged_cells.ranges:
            anchor = ws.cell(rng.min_row, rng.min_col).value
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != (rng.min_row, rng.min_col):
                        m[(r, c)] = anchor
    except Exception:
        pass
    return m


def is_black_fill(cell):
    try:
        f = cell.fill
        if f is None or f.fill_type != "solid":
            return False
        c = f.fgColor
        if c is None:
            return False
        if c.type == "theme":
            return c.theme == 1 and abs(c.tint or 0) < 0.2
        if c.type == "rgb":
            return str(c.rgb).upper() in ("FF000000", "00000000")
        if c.type == "indexed":
            return c.indexed in (8, 64) and c.indexed == 8
    except Exception:
        return False
    return False


# ----------------------------------------------------------------------
#  Header detection and automatic column matching
# ----------------------------------------------------------------------

FIELD_SYNONYMS = OrderedDict([   # the order inside each list = priority
    ("STYLE NUMBER", ["style number", "style no", "style code", "style num", "style #", "style", "item number",
                      "item no", "item #", "item", "sku", "product code", "vpn", "pid", "article number",
                      "article", "material number", "model"]),
    ("NAME", ["style name", "product name", "item name", "name", "product description", "style description",
              "title", "short description", "description", "model"]),
    ("COLOR", ["color name", "colour name", "color", "colour", "vend color", "brand color", "brand colour",
               "color description", "color desc", "colorway"]),
    ("COLOR CODE", ["color code", "colour code", "color no", "color number", "color #", "clr code",
                    "vendor color code"]),
    ("SIZE", ["all sizes grouped together", "sizes", "size", "size range", "size scale", "size run", "size offer",
              "sizing", "available sizes", "size notes", "size description"]),
    ("LINESHEET SEPARATOR", ["linesheet separator", "linesheet separator name", "line sheet separator",
                             "line sheet separator name"]),
    ("DELIVERY MONTH", ["delivery month", "available month", "delivery bucket", "deliv", "sub delivery",
                        "ship month", "delivery mth"]),
    ("LINESHEET NAME", ["linesheet name", "line sheet name", "linesheet", "line sheet", "collection name"]),
    ("SEASON", ["season name", "season", "delivery season", "season description"]),
    ("SEASON YEAR", ["season year", "year"]),
    ("DELIVERY", ["delivery window", "ship window", "delivery dates", "delivery date", "delivery",
                  "ship dates", "ship date"]),
    ("AVAILABLE FROM", ["start ship date", "ship start date", "delivers from", "delivery from", "available from",
                        "start date", "ship start", "delivery start", "ship from", "start ship"]),
    ("AVAILABLE UNTIL", ["end ship date", "ship end date", "delivers to", "delivery to", "delivery end",
                         "available until", "available to", "end date", "ship end", "ship until", "end ship"]),
    ("FABRIC DESCRIPTION", ["fabric description", "fabric content", "fabrication", "fabric", "composition",
                            "material", "materials", "content"]),
    ("US COUNTRY OF SOURCING", ["us country of sourcing", "country of sourcing", "country of origin", "coo",
                                "origin", "made in", "country of manufacture"]),
    ("DEPARTMENT", ["department", "dept", "division name", "division"]),
    ("WHOLESALE PRICE", ["wholesale price", "wholesale price 1", "wholesale", "wsp", "cost ex gst", "cost price",
                         "wholesale cost", "unit cost", "cost", "whsl price", "whsl"]),
    ("RETAIL PRICE", ["suggested retail price", "suggested retail price 1", "retail price", "retail", "msrp",
                      "rrp", "srp", "suggested retail"]),
    ("PRICE CURRENCY", ["wholesale currency", "wholesale currency 1", "price currency", "currency"]),
    ("PRICE TYPE LABEL", ["price type label", "price type label 1"]),
    ("RETAIL CURRENCY", ["retail currency", "retail currency 1"]),
    ("WHOLESALE USD", ["wholesale usd", "usd wholesale", "wsp usd", "whsl usd"]),
    ("RETAIL USD", ["retail usd", "usd retail", "rrp usd", "msrp usd"]),
])
for _cur in ("AUD", "EUR", "GBP", "CAD"):
    _c = _cur.lower()
    FIELD_SYNONYMS["WHOLESALE " + _cur] = [f"wholesale {_c}", f"{_c} wholesale", f"wsp {_c}", f"cost {_c}"]
    FIELD_SYNONYMS["RETAIL " + _cur] = [f"retail {_c}", f"{_c} retail", f"rrp {_c}", f"msrp {_c}"]
FIELD_SYNONYMS["RETAIL NZD"] = ["retail nzd", "nzd retail", "rrp nzd", "msrp nzd"]

_CUR_WORDS = ["usd", "aud", "eur", "gbp", "cad", "nzd", "jpy", "hkd", "currency", "total", "type", "label"]
FIELD_EXCLUDE = {
    "STYLE NUMBER": ["name", "desc", "description", "color", "colour", "size", "id"],
    "NAME": ["number", "code", "color", "colour", "no", "id"],
    "COLOR": ["code", "family", "no", "number", "id"],
    "SIZE": ["code", "chart", "id", "nrf", "eu", "uk", "fit", "information"],
    "DELIVERY": ["season", "id", "minimum", "days", "terms", "month", "bucket"],
    "FABRIC DESCRIPTION": ["id", "code"],
    "US COUNTRY OF SOURCING": ["id", "code"],
    "DEPARTMENT": ["id", "code"],
    "WHOLESALE PRICE": _CUR_WORDS + ["retail"],
    "RETAIL PRICE": _CUR_WORDS + ["wholesale"],
    "PRICE CURRENCY": ["retail", "id"],
}
NO_SUBSTRING = {"SEASON", "SEASON YEAR", "LINESHEET NAME", "LINESHEET SEPARATOR", "DELIVERY MONTH", "PRICE TYPE LABEL", "RETAIL CURRENCY"}

MAP_FIELDS = [   # (internal key, label shown in the window)
    ("STYLE NUMBER", "Style Number  *"),
    ("NAME", "Name  *"),
    ("COLOR", "Color  *"),
    ("COLOR CODE", "Color Code"),
    ("SIZE", "All sizes in one column (e.g. 23-34, XS-XL, list)"),
    ("LINESHEET NAME", "Linesheet Name (column)"),
    ("LINESHEET SEPARATOR", "Linesheet Separator (column)"),
    ("SEASON", "Season (column)"),
    ("SEASON YEAR", "Season Year (column)"),
    ("DELIVERY", "Delivery window (both dates in one column)"),
    ("AVAILABLE FROM", "Available From"),
    ("AVAILABLE UNTIL", "Available Until"),
    ("DELIVERY MONTH", "Delivery Month"),
    ("FABRIC DESCRIPTION", "Fabric Description"),
    ("US COUNTRY OF SOURCING", "Country of Sourcing"),
    ("DEPARTMENT", "Department"),
    ("WHOLESALE PRICE", "Wholesale price (amount)"),
    ("RETAIL PRICE", "Retail price (amount)"),
    ("PRICE CURRENCY", "Wholesale currency (column)"),
    ("PRICE TYPE LABEL", "Price Type Label (column)"),
    ("RETAIL CURRENCY", "Retail currency (column)"),
    ("WHOLESALE USD", "Wholesale USD (already in USD)"),
    ("RETAIL USD", "Retail USD (already in USD)"),
    ("WHOLESALE AUD", "Wholesale AUD"),
    ("RETAIL AUD", "Retail AUD"),
    ("RETAIL NZD", "Retail NZD"),
    ("WHOLESALE EUR", "Wholesale EUR"),
    ("RETAIL EUR", "Retail EUR"),
    ("WHOLESALE GBP", "Wholesale GBP"),
    ("RETAIL GBP", "Retail GBP"),
    ("WHOLESALE CAD", "Wholesale CAD"),
    ("RETAIL CAD", "Retail CAD"),
]

PRICE_ORDER = ["WHOLESALE USD", "RETAIL USD", "WHOLESALE AUD", "RETAIL AUD", "RETAIL NZD",
               "WHOLESALE EUR", "RETAIL EUR", "WHOLESALE GBP", "RETAIL GBP", "WHOLESALE CAD", "RETAIL CAD"]


def detect_header_row(rows):
    """1-based number of the row that most looks like a header."""
    all_syn = {s for lst in FIELD_SYNONYMS.values() for s in lst}
    best_row, best_score = None, 0
    for i, row in enumerate(rows[:40], 1):
        score = sum(1 for v in row if not is_blank(v) and norm_header(v) in all_syn)
        if score > best_score:
            best_row, best_score = i, score
    if best_row:
        return best_row
    for i, row in enumerate(rows[:40], 1):
        if sum(1 for v in row if not is_blank(v)) >= 3:
            return i
    return 1


def make_labels(row):
    """Header row -> unique, readable column names (trailing empty columns dropped)."""
    row = list(row)
    while row and is_blank(row[-1]):
        row.pop()
    labels, seen = [], {}
    for i, v in enumerate(row, 1):
        t = clean_text(v) or f"(Column {get_column_letter(i)})"
        n = seen.get(t.upper(), 0) + 1
        seen[t.upper()] = n
        labels.append(t if n == 1 else f"{t} ({n})")
    return labels


def auto_map(labels):
    """Best guess: NuORDER field -> index of the matching source column."""
    norm = [norm_header(l) for l in labels]
    mapping, used = {}, set()
    for field, syns in FIELD_SYNONYMS.items():           # pass 1: exact names (first synonym wins)
        for syn in syns:
            hit = next((i for i, nh in enumerate(norm) if i not in used and nh == syn), None)
            if hit is not None:
                mapping[field] = hit
                used.add(hit)
                break
    for field, syns in FIELD_SYNONYMS.items():           # pass 2: names that contain a keyword
        if field in mapping or field in NO_SUBSTRING:
            continue
        excl = FIELD_EXCLUDE.get(field, [])
        for i, nh in enumerate(norm):
            if i in used or not nh:
                continue
            words = f" {nh} "
            if any(f" {e} " in words for e in excl):
                continue
            if any(len(s_) >= 5 and f" {s_} " in words for s_ in syns):
                mapping[field] = i
                used.add(i)
                break
    return mapping


def guess_link_columns(labels):
    """Every column holding image links (e.g. JOOR 'Style Image URL_1' ... '_9'), swatches excluded, in order."""
    found = []
    for i, l in enumerate(labels):
        if "swatch" in l.lower() or re.search(r"\bid\b|filename", l, re.I):
            continue
        if re.search(r"image.*(url|link)|photo|picture|\bimg\b|\burl\b", l, re.I):
            m = re.search(r"(\d+)\D*$", l)
            found.append((int(m.group(1)) if m else 0, i))
    return [i for _n, i in sorted(found)]


def guess_link_column(labels):
    """Best column of image links: a '... Image URL_1' style column, never swatches."""
    def ok(l):
        return "swatch" not in l.lower()
    for pat in (r"image.*url", r"url", r"image.*link|photo|picture|img"):
        for i, l in enumerate(labels):
            if ok(l) and re.search(pat, l, re.I) and not re.search(r"\bid\b|filename", l, re.I):
                return i
    return None


SIZE_VALUE_RE = re.compile(r"^\W*D?\s*SIZE[\s_#.\-]*(\d{1,2})\s*$", re.I)


def detect_size_value_columns(labels):
    """'SIZE 1', 'Size 2' ... 'SIZE 12' (also typos like 'DSIZE 7', 'SIZE 8 ') -> column indexes in size order."""
    found = []
    for i, l in enumerate(labels):
        m = SIZE_VALUE_RE.match(re.sub(r"\s*\(\d+\)$", "", l))
        if m:
            found.append((int(m.group(1)), i))
    return [i for _n, i in sorted(found)]


def detect_size_columns(labels):
    """Columns whose header looks like a size (XS, S, M, 70, 75, ONE SIZE ...)."""
    idx = []
    for i, l in enumerate(labels):
        s = re.sub(r"\s*\(\d+\)$", "", l).strip().upper()
        if s in ALPHA_RANK or s in ONE_SIZE_TOKENS or re.fullmatch(r"\d{1,3}(\.5)?", s):
            idx.append(i)
    return idx


# ======================================================================
#  TAB 1 ENGINE - CLEANING
# ======================================================================

def extract_urls(text):
    if not text:
        return []
    pat = r"(?:https?://|www\.)[^\s\"'<>|;]+?(?=,?(?:https?://|www\.)|[\s\"'<>|;]|$)"
    urls = []
    for u in re.findall(pat, str(text)):
        u = u.rstrip(",.)")
        if u.lower().startswith("www."):
            u = "https://" + u
        urls.append(u)
    return urls


def cell_urls(cell, formula_cell=None):
    found = []
    try:
        if cell is not None and cell.hyperlink is not None:
            t = cell.hyperlink.target or cell.hyperlink.location
            if t:
                found += extract_urls(t)
    except Exception:
        pass
    if cell is not None and not is_blank(cell.value):
        found += extract_urls(cell.value)
    if formula_cell is not None and isinstance(formula_cell.value, str):
        found += extract_urls(formula_cell.value)
        try:
            if formula_cell.hyperlink is not None:
                found += extract_urls(formula_cell.hyperlink.target or "")
        except Exception:
            pass
    return list(dict.fromkeys(found))


def auto_fit(ws, min_w=12, max_w=60, extra=3):
    for col in ws.columns:
        longest = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(longest + extra, min_w), max_w)


def last_data_row(ws):
    """Real last row holding a value (some files claim 1,000,000 rows of empty formatting)."""
    try:
        return max((c.row for c in ws._cells.values() if c.value is not None), default=1)
    except Exception:
        return ws.max_row


# ----------------------------------------------------------------------
#  Pictures pasted into the Excel sheet
# ----------------------------------------------------------------------

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _zip_join(base_dir, target):
    import posixpath
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def embedded_pictures(path, sheet_name, read_bytes=True):
    """Pictures floating over the cells of one sheet.
    Returns ([(excel_row, excel_col, extension, bytes_or_None), ...], note)"""
    import posixpath
    import zipfile
    import xml.etree.ElementTree as ET
    if not path.lower().endswith((".xlsx", ".xlsm", ".xlsb")):
        return [], ""
    out, note = [], ""
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            if path.lower().endswith(".xlsb"):
                sheet_path = next((p for n, p, _v in xlsb_sheet_parts(z) if n == sheet_name), None)
                if sheet_path is None:
                    return [], ""
            else:
                wbx = ET.fromstring(z.read("xl/workbook.xml"))
                rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
                rid = None
                for sh in wbx.find(f"{{{_NS_MAIN}}}sheets"):
                    if sh.get("name") == sheet_name:
                        rid = sh.get(f"{{{_NS_REL}}}id")
                if rid is None:
                    return [], ""
                sheet_path = _zip_join("xl", rels[rid])
            rel_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
            if any(n.startswith("xl/richData/") for n in names):
                try:
                    out += _in_cell_pictures(z, names, sheet_path, read_bytes)
                except Exception:
                    note = "some pictures placed INSIDE cells could not be read"
            if rel_path not in names:
                out.sort(key=lambda t: (t[0], t[1]))
                return out, note
            for rel in ET.fromstring(z.read(rel_path)):
                if not rel.get("Type", "").endswith("/drawing"):
                    continue
                dpath = _zip_join(posixpath.dirname(sheet_path), rel.get("Target"))
                drel_path = posixpath.join(posixpath.dirname(dpath), "_rels", posixpath.basename(dpath) + ".rels")
                drels = {}
                if drel_path in names:
                    drels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read(drel_path))}
                for anchor in ET.fromstring(z.read(dpath)):
                    frm = anchor.find(f"{{{_NS_XDR}}}from")
                    pic = anchor.find(f"{{{_NS_XDR}}}pic")
                    if frm is None or pic is None:
                        continue
                    blip = pic.find(f".//{{{_NS_A}}}blip")
                    if blip is None:
                        continue
                    target = drels.get(blip.get(f"{{{_NS_REL}}}embed"))
                    if not target:
                        continue
                    media = _zip_join(posixpath.dirname(dpath), target)
                    if media not in names:
                        continue
                    row = int(frm.find(f"{{{_NS_XDR}}}row").text) + 1
                    col = int(frm.find(f"{{{_NS_XDR}}}col").text) + 1
                    ext = os.path.splitext(media)[1].lower() or ".png"
                    out.append((row, col, ext, z.read(media) if read_bytes else None))
    except Exception:
        return [], note
    out.sort(key=lambda t: (t[0], t[1]))
    return out, note


def _in_cell_pictures(z, names, sheet_path, read_bytes):
    """Pictures placed INSIDE cells (Excel 'Place in cell'). Returns [(row, col, ext, bytes)]."""
    import posixpath
    import struct
    import xml.etree.ElementTree as ET

    def strip(tag):
        return tag.rsplit("}", 1)[-1]

    # rich value -> image file
    rel_list = [el.get("{%s}id" % _NS_REL) for el in ET.fromstring(z.read("xl/richData/richValueRel.xml"))]
    rel_targets = {r.get("Id"): r.get("Target")
                   for r in ET.fromstring(z.read("xl/richData/_rels/richValueRel.xml.rels"))}
    structs = []
    if "xl/richData/rdrichvaluestructure.xml" in names:
        for s_ in ET.fromstring(z.read("xl/richData/rdrichvaluestructure.xml")):
            structs.append([k.get("n") for k in s_])
    rv_image = []
    for rv in ET.fromstring(z.read("xl/richData/rdrichvalue.xml")):
        vals = [v.text for v in rv if strip(v.tag) == "v"]
        keys = structs[int(rv.get("s", 0))] if structs else []
        pos = keys.index("_rvRel:LocalImageIdentifier") if "_rvRel:LocalImageIdentifier" in keys else 0
        media = None
        try:
            rid = rel_list[int(vals[pos])]
            media = _zip_join("xl/richData", rel_targets[rid])
        except Exception:
            pass
        rv_image.append(media)

    # value-metadata index (vm, 1-based) -> rich value index
    vm_to_rv = None
    if "xl/metadata.xml" in names:
        md = ET.fromstring(z.read("xl/metadata.xml"))
        fut = []
        for fm in md:
            if strip(fm.tag) == "futureMetadata" and fm.get("name") == "XLRICHVALUE":
                for bk in fm:
                    rvb = next((e for e in bk.iter() if strip(e.tag) == "rvb"), None)
                    fut.append(int(rvb.get("i")) if rvb is not None else None)
        vm_to_rv = []
        for vmd in md:
            if strip(vmd.tag) == "valueMetadata":
                for bk in vmd:
                    rc = next((e for e in bk if strip(e.tag) == "rc"), None)
                    v = int(rc.get("v")) if rc is not None else None
                    vm_to_rv.append(fut[v] if v is not None and v < len(fut) else None)

    def rv_for(vm):
        if vm_to_rv is not None and 0 < vm <= len(vm_to_rv):
            return vm_to_rv[vm - 1]
        return vm - 1                         # .xlsb: rich values are numbered in the same order

    cells = []
    if sheet_path.endswith(".bin"):
        row, pending = None, None
        for rt, p in _xlsb_records(z.read(sheet_path)):
            if rt == 0:
                row = struct.unpack_from("<I", p, 0)[0] + 1
            elif rt == 50 and len(p) >= 4:        # BrtCellMeta: the next cell holds a rich value
                pending = struct.unpack_from("<I", p, 0)[0]
            elif 1 <= rt <= 11 and pending is not None and row is not None:
                cells.append((row, struct.unpack_from("<I", p, 0)[0] + 1, pending))
                pending = None
    else:
        for _ev, el in ET.iterparse(__import__("io").BytesIO(z.read(sheet_path))):
            if strip(el.tag) == "c" and el.get("vm"):
                ref = el.get("r", "")
                m = re.match(r"([A-Z]+)(\d+)", ref)
                if m:
                    col = 0
                    for ch in m.group(1):
                        col = col * 26 + ord(ch) - 64
                    cells.append((int(m.group(2)), col, int(el.get("vm"))))
            el.clear()
    out = []
    for row, col, vm in cells:
        k = rv_for(vm)
        media = rv_image[k] if k is not None and 0 <= k < len(rv_image) else None
        if media and media in names:
            out.append((row, col, os.path.splitext(media)[1].lower() or ".png",
                        z.read(media) if read_bytes else None))
    return out


# ======================================================================
#  SHARED HELPERS FOR ALL ENGINES
# ======================================================================

# ----------------------------------------------------------------------
#  Image resizing (same as the XnConvert settings: shortest side 1024 px, Lanczos, always, no borders, JPG)
# ----------------------------------------------------------------------

RESIZE_DEFAULT_PX = 1024
JPG_QUALITY = 85
_NO_RESIZE_EXTS = {".emf", ".wmf", ".svg"}


def pillow_available():
    try:
        import PIL.Image  # noqa: F401
        return True
    except Exception:
        return False


def resize_to_jpg(src, target, size):
    """src = file path or bytes. Writes a JPG at `target` whose SHORTEST side is `size` pixels
    (proportions kept, enlarged or reduced, Lanczos, no borders/padding/cropping). Returns True if done."""
    import io
    from PIL import Image, ImageOps
    im = Image.open(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src)
    try:
        im.seek(0)
    except Exception:
        pass
    im = ImageOps.exif_transpose(im)
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))        # transparent parts become white (JPG has no transparency)
        bg.paste(im, mask=im.split()[-1])
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")
    w, h = im.size
    if w < 1 or h < 1:
        return False
    scale = float(size) / min(w, h)
    nw, nh = max(1, round(w * scale)), max(1, round(h * scale))
    if (nw, nh) != (w, h):
        im = im.resize((nw, nh), Image.LANCZOS)
    tmp = target + ".part"
    im.save(tmp, "JPEG", quality=JPG_QUALITY, optimize=True, subsampling=0 if JPG_QUALITY >= 90 else 2)
    os.replace(tmp, target)
    return True


def _resize_px(ctx):
    """Target size when resizing is on for this run (and Pillow is there), else None."""
    px = ctx.get("resize_px")
    if not px:
        return None
    if ctx.get("_pil") is None:
        ctx["_pil"] = pillow_available()
    return px if ctx["_pil"] else None


def _note_resize(ctx, ok):
    st = ctx.setdefault("resize_stats", {"done": 0, "failed": 0})
    st["done" if ok else "failed"] += 1


def resize_existing(path, ctx):
    """Resize a file that was just written (downloads). A non-JPG is replaced by a .jpg. Returns the new path."""
    px = _resize_px(ctx)
    ext = os.path.splitext(path)[1].lower()
    if not px or ext in _NO_RESIZE_EXTS:
        return path
    target = os.path.splitext(path)[0] + ".jpg"
    try:
        resize_to_jpg(path, target, px)
        if os.path.abspath(target) != os.path.abspath(path):
            try:
                os.remove(path)
            except OSError:
                pass
        _note_resize(ctx, True)
        return target
    except Exception:
        _note_resize(ctx, False)
        return path


def resize_log_line(ctx):
    if not ctx.get("resize_px"):
        return None
    if ctx.get("_pil") is False:
        return "NOTE: image resizing is on but the Pillow library is missing - images were saved at their original size."
    st = ctx.get("resize_stats")
    if not st:
        return None
    return (f"Resized: {st['done']} image(s) to {ctx['resize_px']} px on the shortest side (JPG)"
            + (f"; {st['failed']} could not be read and were kept as they were" if st["failed"] else "") + ".")


def save_picture(ctx, name, ext, data):
    """Saves one picture into the shared Images folder. Same picture twice for one product is skipped,
    a different picture for the same product gets _2, _3 ...  Returns the file name or None."""
    img_dir = ctx["img_dir"]
    os.makedirs(img_dir, exist_ok=True)
    key = safe_filename(name).upper()
    h = hashlib.md5(data).hexdigest()
    seen = ctx["pic_hashes"].setdefault(key, set())
    if h in seen:
        ctx["pic_stats"]["repeat"] += 1
        return None
    seen.add(h)
    n = ctx["pic_count"].get(key, 0) + 1
    ctx["pic_count"][key] = n
    px = _resize_px(ctx) if ext.lower() not in _NO_RESIZE_EXTS else None
    fname = f"{safe_filename(name)}_{n:02d}{'.jpg' if px else ext}"
    target = os.path.join(img_dir, fname)
    if ctx["skip_existing"] and os.path.exists(target):
        ctx["pic_stats"]["existing"] += 1
        return fname
    if px:
        try:
            resize_to_jpg(data, target, px)
            _note_resize(ctx, True)
            ctx["pic_stats"]["saved"] += 1
            return fname
        except Exception:
            _note_resize(ctx, False)
            fname = f"{safe_filename(name)}_{n:02d}{ext}"
            target = os.path.join(img_dir, fname)
    with open(target, "wb") as f:
        f.write(data)
    ctx["pic_stats"]["saved"] += 1
    return fname


def extract_row_pictures(path, sheet, row_to_name, ctx, L, only_columns=None):
    """Pictures pasted in the sheet -> Images folder, named after the product of the row they sit on."""
    pics, note = embedded_pictures(path, sheet, True)
    if only_columns:
        pics = [p_ for p_ in pics if p_[1] in only_columns]
    if note:
        L("NOTE: " + note)
    if not pics:
        L("Pictures: none found in this sheet.")
        return 0
    before = ctx["pic_stats"]["saved"] + ctx["pic_stats"]["existing"]
    outside = 0
    for row, _col, ext, data in pics:
        name = row_to_name.get(row)
        if not name:
            outside += 1
            continue
        save_picture(ctx, name, ext, data)
    got = ctx["pic_stats"]["saved"] + ctx["pic_stats"]["existing"] - before
    L(f"Pictures: {len(pics)} found, {got} saved to the Images folder"
      + (f", {outside} outside the product rows (logos etc.) ignored" if outside else "") + ".")
    if any(ext in (".emf", ".wmf") for _r, _c, ext, _d in pics):
        L("NOTE: some pictures are .emf/.wmf files (Windows-only formats).")
    return got


def season_code_from_text(text):
    """Finds a season written inside any text (file name, title row): SS27, SU27, FALL 2027 ... as written."""
    t = str(text or "").upper().replace("_", " ")
    m = re.search(r"(?<![A-Z0-9])(SUMMER|SPRING|FALL|AUTUMN|PRE-?FALL|WINTER|RESORT|CRUISE|HOLIDAY)[\s\-]*(\d{4}|\d{2})(?![0-9])", t)
    if m:
        return re.sub(r"[\s\-]+", " ", m.group(0)).strip()
    m = re.search(r"(?<![A-Z0-9])(SS|FW|AW|SU|SP|PF|FA|WI|HO|RE|F|S)[\s\-]?(\d{2})(?![0-9])", t)
    return (m.group(1) + m.group(2)) if m else ""


def apply_season_map(season, mapping):
    """{'SS': 'SPRING'}: 'SS27' -> 'SPRING 2027'. Anything else is returned unchanged."""
    if not season or not mapping:
        return season
    m = re.fullmatch(r"([A-Z]+)\s*(\d{2}|\d{4})", season.upper().strip())
    if m and m.group(1) in mapping:
        y = m.group(2)
        return f"{mapping[m.group(1)]} {2000 + int(y) if len(y) == 2 else y}"
    return season


def _basic_size_list(text, even_all=False):
    out = []
    for tok in re.split(r"[,;\n]+", str(text)):
        tok = normalize_size(tok)
        if tok:
            rng = expand_size_range(tok, even_all)
            out.extend(rng if rng else [tok])
    return out


def parse_size_scale(text, even_all=False):
    """Size notes -> sizes.
    '5-16 (HALF SIZES OFFERED BUT NO 5.5, 11.5,12.5,13,5)' -> 5, 6, 6.5 ... 16 without the excluded sizes
    '46 - 60  /  S, R, L' -> 46S, 46R, 46L, 48S ... (lengths after the '/')
    '48 - 58' with even_all -> 48, 50 ... 58"""
    t = clean_text(text).upper()
    if not t:
        return []
    if normalize_size(t) == "OS":
        return ["OS"]
    lengths = []
    lm = re.match(r"^(.*?\d)\s*/\s*([A-Z]{1,2}(?:\s*,\s*[A-Z]{1,2})*)\s*$", t)
    if lm:
        t, lengths = lm.group(1), [x.strip() for x in lm.group(2).split(",") if x.strip()]
    m = re.match(r"\s*(\d+(?:\.\d+)?)\s*(?:-|\u2013|\u2014|TO)\s*(\d+(?:\.\d+)?)", t)
    if not m:
        base = _basic_size_list(t, even_all)
    else:
        a_txt, b_txt = m.group(1), m.group(2)
        a, b = float(a_txt), float(b_txt)
        rest = t[m.end():]
        half = "HALF" in rest or a % 1 or b % 1
        excl = set()
        em = re.search(r"\b(?:BUT\s+NO|EXCEPT|EXCLUDING|NO)\b(.*)", rest)
        if em:
            txt = re.sub(r"(?<![\d.])(\d+),5(?![\d.])", r"\1.5", em.group(1))      # 13,5 -> 13.5
            excl = {float(x) for x in re.findall(r"\d+(?:\.\d+)?", txt)}
        if half:
            step = 0.5
        elif a % 2 == 0 and b % 2 == 0 and (b <= 20 or even_all):
            step = 2.0
        else:
            step = 1.0
        base = [a_txt] if re.fullmatch(r"0{2,}", a_txt) else []
        n = a
        while n <= b + 1e-9 and len(base) < 80:
            if n not in excl:
                base.append(str(int(n)) if n == int(n) else str(n))
            n += step
    if lengths:
        return [f"{x}{L}" for x in base for L in lengths]
    return base


def pad_color_code(v):
    """Color codes always have at least 3 digits: 12 -> 012, 5 -> 005 (letters are left alone)."""
    if v is None or isinstance(v, bool):
        return v
    s_ = clean_code(v)
    return s_.zfill(3) if s_.isdigit() and len(s_) < 3 else s_


COLOR_CODE_HEADERS = {"color code", "nrf color code", "colour code"}


class Result:
    """What one source file turned into."""

    def __init__(self, path, label):
        self.path = path
        self.label = label
        self.headers = []          # output headers in order, "SIZES" marks where SIZE 1..N go
        self.rows = []             # [{"values": {header: value}, "sizes": [...]}]
        self.warnings = []
        self.numeric = set()
        self.brand = ""


def validate_rows(res, L):
    """The six key NuORDER fields + duplicates, reported in the log."""
    keys = ["NAME", "STYLE NUMBER", "SEASON", "COLOR", "SIZE 1"]
    if "BRAND_ID" in res.headers:
        keys.append("BRAND_ID")
    for k in keys:
        if k == "SIZE 1":
            n = sum(1 for r in res.rows if not r["sizes"])
        elif k in res.headers:
            n = sum(1 for r in res.rows if r["values"].get(k) in (None, ""))
        else:
            continue
        if n:
            res.warnings.append(f"{k}: empty on {n} product(s)")
    if "BRAND_ID" in res.headers:
        ids = [r["values"].get("BRAND_ID") for r in res.rows if r["values"].get("BRAND_ID")]
        if len(ids) != len(set(ids)):
            res.warnings.append(f"BRAND_ID: {len(ids) - len(set(ids))} duplicate value(s)")
    for w_ in res.warnings:
        L("   CHECK: " + w_)


# ======================================================================
#  ENGINE 1 - GENERAL RULES (any layout, columns chosen in the window)
# ======================================================================

def brand_from_sheet(ws, labels, header_row, last_row):
    """First value of a Brand / Designer / Business column (for the History tab)."""
    want = ("brand", "brand name", "designer", "designer name", "business", "label name")
    for i, l in enumerate(labels):
        if norm_header(l) in want:
            for cells in ws.iter_rows(min_row=header_row + 1, max_row=min(last_row, header_row + 50), max_col=i + 1):
                v = cells[i].value if i < len(cells) else None
                if not is_blank(v) and not isinstance(v, (int, float)):
                    return clean_text(v)
    return ""


def generic_result(cfg, ctx, L, progress, stop):
    path = cfg["path"]
    ws = open_sheet(path, cfg["sheet"], data_only=True)
    ws_f = None
    if cfg["download"] and path.lower().endswith((".xlsx", ".xlsm")):
        ws_f = open_sheet(path, cfg["sheet"], data_only=False)
    last_row = last_data_row(ws)
    L(f"Sheet '{getattr(ws, 'title', cfg['sheet'])}', header on row {cfg['header_row']}, data up to row {last_row}.")

    labels = cfg["labels"]
    ncols = len(labels)
    M = cfg["mapping"]
    merged = merged_value_map(ws)
    default_year = cfg["default_year"]
    base_cur = cfg["price_currency"] if cfg["price_currency"] != "AUTO" else infer_file_currency(M)

    def getv(cells, r, key):
        idx = M.get(key)
        if idx is None or idx >= len(cells):
            return None
        v = cells[idx].value
        if v is None:
            v = merged.get((r, idx + 1))
        return v

    groups = OrderedDict()
    skipped = 0
    total = max(last_row - cfg["header_row"], 1)
    for n, cells in enumerate(ws.iter_rows(min_row=cfg["header_row"] + 1, max_row=last_row, max_col=ncols)):
        r = cfg["header_row"] + 1 + n
        if n % 200 == 0:
            progress(int(30 * n / total))
            if stop.is_set():
                raise UserError("Stopped by user.")
        if all(is_blank(c.value) for c in cells):
            continue
        style = clean_code(getv(cells, r, "STYLE NUMBER"))
        color = clean_text(getv(cells, r, "COLOR"))
        if not style:
            skipped += 1
            continue
        key = (style.upper(), color.upper())
        rec = groups.get(key)
        if rec is None:
            rec = groups[key] = dict(
                style=style, name="", color=color, color_code="", avail_from=None, avail_until=None,
                fabric="", coo="", dept="", linesheet="", separator="", season_txt="", delivery_month="",
                prices={}, gprices={},
                sizes=[], links=[], rows=[], has_pics=False)
        rec["rows"].append(r)

        def fill(attr, val):
            if val not in (None, "") and rec[attr] in (None, ""):
                rec[attr] = val

        fill("name", clean_text(getv(cells, r, "NAME")))
        fill("color_code", clean_code(getv(cells, r, "COLOR CODE")))
        fill("fabric", clean_fabric(getv(cells, r, "FABRIC DESCRIPTION")))
        fill("coo", clean_country(getv(cells, r, "US COUNTRY OF SOURCING")))
        fill("dept", clean_text(getv(cells, r, "DEPARTMENT")))
        fill("delivery_month", clean_text(getv(cells, r, "DELIVERY MONTH")).upper())
        fill("linesheet", linesheet_label(getv(cells, r, "LINESHEET NAME")))
        sep = clean_text(getv(cells, r, "LINESHEET SEPARATOR"))
        if sep and rec["separator"] and sep != rec["separator"]:
            rec["sep_conflict"] = True
        fill("separator", sep)

        sv = clean_text(getv(cells, r, "SEASON"))
        yv = clean_code(getv(cells, r, "SEASON YEAR"))
        if sv and yv and yv not in sv:
            sv = f"{sv} {yv}"
            rec["season_joor"] = True          # JOOR: Season Name + Season Year -> PRE-FALL 2027
        fill("season_txt", sv)

        d_from = parse_dates(getv(cells, r, "AVAILABLE FROM"), default_year)
        d_until = parse_dates(getv(cells, r, "AVAILABLE UNTIL"), default_year)
        d_win = parse_dates(getv(cells, r, "DELIVERY"), default_year)
        if len(d_from) >= 2 and not d_until:       # a whole window typed in "Available From"
            d_win, d_from = d_from, []
        if d_win:
            if len(d_win) >= 2 and d_win[1] < d_win[0] and d_win[1].year == d_win[0].year:
                d_win[1] = d_win[1].replace(year=d_win[1].year + 1)
            if not d_from:
                d_from = d_win[:1]
            if not d_until and len(d_win) >= 2:
                d_until = d_win[1:2]
        fill("avail_from", d_from[0] if d_from else None)
        fill("avail_until", d_until[0] if d_until else None)

        for pf in PRICE_ORDER:
            val = parse_number(getv(cells, r, pf))
            if val and pf not in rec["prices"]:
                rec["prices"][pf] = val
        cur_w = (currency_code(getv(cells, r, "PRICE CURRENCY"))
                 or currency_code(getv(cells, r, "PRICE TYPE LABEL")) or base_cur)
        cur_r = currency_code(getv(cells, r, "RETAIL CURRENCY")) or cur_w
        w = parse_number(getv(cells, r, "WHOLESALE PRICE"))
        rt = parse_number(getv(cells, r, "RETAIL PRICE"))
        if w:
            rec["gprices"].setdefault("WHOLESALE " + cur_w, w)
        if rt:
            rec["gprices"].setdefault("RETAIL " + cur_r, rt)

        if cfg["size_mode"] == "matrix":
            for idx in cfg["size_cols"]:
                if idx >= len(cells):
                    continue
                c = cells[idx]
                if cfg["black_unavailable"] and is_black_fill(c):
                    continue
                if cfg["require_value"] and is_blank(c.value):
                    continue
                lab = normalize_size(re.sub(r"\s*\(\d+\)$", "", labels[idx]))
                if lab:
                    rec["sizes"].append(lab)
        else:
            row_sizes = []
            for idx in cfg.get("size_value_cols", []):
                if idx < len(cells):
                    for s_ in split_sizes(cells[idx].value):
                        if s_ not in row_sizes:
                            row_sizes.append(s_)
            if not row_sizes:                      # 'All sizes' column: 23-34, XXS-XL, lists, notes
                row_sizes = split_sizes(getv(cells, r, "SIZE"))
            rec["sizes"] += row_sizes

        if cfg["download"]:
            for lc in cfg.get("link_cols", []):
                if lc >= len(cells):
                    continue
                fcell = ws_f.cell(r, lc + 1) if ws_f is not None else None
                for u in cell_urls(cells[lc], fcell):
                    if u not in rec["links"]:
                        rec["links"].append(u)

    L(f"Found {len(groups)} unique Style x Color products ({skipped} rows without a style number ignored).")
    if not groups:
        raise UserError("No products were found. Check the Header row number and the Style Number column.")

    override = linesheet_label(cfg.get("linesheet_override"))
    default_ls = sheet_linesheet_name(path, cfg)
    res = Result(path, cfg.get("label") or "Unknown (general rules)")
    res.brand = brand_from_sheet(ws, labels, cfg["header_row"], last_row)
    records = []
    usd_filled = 0
    conflicts = []
    for rec in groups.values():
        rec["name"] = strip_style_from_name(rec["name"], rec["style"])
        sizes = sort_sizes(rec["sizes"])
        rec["sizes"] = ["OS"] if sizes and set(sizes) == {"OS"} else sizes
        if rec.get("season_joor"):
            rec["season"] = season_full(rec["season_txt"], default_year)
        else:
            rec["season"] = rec["season_txt"] or cfg["season"] or season_code_from_text(os.path.basename(path))
        rec["linesheet"] = override or rec["linesheet"] or default_ls
        if rec.get("sep_conflict"):
            conflicts.append(rec["style"])
        second = rec["color"]
        if cfg["brand_source"] == "code" and rec["color_code"]:
            second = rec["color_code"]
        rec["brand_id"] = f"{rec['style']}_{second}" if second else rec["style"]
        if cfg["brand_upper"]:
            rec["brand_id"] = rec["brand_id"].upper()
        p = dict(rec["gprices"])
        p.update(rec["prices"])
        rec["prices"] = p
        if "WHOLESALE USD" in p:
            rec["currency"] = "USD"
        else:
            others = [k.split(" ", 1)[1] for k in p if k.startswith("WHOLESALE ")]
            rec["currency"] = others[0] if others else "USD"
        for k in ("WHOLESALE USD", "RETAIL USD"):
            if not p.get(k):
                p[k] = 0.01
                usd_filled += 1
        records.append(rec)
    if usd_filled:
        L(f"{usd_filled} USD price cells hold the 0.01 placeholder (no USD price, or a price of 0).")
    if conflicts:
        L(f"NOTE: LINESHEET SEPARATOR differs between the rows of {len(conflicts)} product(s); "
          f"the first value was kept: {', '.join(conflicts[:10])}")

    present = {k for r in records for k in r["prices"]}
    price_cols = ["WHOLESALE USD", "RETAIL USD"] + [k for k in PRICE_ORDER[2:] if k in present]
    price_cols += sorted(k for k in present if k not in PRICE_ORDER)
    headers = ["LINESHEET NAME"]
    if M.get("LINESHEET SEPARATOR") is not None:
        headers.append("LINESHEET SEPARATOR")
    headers += ["STYLE NUMBER", "NAME"]
    for key in ("FABRIC DESCRIPTION", "US COUNTRY OF SOURCING", "DEPARTMENT"):
        if M.get(key) is not None:
            headers.append(key)
    headers += ["US VENDOR CURRENCY", "COLOR", "COLOR CODE", "AVAILABLE FROM", "AVAILABLE UNTIL"]
    if M.get("DELIVERY MONTH") is not None:
        headers.append("DELIVERY MONTH")
    headers += ["SIZES"] + price_cols + ["SEASON", "BRAND_ID"]
    res.headers = headers
    res.numeric = set(price_cols)
    for rec in records:
        vals = {
            "LINESHEET NAME": rec["linesheet"], "LINESHEET SEPARATOR": rec["separator"],
            "STYLE NUMBER": rec["style"], "NAME": rec["name"],
            "FABRIC DESCRIPTION": rec["fabric"], "US COUNTRY OF SOURCING": rec["coo"],
            "DEPARTMENT": rec["dept"], "US VENDOR CURRENCY": rec["currency"], "COLOR": rec["color"],
            "DELIVERY MONTH": rec["delivery_month"],
            "COLOR CODE": rec["color_code"], "SEASON": rec["season"], "BRAND_ID": rec["brand_id"],
            "AVAILABLE FROM": format_date(rec["avail_from"], cfg["date_format"]),
            "AVAILABLE UNTIL": format_date(rec["avail_until"], cfg["date_format"]),
        }
        vals.update(rec["prices"])
        res.rows.append({"values": vals, "sizes": rec["sizes"]})

    if cfg.get("extract_pictures"):
        row_to_name = {r: rec["brand_id"] for rec in records for r in rec["rows"]}
        if extract_row_pictures(path, cfg["sheet"], row_to_name, ctx, L):
            for rec in records:
                rec["has_pics"] = True
    if cfg["download"] and cfg.get("link_cols"):
        download_images(records, ctx, L, stop)
    validate_rows(res, L)
    return res


# ======================================================================
#  ENGINE 2 - BRAND PROFILES
# ======================================================================
#
#  A brand profile is a small JSON file in the "Brand Profiles" folder next to the program.
#  It describes one brand's file layout and rules. Built-in profiles are (re)written into
#  that folder whenever their version is newer than the file there.
#
#  Profile keys:
#    name, label, version, description
#    detect          : {"sheet_names": [...], "headers_all": [...], "headers_any": [...]}
#    sheet           : preferred sheet name
#    header_contains : header names that must be on the header row (used to find it)
#    row_key         : source columns that identify one product (rows sharing it are merged);
#                      "row_key_alt" is tried when those columns are missing
#    columns         : output columns in order. Each one is a dict with "header" and one of
#                        "from": "SRC" or ["SRC1", "SRC2"]   first non-empty source column
#                        "join": ["A", "B"], "sep": "_"     several columns joined (used when "from" is empty)
#                        "value": "text"                     fixed text ({linesheet}, {season} allowed)
#                        "date": "from" | "until"            ship window (see "dates")
#                        header "SIZES"                      where SIZE 1 ... SIZE N go
#                      options: "number": true, "usd": true (0 / blank -> 0.01), "optional": true
#                      (column dropped when empty everywhere), "upper": true,
#                      "clean": "country" | "fabric" | "code"
#    sizes           : mode "matrix": {"columns_matching", "label", "available_values", "suffix_from"}
#                      mode "scale_text": {"from": column holding text like '5-16 (HALF SIZES ... BUT NO 5.5)'}
#                      mode "list": {"from": [...] or "columns_matching": regex, "sort": true}
#    dates           : {"window_from": col, "until_from": col, "title_regex": regex on the rows above the header}
#    season          : {"from": "file_name" | "title" | column, "map": {"SS": "SPRING"}}
#    pictures        : {"extract": true, "name_from": OUTPUT header used to name the picture files}

_OPT_EXTRAS = [("COLOR FAMILY", "Color Family"), ("CATEGORY", "Category"), ("SUBCATEGORY", "Subcategory"),
               ("DESCRIPTION", "Description"), ("GENDER", "Gender"), ("COLLECTION NAME", "Collection Name"),
               ("CARE INSTRUCTIONS", "Care Instructions"), ("AGE", "Age")]
_OPT_PRICES = [("WHOLESALE CAD", "Wholesale CAD"), ("RETAIL CAD", "Retail CAD"), ("WHOLESALE AUD", "Wholesale AUD"),
               ("RETAIL AUD", "Retail AUD"), ("RETAIL NZD", "Retail NZD"), ("WHOLESALE EUR", "Wholesale EUR"),
               ("RETAIL EUR", "Retail EUR"), ("WHOLESALE GBP", "Wholesale GBP"), ("RETAIL GBP", "Retail GBP")]

_USD_W = ["WHOLESALE USD", "WHOLESALE", "WHOLESALE PRICE", "WSP", "WSP USD", "COST", "WHSL USD", "WHSL"]
_USD_R = ["RETAIL USD", "RETAIL", "RETAIL PRICE", "MSRP", "M.S.R.P.", "US MSRP", "RRP"]
_SEASON_MAP = {"SS": "SPRING", "FW": "FALL", "AW": "FALL"}


def _armani(name, label, detect, key, key_alt, columns, sizes, linesheet_from=("LINESHEET NAME",), photo_only=None):
    prof = {
        "name": name, "label": label, "brand": "Armani", "version": 1,
        "description": f"{label} linesheets.",
        "detect": detect, "row_key": key, "row_key_alt": key_alt,
        "linesheet_from": list(linesheet_from),
        "columns": [{"header": "LINESHEET NAME", "value": "{linesheet}"}]
        + [c for c in columns if c["header"] != "BRAND_ID"] + [
            {"header": "WHOLESALE USD", "from": _USD_W, "number": True, "usd": True},
            {"header": "RETAIL USD", "from": _USD_R, "number": True, "usd": True},
            {"header": "SIZES"},
        ] + ([{"header": "UPCS"}] if sizes.get("upc_from") else []) + [
            {"header": "SEASON", "value": "{season}"},
        ] + [c for c in columns if c["header"] == "BRAND_ID"],
        "sizes": sizes,
        "season": {"from": ["SEASON"], "map": _SEASON_MAP},
        "pictures": {"extract": True, "name_join": ["MODEL", "FABRIC CODE", "COLOR CODE"], "sep": "_"},
    }
    if photo_only:
        prof["pictures"]["only_columns"] = photo_only
    return prof


BUILTIN_PROFILES = {
    "Ferragamo": {
        "name": "Ferragamo", "label": "Ferragamo", "brand": "Ferragamo", "version": 3,
        "description": "Ferragamo linesheets: NOE Template (English headers) or Italian headers (Colori, Ves, Cll.Nas. ...).",
        "detect": {"sheet_names": ["NOE Template"], "any_of": [
            {"headers_all": ["SKU", "MOD", "PRICE_MOD_SINGLE", "PRICE_MOD_SINGLE_RETAIL"]},
            {"headers_all": ["SKU", "Ves", "Colori"]},
            {"headers_all": ["SKU", "Cll.Nas."]},
        ]},
        "sheet": "NOE Template",
        "row_key": ["SKU"],
        "columns": [
            {"header": "LINESHEET NAME", "value": "{linesheet}"},
            {"header": "CATEGORY", "from": ["T_CAT", "CAT"], "optional": True},
            {"header": "FAMILY NAME", "from": ["Linea", "LINE"], "optional": True},
            {"header": "CUSTOMER STYLE NUMBER", "from": ["Style Number", "MOD"], "optional": True},
            {"header": "NAME", "from": ["Nome", "Style Name"]},
            {"header": "COLOR", "from": ["Colori", "COLOR"]},
            {"header": "STYLE NUMBER", "from": "SKU"},
            {"header": "COLOR CODE", "from": ["Ves"]},
            {"header": "DESCRIPTION", "from": ["DESCR."], "optional": True},
            {"header": "FABRIC DESCRIPTION", "from": ["Materiali", "Composizione", "MATERIAL", "MATERIAL COMPOSITION"]},
            {"header": "DELIVERY MONTH", "from": ["Delivery Bucket", "Deliv", "DELIVERY"], "optional": True},
            {"header": "GARMENT MEASUREMENTS", "from": ["Dimensions", "DIMENSION"], "optional": True},
            {"header": "GENDER", "from": ["Sex"], "optional": True},
            {"header": "GROUP", "from": ["Gruppo"], "optional": True},
            {"header": "TYPOLOGY", "from": ["Tipology", "Tipologia"], "optional": True},
            {"header": "HEEL HIGH", "from": ["Alt.Tac", "HEEL"], "optional": True},
            {"header": "HARDWARE COLOR", "from": ["Finitura", "Finish"], "optional": True},
            {"header": "OCCASION", "from": ["Occasione"], "optional": True},
            {"header": "WHOLESALE USD", "from": ["PRICE_MOD_SINGLE"] + _USD_W, "number": True, "usd": True},
            {"header": "RETAIL USD", "from": ["PRICE_MOD_SINGLE_RETAIL"] + _USD_R, "number": True, "usd": True},
            {"header": "SIZES"},
            {"header": "SEASON", "value": "{season}"},
        ],
        "sizes": {"mode": "matrix", "columns_matching": r"^\s*\d+(\.\d+)?\s*-\s*\d+\s*$",
                  "label": r"^\s*(\d+(?:\.\d+)?)", "available_values": ["0"], "suffix_from": "SIZE_2"},
        "season": {"from": ["Cll.Nas."]},
        "pictures": {"extract": True, "name_from": "STYLE NUMBER"},
    },
    "Gucci": {
        "name": "Gucci", "label": "Gucci", "brand": "Gucci", "version": 2,
        "description": "Gucci master linesheets (title row with SHIP WINDOW, SMC, SIZE SCALE columns).",
        "detect": {"any_of": [{"headers_all": ["UPDATED STYLE NUMBER", "SMC", "SIZE SCALE"]},
                              {"headers_all": ["SMC", "SIZE SCALE", "STYLE", "MATERIAL"]}]},
        "row_key": ["SMC"],
        "row_key_alt": ["UPDATED STYLE NUMBER"],
        "columns": [
            {"header": "LINESHEET NAME", "value": "{linesheet}"},
            {"header": "STYLE NUMBER", "join": ["STYLE", "MATERIAL", ["COLOUR", "COLOR"]], "sep": "_"},
            {"header": "UPDATED STYLE NUMBER", "from": "UPDATED STYLE NUMBER", "optional": True},
            {"header": "Style", "from": "STYLE"},
            {"header": "Material", "from": "MATERIAL"},
            {"header": "Color code", "from": ["COLOUR", "COLOR"]},
            {"header": "NAME", "from": "LINE"},
            {"header": "DESCRIPTION", "from": "STYLE DESCRIPTION"},
            {"header": "MATERIAL COMPOSITION DETAILS", "from": "MATERIAL DESCRIPTION"},
            {"header": "FABRIC DESCRIPTION", "from": "MATERIAL DESCRIPTION"},
            {"header": "COLOR", "from": ["COLOUR DESCRIPTION", "COLOR DESCRIPTION"]},
            {"header": "CATEGORY", "from": "FUNCTION"},
            {"header": "DIVISION", "from": "CO", "optional": True},
            {"header": "SEASONALITY", "from": ["FASHION, SSNL, OR FUTURE CO", "SEASONALITY"], "optional": True},
            {"header": "AVAILABLE FROM", "date": "from"},
            {"header": "AVAILABLE UNTIL", "date": "until"},
            {"header": "DELIVERY MONTH", "from": ["SUB DELIVERY", "TARGET DELIVERY"], "optional": True},
            {"header": "WHOLESALE USD", "from": "WHOLESALE PRICE", "number": True, "usd": True},
            {"header": "RETAIL USD", "from": ["US MSRP", "RETAIL PRICE"], "number": True, "usd": True},
            {"header": "SIZES"},
            {"header": "BRAND_ID", "join": ["STYLE", "MATERIAL", ["COLOUR", "COLOR"]], "sep": "_"},
            {"header": "SEASON", "value": "{season}"},
        ],
        "sizes": {"mode": "scale_text", "from": "SIZE SCALE"},
        "dates": {"title_regex": r"SHIP\s*WINDOW\s*:?\s*(.+)"},
        "season": {"from": "title", "map": _SEASON_MAP},
        "pictures": {"extract": True, "name_from": "BRAND_ID"},
    },
    "MaxMara": {
        "name": "MaxMara", "label": "MaxMara", "brand": "MaxMara", "version": 2,
        "description": "MaxMara bulk linesheets (one sheet per collection: MM 1, MM 2 ...), set up like the brand's NuORDER "
                       "portal data. Add the brand's PDF catalogs to the file list to get the sizes, fabric composition and pictures.",
        "detect": {"headers_all": ["Exit Collection Description", "Ship Window", "VPN", "Model", "Class Description", "Variant"]},
        "row_key": ["VPN"],
        "linesheet_pattern": "{season} - {sheet}",
        "columns": [
            {"header": "LINESHEET NAME", "value": "{linesheet}"},
            {"header": "NAME", "from": "Model"},
            {"header": "STYLE NUMBER", "from": "VPN"},
            {"header": "COLOR", "from": "Color Description"},
            {"header": "COLOR CODE", "from": "Variant"},
            {"header": "DESCRIPTION", "from": "Class Description"},
            {"header": "FABRIC DESCRIPTION", "optional": True},
            {"header": "AVAILABLE FROM", "date": "from"},
            {"header": "AVAILABLE UNTIL", "date": "until"},
            {"header": "WHOLESALE USD", "from": ["Cost"], "number": True, "usd": True},
            {"header": "RETAIL USD", "from": ["Retail"], "number": True, "usd": True},
            {"header": "SIZES"},
            {"header": "SEASON", "value": "{season}"},
            {"header": "BRAND_ID", "join_out": ["STYLE NUMBER", "COLOR"], "sep": "_"},
        ],
        "sizes": {"mode": "list", "from": []},
        "dates": {"window_from": ["SHIP WINDOW"], "anchor_year": "season"},
        "season": {"from": "file_name"},
        "pdf": "maxmara",
        "pictures": {"extract": False, "name_from": "BRAND_ID"},
    },
    "NuORDER Default Data": {
        "name": "NuORDER Default Data", "label": "NuORDER Default Data", "version": 2,
        "description": "Files already in NuORDER's product data template (one or several rows per product).",
        "detect": {"headers_all": ["Season", "Color", "Size 1", "Name"], "headers_any": ["Brand ID", "Style Number"]},
        "sheet": "PRODUCT DATA",
        "row_key": ["Style Number", "Color"],
        "row_key_alt": ["Brand ID"],
        "columns": [
            {"header": "LINESHEET NAME", "value": "{linesheet}"},
            {"header": "STYLE NUMBER", "from": "Style Number"},
            {"header": "NAME", "from": "Name"},
            {"header": "FABRIC DESCRIPTION", "from": "Fabric Description", "optional": True},
            {"header": "US COUNTRY OF SOURCING", "from": "US Country of Sourcing", "optional": True},
            {"header": "DEPARTMENT", "from": "Department", "optional": True},
            {"header": "US VENDOR CURRENCY", "from": "US Vendor Currency", "optional": True},
            {"header": "COLOR", "from": "Color"},
            {"header": "COLOR CODE", "from": "Color Code"},
            {"header": "NRF COLOR CODE", "from": "NRF Color Code", "optional": True},
            {"header": "NRF COLOR DESCRIPTION", "from": "NRF Color Description", "optional": True},
        ] + [{"header": h, "from": s, "optional": True} for h, s in _OPT_EXTRAS] + [
            {"header": "AVAILABLE FROM", "date": "from"},
            {"header": "AVAILABLE UNTIL", "date": "until"},
            {"header": "DELIVERY MONTH", "from": "Delivery Month", "optional": True},
            {"header": "WHOLESALE USD", "from": "Wholesale USD", "number": True, "usd": True},
            {"header": "RETAIL USD", "from": "Retail USD", "number": True, "usd": True},
        ] + [{"header": h, "from": s, "number": True, "optional": True} for h, s in _OPT_PRICES] + [
            {"header": "SIZES"},
            {"header": "SEASON", "from": "Season"},
            {"header": "BRAND_ID", "from": "Brand ID", "join": ["Style Number", "Color"], "sep": "_", "upper": True},
        ],
        "sizes": {"mode": "list", "columns_matching": r"(?i)^\s*d?size[\s_]*\d+\s*$", "sort": True},
        "dates": {"window_from": "Available From", "until_from": "Available Until"},
        "linesheet_from": ["Linesheet Name"],
        "season": {"from": ["Season"]},
        "pictures": {"extract": True, "name_from": "BRAND_ID"},
    },
    "Velvet": {
        "name": "Velvet", "label": "Velvet", "brand": "Velvet", "version": 2,
        "description": "Velvet linesheets (Style Material Color / Product Name / Sizes columns).",
        "detect": {"headers_all": ["Style Material Color", "Product Name", "Color Name", "Sizes"]},
        "row_key": ["Style Material Color"],
        "columns": [
            {"header": "LINESHEET NAME", "value": "{linesheet}"},
            {"header": "STYLE NUMBER", "from": "Style Material Color"},
            {"header": "NAME", "from": "Product Name"},
            {"header": "FABRIC DESCRIPTION", "from": "Material Name"},
            {"header": "COLOR", "from": "Color Name"},
            {"header": "RETAIL USD", "from": "Retail Price", "number": True, "usd": True},
            {"header": "WHOLESALE USD", "from": "Wholesale Price", "number": True, "usd": True},
            {"header": "SIZES"},
            {"header": "SEASON", "value": "{season}"},
            {"header": "BRAND_ID", "join": ["Style Material Color", "Color Name"], "sep": "_"},
        ],
        "sizes": {"mode": "list", "from": ["Sizes"], "sort": True},
        "season": {"from": "file_name", "map": _SEASON_MAP},
        "pictures": {"extract": True, "name_from": "STYLE NUMBER"},
    },
    "Armani - Giorgio Armani Men": _armani(
        "Armani - Giorgio Armani Men", "Armani - Giorgio Armani Men",
        {"headers_all": ["MODEL", "FABRIC", "COLOR DESC.", "FABRIC COMP"]},
        ["STYLE NUMBER"], ["MODEL", "FABRIC", "COLOR"],
        [
            {"header": "CATEGORY", "from": "CATEGORY"},
            {"header": "NAME", "from": "PRODUCT"},
            {"header": "STYLE NUMBER", "from": "STYLE NUMBER", "join": ["MODEL", "FABRIC", "COLOR"], "sep": ""},
            {"header": "COLOR", "from": "COLOR DESC."},
            {"header": "MODEL", "from": "MODEL"},
            {"header": "FABRIC CODE", "from": "FABRIC"},
            {"header": "COLOR CODE", "from": "COLOR"},
            {"header": "SIZE NOTES", "from": "SIZE"},
            {"header": "FABRIC DESCRIPTION", "from": "FABRIC COMP"},
            {"header": "DESCRIPTION", "from": "DESCRIPTION", "optional": True},
            {"header": "BRAND_ID", "join_out": ["STYLE NUMBER", "COLOR"], "sep": "_", "upper": True},
        ],
        {"mode": "scale_text", "from": "SIZE", "even_all": True},
        photo_only=["PHOTO"]),
    "Armani - Emporio Armani Men": _armani(
        "Armani - Emporio Armani Men", "Armani - Emporio Armani Men",
        {"headers_all": ["VPN", "COL/PATTERN DESC"]},
        ["VPN"], ["STYLE", "FABRIC", "COLOR"],
        [
            {"header": "STYLE NUMBER", "from": "VPN", "join": ["STYLE", "FABRIC", "COLOR"], "sep": ""},
            {"header": "NAME", "from": "PRODUCT"},
            {"header": "CATEGORY", "from": "PRODUCT"},
            {"header": "MODEL", "from": "STYLE"},
            {"header": "FABRIC CODE", "from": "FABRIC"},
            {"header": "COLOR CODE", "from": "COLOR"},
            {"header": "COLOR", "from": "COL/PATTERN DESC"},
            {"header": "SIZE NOTES", "from": "SIZING"},
            {"header": "FABRIC DESCRIPTION", "from": "COMPOSITION"},
            {"header": "BRAND_ID", "join_out": ["STYLE NUMBER", "COLOR"], "sep": "_", "upper": True},
        ],
        {"mode": "scale_text", "from": "SIZING", "even_all": True}),
    "Armani - Emporio Armani Women": _armani(
        "Armani - Emporio Armani Women", "Armani - Emporio Armani Women",
        {"headers_all": ["SKU", "MODEL", "STYLE DESCRIPTION", "PRODUCT"]},
        ["SKU"], ["MODEL", "FABRIC#code", "COLOR#code"],
        [
            {"header": "STYLE NUMBER", "from": "SKU", "join": ["MODEL", "FABRIC#code", "COLOR#code"], "sep": ""},
            {"header": "NAME", "from": "STYLE DESCRIPTION"},
            {"header": "CATEGORY", "from": "PRODUCT"},
            {"header": "MODEL", "from": "MODEL"},
            {"header": "FABRIC CODE", "from": "FABRIC#code"},
            {"header": "COLOR CODE", "from": "COLOR#code"},
            {"header": "FABRIC DESCRIPTION", "from": "FABRIC#words"},
            {"header": "COLOR", "from": "COLOR#words"},
            {"header": "SIZE NOTES", "from": "SIZES"},
            {"header": "AVAILABLE FROM", "date": "from"},
            {"header": "AVAILABLE UNTIL", "date": "until"},
            {"header": "BRAND_ID", "join_out": ["STYLE NUMBER"], "sep": "_", "upper": True},
        ],
        {"mode": "scale_text", "from": "SIZES", "even_all": True}),
    "Armani - Emporio Armani Bodywear": _armani(
        "Armani - Emporio Armani Bodywear", "Armani - Emporio Armani Bodywear",
        {"headers_all": ["MODEL NUMBER", "BARCODE", "COLOR NAME"]},
        ["MODEL NUMBER"], ["STYLE", "FABRIC", "COLOR"],
        [
            {"header": "STYLE NUMBER", "from": "MODEL NUMBER", "join": ["STYLE", "FABRIC", "COLOR"], "sep": ""},
            {"header": "NAME", "from": "STYLE NAME"},
            {"header": "MODEL", "from": "STYLE"},
            {"header": "FABRIC CODE", "from": "FABRIC"},
            {"header": "COLOR CODE", "from": "COLOR"},
            {"header": "COLOR", "from": "COLOR NAME", "upper": True},
            {"header": "COUNTRY OF MATERIAL ORIGIN", "from": "COO", "optional": True},
            {"header": "COLLECTION NAME", "from": "FAMILY NAME", "optional": True},
            {"header": "BRAND_ID", "join_out": ["STYLE NUMBER"], "sep": "_", "upper": True},
        ],
        {"mode": "list", "from": ["SIZE"], "upc_from": "BARCODE", "sort": True, "even_all": True},
        linesheet_from=("LINESHEET NAME", "FAMILY NAME")),
    "Armani - Armani Exchange": _armani(
        "Armani - Armani Exchange", "Armani - Armani Exchange",
        {"headers_all": ["REF_MOD_COD", "REF_PAR_COD", "REF_COL_COD"]},
        ["MODEL NUMBER"], ["REF_MOD_COD", "REF_PAR_COD", "REF_COL_COD"],
        [
            {"header": "STYLE NUMBER", "from": "MODEL NUMBER", "join": ["REF_MOD_COD", "REF_PAR_COD", "REF_COL_COD"], "sep": ""},
            {"header": "GENDER", "from": "GENDER", "optional": True},
            {"header": "CATEGORY", "from": "CATEG"},
            {"header": "NAME", "from": "CATEG"},
            {"header": "MODEL", "from": "REF_MOD_COD"},
            {"header": "FABRIC CODE", "from": "REF_PAR_COD"},
            {"header": "COLOR CODE", "from": "REF_COL_COD"},
            {"header": "COLOR", "from": "COLOR"},
            {"header": "DEPARTMENT", "from": "PRODUCT", "optional": True},
            {"header": "FAMILY NAME", "from": "FAMILY NAME", "optional": True},
            {"header": "RAIL/PACK - ACC FAMILY", "from": "RAIL/PACK - ACC FAMILY", "optional": True},
            {"header": "DELIVERY MONTH", "from": "DELIVERY", "optional": True},
            {"header": "BRAND_ID", "join_out": ["STYLE NUMBER", "COLOR"], "sep": "_", "upper": True},
        ],
        {"mode": "list", "from": ["US SIZE"], "upc_from": "EAN", "sort": True}),
}


def app_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def profiles_folder():
    return os.path.join(app_folder(), "Brand Profiles")


def write_builtin_profiles():
    """Creates the Brand Profiles folder (only when you ask for it) with copies of the built-in profiles."""
    folder = profiles_folder()
    try:
        os.makedirs(folder, exist_ok=True)
        for name, prof in BUILTIN_PROFILES.items():
            fp = os.path.join(folder, safe_filename(name) + ".json")
            old_v = -1
            if os.path.exists(fp):
                try:
                    with open(fp, encoding="utf-8") as f:
                        old_v = int(json.load(f).get("version", 0))
                except Exception:
                    old_v = -1
            if old_v < prof.get("version", 0):
                with open(fp, "w", encoding="utf-8") as f:
                    json.dump(prof, f, indent=2, ensure_ascii=False)
    except OSError:
        pass
    return folder


def load_profiles():
    """Built-in profiles (inside the program) + any *.json in the Brand Profiles folder, if that folder exists.
    A file with the same name wins, unless the built-in one is a newer version."""
    profiles = {k: json.loads(json.dumps(v)) for k, v in BUILTIN_PROFILES.items()}
    folder = profiles_folder()
    problems = []
    try:
        files = sorted(os.listdir(folder))
    except OSError:
        files = []
    for fn in files:
        if not fn.lower().endswith(".json"):
            continue
        try:
            with open(os.path.join(folder, fn), encoding="utf-8") as f:
                prof = json.load(f)
            name = prof.get("name") or os.path.splitext(fn)[0]
            prof["name"] = name
            builtin = BUILTIN_PROFILES.get(name)
            if builtin and int(prof.get("version", 0)) < builtin.get("version", 0):
                continue
            profiles[name] = prof
        except Exception as e:  # noqa
            problems.append(f"{fn}: {e}")
    return profiles, problems


def _find_header_row(rows, must):
    want = {norm_header(m) for m in must}
    for i, row in enumerate(rows[:40], 1):
        have = {norm_header(v) for v in row if not is_blank(v)}
        if want <= have:
            return i
    return None


def _detect_sets(det):
    if not det:
        return []
    if det.get("any_of"):
        return [d for d in det["any_of"] if d.get("headers_all") or d.get("headers_any")]
    return [det] if (det.get("headers_all") or det.get("headers_any")) else []


def _profile_matches(rows, det):
    """Row number of the header row if the rows fit one of the profile's header patterns, else None."""
    for d in _detect_sets(det):
        for i, row in enumerate(rows[:40], 1):
            have = {norm_header(v) for v in row if not is_blank(v)}
            if not have:
                continue
            if d.get("headers_all") and not {norm_header(h) for h in d["headers_all"]} <= have:
                continue
            if d.get("headers_any") and not ({norm_header(h) for h in d["headers_any"]} & have):
                continue
            return i
    return None


def _is_joor(rows):
    for row in rows[:5]:
        have = {norm_header(v) for v in row if not is_blank(v)}
        if "joor style id" in have or {"linesheet id", "wholesale currency 1"} <= have:
            return True
    return False


def identify_file(path, profiles):
    """'Data Identifier': returns (label, profile name or None, sheet to use)."""
    try:
        sheets = list_sheets(path, visible_only=True) or list_sheets(path)
    except Exception:
        return "Unreadable file", None, None
    order = [n for n in profiles if n != "NuORDER Default Data"] + (
        ["NuORDER Default Data"] if "NuORDER Default Data" in profiles else [])
    previews = {}
    for name in order:
        prof = profiles[name]
        det = prof.get("detect", {})
        if not _detect_sets(det):
            continue
        wanted = [s_ for s_ in sheets if s_ in det.get("sheet_names", [])]
        for sh in wanted + [s_ for s_ in sheets if s_ not in wanted]:
            if sh not in previews:
                try:
                    previews[sh] = read_preview(path, sh, max_rows=40)
                except Exception:
                    previews[sh] = []
            if _profile_matches(previews[sh], det):
                return prof.get("label") or name, name, sh
    first = sheets[0] if sheets else None
    if first is not None:
        if first not in previews:
            try:
                previews[first] = read_preview(path, first, max_rows=40)
            except Exception:
                previews[first] = []
        if _is_joor(previews[first]):
            return "JOOR export", None, first
    return "Unknown (general rules)", None, first


def _looks_like_code(values):
    """True when the sample values look like codes (T006G, UC99, 001) rather than words (BLACK, 100% WOOL)."""
    vals = [clean_text(v) for v in values if not is_blank(v)]
    if not vals:
        return None
    code = sum(1 for v in vals if " " not in v and re.search(r"\d", v) and len(v) <= 12)
    return code >= len(vals) / 2


def profile_result(cfg, ctx, L, progress, stop):
    prof = cfg["profile"]
    path = cfg["path"]
    sheets = list_sheets(path)
    sheet = cfg.get("sheet") if cfg.get("sheet") in sheets else (prof.get("sheet") if prof.get("sheet") in sheets else sheets[0])
    rows_preview = read_preview(path, sheet, max_rows=60)
    hr = _profile_matches(rows_preview, prof.get("detect")) or _find_header_row(rows_preview, prof.get("header_contains") or [])
    if not hr:
        raise UserError(f"This sheet does not look like a {prof['name']} file (its column names were not found).")
    ws = open_sheet(path, sheet, data_only=True)
    labels = make_labels(rows_preview[hr - 1])
    norms = [norm_header(re.sub(r"\s*\(\d+\)$", "", l)) for l in labels]
    col = {}
    for i, n_ in reversed(list(enumerate(norms))):
        col[n_] = i
    sample = rows_preview[hr:hr + 40]
    last_row = last_data_row(ws)
    L(f"Sheet '{sheet}', header on row {hr}, data up to row {last_row}.")
    above = " | ".join(clean_text(v) for row in rows_preview[:hr - 1] for v in row if not is_blank(v))

    def idx(name):
        if isinstance(name, list):
            for n_ in name:
                i = idx(n_)
                if i is not None:
                    return i
            return None
        if not name:
            return None
        if "#" in name:                     # 'FABRIC#code' / 'FABRIC#words': two columns with the same name
            base, kind = name.split("#", 1)
            cands = [i for i, n_ in enumerate(norms) if n_ == norm_header(base)]
            for i in cands:
                looks = _looks_like_code([r[i] if i < len(r) else None for r in sample])
                if looks is None:
                    continue
                if (kind == "code") == looks:
                    return i
            return None
        return col.get(norm_header(name))

    def cell_text(v):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return clean_code(v)
        return clean_text(v)

    def first_of(vals, names):
        names = names if isinstance(names, list) else [names]
        for n_ in names:
            i = idx(n_)
            if i is not None and i < len(vals) and not is_blank(vals[i]):
                return vals[i]
        return None

    key_cols = prof.get("row_key", [])
    key_idx = [idx(k) for k in key_cols]
    alt_cols = prof.get("row_key_alt") or []
    alt_idx = [idx(k) for k in alt_cols]
    if any(k is None for k in alt_idx):
        alt_idx = []
    if any(k is None for k in key_idx):
        if not alt_idx:
            raise UserError("The product key column(s) " + ", ".join(key_cols) + " were not found.")
        key_cols, key_idx, alt_idx = alt_cols, alt_idx, []

    columns = [dict(c) for c in prof["columns"]]
    if idx("LINESHEET SEPARATOR") is not None or idx("LINESHEET SEPARATOR NAME") is not None:
        if not any(c["header"] == "LINESHEET SEPARATOR" for c in columns):
            pos = next((k for k, c in enumerate(columns) if c["header"] == "LINESHEET NAME"), -1) + 1
            columns.insert(pos, {"header": "LINESHEET SEPARATOR", "from": ["LINESHEET SEPARATOR", "LINESHEET SEPARATOR NAME"]})
    missing = [c["from"] if isinstance(c["from"], str) else c["from"][0]
               for c in columns if c.get("from") and not c.get("optional") and not c.get("join")
               and idx(c["from"]) is None]
    if missing:
        L("NOTE: these columns were not found in this file and stay empty: " + ", ".join(missing))

    # sizes setup
    sz = prof.get("sizes") or {}
    mode = sz.get("mode", "matrix")
    even_all = bool(sz.get("even_all"))
    size_cols = []
    if mode == "matrix" and sz.get("columns_matching"):
        size_re = re.compile(sz["columns_matching"])
        label_re = re.compile(sz.get("label") or r"^(.*)$")
        for i, l in enumerate(labels):
            if size_re.match(l):
                m = label_re.match(l)
                size_cols.append((i, (m.group(1) if m else l).strip()))
    elif mode == "list":
        if sz.get("columns_matching"):
            rx = re.compile(sz["columns_matching"])
            size_cols = [(i, l) for i, l in enumerate(labels) if rx.match(l)]
        for n_ in sz.get("from", []):
            if idx(n_) is not None:
                size_cols.append((idx(n_), n_))
    avail = {str(v).strip().upper() for v in sz.get("available_values", ["0"])}
    suffix_i = idx(sz.get("suffix_from"))
    scale_i = idx(sz.get("from")) if mode == "scale_text" else None
    upc_i = idx(sz.get("upc_from"))

    # dates
    dates = prof.get("dates") or {}
    title_dates = []
    if dates.get("title_regex"):
        m = re.search(dates["title_regex"], above, re.I)
        if m:
            title_dates = parse_dates(m.group(1).split("|")[0], datetime.date.today().year)
    win_src = dates.get("window_from") or [c.get("from") for c in columns if False]
    if not win_src and any(c.get("date") for c in columns):
        win_src = ["DELIVERY WINDOW", "SHIP WINDOW", "DELIVERY DATES"]
    # season
    sea = prof.get("season") or {}
    season_src = sea.get("from", "file_name")
    if season_src == "title":
        file_season = season_code_from_text(above) or season_code_from_text(os.path.basename(path))
    else:
        file_season = season_code_from_text(os.path.basename(path))
    season_col = None if season_src in ("file_name", "title") else idx(season_src)
    ls_cols = prof.get("linesheet_from") or []
    date_year = datetime.date.today().year
    if dates.get("anchor_year") == "season" and file_season:
        date_year = year_from_season(file_season)

    groups = OrderedDict()
    total = max(last_row - hr, 1)
    for n, cells in enumerate(ws.iter_rows(min_row=hr + 1, max_row=last_row, max_col=len(labels))):
        r = hr + 1 + n
        if n % 200 == 0:
            progress(int(40 * n / total))
            if stop.is_set():
                raise UserError("Stopped by user.")
        vals = [c.value for c in cells]
        key = tuple(cell_text(vals[k]) if k < len(vals) else "" for k in key_idx)
        if not all(key) and alt_idx:          # e.g. no VPN on this row -> Style + Fabric + Color
            key = ("alt",) + tuple(cell_text(vals[k]) if k < len(vals) else "" for k in alt_idx)
        if not all(key):
            continue
        g = groups.get(key)
        if g is None:
            g = groups[key] = {"values": {}, "pairs": [], "rows": [], "season": "", "dates": [], "until": [],
                               "linesheet": ""}
        g["rows"].append(r)
        for c in columns:
            h = c["header"]
            if h in g["values"] or h in ("SIZES", "UPCS") or "value" in c or "date" in c or "join_out" in c:
                continue
            v = first_of(vals, c["from"]) if c.get("from") else None
            if v is None and c.get("join"):
                parts = [cell_text(first_of(vals, j)) for j in c["join"]]
                if parts and c.get("strip_prefix") and parts[0].upper().startswith(c["strip_prefix"].upper()):
                    parts[0] = parts[0][len(c["strip_prefix"]):]
                if parts and c.get("title_first"):
                    parts[0] = parts[0].title()
                v = c.get("sep", "_").join(p_ for p_ in parts if p_) if all(parts) else None
            if c.get("number"):
                v = parse_number(v)
                if c.get("usd") and not v:
                    v = None
            elif v is not None:
                cl = c.get("clean")
                v = (clean_country(v) if cl == "country" else clean_fabric(v) if cl == "fabric"
                     else cell_text(v))
                if c.get("upper") and v:
                    v = v.upper()
            if v not in (None, ""):
                g["values"][h] = v
        if season_col is not None and not g["season"] and season_col < len(vals):
            g["season"] = cell_text(vals[season_col])
        if not g["linesheet"] and ls_cols:
            g["linesheet"] = cell_text(first_of(vals, ls_cols))
        if win_src and not g["dates"]:
            g["dates"] = parse_dates(first_of(vals, win_src), date_year)
        if dates.get("until_from") and not g["until"]:
            g["until"] = parse_dates(first_of(vals, dates["until_from"]), datetime.date.today().year)
        # sizes (+ the UPC of each size when the brand sends one row per size)
        upc = cell_text(vals[upc_i]) if upc_i is not None and upc_i < len(vals) else ""
        known = {p_[0] for p_ in g["pairs"]}
        if mode == "matrix":
            width = cell_text(vals[suffix_i]) if suffix_i is not None and suffix_i < len(vals) else ""
            for i, lbl in size_cols:
                if i < len(vals) and cell_text(vals[i]).upper() in avail:
                    label = f"{lbl} {width}".strip()
                    if label not in known:
                        g["pairs"].append((label, ""))
                        known.add(label)
        elif mode == "scale_text":
            if not g["pairs"] and scale_i is not None and scale_i < len(vals):
                g["pairs"] = [(s_, "") for s_ in parse_size_scale(vals[scale_i], even_all=even_all)]
        else:
            for i, _l in size_cols:
                if i < len(vals):
                    for s_ in split_sizes(vals[i], even_all=even_all):
                        if s_ not in known:
                            g["pairs"].append((s_, upc))
                            known.add(s_)

    if not groups:
        raise UserError("No products were found in this sheet.")
    L(f"Found {len(groups)} products (rows merged by {', '.join(k.split('#')[0] for k in key_cols)}).")

    override = linesheet_label(cfg.get("linesheet_override"))
    default_ls = sheet_linesheet_name(path, cfg)
    res = Result(path, prof.get("label") or prof["name"])
    res.brand = prof.get("brand") or prof.get("label") or prof["name"]
    res.numeric = {c["header"] for c in columns if c.get("number")}
    usd_filled = 0
    pic = prof.get("pictures") or {}
    for g in groups.values():
        season = g["season"] or file_season or cfg.get("season", "")
        season = apply_season_map(season, sea.get("map"))
        linesheet = override or linesheet_label(g["linesheet"])
        if not linesheet and prof.get("linesheet_pattern") and cfg.get("sheet"):
            linesheet = linesheet_label(prof["linesheet_pattern"].replace("{season}", season or "").replace(
                "{sheet}", clean_text(cfg["sheet"])).replace(" - -", " -").strip(" -"))
        linesheet = linesheet or default_ls
        win = g["dates"] or title_dates
        d_from = win[0] if win else None
        d_until = (g["until"][0] if g["until"] else (win[1] if len(win) >= 2 else None))
        if d_from and d_until and d_until < d_from and d_until.year == d_from.year:
            if dates.get("anchor_year") == "season":        # 'Dec 1 - Feb 28' of a 2027 season = Dec 2026 - Feb 2027
                d_from = d_from.replace(year=d_from.year - 1)
            else:
                d_until = d_until.replace(year=d_until.year + 1)
        pairs = g["pairs"]
        if sz.get("sort"):
            order = {s_: k for k, s_ in enumerate(sort_sizes([p_[0] for p_ in pairs]))}
            pairs = sorted(pairs, key=lambda p_: order.get(p_[0], 0))
        sizes = [p_[0] for p_ in pairs]
        upcs = [p_[1] for p_ in pairs] if upc_i is not None else []
        if sizes and set(sizes) == {"OS"}:
            sizes, upcs = ["OS"], upcs[:1]
        vals = {}
        for c in columns:
            h = c["header"]
            if h in ("SIZES", "UPCS") or "join_out" in c:
                continue
            if "value" in c:
                vals[h] = str(c["value"]).replace("{linesheet}", linesheet).replace("{season}", season) or None
            elif c.get("date") == "from":
                vals[h] = format_date(d_from, cfg.get("date_format", "EXCEL"))
            elif c.get("date") == "until":
                vals[h] = format_date(d_until, cfg.get("date_format", "EXCEL"))
            else:
                vals[h] = g["values"].get(h)
            if c.get("usd") and not vals.get(h):
                vals[h] = 0.01
                usd_filled += 1
        for c in columns:                       # columns built from other output columns (BRAND_ID ...)
            if "join_out" in c:
                parts = [clean_text(vals.get(j)) for j in c["join_out"]]
                v = c.get("sep", "_").join(p_ for p_ in parts if p_) if all(parts) else None
                vals[c["header"]] = v.upper() if (v and c.get("upper")) else v
        if "SEASON" in vals and not vals["SEASON"]:
            vals["SEASON"] = season or None
        if pic.get("name_join"):
            parts = [clean_text(vals.get(j)) for j in pic["name_join"]]
            pic_name = pic.get("sep", "_").join(parts) if all(parts) else None
        else:
            pic_name = vals.get(pic.get("name_from", "STYLE NUMBER"))
        res.rows.append({"values": vals, "sizes": sizes, "upcs": upcs, "src_rows": g["rows"], "pic_name": pic_name})
    if usd_filled:
        L(f"{usd_filled} USD price cells hold the 0.01 placeholder (price missing or 0).")
    if prof.get("pdf") == "maxmara":
        apply_maxmara_pdf(res, cfg, ctx, L, stop, progress)

    headers = []
    for c in columns:
        h = c["header"]
        if h == "UPCS" and not any(r_["upcs"] and any(r_["upcs"]) for r_ in res.rows):
            continue
        if c.get("optional") and not any(r["values"].get(h) not in (None, "") for r in res.rows):
            continue
        if h not in headers:
            headers.append(h)
    res.headers = headers

    if pic.get("extract") and cfg.get("extract_pictures", True):
        row_to_name = {}
        for r_ in res.rows:
            for src in r_["src_rows"]:
                row_to_name[src] = r_["pic_name"]
        only = None
        if pic.get("only_columns"):
            only = {idx(n_) + 1 for n_ in pic["only_columns"] if idx(n_) is not None}
        extract_row_pictures(path, sheet, row_to_name, ctx, L, only_columns=only)
    validate_rows(res, L)
    return res


def _composition_warnings(products, L):
    """A composition whose percentages do not add up to 100 is probably cut off in the PDF itself."""
    for model, c in sorted(products.items()):
        for seg in (c.get("fabric") or "").split(";"):
            nums = [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)\s*%", seg)]
            if nums and abs(sum(nums) - 100) > 0.01:
                L(f"  CHECK fabric of {model}: '{seg.strip()}' adds up to {sum(nums):g}% - the PDF text may be cut off.")


def apply_maxmara_pdf(res, cfg, ctx, L, stop, progress):
    """Sizes, fabric, article number and pictures of a MaxMara data file come from the brand's PDF catalogs."""
    pdfs = list(cfg.get("pdf_paths") or [])
    if not pdfs:
        if cfg.get("extract_pictures", True):
            L("NOTE: no MaxMara PDF catalog was added, so SIZE 1..., FABRIC DESCRIPTION and the pictures stay empty. "
              "Add the PDF file(s) to the list and run again.")
        return
    if not pymupdf_available():
        L("NOTE: the PDF library (PyMuPDF) is missing - run Build_NuORDER_Client.bat again. Sizes and pictures were skipped.")
        return
    cache = ctx.setdefault("maxmara_pdf", {})
    key = tuple(pdfs)
    if cache.get("key") != key:
        L(f"Reading {len(pdfs)} PDF catalog(s) ...")
        products, images = read_maxmara_catalogs(pdfs, L, stop, progress)
        cache.update(key=key, products=products, images=images, used=set())
        L(f"  {len(products)} products and {len(images)} pictures found in the PDFs.")
        _composition_warnings(products, L)
    products, images, used = cache["products"], cache["images"], cache["used"]
    no_model, no_size, saved = [], [], 0
    for row in res.rows:
        v = row["values"]
        model, var = clean_text(v.get("NAME")), clean_text(v.get("COLOR CODE"))
        c = products.get(model)
        if not c:
            no_model.append(model)
            continue
        used.add(model)
        sizes = c["variants"].get(var)
        if not sizes:
            no_size.append(f"{model}_{var}")
            sizes = [x for x in c["sizes"] if any(x in s_ for s_ in c["variants"].values())]
        row["sizes"] = list(sizes)
        if c.get("fabric"):
            v["FABRIC DESCRIPTION"] = c["fabric"]
        if cfg.get("extract_pictures", True) and model in images and row.get("pic_name"):
            if save_picture(ctx, row["pic_name"], ".jpg", images[model]):
                saved += 1
    L(f"PDF data applied: sizes for {len(res.rows) - len(no_model)} of {len(res.rows)} rows, {saved} picture(s) saved.")
    if no_model:
        L("  NOT in the PDFs (no sizes / fabric / picture): " + ", ".join(sorted(set(no_model))))
    if no_size:
        L("  Color not listed in the PDF size grid (the sizes of the other colors were used): " + ", ".join(no_size))




# ======================================================================
#  UPDATES - new versions are compiled by GitHub and published as a Release
# ======================================================================

def parse_version(text):
    nums = re.findall(r"\d+", str(text or ""))
    return tuple(int(n) for n in nums[:4]) or (0,)


def is_newer(latest, current):
    a, b = list(parse_version(latest)), list(parse_version(current))
    n = max(len(a), len(b))
    a += [0] * (n - len(a))
    b += [0] * (n - len(b))
    return a > b


def _http_get(url, timeout=20, accept="application/vnd.github+json"):
    req = urllib.request.Request(url, headers={"User-Agent": f"{APP_NAME}/{APP_VERSION}", "Accept": accept})
    return urllib.request.urlopen(req, timeout=timeout)


def fetch_latest_release(timeout=20):
    """The newest published version on GitHub: {version, tag, notes, page, exe, py} (exe / py = asset info or None)."""
    try:
        with _http_get(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest", timeout) as r:
            data = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise UserError("No version has been published on GitHub yet.")
        raise UserError(f"GitHub answered with error {e.code}. Try again later.")
    except (urllib.error.URLError, OSError, ValueError):
        raise UserError("Could not reach GitHub. Check the internet connection and try again.")
    tag = str(data.get("tag_name") or "")
    assets = {a.get("name"): a for a in data.get("assets", []) if isinstance(a, dict)}
    return dict(tag=tag, version=tag.lstrip("vV"), notes=str(data.get("body") or "").strip(),
                page=data.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases",
                exe=assets.get(RELEASE_EXE), py=assets.get(RELEASE_PY))


def download_asset(asset, dest, progress=None, stop=None, allow_any_host=False):
    """Downloads one Release file (only from this program's own GitHub repository) to dest."""
    url = str(asset.get("browser_download_url") or "")
    if not allow_any_host and not url.startswith(f"https://github.com/{GITHUB_REPO}/releases/download/"):
        raise UserError("Unexpected download address - the update was cancelled.")
    tmp = dest + ".part"
    try:
        with _http_get(url, 120, "application/octet-stream") as r, open(tmp, "wb") as f:
            total = int(r.headers.get("Content-Length") or asset.get("size") or 0)
            got = 0
            while True:
                if stop is not None and stop.is_set():
                    raise UserError("Stopped by user.")
                chunk = r.read(1 << 16)
                if not chunk:
                    break
                f.write(chunk)
                got += len(chunk)
                if progress:
                    progress(got, total)
        if asset.get("size") and os.path.getsize(tmp) != asset["size"]:
            raise UserError("The download is incomplete. Please try again.")
        os.replace(tmp, dest)
        return dest
    except (urllib.error.URLError, OSError) as e:
        raise UserError(f"The download failed: {e}")
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def install_exe_update(new_path, current=None, check_exe=True):
    """Windows lets a running program be renamed (not deleted): the old file becomes '.old', the new one takes its name."""
    current = current or sys.executable
    if check_exe:
        with open(new_path, "rb") as f:
            if f.read(2) != b"MZ":
                raise UserError("The downloaded file is not a Windows program.")
    old = current + ".old"
    try:
        if os.path.exists(old):
            os.remove(old)
    except OSError:
        pass
    os.replace(current, old)
    try:
        os.replace(new_path, current)
    except Exception:
        os.replace(old, current)
        raise
    return current


def install_py_update(new_path, current):
    """Running the .py file: replace it (a .bak copy of the old one is kept)."""
    with open(new_path, encoding="utf-8") as f:
        compile(f.read(), new_path, "exec")              # refuses a broken file
    shutil.copy2(current, current + ".bak")
    os.replace(new_path, current)
    return current


def cleanup_old_update():
    if getattr(sys, "frozen", False):
        try:
            os.remove(sys.executable + ".old")
        except OSError:
            pass


def relaunch_program():
    env = dict(os.environ)
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    env.pop("_MEIPASS2", None)
    if getattr(sys, "frozen", False):
        subprocess.Popen([sys.executable], env=env, close_fds=True)
    else:
        subprocess.Popen([sys.executable, os.path.abspath(sys.argv[0])], env=env, close_fds=True)



# ======================================================================
#  MASTER FILE: one or several results -> one NuORDER workbook
# ======================================================================

LINESHEET_TINTS = ["FFC7CE", "BDD7EE", "C6EFCE", "FFEB9C", "E4DFEC", "FCE4D6", "CCF2F2", "F8CBDF",
                   "E2EFDA", "D9D2E9", "FFE5CC", "DDEBF7", "EDEDED", "D0F0F8"]
TINT_NAMES = ["red", "blue", "green", "yellow", "purple", "orange", "teal", "pink",
              "light green", "lavender", "peach", "light blue", "grey", "aqua"]


def master_name(paths):
    """One file: its own name.  Several: the part of the names they share (else the first name)."""
    stems = [os.path.splitext(os.path.basename(p))[0] for p in paths]
    if len(stems) == 1:
        return safe_filename(stems[0])
    pre = os.path.commonprefix([s_.upper() for s_ in stems])
    pre = stems[0][:len(pre)]
    if pre and len(pre) < min(len(s_) for s_ in stems):
        cut = max(pre.rfind(ch) for ch in " _-.")
        pre = pre[:cut] if cut > 0 else pre
    pre = pre.strip(" _-.")
    return safe_filename(pre if len(pre) >= 4 else stems[0])


def merge_headers(results):
    union = ["LINESHEET NAME"]
    if any("LINESHEET SEPARATOR" in r.headers for r in results):
        union.append("LINESHEET SEPARATOR")
    for res in results:
        prev = None
        for h in res.headers:
            if h in ("LINESHEET NAME", "LINESHEET SEPARATOR"):
                prev = h
                continue
            if h not in union:
                pos = union.index(prev) + 1 if prev in union else len(union)
                union.insert(pos, h)
            prev = h
    if "SIZES" not in union:
        union.append("SIZES")
    return union


def write_master(results, xlsx_path, L):
    union = merge_headers(results)
    max_sizes = max((len(r["sizes"]) for res in results for r in res.rows), default=1) or 1
    max_upcs = max((len(r.get("upcs") or []) for res in results for r in res.rows), default=0)
    if max_upcs and "UPCS" not in union:
        union.insert(union.index("SIZES") + 1, "UPCS")
    headers = []
    for h in union:
        if h == "SIZES":
            headers += [f"SIZE {i}" for i in range(1, max_sizes + 1)]
        elif h == "UPCS":
            headers += [f"UPC {i}" for i in range(1, max_upcs + 1)]
        else:
            headers.append(h)
    text_cols = {f"SIZE {i}" for i in range(1, max_sizes + 1)} | {f"UPC {i}" for i in range(1, max_upcs + 1)}
    numeric = set().union(*(res.numeric for res in results)) if results else set()
    linesheets = list(dict.fromkeys(r["values"].get("LINESHEET NAME") or "" for res in results for r in res.rows))
    color_rows = len(linesheets) > 1
    tint = {ls: LINESHEET_TINTS[i % len(LINESHEET_TINTS)] for i, ls in enumerate(linesheets)}

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = FONT_HEADER
    r_i = 1
    seen_ids = {}
    dup_report = []
    for res in results:
        for row in res.rows:
            r_i += 1
            vals = dict(row["values"])
            for i in range(max_sizes):
                vals[f"SIZE {i + 1}"] = row["sizes"][i] if i < len(row["sizes"]) else None
            upcs = row.get("upcs") or []
            for i in range(max_upcs):
                vals[f"UPC {i + 1}"] = upcs[i] if i < len(upcs) else None
            ident = vals.get("BRAND_ID") or "|".join(str(vals.get(k) or "") for k in ("STYLE NUMBER", "COLOR"))
            if ident.strip("|"):
                if ident in seen_ids and seen_ids[ident] != res.path:
                    dup_report.append(f"{ident}  (in {os.path.basename(seen_ids[ident])} and {os.path.basename(res.path)})")
                seen_ids.setdefault(ident, res.path)
            fill = PatternFill("solid", fgColor=tint.get(vals.get("LINESHEET NAME") or "")) if color_rows else None
            for c, h in enumerate(headers, 1):
                v = vals.get(h)
                if norm_header(h) in COLOR_CODE_HEADERS:
                    v = pad_color_code(v)
                if v == "":
                    v = None
                cell = ws.cell(row=r_i, column=c, value=v)
                cell.font = FONT_DATA
                if fill is not None:
                    cell.fill = fill
                if h in text_cols and v is not None:
                    cell.value = clean_code(v)           # sizes / UPCs as text, so 00 stays 00
                    cell.number_format = "@"
                elif isinstance(v, (datetime.date, datetime.datetime)):
                    cell.number_format = "m/d/yyyy"      # real date, shown as 1/15/2027
                elif h in numeric and isinstance(v, (int, float)):
                    cell.number_format = "0.00"
                elif isinstance(v, str) and v.isdigit():
                    cell.number_format = "@"
    auto_fit(ws)
    ws.freeze_panes = "A2"

    ws_sum = wb.create_sheet("Linesheet Summary")
    ws_sum.cell(row=1, column=1, value="LINESHEET NAME").font = FONT_HEADER
    for i, nm in enumerate(linesheets, 2):
        cell = ws_sum.cell(row=i, column=1, value=nm)
        cell.font = FONT_DATA
        if color_rows:
            cell.fill = PatternFill("solid", fgColor=tint[nm])
    ws_sum.column_dimensions["A"].width = max((len(n_) for n_ in linesheets), default=10) + 5
    try:
        wb.save(xlsx_path)
    except PermissionError:
        raise UserError(f"Could not save '{os.path.basename(xlsx_path)}'. If it is open in Excel, close it and try again.")
    L(f"Saved: {xlsx_path}")
    L(f"   {r_i - 1} rows, {len(headers)} columns, up to {max_sizes} size column(s), {len(linesheets)} linesheet(s).")
    counts = OrderedDict((nm, 0) for nm in linesheets)
    for res in results:
        for row in res.rows:
            counts[row["values"].get("LINESHEET NAME") or ""] += 1
    if color_rows:
        L("   Each linesheet has its own row color (see the 'Linesheet Summary' tab).")
    if dup_report:
        L(f"CHECK: {len(dup_report)} product(s) appear in more than one file:")
        for d in dup_report[:15]:
            L("   - " + d)
    return r_i - 1, len(linesheets), dup_report, [[nm, counts[nm]] for nm in linesheets]


# ======================================================================
#  IMAGE RENAMING - match loose image files to products and copy them as BRAND_ID_01 ...
# ======================================================================

def _alnum_tokens(s_):
    return [t for t in re.split(r"[^A-Z0-9]+", str(s_ or "").upper()) if t]


def _code_regex(code):
    """Style / color code -> regex that ignores separators, with sensible boundaries:
    N20104 matches 'N20104_BLK' and 'n20104-blk' but not 'N201045'."""
    toks = _alnum_tokens(code)
    if not toks:
        return None
    body = r"[^A-Z0-9]*".join(re.escape(t) for t in toks)
    first, last = toks[0][0], toks[-1][-1]
    before = r"(?<![0-9])" if first.isdigit() else r"(?<![A-Z])"
    after = r"(?![0-9])" if last.isdigit() else r"(?![A-Z])"
    return re.compile(before + body + after)


def _is_abbrev(tok, word):
    """'BLK' is a short form of 'BLACK', 'PRT' of 'PRINT', 'NVY' of 'NAVY'."""
    if tok == word:
        return True
    if len(tok) < 2 or tok[0] != word[0] or len(tok) > len(word):
        return False
    it = iter(word)
    return all(ch in it for ch in tok)


def _color_score(file_tokens, file_text, color, code):
    """How strongly an image name points at one color: color code counts most, then color words."""
    score = 0.0
    if code:
        c = clean_code(code)
        variants = {c, c.lstrip("0") or c, c.zfill(3)}
        for v in variants:
            rx = _code_regex(v)
            if rx and len(v) >= 2 and rx.search(file_text):
                score += 3.0
                break
    words = [w for w in _alnum_tokens(color) if len(w) >= 2]
    if words:
        hit = sum(1 for w in words if any(_is_abbrev(t, w) for t in file_tokens if not t.isdigit()))
        score += 2.0 * hit / len(words)
    return score


def _tok_regex(t):
    before = r"(?<![0-9])" if t[0].isdigit() else r"(?<![A-Z])"
    after = r"(?![0-9])" if t[-1].isdigit() else r"(?![A-Z])"
    return re.compile(before + re.escape(t) + after)


def match_image_names(image_paths, products):
    """products: [{'name': file name to use (BRAND_ID), 'style', 'brand_id', 'style2', 'color', 'code', 'title'}]
    Returns {image path: (product index or None, reason)}.

    1. Anchor: the style number (or BRAND_ID / short Style) written anywhere in the image name, with any
       separators or none. If only part of it is there (e.g. A00QMX_8746 for A00QMX_FAHKE_8746), the first
       part must be present and the product with the most matching parts wins.
    2. Several products left (same style in several colors): the color code / color name decides
       (BLK = BLACK, PRT = PRINT ...), then the product name. Still a tie -> skipped."""
    prepared = []
    for p in products:
        full, parts = [], []
        for code in (p.get("style"), p.get("brand_id"), p.get("style2"), p.get("picname")):
            if not code:
                continue
            toks = _alnum_tokens(code)
            if not toks or len("".join(toks)) < 3:
                continue
            full.append((len("".join(toks)), _code_regex(code)))
            if len(toks) > 1 and len(toks[0]) >= 4:
                parts.append([_tok_regex(t) for t in toks])
        prepared.append((full, parts))
    out = {}
    for path in image_paths:
        stem = os.path.splitext(os.path.basename(path))[0].upper()
        toks = _alnum_tokens(stem)
        scores = {}
        for k, (full, parts) in enumerate(prepared):
            f_ = max((ln for ln, rx in full if rx is not None and rx.search(stem)), default=0)
            p_ = 0
            for regs in parts:
                if regs[0].search(stem):
                    p_ = max(p_, sum(1 for r_ in regs if r_.search(stem)))
            if f_ or p_:
                scores[k] = (f_, p_)
        if not scores:
            out[path] = (None, "no product's style number found in the name")
            continue
        top = max(scores.values())
        cands = [k for k, v in scores.items() if v == top]
        names = {products[k]["name"] for k in cands}
        if len(names) == 1:
            out[path] = (cands[0], "style number")
            continue
        scored = []
        for k in cands:
            p = products[k]
            sc = _color_score(toks, stem, p.get("color"), p.get("code"))
            if p.get("title"):
                tw = [w for w in _alnum_tokens(p["title"]) if len(w) >= 3]
                if tw:
                    sc += 0.5 * sum(1 for w in tw if w in toks) / len(tw)
            scored.append((sc, k))
        scored.sort(reverse=True)
        if scored[0][0] > 0 and scored[0][0] > scored[1][0]:
            out[path] = (scored[0][1], "style number + color")
        else:
            out[path] = (None, f"fits {len(names)} products and the color could not be told apart")
    return out


def rename_images(results, src_dir, recurse, ctx, L, stop=None, progress=None):
    """Copies the images of a folder into the output Images folder, renamed BRAND_ID_01, _02 ..."""
    import shutil
    products = []
    for res in results:
        for r in res.rows:
            v = r["values"]
            name = r.get("pic_name") or v.get("BRAND_ID") or v.get("STYLE NUMBER")
            if not name:
                continue
            products.append({"name": str(name), "brand_id": v.get("BRAND_ID"), "style": v.get("STYLE NUMBER"),
                             "color": v.get("COLOR"),
                             "code": v.get("COLOR CODE") or v.get("Color code") or v.get("NRF COLOR CODE"),
                             "title": v.get("NAME"), "style2": v.get("Style") or v.get("MODEL"),
                             "picname": r.get("pic_name")})
    out_img = os.path.abspath(ctx["img_dir"])
    images = [p for p in list_images(src_dir, recurse) if not os.path.abspath(p).startswith(out_img + os.sep)
              and not is_program_output(p, src_dir)]
    L("")
    L(f"=== Renaming images from: {src_dir} ===")
    L(f"{len(images)} image file(s) found, {len(products)} product(s) to match against.")
    if not images or not products:
        return {"images": len(images), "copied": 0, "skipped": 0, "products": 0}
    matches = match_image_names(images, products)
    by_product = OrderedDict()
    skipped = []
    for path in images:
        k, why = matches[path]
        if k is None:
            skipped.append((path, why))
        else:
            by_product.setdefault(k, []).append(path)
    os.makedirs(ctx["img_dir"], exist_ok=True)
    log_rows = []
    copied = 0
    done_n, total = 0, sum(len(v) for v in by_product.values())
    for k, paths in by_product.items():
        name = products[k]["name"]
        key = safe_filename(name).upper()
        for path in sorted(paths, key=lambda p_: natural_key(os.path.basename(p_))):
            if stop is not None and stop.is_set():
                raise UserError("Stopped by user.")
            n = ctx["pic_count"].get(key, 0) + 1
            ctx["pic_count"][key] = n
            ext = os.path.splitext(path)[1].lower()
            px = _resize_px(ctx) if ext not in _NO_RESIZE_EXTS else None
            new = f"{safe_filename(name)}_{n:02d}{'.jpg' if px else ext}"
            target = os.path.join(ctx["img_dir"], new)
            done_n += 1
            if progress is not None:
                progress(done_n, total)
            if ctx["skip_existing"] and os.path.exists(target):
                ctx["pic_stats"]["existing"] += 1
                log_rows.append((os.path.relpath(path, src_dir), new, "already there"))
                continue
            how = "copied"
            if px:
                try:
                    resize_to_jpg(path, target, px)
                    _note_resize(ctx, True)
                    how = f"copied + resized to {px} px"
                except Exception:
                    _note_resize(ctx, False)
                    new = f"{safe_filename(name)}_{n:02d}{ext}"
                    target = os.path.join(ctx["img_dir"], new)
                    shutil.copy2(path, target)
                    how = "copied (could not be resized)"
            else:
                shutil.copy2(path, target)
            copied += 1
            log_rows.append((os.path.relpath(path, src_dir), new, how + " (" + matches[path][1] + ")"))
    for path, why in skipped:
        log_rows.append((os.path.relpath(path, src_dir), "", "SKIPPED - " + why))
    ctx["pic_stats"]["saved"] += copied
    no_image = [p["name"] for k, p in enumerate(products) if k not in by_product
                and not ctx["pic_count"].get(safe_filename(p["name"]).upper())]
    L(f"Renamed copies: {copied} image(s) for {len(by_product)} product(s).")
    if skipped:
        L(f"CHECK: {len(skipped)} image(s) were NOT copied because they could not be matched with certainty:")
        for path, why in skipped[:15]:
            L(f"   - {os.path.basename(path)}  ({why})")
        if len(skipped) > 15:
            L(f"   ... and {len(skipped) - 15} more (see Image_Renaming_Log.csv)")
    if no_image:
        L(f"{len(no_image)} product(s) have no image: " + ", ".join(no_image[:10]) + (" ..." if len(no_image) > 10 else ""))
    try:
        with open(os.path.join(ctx["out_dir"], "Image_Renaming_Log.csv"), "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["ORIGINAL FILE", "NEW NAME", "RESULT"])
            w.writerows(log_rows)
        L("Every old -> new name is listed in Image_Renaming_Log.csv")
    except OSError:
        pass
    return {"images": len(images), "copied": copied, "skipped": len(skipped), "products": len(by_product)}


# ----------------------------------------------------------------------
#  Brand folder scan - find every data file and image in a folder the brand sent
# ----------------------------------------------------------------------

DATA_EXTS = (".xlsx", ".xlsm", ".xlsb", ".csv")
_OUTPUT_SUFFIXES = ("_processed", "_updated", "_changes")
_OUTPUT_FILES = {"image_renaming_log.csv", "failed_downloads.csv"}


def is_program_output(path, root=None):
    """True for files this program made itself (so a re-scan never picks up its own results)."""
    rel = os.path.relpath(path, root) if root else path
    parts = [p_.lower() for p_ in re.split(r"[\\/]+", rel)]
    for d in parts[:-1]:
        if d.endswith("_nuorder") or d == "renamed images":
            return True
    name = parts[-1]
    stem = os.path.splitext(name)[0]
    if name in _OUTPUT_FILES:
        return True
    return os.path.splitext(name)[1] in DATA_EXTS + (".txt",) and (stem.endswith(_OUTPUT_SUFFIXES) or stem.endswith("_log"))


def scan_brand_folder(folder, recurse=True):
    """-> (data files, number of images) found in a brand's folder (and its sub-folders)."""
    data, n_img = [], 0
    walker = os.walk(folder) if recurse else [(folder, [], os.listdir(folder))]
    for root, _dirs, files in walker:
        for f in files:
            full = os.path.join(root, f)
            if f.startswith("~$") or f.startswith(".") or is_program_output(full, folder):
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext in DATA_EXTS:
                data.append(full)
            elif ext in IMAGE_EXTS:
                n_img += 1
    return sorted(data, key=lambda p_: natural_key(os.path.basename(p_))), n_img


def scan_brand_pdfs(folder, recurse=True):
    """The PDF files of a brand's folder (MaxMara catalogs ...)."""
    out = []
    walker = os.walk(folder) if recurse else [(folder, [], os.listdir(folder))]
    for root, _dirs, files in walker:
        for f in files:
            if f.lower().endswith(".pdf") and not f.startswith(("~$", ".")):
                out.append(os.path.join(root, f))
    return sorted(out, key=lambda p_: natural_key(os.path.basename(p_)))


def auto_file_settings(path, profiles):
    """Everything the program can work out by itself about one file (the window lets you change it)."""
    label, prof_name, sheet = identify_file(path, profiles)
    st = {"path": path, "label": label, "profile_name": prof_name, "sheet": sheet}
    try:
        rows = read_preview(path, sheet)
    except Exception as e:  # noqa
        st["error"] = str(e)
        rows = []
    if prof_name:
        st["labels"] = []
        det = profiles[prof_name].get("detect", {})
        st["more_sheets"] = []
        for sh in (list_sheets(path, visible_only=True) if sheet else []):
            if sh != sheet:
                try:
                    if _profile_matches(read_preview(path, sh, max_rows=40), det):
                        st["more_sheets"].append(sh)
                except Exception:
                    pass
        return st
    h = detect_header_row(rows) if rows else 1
    labels = make_labels(rows[h - 1]) if rows else []
    guess = auto_map(labels)
    sc = detect_size_columns(labels)
    st.update(header_row=h, labels=labels, mapping={k: guess.get(k) for k, _ in MAP_FIELDS},
              size_mode="matrix" if len(sc) >= 3 else "single", size_cols=sc,
              size_value_cols=detect_size_value_columns(labels), link_col=default_link_choice(labels))
    st["more_sheets"] = same_layout_sheets(path, sheet, labels, h)
    return st


def same_layout_sheets(path, sheet, labels, header_row):
    """Other visible sheets of the file whose column names are the same as the chosen sheet's."""
    want = [norm_header(l) for l in labels]
    out = []
    try:
        sheets = list_sheets(path, visible_only=True)
    except Exception:
        return out
    for sh in sheets:
        if sh == sheet:
            continue
        try:
            rows = read_preview(path, sh, max_rows=max(header_row, 5))
        except Exception:
            continue
        if len(rows) >= header_row and [norm_header(l) for l in make_labels(rows[header_row - 1])] == want:
            out.append(sh)
    return out


def default_link_choice(labels):
    cols = guess_link_columns(labels)
    if len(cols) > 1:
        return "ALL"
    return cols[0] if cols else guess_link_column(labels)


def link_cols_of(st):
    lc = st.get("link_col")
    if lc == "ALL":
        return guess_link_columns(st.get("labels", []))
    return [lc] if isinstance(lc, int) else []


def file_job(st, opts, profiles):
    """Per-file settings + window options -> what the engines need."""
    base = dict(path=st["path"], sheet=st.get("sheet"), label=st.get("label"),
                linesheet_override=opts.get("linesheet_override", ""), season=opts.get("season", ""),
                date_format=opts.get("date_format", "EXCEL"),
                extract_pictures=opts.get("extract_pictures", True), skip_existing=opts.get("skip_existing", True),
                pdf_paths=list(opts.get("pdf_paths") or []))
    if st.get("profile_name") and st["profile_name"] in profiles:
        base["profile"] = profiles[st["profile_name"]]
        return base
    season = opts.get("season", "") or season_code_from_text(os.path.basename(st["path"]))
    try:
        year = int(opts.get("year"))
    except (TypeError, ValueError):
        year = year_from_season(season)
    base.update(header_row=st["header_row"], labels=st["labels"], mapping=st["mapping"],
                size_mode=st["size_mode"], size_cols=st["size_cols"], size_value_cols=st.get("size_value_cols", []),
                black_unavailable=opts.get("black_unavailable", True), require_value=opts.get("require_value", False),
                default_year=year, date_format=opts.get("date_format", "EXCEL"),
                brand_source=opts.get("brand_source", "name"), brand_upper=opts.get("brand_upper", True),
                price_currency=opts.get("price_currency", "AUTO"),
                download=bool(opts.get("download", True) and link_cols_of(st)),
                link_cols=link_cols_of(st))
    return base


def file_jobs(st, opts, profiles):
    """One job per sheet: the chosen sheet plus every other sheet with the same layout."""
    first = file_job(st, opts, profiles)
    sheets = [st.get("sheet")] + list(st.get("more_sheets") or [])
    if len(sheets) == 1:
        return [first]
    jobs = []
    for sh in sheets:
        j = dict(first)
        j["sheet"] = sh
        j["multi_sheet"] = True
        jobs.append(j)
    return jobs


def history_files(files):
    """One entry per source file (a file with 5 sheets is listed once, '5 sheets')."""
    out = OrderedDict()
    for f in files:
        e = out.setdefault(f["path"], {"name": os.path.basename(f["path"]), "type": f.get("label", ""), "sheets": 0})
        e["sheets"] += 1
    for e in out.values():
        if e["sheets"] > 1:
            e["name"] += f"  [{e['sheets']} sheets]"
    return list(out.values())


def run_job(cfg, log, progress, stop):
    """cfg["files"] = list of per-file settings (each with "profile" or the general-rule settings)."""
    files = cfg["files"]
    name = master_name([f["path"] for f in files])
    out_dir = os.path.join(cfg["out_root"], name + "_NuORDER")
    os.makedirs(out_dir, exist_ok=True)
    lines = []

    def L(msg):
        lines.append(msg)
        log(msg)

    ctx = {"out_dir": out_dir, "img_dir": os.path.join(out_dir, "Images"), "skip_existing": cfg["skip_existing"],
           "pic_hashes": {}, "pic_count": {}, "pic_stats": {"saved": 0, "existing": 0, "repeat": 0},
           "resize_px": cfg.get("resize_px")}
    L(f"=== {APP_NAME} {APP_VERSION} ===")
    L(f"Output folder: {out_dir}")
    if cfg.get("resize_px"):
        L(f"Images are resized to {cfg['resize_px']} px on the shortest side and saved as JPG"
          + ("" if pillow_available() else "  (NOT possible: the Pillow library is missing)") + ".")
    results, failed = [], []
    for k, f in enumerate(files, 1):
        L("")
        L(f"--- File {k} of {len(files)}: {os.path.basename(f['path'])}"
          + (f"  (sheet '{f['sheet']}')" if f.get("multi_sheet") else "") + f"   [Data type: {f.get('label', '?')}] ---")

        def prog(v, base=k - 1):
            progress(int(90 * (base + v / 100.0) / len(files)), 100)

        try:
            res = (profile_result if f.get("profile") else generic_result)(f, ctx, L, prog, stop)
            results.append(res)
            L(f"OK: {len(res.rows)} products from this file.")
        except UserError as e:
            if stop.is_set():
                raise
            failed.append((f["path"], str(e)))
            L(f"SKIPPED THIS FILE: {e}")
        except Exception as e:  # noqa
            failed.append((f["path"], str(e)))
            L("SKIPPED THIS FILE - unexpected error:\n" + traceback.format_exc())
    if not results:
        raise UserError("None of the files could be processed - see the log.")
    L("")
    L("=== Master file ===")
    xlsx_path = os.path.join(out_dir, f"{name}_Processed.xlsx")
    n_rows, n_ls, dups, ls_counts = write_master(results, xlsx_path, L)
    if cfg.get("rename_dir"):
        rename_images(results, cfg["rename_dir"], cfg.get("rename_recurse", True), ctx, L, stop)
    st = ctx["pic_stats"]
    if st["saved"] or st["existing"]:
        L(f"Images folder: {st['saved']} picture(s) saved, {st['existing']} already there"
          + (f", {st['repeat']} repeated copies skipped" if st["repeat"] else "") + ".")
    rl = resize_log_line(ctx)
    if rl:
        L(rl)
    L("")
    L(f"Linesheets processed ({len(ls_counts)}):")
    for i, (nm, cnt) in enumerate(ls_counts):
        L(f"   • {nm or '(no linesheet name)'}  —  {cnt} product(s)"
          + (f"   [{TINT_NAMES[i % len(TINT_NAMES)]} rows]" if len(ls_counts) > 1 else ""))
    progress(100, 100)
    try:
        with open(os.path.join(out_dir, f"{name}_Log.txt"), "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines))
    except OSError:
        pass
    warn_n = sum(len(r.warnings) for r in results) + (1 if dups else 0)
    n_files = len({f["path"] for f in files})
    if cfg.get("record_history", True):
        brands = list(dict.fromkeys(b for b in ((r.brand or "") for r in results) if b))
        run_id = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
        saved = save_history_copy(xlsx_path, run_id, brands[0] if brands else name)
        add_history({
            "id": run_id, "kind": "clean", "saved_copy": saved, "linesheet_names": ls_counts,
            "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "brand": ", ".join(brands),
            "files": history_files(files),
            "output": xlsx_path, "folder": out_dir, "log": os.path.join(out_dir, f"{name}_Log.txt"),
            "products": n_rows, "linesheets": n_ls, "images": st["saved"] + st["existing"],
            "seasons": sorted({str(r["values"].get("SEASON")) for res in results for r in res.rows if r["values"].get("SEASON")}),
            "status": "Files skipped" if failed else ("OK with warnings" if warn_n else "OK"),
            "notes": "; ".join([f"{len(failed)} file(s) skipped"] if failed else [] +
                               ([f"{warn_n} warning(s)"] if warn_n else [])),
        })
    summary = [f"{os.path.basename(xlsx_path)}", f"{n_rows} products from {n_files} file(s), {n_ls} linesheet(s)."]
    if st["saved"] or st["existing"]:
        summary.append(f"Images: {st['saved'] + st['existing']} in the Images folder.")
    if failed:
        summary.append(f"{len(failed)} file(s) could not be processed - see the log.")
    if warn_n:
        summary.append(f"{warn_n} warning(s) - see the log.")
    return "\n".join(summary), out_dir


# ======================================================================
#  IMAGE RENAMING TAB - only rename (and resize) a folder of images, using a product file
# ======================================================================

def _processed_products(path, L):
    """A file that is already in NuORDER layout (processed by this client, or a NuORDER data export):
    read the products straight from it. Returns a Result, or None if the file is not in that layout."""
    try:
        sheets = list_sheets(path, visible_only=True) or list_sheets(path)
    except Exception:
        return None
    res = None
    for sh in sheets:
        rows = read_preview(path, sh, max_rows=10 ** 6)
        if not rows:
            continue
        h = detect_header_row(rows)
        labels = make_labels(rows[h - 1])
        norms = [norm_header(re.sub(r"\s*\(\d+\)$", "", l)) for l in labels]
        body = rows[h:]
        bi = norms.index("brand id") if "brand id" in norms else None
        filled = bi is not None and sum(1 for r_ in body if bi < len(r_) and not is_blank(r_[bi])) >= max(1, len(body) // 2)
        if "style number" not in norms or not (filled or (bi is None and "linesheet name" in norms)):
            if res is None and sh == sheets[0]:
                return None            # e.g. an empty NuORDER template: the brand rules build its BRAND_IDs
            continue

        def col(*names):
            return next((norms.index(n_) for n_ in names if n_ in norms), None)
        c = {"BRAND_ID": col("brand id"), "STYLE NUMBER": col("style number"), "COLOR": col("color", "colour"),
             "COLOR CODE": col("color code", "colour code", "nrf color code"), "NAME": col("name"),
             "MODEL": col("model"), "FABRIC CODE": col("fabric code"),
             "Style": col("customer style number", "updated style number"),
             "BRAND": col("brand")}
        armani = c["MODEL"] is not None and c["FABRIC CODE"] is not None and c["COLOR CODE"] is not None
        if res is None:
            res = Result(path, "NuORDER layout (processed file / data export)")
        for row in rows[h:]:
            vals = {k: (row[i] if i is not None and i < len(row) else None) for k, i in c.items()}
            vals = {k: (clean_code(v) if k in ("STYLE NUMBER", "COLOR CODE", "MODEL", "FABRIC CODE") else clean_text(v))
                    for k, v in vals.items() if not is_blank(v)}
            if not vals.get("STYLE NUMBER") and not vals.get("BRAND_ID"):
                continue
            pic = None
            if armani and all(vals.get(k) for k in ("MODEL", "FABRIC CODE", "COLOR CODE")):
                pic = "_".join([vals["MODEL"], vals["FABRIC CODE"], pad_color_code(vals["COLOR CODE"])])
            if not res.brand and vals.get("BRAND"):
                res.brand = vals["BRAND"]
            res.rows.append({"values": vals, "sizes": [], "upcs": [], "src_rows": [], "pic_name": pic})
    if res is not None:
        L(f"{os.path.basename(path)}: NuORDER layout - {len(res.rows)} product(s) read. "
          "Images are named by " + ("MODEL_FABRIC CODE_COLOR CODE" if any(r["pic_name"] for r in res.rows)
                                     else "BRAND_ID" if any(r["values"].get("BRAND_ID") for r in res.rows)
                                     else "STYLE NUMBER") + ".")
    return res


def products_for_renaming(path, profiles, L, stop):
    """Products (with the names their images get) from a processed file OR a raw brand file.
    Nothing is saved - a raw file is cleaned in memory only."""
    check_extension(path)
    label, prof, _sheet = identify_file(path, profiles)
    if not ((prof and prof != "NuORDER Default Data") or label == "JOOR export"):
        res = _processed_products(path, L)
        if res is not None:
            return [res]
    st = auto_file_settings(path, profiles)
    if not st.get("profile_name"):
        if st.get("error"):
            raise UserError(st["error"])
        if not st.get("labels") or st.get("mapping", {}).get("STYLE NUMBER") is None:
            raise UserError("no Style Number column was found (use the 'Clean Data & Images' tab for this file).")
    L(f"{os.path.basename(path)}: detected as {st['label']}"
      + (f" ({1 + len(st.get('more_sheets') or [])} sheets)" if st.get("more_sheets") else ""))
    opts = dict(download=False, extract_pictures=False, skip_existing=True)
    ctx = {"out_dir": None, "img_dir": None, "skip_existing": True, "pic_hashes": {}, "pic_count": {},
           "pic_stats": {"saved": 0, "existing": 0, "repeat": 0}}
    out = []
    quiet = lambda m: None  # noqa: E731 - the cleaning details are not needed here
    for job in file_jobs(st, opts, profiles):
        job["download"] = False
        job["extract_pictures"] = False
        res = (profile_result if job.get("profile") else generic_result)(job, ctx, quiet, lambda *a: None, stop)
        out.append(res)
    L(f"   {sum(len(r.rows) for r in out)} product(s) found.")
    return out


def run_rename(cfg, log, progress, stop):
    lines = []

    def L(msg):
        lines.append(msg)
        log(msg)

    img_folder = cfg["images_dir"]
    out_dir = os.path.join(img_folder, "Renamed Images")
    os.makedirs(out_dir, exist_ok=True)
    L(f"=== {APP_NAME} {APP_VERSION} - Image Renaming ===")
    L(f"Images folder : {img_folder}" + ("  (and its sub-folders)" if cfg.get("recurse", True) else ""))
    L(f"Renamed copies: {out_dir}")
    if cfg.get("resize_px"):
        L(f"Images are resized to {cfg['resize_px']} px on the shortest side and saved as JPG"
          + ("" if pillow_available() else "  (NOT possible: the Pillow library is missing)") + ".")
    L("")
    profiles, _problems = load_profiles()
    results, failed = [], []
    for k, path in enumerate(cfg["data_paths"], 1):
        if stop.is_set():
            raise UserError("Stopped by user.")
        progress(int(20 * k / len(cfg["data_paths"])), 100)
        try:
            results.extend(products_for_renaming(path, profiles, L, stop))
        except UserError as e:
            if stop.is_set():
                raise
            failed.append(path)
            L(f"SKIPPED {os.path.basename(path)}: {e}")
        except Exception:  # noqa
            failed.append(path)
            L(f"SKIPPED {os.path.basename(path)} - unexpected error:\n" + traceback.format_exc())
    if not any(r.rows for r in results):
        raise UserError("No products could be read from the file(s) - see the log.")
    ctx = {"out_dir": out_dir, "img_dir": out_dir, "skip_existing": cfg.get("skip_existing", True),
           "pic_hashes": {}, "pic_count": {}, "pic_stats": {"saved": 0, "existing": 0, "repeat": 0},
           "resize_px": cfg.get("resize_px")}
    stats = rename_images(results, img_folder, cfg.get("recurse", True), ctx, L, stop,
                          lambda d, t: progress(20 + int(80 * d / max(t, 1)), 100)) or {}
    rl = resize_log_line(ctx)
    if rl:
        L(rl)
    progress(100, 100)
    log_path = os.path.join(out_dir, "Renaming_Log.txt")
    try:
        with open(log_path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines))
    except OSError:
        log_path = None
    n_products = sum(len(r.rows) for r in results)
    brands = list(dict.fromkeys(r.brand for r in results if r.brand))
    if cfg.get("record_history", True):
        add_history({
            "id": datetime.datetime.now().strftime("%Y%m%d%H%M%S%f"), "kind": "rename",
            "time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "brand": ", ".join(brands),
            "files": [{"name": os.path.basename(p_), "type": "Image renaming"} for p_ in cfg["data_paths"]],
            "output": out_dir, "folder": out_dir, "log": log_path,
            "products": n_products, "linesheets": "", "images": stats.get("copied", 0) + ctx["pic_stats"]["existing"],
            "seasons": [], "status": "Image renaming" + (" - files skipped" if failed else ""),
            "notes": f"{stats.get('copied', 0)} image(s) renamed, {stats.get('skipped', 0)} skipped as unclear, "
                     f"{stats.get('images', 0)} found",
        })
    summary = [f"{stats.get('copied', 0)} image(s) renamed into 'Renamed Images'.",
               f"{stats.get('images', 0)} image(s) found, {n_products} product(s) read."]
    if ctx["pic_stats"]["existing"]:
        summary.append(f"{ctx['pic_stats']['existing']} were already there (skipped).")
    if stats.get("skipped"):
        summary.append(f"{stats['skipped']} image(s) could not be matched with certainty - see the log.")
    if failed:
        summary.append(f"{len(failed)} file(s) could not be read - see the log.")
    return "\n".join(summary), out_dir


# ======================================================================
#  MAXMARA PDF IMAGES - pull the product pictures out of the brand's PDF catalogs
#  (logic by Randy: the product code "MML..." is found as text, and the picture just below it is saved as CODE.jpg)
# ======================================================================

def _load_pymupdf():
    try:
        import pymupdf as fitz_      # new name
        return fitz_
    except ImportError:
        import fitz as fitz_         # old name
        return fitz_


def pymupdf_available():
    try:
        _load_pymupdf()
        return True
    except Exception:
        return False


def _pdf_pick_image(fitz, page, images, position):
    """Randy's rule: the one picture that sits just below the product code (20-60 pt) and slightly to its left."""
    found = []
    for image in images:
        box = fitz.Rect(image["bbox"]) * page.rotation_matrix
        dx = (box.x0 + box.x1 - position.x0 - position.x1) / 2
        dy = box.y0 - position.y0
        if 20 < dy < 60 and -100 < dx < 10 and box.width > 45:
            found.append(image)
    return found


def _pdf_lines(words, gap=1.5):
    """[(rect, text)] -> list of text lines (words with nearly the same top edge belong together)."""
    lines, cur, last = [], [], None
    for r, t in sorted(words, key=lambda w: (round(w[0].y0, 1), w[0].x0)):
        if last is not None and r.y0 - last > gap:
            lines.append(cur)
            cur = []
        cur.append((r, t))
        last = r.y0
    if cur:
        lines.append(cur)
    return [" ".join(t for _r, t in sorted(l, key=lambda w: w[0].x0)) for l in lines]


def _fabric_from_lines(lines):
    """Keeps the composition lines (the ones with a %), glues a wrapped word back on, drops repeats."""
    out = []
    for ln in lines:
        if "%" in ln:
            if ln not in out:
                out.append(ln)
        elif out and re.search(r"[%,]\s*$", out[-1]):
            out[-1] = out[-1] + " " + ln
    txt = "; ".join(out)
    return re.sub(r"\s*:\s*", ": ", txt).strip()


def _pdf_catalog_cells(fitz, page):
    """Reads every product box of a MaxMara catalog page: model, article, fabric, sizes per color."""
    words = [(fitz.Rect(w[:4]) * page.rotation_matrix, w[4]) for w in page.get_text("words")]
    anchors = [(r, t) for r, t in words if re.fullmatch(r"\d{10,14}", t)]
    cols = [(r, t) for r, t in words if t == "Col."]
    cells = []
    for ar, root in anchors:
        mine = [(r, t) for r, t in cols if r.y0 > ar.y0 and 0 <= ar.x0 - r.x0 < 150]
        if not mine:
            continue
        colr = min(mine, key=lambda w: w[0].y0 - ar.y0)[0]
        cells.append(dict(root=root, a=ar, col=colr, left=colr.x0 - 4))
    page_w = page.rect.width if page.rotation in (0, 180) else page.rect.height
    page_h = page.rect.height if page.rotation in (0, 180) else page.rect.width
    page_w, page_h = max(page_w, page.rect.width), max(page_h, page.rect.height)
    out = []
    for c in cells:
        row = sorted([o for o in cells if abs(o["a"].y0 - c["a"].y0) < 6], key=lambda o: o["left"])
        nxt = [o["left"] for o in row if o["left"] > c["left"] + 20]
        right = min(nxt) if nxt else page_w
        below = [o["a"].y0 for o in cells if o["a"].y0 > c["a"].y0 + 50 and abs(o["left"] - c["left"]) < 30]
        top, bottom = c["a"].y0 - 6, (min(below) - 6 if below else page_h)
        inside = [(r, t) for r, t in words if c["left"] <= r.x0 < right and top <= r.y0 < bottom]
        info = dict(root=c["root"], model="", art="", fabric="", sizes=[], variants={})
        for r, t in inside:
            if re.fullmatch(r"MML[A-Z0-9]+", t) and abs(r.y0 - c["a"].y0) < 4:
                info["model"] = t
        arts = [(r, t) for r, t in inside if t == "Art." and abs(r.y0 - c["a"].y0) < 4]
        if arts:
            ar_ = arts[0][0]
            nxt_w = [(r, t) for r, t in inside if abs(r.y0 - ar_.y0) < 4 and r.x0 > ar_.x1]
            if nxt_w:
                info["art"] = min(nxt_w, key=lambda w: w[0].x0)[1]
        sell = [r for r, t in inside if t == "Sell-in"]
        rcol = min(r.x0 for r in sell) if sell else c["left"] + 0.55 * (right - c["left"])
        block = [(r, t) for r, t in inside if c["a"].y0 + 6 < r.y0 < c["a"].y0 + 100 and r.x0 < rcol - 2]
        info["fabric"] = _fabric_from_lines(_pdf_lines(block))
        # size grid
        cy = c["col"].y0
        tot = [r for r, t in inside if t == "Totali" and abs(r.y0 - cy) < 3]
        tot_x = min(r.x0 for r in tot) if tot else right
        head = sorted([(r, t) for r, t in inside if abs(r.y0 - cy) < 3 and c["col"].x1 <= r.x0 < tot_x - 2],
                      key=lambda w: w[0].x0)
        info["sizes"] = [t for _r, t in head]
        for r, t in inside:
            if r.y0 > cy + 3 and abs(r.x0 - c["col"].x0) < 8 and re.fullmatch(r"[0-9A-Z]{2,4}", t):
                qty = [(r2, t2) for r2, t2 in inside if abs(r2.y0 - r.y0) < 3 and r2.x0 > r.x1 and r2.x0 < tot_x - 2]
                have = []
                for r2, t2 in qty:
                    if not re.fullmatch(r"\d+", t2) or int(t2) <= 0 or not head:
                        continue
                    best = min(head, key=lambda h_: abs((h_[0].x0 + h_[0].x1) / 2 - (r2.x0 + r2.x1) / 2))
                    have.append(best[1])
                info["variants"][t] = [x for x in info["sizes"] if x in have]
        out.append(info)
    return out


def read_maxmara_catalogs(pdfs, L, stop, progress=None):
    """-> (products, images): products[model] = {art, fabric, sizes, variants{code: [sizes]}},
    images[model] = JPG bytes of its picture.  The first appearance of a model wins."""
    fitz = _load_pymupdf()
    from PIL import Image
    products, images = {}, {}
    for fi, source in enumerate(pdfs, 1):
        name = os.path.basename(source)
        try:
            with fitz.open(source) as doc:
                if doc.needs_pass:
                    raise ValueError("the PDF is protected with a password.")
                for index, page in enumerate(doc):
                    if stop.is_set():
                        raise UserError("Stopped by user.")
                    if progress:
                        progress(int(10 * ((fi - 1) + (index + 1) / max(len(doc), 1)) / len(pdfs)), 100)
                    for cell in _pdf_catalog_cells(fitz, page):
                        m_ = cell["model"]
                        if not m_:
                            continue
                        if m_ not in products:
                            products[m_] = cell
                        else:                       # the same model on another page: add its other colors
                            for v_, sz_ in cell["variants"].items():
                                products[m_]["variants"].setdefault(v_, sz_)
                            for f_ in ("art", "fabric"):
                                if not products[m_][f_] and cell[f_]:
                                    products[m_][f_] = cell[f_]
                    names = [(w[4], fitz.Rect(w[:4]) * page.rotation_matrix)
                             for w in page.get_text("words") if re.fullmatch(r"MML[A-Z0-9]+", w[4])]
                    imgs = page.get_image_info(xrefs=True)
                    for code, position in names:
                        if code in images:
                            continue
                        found = _pdf_pick_image(fitz, page, imgs, position)
                        if len(found) != 1 or not found[0]["xref"]:
                            L(f"  CHECK {code} ({name}, page {index + 1}): could not identify one single picture.")
                            continue
                        try:
                            data = doc.extract_image(found[0]["xref"])["image"]
                            with Image.open(io.BytesIO(data)) as original:
                                buf = io.BytesIO()
                                original.convert("RGB").save(buf, format="JPEG", quality=95, subsampling=0)
                            images[code] = buf.getvalue()
                        except Exception as e:  # noqa
                            L(f"  ERROR {code} ({name}, page {index + 1}): {e}")
        except UserError:
            raise
        except Exception as e:  # noqa
            L(f"ERROR in {name}: {e}")
    return products, images


def extract_pdf_images(pdfs, dest, L, stop, progress=None, resize_px=None):
    """Save one JPG per MML code found in the PDFs. The first appearance of a code wins and existing files are never
    overwritten. Returns a dict of counters."""
    fitz = _load_pymupdf()
    stats = dict(saved=0, duplicates=0, existing=0, warnings=0, pages=0, review=[])
    seen = set()
    os.makedirs(dest, exist_ok=True)
    for fi, source in enumerate(pdfs, 1):
        if stop.is_set():
            raise UserError("Stopped by user.")
        name = os.path.basename(source)
        try:
            with fitz.open(source) as document:
                if document.needs_pass:
                    raise ValueError("the PDF is protected with a password.")
                total = len(document)
                for index, page in enumerate(document):
                    if stop.is_set():
                        raise UserError("Stopped by user.")
                    stats["pages"] += 1
                    if progress:
                        progress(int(100 * ((fi - 1) + (index + 1) / max(total, 1)) / len(pdfs)), 100)
                    names = [(w[4], fitz.Rect(w[:4]) * page.rotation_matrix)
                             for w in page.get_text("words") if re.fullmatch(r"MML[A-Z0-9]+", w[4])]
                    names.sort(key=lambda item: (round(item[1].y0), item[1].x0))
                    images = page.get_image_info(xrefs=True)
                    if not names:
                        stats["warnings"] += 1
                        L(f"{name} - page {index + 1}/{total}: no MML codes found on this page.")
                        continue
                    for code, position in names:
                        if code in seen:
                            stats["duplicates"] += 1
                            continue
                        seen.add(code)                       # the first appearance has priority
                        target = os.path.join(dest, f"{code}.jpg")
                        if os.path.exists(target):
                            stats["existing"] += 1
                            L(f"  {code}: already exists - kept.")
                            continue
                        candidates = _pdf_pick_image(fitz, page, images, position)
                        if len(candidates) != 1 or not candidates[0]["xref"]:
                            stats["warnings"] += 1
                            stats["review"].append(f"{code}  ({name}, page {index + 1})")
                            L(f"  CHECK {code} ({name}, page {index + 1}): could not identify one single picture.")
                            continue
                        try:
                            data = document.extract_image(candidates[0]["xref"])["image"]
                            from PIL import Image
                            with Image.open(io.BytesIO(data)) as original:
                                buf = io.BytesIO()
                                original.convert("RGB").save(buf, format="JPEG", quality=95, subsampling=0)
                            try:
                                with open(target, "xb") as out:
                                    out.write(buf.getvalue())
                            except FileExistsError:
                                stats["existing"] += 1
                                continue
                            stats["saved"] += 1
                            L(f"  saved {code}.jpg")
                            if resize_px and pillow_available():
                                resize_to_jpg(target, target, resize_px)
                        except UserError:
                            raise
                        except Exception as e:  # noqa
                            stats["warnings"] += 1
                            stats["review"].append(f"{code}  ({name}, page {index + 1}) - {e}")
                            L(f"  ERROR {code}: {e}")
        except UserError:
            raise
        except Exception as e:  # noqa
            stats["warnings"] += 1
            L(f"ERROR in {name}: {e}")
    return stats


def run_pdf_images(cfg, log, progress, stop):
    lines = []

    def L(msg):
        lines.append(msg)
        log(msg)

    dest = cfg["dest"]
    L(f"=== {APP_NAME} {APP_VERSION} - MaxMara PDF images ===")
    L(f"PDF files : {len(cfg['pdfs'])}")
    L(f"Images go to: {dest}")
    L("")
    if not pymupdf_available():
        raise UserError("The PDF library (PyMuPDF) is missing. Run Build_NuORDER_Client.bat again "
                        "(or:  py -m pip install pymupdf pillow).")
    if not pillow_available():
        raise UserError("The Pillow library is missing. Run Build_NuORDER_Client.bat again.")
    st = extract_pdf_images(cfg["pdfs"], dest, L, stop, progress, cfg.get("resize_px"))
    L("")
    L(f"Saved: {st['saved']}   Repeated codes skipped: {st['duplicates']}   Already there: {st['existing']}   "
      f"Warnings: {st['warnings']}")
    if st["review"]:
        L("")
        L(f"Codes to check by hand ({len(st['review'])}):")
        for r_ in st["review"]:
            L("  - " + r_)
    log_path = os.path.join(dest, "PDF_Images_Log.txt")
    try:
        with open(log_path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines))
    except OSError:
        pass
    progress(100, 100)
    summary = [f"{st['saved']} image(s) saved as CODE.jpg."]
    if st["existing"]:
        summary.append(f"{st['existing']} already existed (kept).")
    if st["duplicates"]:
        summary.append(f"{st['duplicates']} repeated code(s) skipped.")
    if st["warnings"]:
        summary.append(f"{st['warnings']} warning(s) - see the log.")
    return "\n".join(summary), dest


# ======================================================================
#  LINESHEET UPDATER - apply a brand's update file to a NuORDER data export / processed file
# ======================================================================

UPDATE_SCOPES = OrderedDict([
    ("Everything found", "all"),
    ("Pricing only", "pricing"),
    ("Style Numbers", "style"),
    ("Names", "name"),
    ("Season", "season"),
    ("Dates", "dates"),
    ("Sizes", "sizes"),
    ("Update UPCs", "upcs"),
    ("Pick columns...", "custom"),
])

_UPC_NAMES = {"upc", "upcs", "upc code", "upc number", "ean", "ean code", "ean13", "ean 13", "barcode", "bar code",
              "barcodes", "gtin", "upc ean", "ean upc", "new upc", "updated upc", "upd upc", "new ean", "upc 12"}
_SIZE_NAMES = {"size", "sizes", "size code", "us size", "size description", "size name", "size label", "taglia",
               "talla", "size us", "eu size", "size value"}


def _upc_number(name):
    """'upc 3' / 'ean 3' -> 3 (the numbered UPC columns), else None."""
    m = re.fullmatch(r"(?:upc|ean|barcode|gtin)\s*(\d+)", name)
    return int(m.group(1)) if m else None


def _size_key(v):
    """Sizes compared loosely: ' m ' = 'M', '8.0' = 8 = '8', '10 1/2' = '10 1/2'."""
    if is_blank(v):
        return ""
    n = _as_number(v)
    if n is not None:
        return ("%g" % n)
    return re.sub(r"\s+", "", clean_code(v).upper())


def _upc_text(v):
    t = clean_code(v)
    if isinstance(v, (int, float)) and t.isdigit() and len(t) == 11:
        t = t.zfill(12)                     # a UPC typed as a number lost its leading 0
    return t


_PROTECTED = {"id", "sort", "created on", "media", "banners", "archived", "disabled", "cancelled"}
_STYLE_HEADERS = ["style number", "style", "style no", "style #", "vpn", "sku", "model number", "style code",
                  "item number", "article"]
_UPD_GROUPS = [
    # (data column names, update column names) - first data name found in the data file is the target
    (["wholesale usd", "wholesale", "wholesale price"],
     ["wholesale usd", "wholesale", "wholesale price", "wsp", "wsp usd", "cost", "new cost", "whsl", "whsl usd",
      "new wholesale", "usd wholesale", "cost price", "new wsp", "updated wholesale", "revised wholesale"]),
    (["retail usd", "retail", "retail price"],
     ["retail usd", "retail", "retail price", "msrp", "m s r p", "new retail", "rrp", "new msrp", "us msrp",
      "suggested retail price", "usd retail", "updated retail", "revised retail", "new rrp"]),
    (["style number", "style"],
     ["upd vpn", "new vpn", "updated vpn", "new style number", "new style", "new sku", "upd style",
      "updated style number"]),
    (["name"], ["name", "style name", "product name", "new name", "updated name"]),
    (["season"], ["season", "new season", "updated season"]),
    (["available from"], ["available from", "start ship date", "delivery from", "ship start", "new available from"]),
    (["available until"], ["available until", "end ship date", "delivery to", "ship end", "new available until"]),
]


def _sq(v):
    return re.sub(r"[^A-Z0-9]", "", clean_code(v).upper()) if not is_blank(v) else ""


def _scope_ok(scope, target_norm, custom):
    if scope == "all":
        return True
    if scope == "custom":
        return target_norm in custom
    if scope == "pricing":
        return any(w in target_norm for w in ("wholesale", "retail", "price", "cost", "msrp", "special cost"))
    if scope == "style":
        return target_norm in ("style number", "style", "brand id", "customer style number")
    if scope == "name":
        return target_norm in ("name", "style name")
    if scope == "season":
        return target_norm in ("season", "seasons")
    if scope == "dates":
        return target_norm in ("available from", "available until", "order closing", "delivery month")
    if scope == "sizes":
        return bool(re.fullmatch(r"size \d+", target_norm))
    if scope == "upcs":
        return bool(re.fullmatch(r"upc \d+", target_norm))
    return False


_NUMLIKE = re.compile(r"\s*[$€£]?\s*-?[\d.,]+\s*")


def _as_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str) and _NUMLIKE.fullmatch(v):
        return parse_number(v)
    return None


def _values_differ(old, new):
    """True when the update really changes the value (333 = 333.0 = '333'; dates compared as dates)."""
    if is_blank(old) and is_blank(new):
        return False
    if isinstance(old, (datetime.date, datetime.datetime)) or isinstance(new, (datetime.date, datetime.datetime)):
        da = parse_dates(old, datetime.date.today().year)
        db = parse_dates(new, datetime.date.today().year)
        return (da[:1] or [None]) != (db[:1] or [None])
    a, b = _as_number(old), _as_number(new)
    if a is not None and b is not None:
        return abs(a - b) > 1e-9
    return clean_text(old) != clean_text(new)


def data_file_columns(path):
    """Column names of a data file (for the 'Pick columns' list in the window)."""
    sheets = list_sheets(path, visible_only=True) or list_sheets(path)
    rows = read_preview(path, sheets[0], max_rows=30)
    h = detect_header_row(rows) if rows else 1
    return make_labels(rows[h - 1]) if rows else []


def run_update(cfg, log, progress, stop):
    data_path = cfg["data_path"]
    scope = cfg.get("scope", "all")
    custom = {norm_header(c) for c in cfg.get("custom_cols", [])}
    lines = []

    def L(msg):
        lines.append(msg)
        log(msg)

    L(f"=== {APP_NAME} {APP_VERSION} - Linesheet Updater ===")
    L(f"Data file   : {data_path}")
    L(f"What to update: {next(k for k, v in UPDATE_SCOPES.items() if v == scope)}")
    progress(0, 100)

    # ---------- data file ----------
    ext = check_extension(data_path)
    sheets = list_sheets(data_path, visible_only=True) or list_sheets(data_path)
    sheet = sheets[0]
    preview = read_preview(data_path, sheet, max_rows=30)
    hr = detect_header_row(preview) if preview else 1
    labels = make_labels(preview[hr - 1])
    norms = [norm_header(re.sub(r"\s*\(\d+\)$", "", l)) for l in labels]
    if ext in (".xlsx", ".xlsm"):
        wb = load_workbook(data_path, keep_vba=(ext == ".xlsm"))
        ws = wb[sheet]
        ws_vals = load_workbook(data_path, data_only=True, read_only=False)[sheet]
    else:
        src = open_sheet(data_path, sheet)
        wb = Workbook()
        ws = wb.active
        ws.title = sheet[:31] or "Sheet1"
        for r in range(1, src.max_row + 1):
            for c in range(1, src.max_column + 1):
                v = src.cell(r, c).value
                if v not in (None, ""):
                    ws.cell(row=r, column=c, value=v)
        ws_vals = ws
    last_row = last_data_row(ws_vals)

    def col_of(names):
        for n_ in names:
            if n_ in norms:
                return norms.index(n_)
        return None

    c_style = col_of(_STYLE_HEADERS)
    c_brand = col_of(["brand id", "brandid"])
    c_color = col_of(["color", "colour", "color name"])
    c_ccode = col_of(["color code", "colour code", "nrf color code"])
    c_fab = col_of(["fab code", "fabric code", "material code", "material"])
    c_cust = col_of(["customer style number", "updated style number"])
    d_size = {int(m.group(1)): i for i, n_ in enumerate(norms) for m in [re.fullmatch(r"size (\d+)", n_)] if m}
    d_upc = {int(m.group(1)): i for i, n_ in enumerate(norms) for m in [re.fullmatch(r"upc (\d+)", n_)] if m}
    upc_on = scope in ("all", "upcs")
    if scope == "upcs" and not d_upc:
        raise UserError("The data file has no UPC 1, UPC 2 ... columns, so there is nowhere to put the UPCs.")
    if c_style is None and c_brand is None:
        raise UserError("The data file has no Style Number or brand_id column.")
    L(f"Sheet '{sheet}', header on row {hr}, {last_row - hr} product rows. "
      f"Matching on: {', '.join(labels[c] for c in (c_style, c_brand, c_fab, c_ccode, c_cust) if c is not None)}")

    full_idx, group_idx = {}, {}
    data_rows = []
    for r in range(hr + 1, last_row + 1):
        def v(c):
            return ws_vals.cell(r, c + 1).value if c is not None else None
        style, bid, cc, fab, cust = _sq(v(c_style)), str(v(c_brand) or ""), _sq(v(c_ccode)), _sq(v(c_fab)), _sq(v(c_cust))
        if not (style or bid):
            continue
        data_rows.append(r)
        full, group = set(), set()
        if bid:
            pre = _sq(bid.split("_")[0])
            full |= {_sq(bid), pre}
            if cc and pre.endswith(cc) and len(pre) > len(cc):
                group.add(pre[:-len(cc)])                       # style + material, no color
        if style:
            group.add(style)
            if cc:
                full.add(style + cc)
            if fab:
                group.add(style + fab)
                if cc:
                    full.add(style + fab + cc)
        if cust:
            full.add(cust)
        for k in full:
            if len(k) >= 3:
                full_idx.setdefault(k, set()).add(r)
        for k in group - full:
            if len(k) >= 3:
                group_idx.setdefault(k, set()).add(r)

    def lookup(key):
        if key in full_idx:
            return full_idx[key], "exact"
        if key in group_idx:
            return group_idx[key], "style (all colors)"
        return set(), None

    # ---------- update files ----------
    changes, not_found, unclear, pairing, conflicts = [], [], [], [], []
    upc_unplaced = []
    text_cells = set()
    changed_cells = {}
    covered = set()
    matched_rows_total = 0
    yellow = PatternFill("solid", fgColor="FFFF00")
    upd_files = cfg["update_paths"]
    for fi, up in enumerate(upd_files, 1):
        try:
            usheets = list_sheets(up, visible_only=True) or list_sheets(up)
        except Exception as e:  # noqa
            L(f"SKIPPED {os.path.basename(up)}: {e}")
            continue
        for ush in usheets:
            if stop.is_set():
                raise UserError("Stopped by user. Nothing was saved.")
            urows = read_preview(up, ush, max_rows=10 ** 6)
            if not urows:
                continue
            uh = detect_header_row(urows)
            ulabels = make_labels(urows[uh - 1])
            unorms = [norm_header(re.sub(r"\s*\(\d+\)$", "", l)) for l in ulabels]
            body = urows[uh:]
            # which column holds the style numbers? the one whose values match the data file best
            best, best_hits = None, 0
            for i in range(len(ulabels)):
                hits = sum(1 for row in body[:3000] if i < len(row) and lookup(_sq(row[i]))[0])
                bonus = 0.5 if unorms[i] in _STYLE_HEADERS or unorms[i] in ("style number", "brand id") else 0
                if hits + bonus > best_hits:
                    best, best_hits = i, hits + bonus
            tag = f"{os.path.basename(up)}" + (f" / {ush}" if len(usheets) > 1 else "")
            if best is None or best_hits < 1:
                L(f"--- {tag}: no style numbers matching the data file were found - skipped.")
                continue
            u_color = next((i for i, n_ in enumerate(unorms) if n_ in ("color", "colour", "color name") and i != best), None)
            u_ccode = next((i for i, n_ in enumerate(unorms) if n_ in ("color code", "colour code") and i != best), None)
            # UPCs: one UPC column (+ a size column), or numbered UPC 1 / SIZE 1 ... columns
            u_upc = next((i for i, n_ in enumerate(unorms) if n_ in _UPC_NAMES and i != best), None)
            u_upc_n = {_upc_number(n_): i for i, n_ in enumerate(unorms) if _upc_number(n_) and i != best}
            u_size = next((i for i, n_ in enumerate(unorms) if n_ in _SIZE_NAMES and i != best), None)
            u_size_n = {int(m.group(1)): i for i, n_ in enumerate(unorms) for m in [re.fullmatch(r"size (\d+)", n_)] if m}
            upc_here = upc_on and bool(d_upc) and (u_upc is not None or bool(u_upc_n))
            upc_cols = ({u_upc} if u_upc is not None else set()) | set(u_upc_n.values())
            # pair update columns with data columns
            targets = {}
            for i, un in enumerate(unorms):
                if i == best or not un:
                    continue
                if i in upc_cols or un in _UPC_NAMES or _upc_number(un):
                    pairing.append((tag, ulabels[i], "UPC n  (placed by size)" if upc_here else "(not used)"))
                    continue
                if upc_here and scope == "upcs":
                    continue
                t = None
                for dnames, unames in _UPD_GROUPS:
                    if un in unames:
                        t = col_of(dnames)
                        if t is not None:
                            break
                if t is None and un in norms:
                    t = norms.index(un)
                if t is None or norms[t] in _PROTECTED or t in (c_brand,):
                    pairing.append((tag, ulabels[i], "(not used)"))
                    continue
                if not _scope_ok(scope, norms[t], custom):
                    pairing.append((tag, ulabels[i], f"{labels[t]}  (skipped: outside '{scope}')"))
                    continue
                newish = bool(re.search(r"\b(new|upd|updated|revised)\b", un))
                prev = targets.get(t)
                if prev is None or (newish and not prev[1]):
                    if prev is not None:
                        pairing.append((tag, ulabels[prev[0]], f"{labels[t]}  (replaced by '{ulabels[i]}')"))
                    targets[t] = (i, newish)
                else:
                    pairing.append((tag, ulabels[i], f"{labels[t]}  (not used: '{ulabels[prev[0]]}' wins)"))
            for t, (i, _n) in targets.items():
                pairing.append((tag, ulabels[i], labels[t]))
            upc_desc = []
            if upc_here:
                upc_desc = [(f"{ulabels[u_upc]}" + (f" + size from '{ulabels[u_size]}'" if u_size is not None else ""))
                            if u_upc is not None else f"{len(u_upc_n)} numbered UPC column(s)"]
                upc_desc = [upc_desc[0] + " -> UPC 1, UPC 2 ... (by size)"]
            L(f"--- {tag}: style numbers in '{ulabels[best]}', {len(body)} rows; updating: "
              + (", ".join([f"{ulabels[i]} -> {labels[t]}" for t, (i, _n) in targets.items()] + upc_desc)
                 or "nothing (no matching columns)"))
            for k, row in enumerate(body, uh + 1):
                if best >= len(row) or is_blank(row[best]):
                    continue
                rows_hit, how = lookup(_sq(row[best]))
                if not rows_hit:
                    not_found.append((tag, k, clean_code(row[best])))
                    continue
                if len(rows_hit) > 1 and (u_color is not None or u_ccode is not None):
                    want_c = _sq(row[u_color]) if u_color is not None and u_color < len(row) else ""
                    want_cc = _sq(row[u_ccode]) if u_ccode is not None and u_ccode < len(row) else ""
                    narrowed = {r for r in rows_hit
                                if (want_cc and c_ccode is not None and _sq(ws_vals.cell(r, c_ccode + 1).value) == want_cc)
                                or (want_c and c_color is not None and _sq(ws_vals.cell(r, c_color + 1).value) == want_c)}
                    if narrowed:
                        rows_hit = narrowed
                        how = "style + color" if how != "exact" else how
                matched_rows_total += 1
                if upc_here:
                    pairs = []
                    if u_upc is not None and u_upc < len(row) and not is_blank(row[u_upc]):
                        pairs.append((row[u_size] if u_size is not None and u_size < len(row) else None, row[u_upc], None))
                    for n_, ci in u_upc_n.items():
                        if ci < len(row) and not is_blank(row[ci]):
                            si = u_size_n.get(n_)
                            pairs.append((row[si] if si is not None and si < len(row) else None, row[ci], n_))
                    sty = clean_code(row[best])
                    for size_v, upc_v, pos in pairs:
                        upc = _upc_text(upc_v)
                        if len(rows_hit) != 1:
                            upc_unplaced.append((tag, k, sty, clean_code(size_v), upc,
                                                 f"the style matches {len(rows_hit)} colors - the update file needs a color column"))
                            continue
                        r = next(iter(rows_hit))
                        n_ = None
                        if not is_blank(size_v):
                            want = _size_key(size_v)
                            n_ = next((n for n, ci in sorted(d_size.items())
                                       if _size_key(ws_vals.cell(r, ci + 1).value) == want), None)
                            if n_ is None:
                                upc_unplaced.append((tag, k, sty, clean_code(size_v), upc,
                                                     "this size is not in the data file for this product"))
                                continue
                        elif pos is not None:
                            n_ = pos
                        else:
                            have = [n for n, ci in d_size.items() if not is_blank(ws_vals.cell(r, ci + 1).value)]
                            if len(have) == 1:
                                n_ = have[0]
                            else:
                                upc_unplaced.append((tag, k, sty, "", upc, "no size given in the update file"))
                                continue
                        t = d_upc.get(n_)
                        if t is None:
                            upc_unplaced.append((tag, k, sty, clean_code(size_v), upc, f"the data file has no UPC {n_} column"))
                            continue
                        covered.add(r)
                        old = ws_vals.cell(r, t + 1).value
                        if not _values_differ(old, upc):
                            continue
                        if (r, t) in changed_cells:
                            if _values_differ(changed_cells[(r, t)], upc):
                                conflicts.append((tag, k, sty, labels[t], changed_cells[(r, t)], upc))
                            continue
                        cell = ws.cell(r, t + 1)
                        cell.value = upc
                        cell.number_format = "@"
                        cell.fill = yellow
                        changed_cells[(r, t)] = upc
                        size_txt = clean_code(ws_vals.cell(r, d_size[n_] + 1).value) if n_ in d_size else ""
                        changes.append((
                            clean_code(ws_vals.cell(r, c_style + 1).value) if c_style is not None else "",
                            clean_text(ws_vals.cell(r, c_brand + 1).value) if c_brand is not None else "",
                            clean_text(ws_vals.cell(r, c_color + 1).value) if c_color is not None else "",
                            labels[t] + (f"  (size {size_txt})" if size_txt else ""), old, upc, tag, how, r))
                for r in rows_hit:
                    if scope == "upcs":
                        break
                    covered.add(r)
                    for t, (i, _n) in targets.items():
                        if i >= len(row) or is_blank(row[i]):
                            continue                         # a blank never erases data
                        new = row[i]
                        old = ws_vals.cell(r, t + 1).value
                        if not _values_differ(old, new):
                            continue
                        if (r, t) in changed_cells and _values_differ(changed_cells[(r, t)], new):
                            conflicts.append((tag, k, clean_code(row[best]), labels[t], changed_cells[(r, t)], new))
                            continue
                        if (r, t) in changed_cells:
                            continue
                        num = _as_number(new)
                        if num is not None and (isinstance(old, (int, float)) or is_blank(old)) and \
                                any(w in norms[t] for w in ("wholesale", "retail", "price", "cost", "msrp")):
                            value = int(num) if float(num).is_integer() else num
                        elif isinstance(new, (datetime.date, datetime.datetime)):
                            value = new
                        else:
                            value = clean_code(new) if isinstance(new, (int, float)) else clean_text(new)
                        cell = ws.cell(r, t + 1)
                        cell.value = value
                        cell.fill = yellow
                        changed_cells[(r, t)] = value
                        changes.append((
                            clean_code(ws_vals.cell(r, c_style + 1).value) if c_style is not None else "",
                            clean_text(ws_vals.cell(r, c_brand + 1).value) if c_brand is not None else "",
                            clean_text(ws_vals.cell(r, c_color + 1).value) if c_color is not None else "",
                            labels[t], old, value, tag, how, r))
            progress(int(90 * fi / max(len(upd_files), 1)), 100)

    # ---------- save ----------
    stem = os.path.splitext(os.path.basename(data_path))[0]
    out_dir = cfg.get("out_dir") or os.path.dirname(data_path)
    out_path = os.path.join(out_dir, f"{safe_filename(stem)}_Updated.xlsx")
    log_path = os.path.join(out_dir, f"{safe_filename(stem)}_Changes.xlsx")
    if os.path.abspath(out_path) == os.path.abspath(data_path):
        raise UserError("The updated file would overwrite the data file - choose another folder.")
    try:
        wb.save(out_path)
    except PermissionError:
        raise UserError(f"Could not save '{os.path.basename(out_path)}'. If it is open in Excel, close it and try again.")

    lw = Workbook()
    s1 = lw.active
    s1.title = "Changes"
    s1.append(["STYLE NUMBER", "BRAND_ID", "COLOR", "COLUMN", "OLD VALUE", "NEW VALUE", "UPDATE FILE", "MATCHED BY", "ROW"])
    for ch in changes:
        s1.append(list(ch))
    s2 = lw.create_sheet("Not in data file")
    s2.append(["UPDATE FILE", "ROW", "STYLE NUMBER IN UPDATE FILE"])
    for x in not_found:
        s2.append(list(x))
    s3 = lw.create_sheet("Not in update")
    s3.append(["ROW", "STYLE NUMBER", "BRAND_ID", "COLOR"])
    for r in data_rows:
        if r not in covered:
            s3.append([r] + [clean_text(ws_vals.cell(r, c + 1).value) if c is not None else ""
                             for c in (c_style, c_brand, c_color)])
    s4 = lw.create_sheet("Column pairing")
    s4.append(["UPDATE FILE", "UPDATE COLUMN", "DATA FILE COLUMN"])
    for x in pairing:
        s4.append(list(x))
    if upc_unplaced:
        s6 = lw.create_sheet("UPCs not placed")
        s6.append(["UPDATE FILE", "ROW", "STYLE NUMBER", "SIZE", "UPC", "WHY"])
        for x in upc_unplaced:
            s6.append(list(x))
        for row_ in s6.iter_rows(min_row=2, min_col=5, max_col=5):
            for c_ in row_:
                c_.number_format = "@"
    if conflicts:
        s5 = lw.create_sheet("Conflicts")
        s5.append(["UPDATE FILE", "ROW", "STYLE NUMBER", "COLUMN", "KEPT VALUE", "IGNORED VALUE"])
        for x in conflicts:
            s5.append(list(x))
    for sh in lw.worksheets:
        for c in sh[1]:
            c.font = FONT_HEADER
        auto_fit(sh, max_w=50)
        sh.freeze_panes = "A2"
    try:
        lw.save(log_path)
    except PermissionError:
        L(f"Could not save the change log '{os.path.basename(log_path)}' (is it open in Excel?).")

    n_rows_changed = len({c[-1] for c in changes})
    L("")
    L(f"Products matched: {len(covered)} of {len(data_rows)}   |   cells changed: {len(changes)} "
      f"on {n_rows_changed} product(s) (highlighted yellow)")
    L(f"Update rows not found in the data file: {len(not_found)}"
      + (f"  e.g. {', '.join(x[2] for x in not_found[:8])}" if not_found else ""))
    L(f"Data-file products the update did not mention: {len(data_rows) - len(covered)}")
    if upc_on and d_upc:
        n_upc = sum(1 for c_ in changes if re.match(r"UPC \d+", str(c_[3]), re.I))
        if n_upc or upc_unplaced or scope == "upcs":
            L(f"UPCs updated: {n_upc}" + (f"   |   CHECK: {len(upc_unplaced)} UPC(s) could not be placed "
                                          "(see 'UPCs not placed' in the change log)" if upc_unplaced else ""))
    if conflicts:
        L(f"CHECK: {len(conflicts)} cell(s) got two different values from the update files - the first was kept (see Conflicts).")
    L(f"Saved: {out_path}")
    L(f"Change log: {log_path}")
    progress(100, 100)
    summary = (f"{os.path.basename(out_path)}\n{len(covered)} of {len(data_rows)} products matched, "
               f"{len(changes)} cell(s) changed.\n{len(not_found)} update row(s) not found in the data file.")
    return summary, out_dir


# ----------------------------------------------------------------------
#  Image download
# ----------------------------------------------------------------------

EXT_BY_CTYPE = {
    "image/jpeg": ".jpg", "image/jpg": ".jpg", "image/png": ".png", "image/gif": ".gif",
    "image/webp": ".webp", "image/tiff": ".tif", "image/bmp": ".bmp",
}


def sniff_ext(data):
    if data[:3] == b"\xff\xd8\xff":
        return ".jpg"
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return ".png"
    if data[:4] == b"GIF8":
        return ".gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return ".webp"
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return ".tif"
    if data[:2] == b"BM":
        return ".bmp"
    return None


def fetch_image(url, dest_base, timeout=30, retries=2):
    """Download one image, return the saved path. Raises on failure."""
    last = None
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                              "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
                "Accept": "image/*,*/*;q=0.8"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = resp.read()
                ctype = (resp.headers.get("Content-Type") or "").split(";")[0].strip().lower()
            if not data:
                raise ValueError("empty response")
            ext = sniff_ext(data)
            if ext is None:
                if ctype in EXT_BY_CTYPE:
                    ext = EXT_BY_CTYPE[ctype]
                else:
                    raise ValueError(f"not an image (got {ctype or 'unknown type'})")
            path = dest_base + ext
            with open(path, "wb") as f:
                f.write(data)
            return path
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}"
            if e.code in (400, 401, 403, 404, 410):
                break
        except Exception as e:  # noqa
            last = str(e) or e.__class__.__name__
    raise RuntimeError(last or "download failed")


def download_images(records, ctx, L, stop):
    """Downloads every link of every product: BRAND_ID_01, BRAND_ID_02 ... (link 1 = _01, link 2 = _02 ...)."""
    img_dir = ctx["img_dir"]
    os.makedirs(img_dir, exist_ok=True)
    out_dir = ctx["out_dir"]
    tasks = []
    no_link = 0
    for rec in records:
        if not rec["links"]:
            no_link += 1
            continue
        key = safe_filename(rec["brand_id"]).upper()
        for url in rec["links"]:
            n = ctx["pic_count"].get(key, 0) + 1
            ctx["pic_count"][key] = n
            tasks.append((rec["brand_id"], f"{safe_filename(rec['brand_id'])}_{n:02d}", url))
    L(f"Image download: {len(tasks)} link(s) for {len(records) - no_link} product(s) ({no_link} without a link).")
    if not tasks:
        return "Images: no links found."
    existing = {os.path.splitext(f)[0].upper() for f in os.listdir(img_dir)}
    ok = skipped = 0
    failed = []
    done = 0

    def work(t):
        brand_id, fname, url = t
        if stop.is_set():
            return t, "stopped", None
        if ctx["skip_existing"] and fname.upper() in existing:
            return t, "skipped", fname
        try:
            p_ = fetch_image(url, os.path.join(img_dir, fname))
            p_ = resize_existing(p_, ctx)
            return t, "ok", os.path.basename(p_)
        except Exception as e:  # noqa
            return t, "fail", str(e)

    def progress(*_a):
        pass

    with ThreadPoolExecutor(max_workers=6) as pool:
        futures = [pool.submit(work, t) for t in tasks]
        for fut in as_completed(futures):
            t, status, info = fut.result()
            done += 1
            progress(40 + int(60 * done / len(tasks)), 100)
            if status == "ok":
                ok += 1
                L(f"[{done}/{len(tasks)}] OK      {t[0]}  ->  {info}")
            elif status == "skipped":
                skipped += 1
                L(f"[{done}/{len(tasks)}] SKIPPED {t[0]}  (already downloaded)")
            elif status == "fail":
                failed.append((t[0], t[2], info))
                L(f"[{done}/{len(tasks)}] FAILED  {t[0]}  - {info}")
            if stop.is_set():
                for f in futures:
                    f.cancel()
                L("Stopped by user.")
                break

    ctx["pic_stats"]["saved"] += ok
    ctx["pic_stats"]["existing"] += skipped
    if failed:
        fpath = os.path.join(out_dir, "Failed_Downloads.csv")
        new = not os.path.exists(fpath) or "failed_written" not in ctx
        ctx["failed_written"] = True
        try:
            with open(fpath, "w" if new else "a", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                if new:
                    w.writerow(["BRAND_ID", "URL", "REASON"])
                w.writerows(failed)
            L(f"A list of the failed links was saved to: {fpath}")
        except OSError:
            pass
    L(f"Download finished: {ok} downloaded, {skipped} skipped, {len(failed)} failed.")
    return f"Images: {ok} downloaded, {skipped} skipped, {len(failed)} failed."


# ======================================================================
#  IMAGE FILE HELPERS
# ======================================================================

def tokens_of(s):
    return re.sub(r"[^A-Z0-9]+", " ", str(s).upper()).strip()


def squash(s):
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def list_images(folder, recurse):
    out = []
    if recurse:
        for root, _dirs, files in os.walk(folder):
            for f in files:
                if os.path.splitext(f)[1].lower() in IMAGE_EXTS:
                    out.append(os.path.join(root, f))
    else:
        for f in os.listdir(folder):
            p = os.path.join(folder, f)
            if os.path.isfile(p) and os.path.splitext(f)[1].lower() in IMAGE_EXTS:
                out.append(p)
    return sorted(out, key=lambda p: natural_key(os.path.basename(p)))


# ======================================================================
#  HISTORY - one line per run, kept forever in the Windows user folder (never the files themselves)
# ======================================================================

def history_file():
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    folder = os.path.join(base, "NuORDER Client")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "history.jsonl")


def load_history():
    items = []
    try:
        with open(history_file(), encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        items.append(json.loads(line))
                    except ValueError:
                        pass
    except OSError:
        pass
    return items


def save_history(items):
    path = history_file()
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        for it in items:
            f.write(json.dumps(it, ensure_ascii=False, default=str) + "\n")
    os.replace(tmp, path)


def saved_copies_folder():
    folder = os.path.join(os.path.dirname(history_file()), "Saved Files")
    os.makedirs(folder, exist_ok=True)
    return folder


def save_history_copy(xlsx_path, run_id, label):
    """A copy of the cleaned data file (no images) kept with the history entry, in case it is needed later."""
    import shutil
    try:
        stamp = datetime.datetime.strptime(run_id[:12], "%Y%m%d%H%M").strftime("%Y-%m-%d_%H%M")
    except ValueError:
        stamp = run_id[:12]
    target = os.path.join(saved_copies_folder(),
                          f"{stamp}_{run_id[12:16]}_{safe_filename(label)[:60]}_Processed.xlsx")
    try:
        shutil.copy2(xlsx_path, target)
        return target
    except OSError:
        return None


def delete_history_copies(items):
    """Deletes the saved copies that belong to these history entries (only files inside 'Saved Files')."""
    folder = os.path.abspath(saved_copies_folder())
    n = 0
    for it in items:
        p_ = it.get("saved_copy")
        if p_ and os.path.abspath(p_).startswith(folder + os.sep) and os.path.isfile(p_):
            try:
                os.remove(p_)
                n += 1
            except OSError:
                pass
    return n


def add_history(rec):
    try:
        with open(history_file(), "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
    except OSError:
        pass


# ======================================================================
#  USER INTERFACE
# ======================================================================

def open_url(url):
    try:
        import webbrowser
        webbrowser.open(url)
    except Exception:
        pass


def open_folder(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                      # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ======================================================================
#  LOOK & FEEL - a modern light / dark theme (no extra libraries)
# ======================================================================

THEMES = {
    "light": dict(page="#EDF0F5", card="#FFFFFF", side="#FFFFFF", fg="#1C2433", muted="#6B7586", faint="#A0A8B8",
                  border="#E1E6EE", line="#D6DCE6", field="#FFFFFF", btn="#FFFFFF", btn_hover="#F3F6FB", log="#FAFBFD",
                  trough="#E4E8EF", warn="#B54708", good="#15803D", note="#B45309", bad="#B42318"),
    "dark": dict(page="#12161D", card="#1C212B", side="#181C25", fg="#E5E8EE", muted="#98A2B3", faint="#6B7486",
                 border="#2E3544", line="#394254", field="#232936", btn="#242A36", btn_hover="#2E3646", log="#181D26",
                 trough="#2A3141", warn="#F79009", good="#4ADE80", note="#FBBF24", bad="#F87171"),
}
# accent -> mode -> colors.  accent = buttons / progress / links,  a2 = step numbers / badge (same as accent
# except for the two-color Olive & Old Rose),  soft = selected sidebar item,  select = selected rows.
ACCENTS = OrderedDict([
    ("Red", {"light": dict(accent="#D92D20", hover="#B42318", soft="#FDECEA", select="#FEE4E2", a2="#D92D20", a2t="#B42318", soft2="#FDECEA"),
             "dark": dict(accent="#F04438", hover="#F97066", soft="#3A1F21", select="#4A2427", a2="#F04438", a2t="#F97066", soft2="#3A1F21")}),
    ("Purple", {"light": dict(accent="#7A3FD1", hover="#6029B0", soft="#F1EAFD", select="#EBE0FC", a2="#7A3FD1", a2t="#6029B0", soft2="#F1EAFD"),
                "dark": dict(accent="#9B6CF0", hover="#B18CF5", soft="#2B2240", select="#362A52", a2="#9B6CF0", a2t="#B18CF5", soft2="#2B2240")}),
    ("Pink", {"light": dict(accent="#D6338A", hover="#B4246F", soft="#FCE9F3", select="#FAD9EA", a2="#D6338A", a2t="#B4246F", soft2="#FCE9F3"),
              "dark": dict(accent="#EC4899", hover="#F472B6", soft="#3B1F30", select="#4A263C", a2="#EC4899", a2t="#F472B6", soft2="#3B1F30")}),
    ("Blue", {"light": dict(accent="#2453C7", hover="#1C44A8", soft="#EAF0FD", select="#DCE6FB", a2="#2453C7", a2t="#2453C7", soft2="#EAF0FD"),
              "dark": dict(accent="#4F7FF0", hover="#6A93F5", soft="#1F2A44", select="#2D4170", a2="#4F7FF0", a2t="#8EAEFF", soft2="#1F2A44")}),
    ("Olive & Old Rose", {"light": dict(accent="#66753A", hover="#55622F", soft="#F5E7E8", select="#F2E1E2", a2="#B97A80", a2t="#9A5A61", soft2="#F5E7E8"),
                          "dark": dict(accent="#8A9A55", hover="#9EAE68", soft="#3A2A2D", select="#45323A", a2="#C98F95", a2t="#E0A9AE", soft2="#3A2A2D")}),
])
DEFAULT_ACCENT = "Blue"
THEME_CHOICES = OrderedDict([("Auto", "system"), ("Light", "light"), ("Dark", "dark")])


def theme_colors(mode, accent):
    c = dict(THEMES[mode])
    c.update(ACCENTS.get(accent, ACCENTS[DEFAULT_ACCENT])[mode])
    c["info"] = c["accent"] if mode == "light" else c["hover"]
    return c


def settings_path():
    return os.path.join(os.path.dirname(history_file()), "settings.json")


def load_settings():
    try:
        with open(settings_path(), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        cur = load_settings()
        cur.update(data)
        with open(settings_path(), "w", encoding="utf-8") as f:
            json.dump(cur, f, indent=1)
    except Exception:
        pass


def windows_uses_dark():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except Exception:
        return False


def resolve_theme(choice):
    if choice == "system":
        return "dark" if windows_uses_dark() else "light"
    return choice if choice in THEMES else "light"


if tk is not None:

    def W(n):
        """Pixels that follow the Windows zoom level (100% / 125% / 150% ...)."""
        return int(round(n * UI.get("scale", 1.0)))

    UI = {"mode": "light", "accent": DEFAULT_ACCENT, "c": theme_colors("light", DEFAULT_ACCENT),
          "font": "Segoe UI", "mono": "Consolas", "nums": [], "hooks": []}

    def _pick_fonts(root):
        from tkinter import font as tkfont
        try:
            UI["scale"] = max(1.0, min(3.0, root.winfo_fpixels("1i") / 96.0))
        except Exception:
            UI["scale"] = 1.0
        fams = set(tkfont.families(root))
        UI["font"] = next((f for f in ("Segoe UI", "Inter", "Helvetica Neue", "DejaVu Sans") if f in fams), "TkDefaultFont")
        UI["mono"] = next((f for f in ("Cascadia Mono", "Consolas", "DejaVu Sans Mono") if f in fams), "TkFixedFont")
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
            try:
                tkfont.nametofont(name).configure(family=UI["font"], size=10)
            except Exception:
                pass

    def apply_theme(root, mode, accent=None):
        """(Re)paints every ttk style and every classic Tk widget: light / dark + accent color."""
        accent = accent if accent in ACCENTS else UI.get("accent", DEFAULT_ACCENT)
        c = theme_colors(mode, accent)
        UI["mode"], UI["accent"], UI["c"] = mode, accent, c
        F, B = UI["font"], (UI["font"], 10, "bold")

        def P(*v):
            return tuple(W(x) for x in v)
        st = ttk.Style(root)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure(".", background=c["card"], foreground=c["fg"], bordercolor=c["border"], lightcolor=c["card"],
                     darkcolor=c["card"], troughcolor=c["trough"], focuscolor=c["accent"], selectbackground=c["select"],
                     selectforeground=c["fg"], fieldbackground=c["field"], insertcolor=c["fg"], font=(F, 10),
                     arrowcolor=c["muted"])
        st.map(".", background=[("disabled", c["card"])], foreground=[("disabled", c["faint"])])
        # frames & labels
        for name, bg in (("TFrame", "card"), ("Card.TFrame", "card"), ("Page.TFrame", "page"), ("Side.TFrame", "side"),
                         ("Line.TFrame", "border"), ("Seg.TFrame", "page")):
            st.configure(name, background=c[bg])
        st.configure("TLabel", background=c["card"], foreground=c["fg"])
        st.configure("Page.TLabel", background=c["page"], foreground=c["fg"])
        st.configure("Muted.TLabel", background=c["card"], foreground=c["muted"])
        st.configure("PageMuted.TLabel", background=c["page"], foreground=c["muted"])
        st.configure("Info.TLabel", background=c["card"], foreground=c["info"])
        st.configure("Warn.TLabel", background=c["card"], foreground=c["warn"])
        st.configure("Bold.TLabel", background=c["card"], foreground=c["fg"], font=B)
        st.configure("Status.TLabel", background=c["card"], foreground=c["muted"], font=(F, 10, "italic"))
        st.configure("CardTitle.TLabel", background=c["card"], foreground=c["fg"], font=(F, 12, "bold"))
        st.configure("CardSub.TLabel", background=c["card"], foreground=c["muted"], font=(F, 9))
        st.configure("PageTitle.TLabel", background=c["card"], foreground=c["fg"], font=(F, 17, "bold"))
        st.configure("PageSub.TLabel", background=c["card"], foreground=c["muted"], font=(F, 10))
        st.configure("Badge.TLabel", background=c["soft2"], foreground=c["a2t"], font=(F, 9, "bold"), padding=P(12, 5))
        st.configure("Logo.TLabel", background=c["accent"], foreground="#FFFFFF", font=(F, 14, "bold"), padding=P(10, 3),
                     anchor="center")
        st.configure("App.TLabel", background=c["side"], foreground=c["fg"], font=(F, 13, "bold"))
        st.configure("SideMuted.TLabel", background=c["side"], foreground=c["muted"], font=(F, 9))
        st.configure("Section.TLabel", background=c["side"], foreground=c["faint"], font=(F, 8, "bold"))
        # cards = white blocks on the grey page; the title (number + name + description) sits inside
        st.configure("TLabelframe", background=c["card"], bordercolor=c["card"], lightcolor=c["card"],
                     darkcolor=c["card"], relief="flat", borderwidth=0)
        st.configure("TLabelframe.Label", background=c["card"], foreground=c["fg"], font=(F, 12, "bold"))
        st.configure("Inner.TLabelframe", background=c["card"], bordercolor=c["border"], lightcolor=c["card"],
                     darkcolor=c["card"], relief="solid", borderwidth=1)
        st.configure("Inner.TLabelframe.Label", background=c["card"], foreground=c["muted"], font=(F, 10, "bold"))
        # buttons
        st.configure("TButton", background=c["btn"], foreground=c["fg"], bordercolor=c["line"],
                     lightcolor=c["btn"], darkcolor=c["btn"], relief="raised", borderwidth=1, padding=P(14, 6),
                     focusthickness=0, focuscolor=c["btn"])
        st.map("TButton", background=[("disabled", c["card"]), ("pressed", c["trough"]), ("active", c["btn_hover"])],
               bordercolor=[("disabled", c["border"]), ("active", c["accent"])], foreground=[("disabled", c["faint"]),
                                                                                            ("active", c["accent"])],
               lightcolor=[("active", c["btn_hover"])], darkcolor=[("active", c["btn_hover"])])
        st.configure("Accent.TButton", background=c["accent"], foreground="#FFFFFF", bordercolor=c["accent"],
                     lightcolor=c["accent"], darkcolor=c["accent"], font=B, padding=P(20, 7), focuscolor=c["accent"])
        st.map("Accent.TButton",
               background=[("disabled", c["trough"]), ("pressed", c["hover"]), ("active", c["hover"])],
               lightcolor=[("disabled", c["trough"]), ("active", c["hover"])],
               darkcolor=[("disabled", c["trough"]), ("active", c["hover"])],
               bordercolor=[("disabled", c["trough"]), ("active", c["hover"])],
               foreground=[("disabled", c["faint"])])
        st.configure("CollIcon.TLabel", background=c["soft"], foreground=c["accent"], font=(F, 12, "bold"),
                     padding=P(8, 0), anchor="center")
        st.configure("CollTitle.TLabel", background=c["card"], foreground=c["fg"], font=(F, 11, "bold"))
        st.configure("CollDetail.TLabel", background=c["card"], foreground=c["muted"], font=(F, 9))
        st.configure("CollLink.TLabel", background=c["card"], foreground=c["info"], font=(F, 9, "bold"))
        st.configure("Collapse.TButton", background=c["card"], foreground=c["fg"], bordercolor=c["card"],
                     lightcolor=c["card"], darkcolor=c["card"], anchor="w", padding=P(20, 13), font=B,
                     focuscolor=c["card"])
        st.map("Collapse.TButton", background=[("active", c["soft"])], bordercolor=[("active", c["soft"])],
               lightcolor=[("active", c["soft"])], darkcolor=[("active", c["soft"])])
        for name, bg, fg, font in (("Nav.TButton", c["side"], c["muted"], (F, 10)),
                                   ("NavOn.TButton", c["soft"], c["accent"], (F, 10, "bold"))):
            st.configure(name, background=bg, foreground=fg, bordercolor=bg, lightcolor=bg, darkcolor=bg,
                         anchor="w", padding=P(16, 11), font=font, focuscolor=bg, relief="flat")
            st.map(name, background=[("active", c["soft"] if name == "NavOn.TButton" else c["btn"])],
                   bordercolor=[("active", c["soft"] if name == "NavOn.TButton" else c["btn"])],
                   lightcolor=[("active", c["soft"] if name == "NavOn.TButton" else c["btn"])],
                   darkcolor=[("active", c["soft"] if name == "NavOn.TButton" else c["btn"])],
                   foreground=[("active", c["accent"] if name == "NavOn.TButton" else c["fg"])])
        for name, bg, fg in (("Seg.TButton", c["page"], c["muted"]), ("SegOn.TButton", c["card"], c["fg"])):
            st.configure(name, background=bg, foreground=fg, bordercolor=c["border"] if name == "SegOn.TButton" else bg,
                         lightcolor=bg, darkcolor=bg, padding=P(2, 4), font=(F, 9, "bold" if name == "SegOn.TButton" else "normal"),
                         focuscolor=bg, anchor="center")
            st.map(name, background=[("active", c["card"])], lightcolor=[("active", c["card"])],
                   darkcolor=[("active", c["card"])])
        # inputs
        for w in ("TEntry", "TCombobox", "TSpinbox"):
            st.configure(w, fieldbackground=c["field"], foreground=c["fg"], bordercolor=c["line"],
                         lightcolor=c["field"], darkcolor=c["field"], background=c["btn"], arrowcolor=c["muted"],
                         insertcolor=c["fg"], padding=P(7, 5))
            st.map(w, bordercolor=[("focus", c["accent"]), ("hover", c["accent"])],
                   lightcolor=[("focus", c["field"])], fieldbackground=[("readonly", c["field"]), ("disabled", c["card"])],
                   foreground=[("disabled", c["faint"])], background=[("active", c["btn_hover"])],
                   selectbackground=[("readonly", c["field"])], selectforeground=[("readonly", c["fg"])])
        for w in ("TCheckbutton", "TRadiobutton"):
            st.configure(w, background=c["card"], foreground=c["fg"], indicatorbackground=c["field"],
                         indicatorforeground="#FFFFFF", upperbordercolor=c["border"], lowerbordercolor=c["border"],
                         indicatormargin=(0, 0, 8, 0), padding=P(2, 3), focuscolor=c["card"])
            st.map(w, background=[("active", c["card"])],
                   indicatorbackground=[("selected", c["accent"]), ("active", c["btn_hover"])],
                   upperbordercolor=[("selected", c["accent"])], lowerbordercolor=[("selected", c["accent"])])
        # lists / progress / scrollbars
        st.configure("Treeview", background=c["field"], fieldbackground=c["field"], foreground=c["fg"],
                     bordercolor=c["border"], lightcolor=c["field"], darkcolor=c["field"], rowheight=W(30))
        st.map("Treeview", background=[("selected", c["select"])], foreground=[("selected", c["fg"])])
        st.configure("Treeview.Heading", background=c["page"], foreground=c["muted"], bordercolor=c["border"],
                     lightcolor=c["btn"], darkcolor=c["btn"], relief="flat", font=(F, 9, "bold"), padding=P(10, 7))
        st.map("Treeview.Heading", background=[("active", c["page"])])
        st.configure("Horizontal.TProgressbar", background=c["accent"], troughcolor=c["trough"],
                     bordercolor=c["trough"], lightcolor=c["accent"], darkcolor=c["accent"], thickness=W(8))
        for w in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            st.configure(w, background=c["trough"], troughcolor=c["page"], bordercolor=c["page"],
                         lightcolor=c["trough"], darkcolor=c["trough"], arrowcolor=c["muted"], gripcount=0)
            st.map(w, background=[("active", c["border"])])
        root.configure(background=c["page"])
        for key, val in (("background", c["field"]), ("foreground", c["fg"]),
                         ("selectBackground", c["select"]), ("selectForeground", c["fg"])):
            root.option_add("*TCombobox*Listbox." + key, val)
        paint_classic(root)
        for cv, txt in list(UI["nums"]):
            _draw_num(cv, txt)
        for hook in list(UI["hooks"]):
            try:
                hook()
            except Exception:
                pass
        _title_bar(root, mode)

    def _draw_num(cv, txt):
        c = UI["c"]
        try:
            cv.configure(background=c["card"])
            cv.delete("all")
            d = W(26)
            cv.create_oval(1, 1, d - 1, d - 1, fill=c["a2"], outline=c["a2"])
            cv.create_text(d / 2, d / 2, text=txt, fill="#FFFFFF", font=(UI["font"], 9, "bold"))
        except Exception:
            pass

    def step_number(parent, txt):
        cv = tk.Canvas(parent, width=W(26), height=W(26), highlightthickness=0, borderwidth=0)
        cv._keep = True
        UI["nums"].append((cv, txt))
        _draw_num(cv, txt)
        return cv

    def _split_title(text):
        t = " ".join(str(text or "").split())
        m = re.match(r"(\d+)\.\s*(.*)$", t)
        num, t = (m.group(1), m.group(2)) if m else ("", t)
        sub = ""
        if "—" in t:
            t, sub = [x.strip() for x in t.split("—", 1)]
        else:
            m = re.match(r"(.*?)\s*\((.*)\)\s*$", t)
            if m:
                t, sub = m.group(1), m.group(2)
        return num, t, sub[:1].upper() + sub[1:] if sub else ""

    class CardFrame(ttk.LabelFrame):
        """A LabelFrame inside a 1-pixel border frame; pack / grid act on the border frame."""

        def __init__(self, parent, **kw):
            c = UI["c"]
            self._outer = tk.Frame(parent, highlightthickness=1, borderwidth=0, background=c["card"],
                                   highlightbackground=c["border"], highlightcolor=c["border"])
            super().__init__(self._outer, **kw)
            tk.Pack.pack_configure(self, fill="both", expand=True)

        def pack(self, **kw):
            self._outer.pack(**kw)
        pack_configure = pack

        def pack_forget(self):
            self._outer.pack_forget()

        def grid(self, **kw):
            self._outer.grid(**kw)
        grid_configure = grid

    def Card(parent, text="", padding=12, style=None, **kw):
        """A white card: number circle + title + description on top, contents below."""
        if style:
            return ttk.LabelFrame(parent, text=text, padding=padding, style=style, **kw)
        lf = CardFrame(parent, padding=(W(22), W(4), W(22), W(20)), **kw)
        num, title, sub = _split_title(text)
        hdr = ttk.Frame(lf, style="Card.TFrame", padding=(W(22), W(18), W(22), W(6)))
        col = 0
        if num:
            step_number(hdr, num).grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, W(12)))
            col = 1
        ttk.Label(hdr, text=title, style="CardTitle.TLabel").grid(row=0, column=col, sticky="w")
        if sub:
            ttk.Label(hdr, text=sub, style="CardSub.TLabel").grid(row=1, column=col, sticky="w", pady=(1, 0))
        lf.configure(labelwidget=hdr)
        return lf

    def paint_classic(widget):
        """Classic Tk widgets (lists, the log, canvases, pop-up windows) don't follow ttk styles."""
        c = UI["c"]
        cls = widget.winfo_class()
        try:
            if cls in ("Tk", "Toplevel"):
                widget.configure(background=c["page"])
            elif cls == "Canvas" and not getattr(widget, "_keep", False):
                widget.configure(background=c["page"], highlightthickness=0)
            elif cls == "Listbox":
                widget.configure(background=c["field"], foreground=c["fg"], selectbackground=c["select"],
                                 selectforeground=c["fg"], highlightthickness=1, highlightbackground=c["border"],
                                 highlightcolor=c["accent"], relief="flat", borderwidth=0, activestyle="none")
            elif cls == "Text":
                widget.configure(background=c["log"], foreground=c["fg"], insertbackground=c["fg"],
                                 selectbackground=c["select"], selectforeground=c["fg"], relief="flat",
                                 highlightthickness=1, highlightbackground=c["border"], highlightcolor=c["border"],
                                 borderwidth=0, padx=12, pady=10, font=(UI["mono"], 9))
            elif cls == "Frame":
                widget.configure(background=c["card"], highlightbackground=c["border"], highlightcolor=c["border"])
            elif cls == "TCombobox":
                try:
                    pd = widget.tk.call("ttk::combobox::PopdownWindow", widget)
                    widget.tk.call(f"{pd}.f.l", "configure", "-background", c["field"], "-foreground", c["fg"],
                                   "-selectbackground", c["select"], "-selectforeground", c["fg"])
                except Exception:
                    pass
        except Exception:
            pass
        for ch in widget.winfo_children():
            paint_classic(ch)

    def _title_bar(root, mode):
        """Dark window title bar on Windows 10/11 when the dark theme is on."""
        try:
            import ctypes
            root.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
            val = ctypes.c_int(1 if mode == "dark" else 0)
            for attr in (20, 19):
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(val), ctypes.sizeof(val)) == 0:
                    break
        except Exception:
            pass


if tk is not None:

    class ScrollFrame(ttk.Frame):
        """A frame with a vertical scrollbar (so small screens still work)."""

        def __init__(self, parent):
            super().__init__(parent, style="Page.TFrame")
            self.canvas = tk.Canvas(self, highlightthickness=0, background=UI["c"]["page"])
            vs = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
            self.inner = ttk.Frame(self.canvas, style="Page.TFrame", padding=(W(22), W(16), W(22), W(18)))
            self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
            self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
            self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
            self.canvas.configure(yscrollcommand=vs.set)
            self.canvas.pack(side="left", fill="both", expand=True)
            vs.pack(side="right", fill="y")
            self.canvas.bind("<Enter>", self._bind_wheel)
            self.canvas.bind("<Leave>", self._unbind_wheel)

        def _bind_wheel(self, _e):
            self.canvas.bind_all("<MouseWheel>", self._on_wheel)
            self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-2, "units"))
            self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(2, "units"))

        def _unbind_wheel(self, _e):
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")

        def _on_wheel(self, e):
            self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    class Collapsible(ttk.Frame):
        """A section that opens and closes when its bar is clicked:   [+]  Title   details ......  Show"""

        def __init__(self, parent, title, opened=False):
            super().__init__(parent, style="Page.TFrame")
            self.title, self.detail = title, ""
            c = UI["c"]
            self.bar = tk.Frame(self, highlightthickness=1, borderwidth=0, background=c["card"],
                                highlightbackground=c["border"], cursor="hand2")
            self.bar.pack(fill="x")
            inner = ttk.Frame(self.bar, padding=(W(22), W(14), W(22), W(14)), cursor="hand2")
            inner.pack(fill="x")
            self.icon = ttk.Label(inner, style="CollIcon.TLabel", width=2, cursor="hand2")
            self.icon.pack(side="left")
            self.lbl = ttk.Label(inner, style="CollTitle.TLabel", cursor="hand2")
            self.lbl.pack(side="left", padx=(12, 0))
            self.det = ttk.Label(inner, style="CollDetail.TLabel", cursor="hand2")
            self.det.pack(side="left", padx=(14, 0))
            self.link = ttk.Label(inner, style="CollLink.TLabel", cursor="hand2")
            self.link.pack(side="right")
            for w in (self.bar, inner, self.icon, self.lbl, self.det, self.link):
                w.bind("<Button-1>", lambda e: self.toggle())
            self.body = ttk.Frame(self, style="Page.TFrame", padding=(0, W(10), 0, 0))
            self._open = False
            self.set_open(opened)

        def set_title(self, title=None, detail=None):
            if title is not None:
                self.title = title
            if detail is not None:
                self.detail = detail
            self._paint()

        def _paint(self):
            self.icon.configure(text="−" if self._open else "+")
            self.lbl.configure(text=self.title)
            self.det.configure(text=self.detail)
            self.link.configure(text="Hide" if self._open else "Show")

        def set_open(self, value):
            self._open = bool(value)
            self._paint()
            if self._open:
                self.body.pack(fill="x")
            else:
                self.body.pack_forget()

        def toggle(self):
            self.set_open(not self._open)

    def resize_row(parent, v_on, v_px):
        """'Resize images' checkbox + size box (shared by the Clean and Image Renaming tabs)."""
        fr = ttk.Frame(parent)
        ttk.Checkbutton(fr, text="Resize images  —  shortest side", variable=v_on).pack(side="left")
        ttk.Spinbox(fr, from_=500, to=6000, increment=1, width=6, textvariable=v_px).pack(side="left", padx=6)
        ttk.Label(fr, text="px   (proportions kept, no borders, Lanczos, saved as JPG  —  same as XnConvert)",
                  style="Muted.TLabel").pack(side="left")
        if not pillow_available():
            ttk.Label(fr, text="   Pillow is missing: run Build_NuORDER_Client.bat again.",
                      style="Warn.TLabel").pack(side="left")
        return fr

    def resize_value(v_on, v_px):
        """None = no resizing, a number = resize to it, False = invalid (a message was shown)."""
        if not v_on.get():
            return None
        try:
            px = int(str(v_px.get()).strip())
        except ValueError:
            px = 0
        if px < 500 or px > 6000:
            messagebox.showwarning(APP_NAME, "The resize size must be a number between 500 and 6000 pixels "
                                             "(NuORDER needs at least 500 x 500).")
            return False
        return px

    class JobTab(ttk.Frame):
        """Shared behaviour: Run / Stop buttons, progress bar, live log, background thread."""

        def __init__(self, parent):
            super().__init__(parent, style="Page.TFrame")
            self.q = queue.Queue()
            self.stop_event = threading.Event()
            self.result_dir = None
            self.running = False

        # ---- widgets ----
        def build_run_area(self, parent):
            box = Card(parent, text=" Run & status ")
            box.pack(fill="both", expand=True, padx=W(8), pady=(W(14), W(8)))
            top = ttk.Frame(box)
            top.pack(fill="x")
            self.run_btn = ttk.Button(top, text="▶   Run", command=self.start, style="Accent.TButton", width=12)
            self.run_btn.pack(side="left")
            self.stop_btn = ttk.Button(top, text="■  Stop", command=self.stop_event.set, width=9, state="disabled")
            self.stop_btn.pack(side="left", padx=8)
            self.open_btn = ttk.Button(top, text="Open output folder", command=self.open_result, state="disabled")
            self.open_btn.pack(side="left")
            if hasattr(self, "reset"):
                ttk.Button(top, text="↺  Start over", command=self.ask_reset).pack(side="right")
            self.status = tk.StringVar(value="Ready.")
            ttk.Label(top, textvariable=self.status, style="Status.TLabel").pack(side="left", padx=14)
            self.bar = ttk.Progressbar(box, mode="determinate", maximum=100)
            self.bar.pack(fill="x", pady=(12, 10))
            lf = ttk.Frame(box)
            lf.pack(fill="both", expand=True)
            self.log_box = tk.Text(lf, height=14, wrap="word", state="disabled", font=(UI["mono"], 9))
            lsb = ttk.Scrollbar(lf, orient="vertical", command=self.log_box.yview)
            self.log_box.configure(yscrollcommand=lsb.set)
            lsb.pack(side="right", fill="y")
            self.log_box.pack(side="left", fill="both", expand=True)
            paint_classic(box)

        def write_log(self, text):
            self.log_box.configure(state="normal")
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        def clear_log(self):
            self.log_box.configure(state="normal")
            self.log_box.delete("1.0", "end")
            self.log_box.configure(state="disabled")

        def open_result(self):
            if self.result_dir:
                open_folder(self.result_dir)

        def ask_reset(self):
            if self.running:
                messagebox.showinfo(APP_NAME, "Please wait until the current run has finished (or press Stop).")
                return
            if messagebox.askyesno(APP_NAME, "Start over?\n\nThis clears the file list and the log and resets the settings "
                                   "(the 'Save in this folder' location is kept).", parent=self.winfo_toplevel()):
                self.reset()
                self.clear_log()
                self.bar["value"] = 0
                self.result_dir = None
                self.open_btn.configure(state="disabled")
                self.status.set("Ready.")

        # ---- running ----
        def start(self):
            if self.running:
                return
            cfg = self.collect_cfg()
            if cfg is None:
                return
            self.clear_log()
            self.stop_event.clear()
            self.running = True
            self.result_dir = None
            self.run_btn.configure(state="disabled")
            self.stop_btn.configure(state="normal")
            self.open_btn.configure(state="disabled")
            self.bar["value"] = 0
            self.status.set("Working ...")
            threading.Thread(target=self._worker, args=(cfg,), daemon=True).start()
            self.after(100, self._poll)

        def _worker(self, cfg):
            def log(m):
                self.q.put(("log", m))

            def progress(v, m=100):
                self.q.put(("progress", v, m))

            try:
                summary, folder = self.work(cfg, log, progress, self.stop_event)
                self.q.put(("done", summary, folder, True))
            except UserError as e:
                self.q.put(("log", f"\nSTOPPED: {e}"))
                self.q.put(("done", str(e), None, False))
            except Exception as e:  # noqa
                self.q.put(("log", "\nUNEXPECTED ERROR - please send this text to whoever maintains the program:\n"
                            + traceback.format_exc()))
                self.q.put(("done", f"Unexpected error: {e}", None, False))

        def _poll(self):
            finished = False
            try:
                while True:
                    msg = self.q.get_nowait()
                    if msg[0] == "log":
                        self.write_log(msg[1])
                    elif msg[0] == "progress":
                        self.bar["maximum"] = msg[2] or 100
                        self.bar["value"] = msg[1]
                    elif msg[0] == "done":
                        finished = True
                        _, summary, folder, ok = msg
                        self.running = False
                        self.run_btn.configure(state="normal")
                        self.stop_btn.configure(state="disabled")
                        self.status.set("Finished." if ok else "Stopped - see the log.")
                        if folder:
                            self.result_dir = folder
                            self.open_btn.configure(state="normal")
                        if hasattr(self, "after_run"):
                            self.after_run(ok)
                        (messagebox.showinfo if ok else messagebox.showwarning)(
                            APP_NAME, summary, parent=self.winfo_toplevel())
            except queue.Empty:
                pass
            if not finished:
                self.after(100, self._poll)

        # to be provided by the tabs
        def collect_cfg(self):
            raise NotImplementedError

        def work(self, cfg, log, progress, stop):
            raise NotImplementedError

    # ------------------------------------------------------------------
    #  TAB 1
    # ------------------------------------------------------------------
    class CleanTab(JobTab):
        NONE = "(none)"
        GENERAL = "General rules (column settings under Advanced)"

        def __init__(self, parent):
            super().__init__(parent)
            self.profiles, problems = load_profiles()
            self.states = OrderedDict()          # path -> per-file settings
            self.pdfs = []                       # PDF catalogs (sizes / fabric / pictures for MaxMara files)
            self.iid_path = {}
            self.cur = None
            self.labels = []
            self.combos = {}
            self.on_history_change = None
            self._loading = False
            self._link_offset = 1
            sf = ScrollFrame(self)
            sf.pack(fill="both", expand=True)
            body = sf.inner

            # --- 1. files ---
            f1 = Card(body, text=" 1.  Source files   —   one file, several files (one master file), or a whole brand folder   (MaxMara: add its PDFs too) ", padding=12)
            f1.pack(fill="x", padx=W(8), pady=(W(14), 0))
            bf = ttk.Frame(f1)
            bf.pack(fill="x")
            ttk.Button(bf, text="Add files...", command=self.add_files).pack(side="left")
            ttk.Button(bf, text="Pick a brand folder...", command=self.add_brand_folder).pack(side="left", padx=6)
            ttk.Button(bf, text="Remove selected", command=self.remove_selected).pack(side="left")
            ttk.Button(bf, text="Clear list", command=self.clear_files).pack(side="left", padx=6)
            ttk.Button(bf, text="Open profiles folder", command=lambda: open_folder(write_builtin_profiles())).pack(side="right")
            tf = ttk.Frame(f1)
            tf.pack(fill="x", pady=6)
            self.tree = ttk.Treeview(tf, columns=("file", "type", "sheet"), show="headings", height=5, selectmode="browse")
            for c_, t_, w_ in (("file", "File", 520), ("type", "Data type (detected)", 220), ("sheet", "Sheet", 160)):
                self.tree.heading(c_, text=t_, anchor="w")
                self.tree.column(c_, width=w_, anchor="w")
            self.tree.pack(side="left", fill="x", expand=True)
            tsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
            tsb.pack(side="left", fill="y")
            self.tree.configure(yscrollcommand=tsb.set)
            self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_select())

            lf = ttk.Frame(f1)
            lf.pack(fill="x", pady=(2, 4))
            ttk.Label(lf, text="Linesheet Name:").pack(side="left")
            self.v_linesheet = tk.StringVar()
            self.e_linesheet = ttk.Entry(lf, textvariable=self.v_linesheet, width=60)
            self.e_linesheet.pack(side="left", padx=6)
            self.v_lsinfo = tk.StringVar(value="Optional. Leave empty to use the file's own linesheet name.  ' RP' is added automatically.")
            ttk.Label(lf, textvariable=self.v_lsinfo, style="Muted.TLabel", wraplength=W(430), justify="left").pack(side="left")

            sel = ttk.LabelFrame(f1, text=" Selected file ", padding=10, style="Inner.TLabelframe")
            sel.pack(fill="x", pady=(4, 0))
            self.v_selname = tk.StringVar(value="(add files above)")
            ttk.Label(sel, textvariable=self.v_selname, style="Bold.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
            ttk.Label(sel, text="Data type:").grid(row=1, column=0, sticky="w", pady=(4, 0))
            self.v_type = tk.StringVar()
            self.cb_type = ttk.Combobox(sel, textvariable=self.v_type, state="readonly", width=40,
                                        values=sorted(self.profiles) + [self.GENERAL])
            self.cb_type.grid(row=1, column=1, sticky="w", padx=6, pady=(4, 0))
            self.cb_type.bind("<<ComboboxSelected>>", lambda e: self.type_changed())
            self.v_typeinfo = tk.StringVar()
            ttk.Label(sel, textvariable=self.v_typeinfo, style="Info.TLabel", wraplength=W(470), justify="left").grid(row=1, column=2, columnspan=2, sticky="w")
            ttk.Label(sel, text="Sheet:").grid(row=2, column=0, sticky="w", pady=(4, 0))
            self.v_sheet = tk.StringVar()
            self.cb_sheet = ttk.Combobox(sel, textvariable=self.v_sheet, state="readonly", width=30)
            self.cb_sheet.grid(row=2, column=1, sticky="w", padx=6, pady=(4, 0))
            self.cb_sheet.bind("<<ComboboxSelected>>", lambda e: self.reload_columns(redetect=True))
            hf = ttk.Frame(sel)
            hf.grid(row=2, column=2, columnspan=2, sticky="w", pady=(4, 0))
            ttk.Label(hf, text="Header row:").pack(side="left")
            self.v_header = tk.IntVar(value=1)
            ttk.Spinbox(hf, from_=1, to=200, width=6, textvariable=self.v_header,
                        command=lambda: self.reload_columns(redetect=False)).pack(side="left", padx=4)
            ttk.Button(hf, text="Reload columns", command=lambda: self.reload_columns(redetect=False)).pack(side="left", padx=4)

            # --- advanced (collapsed): column mapping + sizes ---
            self.adv = Collapsible(body, "Advanced")
            self.adv.set_title(detail="column & size settings for the selected file  ·  only needed if something was not recognised")
            self.adv.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.f2 = Card(self.adv.body, text=" Which source column is which?   —   guessed automatically;  * = required ", padding=12)
            self.f2.pack(fill="x", pady=4)
            for i, (key, label) in enumerate(MAP_FIELDS):
                r, c = divmod(i, 2)
                ttk.Label(self.f2, text=label + ":").grid(row=r, column=c * 2, sticky="w", padx=(0, 6), pady=2)
                cb = ttk.Combobox(self.f2, state="readonly", width=28, values=[self.NONE])
                cb.set(self.NONE)
                cb.grid(row=r, column=c * 2 + 1, sticky="w", padx=(0, 18), pady=2)
                self.combos[key] = cb

            # --- 3. sizes ---
            f3 = self.f3 = Card(self.adv.body, text=" Sizes ", padding=12)
            f3.pack(fill="x", pady=4)
            self.v_size_mode = tk.StringVar(value="single")
            ttk.Radiobutton(f3, text="Sizes in SIZE 1, SIZE 2 ... columns and/or one 'all sizes' column (23-34, XXS-XL, lists)",
                            variable=self.v_size_mode, value="single").grid(row=0, column=0, columnspan=3, sticky="w")
            ttk.Radiobutton(f3, text="Sizes are ACROSS several columns (matrix) - select those columns below  (Ctrl+click for several):",
                            variable=self.v_size_mode, value="matrix").grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
            self.lb_sizes = tk.Listbox(f3, selectmode="multiple", height=5, exportselection=False, width=34)
            self.lb_sizes.grid(row=2, column=0, sticky="w", pady=4)
            sb = ttk.Scrollbar(f3, orient="vertical", command=self.lb_sizes.yview)
            self.lb_sizes.configure(yscrollcommand=sb.set)
            sb.grid(row=2, column=1, sticky="ns", pady=4)
            self.v_black = tk.BooleanVar(value=True)
            self.v_reqval = tk.BooleanVar(value=False)
            opts = ttk.Frame(f3)
            opts.grid(row=2, column=2, sticky="w", padx=14)
            ttk.Checkbutton(opts, text="Black-filled cells = size NOT available", variable=self.v_black).pack(anchor="w")
            ttk.Checkbutton(opts, text="A size only counts if its cell has a value in it", variable=self.v_reqval).pack(anchor="w")

            # --- 4. settings ---
            self.more = Collapsible(body, "More settings")
            self.more.set_title(detail="season, dates, BRAND_ID, currency  ·  all files")
            self.more.pack(fill="x", padx=W(8), pady=(W(14), 0))
            f4 = self.f4 = Card(self.more.body, text=" NuORDER settings ", padding=12)
            f4.pack(fill="x", pady=4)
            self.v_season = tk.StringVar()
            self.v_year = tk.StringVar()
            self.v_datefmt = tk.StringVar(value=list(DATE_FORMAT_CHOICES)[0])
            self.v_brandsrc = tk.StringVar(value="Color name   (STYLE_COLOR)")
            self.v_pricecur = tk.StringVar(value="Automatic")
            self.v_upper = tk.BooleanVar(value=True)
            rows = [
                ("Season (only for files that have no season):", ttk.Entry(f4, textvariable=self.v_season, width=24)),
                ("Year for dates without a year:", ttk.Entry(f4, textvariable=self.v_year, width=8)),
                ("Date format (general rules):", ttk.Combobox(f4, textvariable=self.v_datefmt, state="readonly", width=46,
                                                              values=list(DATE_FORMAT_CHOICES))),
                ("BRAND_ID is built from (general rules):", ttk.Combobox(
                    f4, textvariable=self.v_brandsrc, state="readonly", width=30,
                    values=["Color name   (STYLE_COLOR)", "Color code   (STYLE_123)"])),
                ("Currency of prices that don't state one:", ttk.Combobox(
                    f4, textvariable=self.v_pricecur, state="readonly", width=14,
                    values=["Automatic", "USD", "AUD", "NZD", "EUR", "GBP", "CAD"])),
            ]
            for i, (txt, w) in enumerate(rows):
                ttk.Label(f4, text=txt).grid(row=i, column=0, sticky="w", pady=2)
                w.grid(row=i, column=1, sticky="w", padx=6, pady=2)
            ttk.Checkbutton(f4, text="Make BRAND_ID all UPPERCASE (NuORDER requires one case only)",
                            variable=self.v_upper).grid(row=5, column=0, columnspan=2, sticky="w", pady=(4, 0))
            ttk.Label(f4, text="Always applied: season kept as written in the file (JOOR: Season Name + Season Year), "
                               "dates as real dates shown 1/15/2027, sizes as text (00 stays 00), every sheet with the same layout, "
                               "' RP' after every linesheet name, color codes padded to 3 digits (12 -> 012), "
                               "WHOLESALE USD / RETAIL USD = 0.01 when there is no USD price or the price is 0. "
                               "Several files: one master file, each linesheet in its own light color.",
                      wraplength=W(880), justify="left").grid(row=6, column=0, columnspan=2, sticky="w", pady=(4, 0))

            # --- 5. images ---
            self.imgsec = Collapsible(body, "Images")
            self.imgsec.pack(fill="x", padx=W(8), pady=(W(14), 0))
            f5 = Card(self.imgsec.body, text=" Images   —   pictures in the file, image links, renaming and resizing ", padding=12)
            f5.pack(fill="x", pady=4)
            self.v_download = tk.BooleanVar(value=True)
            self.v_pics = tk.BooleanVar(value=True)
            self.v_skip = tk.BooleanVar(value=True)
            self.v_resize = tk.BooleanVar(value=True)
            self.v_resize_px = tk.StringVar(value=str(RESIZE_DEFAULT_PX))
            ttk.Checkbutton(f5, text="Download images from a column of links (when a link column is found)",
                            variable=self.v_download).grid(row=0, column=0, columnspan=3, sticky="w")
            ttk.Label(f5, text="Link column(s) of the selected file:").grid(row=1, column=0, sticky="w", pady=4, padx=(20, 0))
            self.cb_link = ttk.Combobox(f5, state="readonly", width=40, values=[self.NONE])
            self.cb_link.set(self.NONE)
            self.cb_link.grid(row=1, column=1, sticky="w", padx=6)
            ttk.Checkbutton(f5, text="Extract the pictures pasted into the Excel sheets (when found)",
                            variable=self.v_pics).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
            self.v_picinfo = tk.StringVar(value="")
            ttk.Label(f5, textvariable=self.v_picinfo, style="Muted.TLabel").grid(row=3, column=0, columnspan=3, sticky="w", padx=(20, 0))
            ttk.Label(f5, text="Rename images from this folder (optional):").grid(row=4, column=0, sticky="w", pady=(8, 0))
            self.v_rename = tk.StringVar()
            ttk.Entry(f5, textvariable=self.v_rename, width=70).grid(row=4, column=1, sticky="ew", padx=6, pady=(8, 0))
            ttk.Button(f5, text="Browse...", command=self.browse_rename).grid(row=4, column=2, pady=(8, 0))
            self.v_rename_sub = tk.BooleanVar(value=True)
            ttk.Checkbutton(f5, text="Also look inside sub-folders", variable=self.v_rename_sub).grid(row=5, column=1, sticky="w", padx=6)
            ttk.Label(f5, text="The images are matched to the products automatically (Style Number, Color, Color Code, Name) and COPIED, "
                               "renamed, into the Images folder. Your original images are not changed; unclear matches are skipped and listed.",
                      style="Muted.TLabel", wraplength=W(880), justify="left").grid(row=6, column=0, columnspan=3, sticky="w", padx=(20, 0))
            ttk.Checkbutton(f5, text="Skip images that are already in the output folder", variable=self.v_skip).grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 0))
            ttk.Label(f5, text="All images are saved in the 'Images' sub-folder as BRAND_ID_01, BRAND_ID_02 ... "
                               "(Ferragamo: STYLE NUMBER_01 ...). Several link columns: link 1 = _01, link 2 = _02 ...",
                      style="Muted.TLabel").grid(row=8, column=0, columnspan=3, sticky="w")
            resize_row(f5, self.v_resize, self.v_resize_px).grid(row=9, column=0, columnspan=3, sticky="w", pady=(10, 0))
            f5.columnconfigure(1, weight=1)
            for v_ in (self.v_download, self.v_pics, self.v_rename, self.v_resize, self.v_resize_px):
                v_.trace_add("write", lambda *a: self.update_image_summary())
            self.update_image_summary()

            # --- 6. output ---
            f6 = Card(body, text=" 3.  Where to save ", padding=12)
            f6.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.v_out = tk.StringVar()
            ttk.Label(f6, text="Save in this folder:").grid(row=0, column=0, sticky="w")
            ttk.Entry(f6, textvariable=self.v_out).grid(row=0, column=1, sticky="ew", padx=6)
            ttk.Button(f6, text="Browse...", command=self.browse_out).grid(row=0, column=2)
            ttk.Label(f6, text="A new folder  <file name>_NuORDER  is created there, holding  <file name>_Processed.xlsx, "
                               "the Images sub-folder and a log.").grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
            f6.columnconfigure(1, weight=1)

            self.build_run_area(body)
            UI["hooks"].append(self.tint_tree)
            self.tint_tree()
            if problems:
                self.write_log("Some brand profile files could not be read: " + "; ".join(problems))
            self.update_linesheet_box()

        # ---------------- file list ----------------
        def add_files(self):
            paths = filedialog.askopenfilenames(title="Choose one or several Excel / CSV files (and the brand's PDF catalogs)",
                                                filetypes=[("Excel / CSV / PDF", "*.xlsx *.xlsm *.xlsb *.csv *.pdf"),
                                                           ("All files", "*.*")])
            if paths:
                self.add_paths(paths)

        def add_brand_folder(self):
            folder = filedialog.askdirectory(title="Choose the brand's folder (data files and images)")
            if not folder:
                return
            self.status.set("Scanning the folder ...")
            self.update_idletasks()
            data, n_img = scan_brand_folder(folder)
            pdfs = scan_brand_pdfs(folder)
            if not data and not n_img:
                messagebox.showinfo(APP_NAME, "No data files (.xlsx .xlsm .xlsb .csv) or images were found in that folder.")
                self.status.set("Ready.")
                return
            if not self.v_out.get():
                self.v_out.set(folder)
            if n_img:
                self.v_rename.set(folder)
                self.v_rename_sub.set(True)
            if data:
                self.add_paths(data + pdfs)
            msg = (f"Brand folder: {len(data)} data file(s), {len(pdfs)} PDF(s) and {n_img} image(s) found in {folder} (sub-folders included)."
                   + ("  The images will be matched and renamed (see '2. Images')." if n_img else ""))
            self.write_log(msg)
            self.status.set(f"Found {len(data)} data file(s) and {n_img} image(s).")

        @staticmethod
        def kind_of(st):
            if st.get("profile_name"):
                return "profile"
            lab = st.get("label") or ""
            if lab.startswith("JOOR"):
                return "joor"
            if lab.startswith("Unreadable") or st.get("error"):
                return "bad"
            return "general"

        def tint_tree(self):
            c = UI["c"]
            for tag, col in (("profile", c["good"]), ("joor", c["info"]), ("general", c["fg"]), ("bad", c["bad"]), ("pdf", c["info"])):
                self.tree.tag_configure(tag, foreground=col)

        def badge_text(self):
            n = len(self.states)
            if not n:
                return f"{len(self.pdfs)} PDF{'s' if len(self.pdfs) != 1 else ''}" if self.pdfs else ""
            brands = {st.get("label") for st in self.states.values()
                      if st.get("profile_name") or (st.get("label") or "").startswith("JOOR")}
            return (f"{n} file{'s' if n != 1 else ''}" + (f"  ·  {len(self.pdfs)} PDF{'s' if len(self.pdfs) != 1 else ''}"
                                                          if self.pdfs else "")
                    + (f"  ·  {len(brands)} brand{'s' if len(brands) != 1 else ''} detected" if brands else ""))

        def update_image_summary(self):
            parts = []
            if self.v_pics.get():
                parts.append("pictures in the file")
            if self.v_download.get():
                parts.append("image links")
            if self.v_rename.get().strip():
                parts.append("rename from folder")
            if self.v_resize.get():
                parts.append(f"resize {self.v_resize_px.get().strip() or RESIZE_DEFAULT_PX} px")
            self.imgsec.set_title(detail=" · ".join(parts) if parts else "no images")

        def add_paths(self, paths):
            self.save_current()
            first_new = None
            for p in paths:
                if p in self.states or p in self.pdfs:
                    continue
                if os.path.splitext(p)[1].lower() == ".pdf":
                    self.pdfs.append(p)
                    iid = self.tree.insert("", "end", values=(os.path.basename(p), "PDF catalog", "support file"),
                                           tags=("pdf",))
                    self.iid_path[iid] = p
                    first_new = first_new or iid
                    continue
                try:
                    check_extension(p)
                except UserError as e:
                    messagebox.showwarning(APP_NAME, f"{os.path.basename(p)}:\n{e}")
                    continue
                self.status.set(f"Reading {os.path.basename(p)} ...")
                self.update_idletasks()
                st = auto_file_settings(p, self.profiles)
                self.states[p] = st
                n_more = len(st.get("more_sheets") or [])
                iid = self.tree.insert("", "end", values=(os.path.basename(p), st["label"],
                                                          (st.get("sheet") or "") + (f"  (+{n_more} more)" if n_more else "")),
                                       tags=(self.kind_of(st),))
                self.iid_path[iid] = p
                first_new = first_new or iid
            if not self.v_out.get() and self.states:
                self.v_out.set(os.path.dirname(next(iter(self.states))))
            self.status.set(f"{len(self.states)} file(s) in the list" + (f" + {len(self.pdfs)} PDF(s)." if self.pdfs else "."))
            self.update_linesheet_box()
            if first_new:
                self.tree.selection_set(first_new)
                self.tree.see(first_new)

        def remove_selected(self):
            for iid in self.tree.selection():
                p = self.iid_path.pop(iid, None)
                self.states.pop(p, None)
                if p in self.pdfs:
                    self.pdfs.remove(p)
                self.tree.delete(iid)
                if p == self.cur:
                    self.cur = None
            self.update_linesheet_box()
            kids = self.tree.get_children()
            if kids:
                self.tree.selection_set(kids[0])
            else:
                self.show_empty()

        def clear_files(self):
            for iid in self.tree.get_children():
                self.tree.delete(iid)
            self.states.clear()
            self.pdfs.clear()
            self.iid_path.clear()
            self.cur = None
            self.update_linesheet_box()
            self.show_empty()

        def show_empty(self):
            self.v_selname.set("(add files above)")
            self.v_type.set("")
            self.v_typeinfo.set("")
            self.v_picinfo.set("")

        def update_linesheet_box(self):
            n = len(self.states)
            if n > 1:
                self.e_linesheet.configure(state="disabled")
                self.v_lsinfo.set(f"{n} files picked: each file keeps its own linesheet name (the box is only used for one file).")
            else:
                self.e_linesheet.configure(state="normal")
                self.v_lsinfo.set("Optional. Leave empty to use the file's own linesheet name.  ' RP' is added automatically.")

        def on_select(self):
            sel = self.tree.selection()
            if not sel:
                return
            p = self.iid_path.get(sel[0])
            if p == self.cur:
                return
            self.save_current()
            if p in self.pdfs:
                self.cur = None
                self.v_selname.set(os.path.basename(p))
                self.v_type.set("PDF catalog")
                self.v_typeinfo.set("The brand's PDF catalog. MaxMara files use it to fill the sizes, the fabric composition "
                                    "and the pictures (named BRAND_ID_01). Other brands ignore it.")
                self.v_picinfo.set("")
                self.cb_sheet.configure(values=[])
                self.v_sheet.set("")
                self.show_mapping(False)
                return
            self.load_state(p)

        def _row_of(self, path):
            return next((i for i, pp in self.iid_path.items() if pp == path), None)

        # ---------------- per-file settings ----------------
        def _idx(self, cb):
            i = cb.current()
            return None if i <= 0 else i - 1

        def save_current(self):
            st = self.states.get(self.cur)
            if not st or self._loading or st.get("profile_name"):
                return
            if not st.get("labels"):
                return
            st["sheet"] = self.v_sheet.get() or st.get("sheet")
            try:
                st["header_row"] = int(self.v_header.get())
            except (tk.TclError, ValueError):
                pass
            st["mapping"] = {k: self._idx(cb) for k, cb in self.combos.items()}
            st["size_mode"] = self.v_size_mode.get()
            st["size_cols"] = list(self.lb_sizes.curselection())
            if self.cb_link.get().startswith("All image link columns"):
                st["link_col"] = "ALL"
            else:
                i = self.cb_link.current() - self._link_offset
                st["link_col"] = i if i >= 0 else None

        def load_state(self, path):
            st = self.states.get(path)
            self.cur = path
            if not st:
                return
            self._loading = True
            try:
                self.v_selname.set(os.path.basename(path))
                prof = st.get("profile_name")
                self.v_type.set(prof if prof else self.GENERAL)
                try:
                    sheets = list_sheets(path)
                except Exception:
                    sheets = [st.get("sheet") or ""]
                self.cb_sheet.configure(values=sheets)
                self.v_sheet.set(st.get("sheet") or (sheets[0] if sheets else ""))
                if prof:
                    self.v_typeinfo.set(f"Detected: {st['label']}. The {prof} rules are used automatically "
                                        f"(column and size settings are hidden for this file).")
                    self.fill_widgets([], {}, "single", [], None)
                    self.show_mapping(False)
                else:
                    self.v_typeinfo.set(f"Detected: {st['label']}" + ("  -  " + st["error"] if st.get("error") else ""))
                    self.show_mapping(True)
                    if not st.get("labels"):
                        self.reload_columns(redetect=True, keep=False)
                    self.v_header.set(st.get("header_row", 1))
                    self.fill_widgets(st.get("labels", []), st.get("mapping", {}), st.get("size_mode", "single"),
                                      st.get("size_cols", []), st.get("link_col"))
                    if st.get("labels") and st.get("mapping", {}).get("STYLE NUMBER") is None:
                        self.adv.set_open(True)
                        self.v_typeinfo.set(self.v_typeinfo.get() + "  -  Style Number column not found: please pick it below.")
                try:
                    pics, note = embedded_pictures(path, self.v_sheet.get(), read_bytes=False)
                except Exception:
                    pics, note = [], ""
                self.v_picinfo.set((f"{len(pics)} picture(s) detected in the selected file's sheet." if pics
                                    else "No pictures detected in the selected file's sheet.") + (f"  ({note})" if note else ""))
            finally:
                self._loading = False

        def show_mapping(self, show):
            """The Advanced bar is only needed for files handled by the general rules."""
            if show and not self.adv.winfo_ismapped():
                self.adv.pack(fill="x", padx=W(8), pady=(W(14), 0), before=self.more)
            elif not show and self.adv.winfo_ismapped():
                self.adv.pack_forget()

        def fill_widgets(self, labels, mapping, size_mode, size_cols, link_col):
            self.labels = labels
            values = [self.NONE] + labels
            for key, cb in self.combos.items():
                cb.configure(values=values)
                i = mapping.get(key)
                cb.set(values[i + 1] if i is not None and i + 1 < len(values) else self.NONE)
            n_links = len(guess_link_columns(labels))
            extra = [f"All image link columns ({n_links} found)"] if n_links > 1 else []
            self._link_offset = 1 + len(extra)
            self.cb_link.configure(values=[self.NONE] + extra + labels)
            if link_col == "ALL" and extra:
                self.cb_link.set(extra[0])
            elif isinstance(link_col, int) and link_col < len(labels):
                self.cb_link.set(labels[link_col])
            else:
                self.cb_link.set(self.NONE)
            self.lb_sizes.delete(0, "end")
            for l in labels:
                self.lb_sizes.insert("end", l)
            for i in size_cols:
                if i < len(labels):
                    self.lb_sizes.selection_set(i)
            self.v_size_mode.set(size_mode)

        def reload_columns(self, redetect=True, keep=True):
            st = self.states.get(self.cur)
            if not st:
                return
            path = st["path"]
            sheet = self.v_sheet.get() or st.get("sheet")
            try:
                rows = read_preview(path, sheet)
            except Exception as e:  # noqa
                messagebox.showerror(APP_NAME, f"Could not read the file:\n{e}")
                return
            if redetect or not keep:
                h = detect_header_row(rows) if rows else 1
            else:
                try:
                    h = int(self.v_header.get())
                except (tk.TclError, ValueError):
                    h = 1
            if h < 1 or h > len(rows):
                return
            labels = make_labels(rows[h - 1])
            guess = auto_map(labels)
            sc = detect_size_columns(labels)
            st.update(sheet=sheet, header_row=h, labels=labels, mapping={k: guess.get(k) for k, _ in MAP_FIELDS},
                      size_mode="matrix" if len(sc) >= 3 else "single", size_cols=sc, link_col=guess_link_column(labels))
            self.v_header.set(h)
            self.fill_widgets(labels, st["mapping"], st["size_mode"], sc, st["link_col"])
            iid = self._row_of(path)
            if iid:
                self.tree.set(iid, "sheet", sheet)
            self.status.set(f"{len(labels)} columns found in '{sheet}'.")

        def type_changed(self):
            st = self.states.get(self.cur)
            if not st:
                return
            choice = self.v_type.get()
            if choice == self.GENERAL:
                st["profile_name"] = None
                st["label"] = "General rules"
                if not st.get("labels"):
                    self.reload_columns(redetect=True, keep=False)
            else:
                st["profile_name"] = choice
                st["label"] = self.profiles[choice].get("label") or choice
            iid = self._row_of(self.cur)
            if iid:
                self.tree.set(iid, "type", st["label"])
                self.tree.item(iid, tags=(self.kind_of(st),))
            self.cur = None
            self.load_state(st["path"])

        def browse_rename(self):
            p = filedialog.askdirectory(title="Choose the folder with the images to rename")
            if p:
                self.v_rename.set(p)

        def browse_out(self):
            p = filedialog.askdirectory(title="Choose where to save the results")
            if p:
                self.v_out.set(p)

        # ---------------- run ----------------
        def collect_cfg(self):
            self.save_current()
            if not self.states:
                messagebox.showwarning(APP_NAME, "Please add at least one Excel / CSV file (section 1)."
                                       + ("\n\nThe PDF catalogs are only used together with the brand's data file." if self.pdfs else ""))
                return None
            problems = []
            first_bad = None
            for st in self.states.values():
                if st.get("profile_name"):
                    continue
                bad = None
                if not st.get("labels"):
                    bad = f"{os.path.basename(st['path'])}: no columns could be read."
                elif st["mapping"].get("STYLE NUMBER") is None:
                    bad = f"{os.path.basename(st['path'])}: pick its Style Number column (Advanced settings)."
                elif st["size_mode"] == "matrix" and not st["size_cols"]:
                    bad = f"{os.path.basename(st['path'])}: select its size columns (Advanced settings)."
                if bad:
                    problems.append(bad)
                    first_bad = first_bad or st["path"]
            if problems:
                iid = self._row_of(first_bad)
                if iid:
                    self.tree.selection_set(iid)
                    self.on_select()
                self.adv.set_open(True)
                messagebox.showwarning(APP_NAME, "Please fix these first:\n\n" + "\n".join(problems))
                return None
            out_root = self.v_out.get().strip() or os.path.dirname(next(iter(self.states)))
            if not os.path.isdir(out_root):
                messagebox.showwarning(APP_NAME, "The output folder does not exist.")
                return None
            opts = dict(
                linesheet_override=self.v_linesheet.get().strip() if len(self.states) == 1 else "",
                season=clean_text(self.v_season.get()), year=self.v_year.get(),
                date_format=DATE_FORMAT_CHOICES[self.v_datefmt.get()],
                brand_source="code" if self.v_brandsrc.get().startswith("Color code") else "name",
                brand_upper=self.v_upper.get(), black_unavailable=self.v_black.get(), require_value=self.v_reqval.get(),
                price_currency="AUTO" if self.v_pricecur.get() == "Automatic" else self.v_pricecur.get(),
                download=self.v_download.get(), extract_pictures=self.v_pics.get(), skip_existing=self.v_skip.get(),
                pdf_paths=list(self.pdfs))
            rename_dir = self.v_rename.get().strip()
            if rename_dir and not os.path.isdir(rename_dir):
                messagebox.showwarning(APP_NAME, "The folder of images to rename does not exist (section 2. Images).")
                return None
            px = resize_value(self.v_resize, self.v_resize_px)
            if px is False:
                return None
            files = [job for st in self.states.values() for job in file_jobs(st, opts, self.profiles)]
            return dict(files=files, out_root=out_root, skip_existing=self.v_skip.get(),
                        rename_dir=rename_dir, rename_recurse=self.v_rename_sub.get(), resize_px=px)

        def work(self, cfg, log, progress, stop):
            return run_job(cfg, log, progress, stop)

        def after_run(self, ok):
            if self.on_history_change:
                self.on_history_change()

        def reset(self):
            self.clear_files()
            self.v_linesheet.set("")
            self.v_rename.set("")
            self.v_rename_sub.set(True)
            self.v_season.set("")
            self.v_year.set("")
            self.v_datefmt.set(list(DATE_FORMAT_CHOICES)[0])
            self.v_brandsrc.set("Color name   (STYLE_COLOR)")
            self.v_pricecur.set("Automatic")
            self.v_upper.set(True)
            self.v_black.set(True)
            self.v_reqval.set(False)
            self.v_download.set(True)
            self.v_pics.set(True)
            self.v_skip.set(True)
            self.v_resize.set(True)
            self.v_resize_px.set(str(RESIZE_DEFAULT_PX))
            self.fill_widgets([], {}, "single", [], None)
            self.adv.set_open(False)
            self.more.set_open(False)
            self.imgsec.set_open(False)
            self.show_mapping(True)

    # ------------------------------------------------------------------
    #  TAB 2 - LINESHEET UPDATER
    # ------------------------------------------------------------------
    class UpdaterTab(JobTab):
        def __init__(self, parent):
            super().__init__(parent)
            self.custom_cols = []
            sf = ScrollFrame(self)
            sf.pack(fill="both", expand=True)
            body = sf.inner

            f1 = Card(body, text=" 1.  Data file / Data Export   —   the linesheet to update ", padding=12)
            f1.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.v_data = tk.StringVar()
            ttk.Entry(f1, textvariable=self.v_data).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            ttk.Button(f1, text="Browse...", command=self.browse_data).grid(row=0, column=1)
            ttk.Label(f1, text="A NuORDER portal Data Export, or a file this client already processed. It is never "
                               "overwritten: the result is saved as  <name>_Updated.xlsx.", style="Muted.TLabel"
                      ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
            f1.columnconfigure(0, weight=1)

            f2 = Card(body, text=" 2.  Updated file(s)   —   what the brand sent ", padding=12)
            f2.pack(fill="x", padx=W(8), pady=(W(14), 0))
            bf = ttk.Frame(f2)
            bf.pack(fill="x")
            ttk.Button(bf, text="Add files...", command=self.add_updates).pack(side="left")
            ttk.Button(bf, text="Remove selected", command=self.remove_update).pack(side="left", padx=6)
            ttk.Button(bf, text="Clear list", command=lambda: self.lb_upd.delete(0, "end")).pack(side="left")
            self.lb_upd = tk.Listbox(f2, height=4, selectmode="extended", exportselection=False)
            self.lb_upd.pack(fill="x", pady=8)
            ttk.Label(f2, text="Every sheet and column is scanned. Style numbers are matched even with different spaces, "
                               "dashes or capitals (634437DTI0W = 634437 DTI0W); a style without a color updates all its colors. "
                               "Blank cells never erase data.", style="Muted.TLabel", wraplength=W(880), justify="left").pack(anchor="w")

            f3 = Card(body, text=" 3.  What should be updated?   (optional) ", padding=12)
            f3.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.v_scope = tk.StringVar(value=list(UPDATE_SCOPES)[0])
            cb = ttk.Combobox(f3, textvariable=self.v_scope, state="readonly", width=24, values=list(UPDATE_SCOPES))
            cb.grid(row=0, column=0, sticky="w")
            cb.bind("<<ComboboxSelected>>", lambda e: self.scope_changed())
            self.v_scopeinfo = tk.StringVar(value="Everything the update file and the data file have in common is updated.")
            ttk.Label(f3, textvariable=self.v_scopeinfo, style="Info.TLabel").grid(row=0, column=1, sticky="w", padx=10)
            ttk.Label(f3, text="Tip: for a price list, choose 'Pricing only' so its other columns (descriptions, departments ...) "
                               "are not copied over.", style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

            f4 = Card(body, text=" 4.  Where to save ", padding=12)
            f4.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.v_out = tk.StringVar()
            ttk.Label(f4, text="Save in this folder:").grid(row=0, column=0, sticky="w")
            ttk.Entry(f4, textvariable=self.v_out).grid(row=0, column=1, sticky="ew", padx=6)
            ttk.Button(f4, text="Browse...", command=self.browse_out).grid(row=0, column=2)
            ttk.Label(f4, text="Results: <name>_Updated.xlsx (changed cells highlighted yellow) and <name>_Changes.xlsx "
                               "(every change, rows not found, column pairing).", style="Muted.TLabel"
                      ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 0))
            f4.columnconfigure(1, weight=1)
            self.build_run_area(body)

        def badge_text(self):
            n = self.lb_upd.size()
            parts = []
            if self.v_data.get().strip():
                parts.append(os.path.basename(self.v_data.get().strip()))
            if n:
                parts.append(f"{n} update file{'s' if n != 1 else ''}")
            return "  ·  ".join(parts)

        def browse_data(self):
            p = filedialog.askopenfilename(title="Choose the Data file / Data Export",
                                           filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.xlsb *.csv"), ("All files", "*.*")])
            if p:
                self.v_data.set(p)
                if not self.v_out.get():
                    self.v_out.set(os.path.dirname(p))
                self.custom_cols = []

        def add_updates(self):
            paths = filedialog.askopenfilenames(title="Choose the updated file(s)",
                                                filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.xlsb *.csv"), ("All files", "*.*")])
            have = set(self.lb_upd.get(0, "end"))
            for p in paths:
                if p not in have:
                    self.lb_upd.insert("end", p)

        def remove_update(self):
            for i in reversed(self.lb_upd.curselection()):
                self.lb_upd.delete(i)

        def browse_out(self):
            p = filedialog.askdirectory(title="Choose where to save the results")
            if p:
                self.v_out.set(p)

        def scope_changed(self):
            sc = UPDATE_SCOPES[self.v_scope.get()]
            texts = {"all": "Everything the update file and the data file have in common is updated.",
                     "pricing": "Only Wholesale / Retail prices are updated.",
                     "style": "Only Style Numbers are updated (e.g. from an 'UPD VPN' column).",
                     "name": "Only product Names are updated.", "season": "Only Season is updated.",
                     "dates": "Only Available From / Until (and Delivery Month) are updated.",
                     "sizes": "Only SIZE 1, SIZE 2 ... are updated.",
                     "upcs": "Only UPCs: each UPC (or EAN / barcode) goes into the UPC column of its size "
                             "(size M in SIZE 3 -> UPC 3)."}
            if sc == "custom":
                self.pick_columns()
            else:
                self.v_scopeinfo.set(texts.get(sc, ""))

        def pick_columns(self):
            path = self.v_data.get().strip()
            if not path or not os.path.isfile(path):
                messagebox.showwarning(APP_NAME, "Choose the Data file first (section 1).")
                self.v_scope.set(list(UPDATE_SCOPES)[0])
                return
            try:
                cols = [c for c in data_file_columns(path) if norm_header(c) not in _PROTECTED]
            except Exception as e:  # noqa
                messagebox.showerror(APP_NAME, f"Could not read the data file:\n{e}")
                return
            win = tk.Toplevel(self)
            win.title("Pick the columns that may be updated")
            win.geometry("520x560")
            sf = ScrollFrame(win)
            sf.pack(fill="both", expand=True)
            vars_ = []
            for c in cols:
                v = tk.BooleanVar(value=c in self.custom_cols)
                vars_.append((c, v))
                ttk.Checkbutton(sf.inner, text=c, variable=v).pack(anchor="w", padx=10)

            def done():
                self.custom_cols = [c for c, v in vars_ if v.get()]
                self.v_scopeinfo.set(f"{len(self.custom_cols)} column(s) picked: " + ", ".join(self.custom_cols[:6])
                                     + (" ..." if len(self.custom_cols) > 6 else ""))
                win.destroy()
            ttk.Button(win, text="OK", command=done, style="Accent.TButton").pack(pady=10)
            paint_classic(win)
            win.transient(self.winfo_toplevel())
            win.grab_set()

        def collect_cfg(self):
            data = self.v_data.get().strip()
            if not data or not os.path.isfile(data):
                messagebox.showwarning(APP_NAME, "Please choose the Data file / Data Export (section 1).")
                return None
            ups = list(self.lb_upd.get(0, "end"))
            if not ups:
                messagebox.showwarning(APP_NAME, "Please add the updated file(s) (section 2).")
                return None
            scope = UPDATE_SCOPES[self.v_scope.get()]
            if scope == "custom" and not self.custom_cols:
                messagebox.showwarning(APP_NAME, "'Pick columns' is selected but no column was picked (section 3).")
                return None
            out = self.v_out.get().strip() or os.path.dirname(data)
            if not os.path.isdir(out):
                messagebox.showwarning(APP_NAME, "The output folder does not exist.")
                return None
            return dict(data_path=data, update_paths=ups, scope=scope, custom_cols=self.custom_cols, out_dir=out)

        def work(self, cfg, log, progress, stop):
            return run_update(cfg, log, progress, stop)

        def reset(self):
            self.v_data.set("")
            self.lb_upd.delete(0, "end")
            self.v_scope.set(list(UPDATE_SCOPES)[0])
            self.custom_cols = []
            self.scope_changed()

    # ------------------------------------------------------------------
    #  TAB 3 - IMAGE RENAMING
    # ------------------------------------------------------------------
    class RenameTab(JobTab):
        def __init__(self, parent):
            super().__init__(parent)
            self.on_history_change = None
            sf = ScrollFrame(self)
            sf.pack(fill="both", expand=True)
            body = sf.inner

            f1 = Card(body, text=" 1.  Product file(s)   —   a processed file, a data export, or the brand's raw file ",
                                padding=12)
            f1.pack(fill="x", padx=W(8), pady=(W(14), 0))
            bf = ttk.Frame(f1)
            bf.pack(fill="x")
            ttk.Button(bf, text="Add files...", command=self.add_files).pack(side="left")
            ttk.Button(bf, text="Pick a brand folder...", command=self.pick_brand_folder).pack(side="left", padx=6)
            ttk.Button(bf, text="Remove selected", command=self.remove_files).pack(side="left")
            ttk.Button(bf, text="Clear list", command=lambda: self.lb_files.delete(0, "end")).pack(side="left", padx=6)
            self.lb_files = tk.Listbox(f1, height=4, selectmode="extended", exportselection=False)
            self.lb_files.pack(fill="x", pady=8)
            ttk.Label(f1, text="The brand is detected automatically. A raw file is only read to work out the image names "
                               "- no data file is saved.  'Pick a brand folder' adds every data file of the folder and "
                               "uses the folder for the images.", style="Muted.TLabel", wraplength=W(880),
                      justify="left").pack(anchor="w")

            f2 = Card(body, text=" 2.  Images folder ", padding=12)
            f2.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.v_imgdir = tk.StringVar()
            ttk.Entry(f2, textvariable=self.v_imgdir).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            ttk.Button(f2, text="Browse...", command=self.browse_images).grid(row=0, column=1)
            self.v_sub = tk.BooleanVar(value=True)
            ttk.Checkbutton(f2, text="Also look inside sub-folders", variable=self.v_sub).grid(row=1, column=0, sticky="w", pady=(6, 0))
            self.v_imginfo = tk.StringVar(value="")
            ttk.Label(f2, textvariable=self.v_imginfo, style="Info.TLabel").grid(row=2, column=0, columnspan=2, sticky="w")
            f2.columnconfigure(0, weight=1)

            f3 = Card(body, text=" 3.  Options ", padding=12)
            f3.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.v_resize = tk.BooleanVar(value=True)
            self.v_resize_px = tk.StringVar(value=str(RESIZE_DEFAULT_PX))
            resize_row(f3, self.v_resize, self.v_resize_px).pack(anchor="w")
            self.v_skip = tk.BooleanVar(value=True)
            ttk.Checkbutton(f3, text="Skip images that are already in the 'Renamed Images' folder",
                            variable=self.v_skip).pack(anchor="w", pady=(6, 0))
            ttk.Label(f3, text="Images are matched with the same logic as the first tab (Style Number, Color, Color Code, Name) "
                               "and COPIED as BRAND_ID_01, _02 ...  (Ferragamo: STYLE NUMBER_01, Armani: MODEL_FABRIC_COLOR_01) "
                               "into a 'Renamed Images' folder inside the images folder. Your originals are never changed; "
                               "unclear images are skipped and listed in Image_Renaming_Log.csv.",
                      style="Muted.TLabel", wraplength=W(880), justify="left").pack(anchor="w", pady=(8, 0))
            self.build_run_area(body)

        def badge_text(self):
            n = self.lb_files.size()
            parts = [f"{n} file{'s' if n != 1 else ''}"] if n else []
            if self.v_imgdir.get().strip():
                parts.append("images folder: " + os.path.basename(self.v_imgdir.get().strip().rstrip("/\\")))
            return "  ·  ".join(parts)

        def _add(self, paths):
            have = set(self.lb_files.get(0, "end"))
            for p_ in paths:
                if p_ not in have:
                    self.lb_files.insert("end", p_)

        def add_files(self):
            paths = filedialog.askopenfilenames(title="Choose the product file(s)",
                                                filetypes=[("Excel / CSV", "*.xlsx *.xlsm *.xlsb *.csv"), ("All files", "*.*")])
            self._add(paths)

        def remove_files(self):
            for i in reversed(self.lb_files.curselection()):
                self.lb_files.delete(i)

        def pick_brand_folder(self):
            folder = filedialog.askdirectory(title="Choose the brand's folder (data files and images)")
            if not folder:
                return
            data, n_img = scan_brand_folder(folder, self.v_sub.get())
            if not data and not n_img:
                messagebox.showinfo(APP_NAME, "No data files (.xlsx .xlsm .xlsb .csv) or images were found in that folder.")
                return
            self._add(data)
            if n_img:
                self.v_imgdir.set(folder)
            self.v_imginfo.set(f"Brand folder: {len(data)} data file(s) and {n_img} image(s) found.")
            self.status.set(f"Found {len(data)} data file(s) and {n_img} image(s).")

        def browse_images(self):
            p_ = filedialog.askdirectory(title="Choose the folder with the images to rename")
            if p_:
                self.v_imgdir.set(p_)
                n = len([x for x in list_images(p_, self.v_sub.get()) if not is_program_output(x, p_)])
                self.v_imginfo.set(f"{n} image(s) in that folder" + (" (sub-folders included)." if self.v_sub.get() else "."))

        def collect_cfg(self):
            files = list(self.lb_files.get(0, "end"))
            if not files:
                messagebox.showwarning(APP_NAME, "Please add the product file(s) (section 1).")
                return None
            img = self.v_imgdir.get().strip()
            if not img or not os.path.isdir(img):
                messagebox.showwarning(APP_NAME, "Please choose the images folder (section 2).")
                return None
            px = resize_value(self.v_resize, self.v_resize_px)
            if px is False:
                return None
            return dict(data_paths=files, images_dir=img, recurse=self.v_sub.get(), resize_px=px,
                        skip_existing=self.v_skip.get())

        def work(self, cfg, log, progress, stop):
            return run_rename(cfg, log, progress, stop)

        def after_run(self, ok):
            if self.on_history_change:
                self.on_history_change()

        def reset(self):
            self.lb_files.delete(0, "end")
            self.v_imgdir.set("")
            self.v_imginfo.set("")
            self.v_sub.set(True)
            self.v_resize.set(True)
            self.v_resize_px.set(str(RESIZE_DEFAULT_PX))
            self.v_skip.set(True)

    # ------------------------------------------------------------------
    #  MAXMARA PDF IMAGES
    # ------------------------------------------------------------------
    class PdfTab(JobTab):
        def __init__(self, parent):
            super().__init__(parent)
            sf = ScrollFrame(self)
            sf.pack(fill="both", expand=True)
            body = sf.inner

            f1 = Card(body, text=" 1.  PDF catalogs   —   they are processed in the order of this list ", padding=12)
            f1.pack(fill="x", padx=W(8), pady=(W(14), 0))
            bf = ttk.Frame(f1)
            bf.pack(fill="x")
            ttk.Button(bf, text="Add PDFs...", command=self.add_files).pack(side="left")
            ttk.Button(bf, text="Remove selected", command=self.remove_files).pack(side="left", padx=6)
            ttk.Button(bf, text="Move up", command=lambda: self.move(-1)).pack(side="left")
            ttk.Button(bf, text="Move down", command=lambda: self.move(1)).pack(side="left", padx=6)
            ttk.Button(bf, text="Clear list", command=lambda: self.lb_files.delete(0, "end")).pack(side="left")
            self.lb_files = tk.Listbox(f1, height=5, selectmode="extended", exportselection=False)
            self.lb_files.pack(fill="x", pady=8)
            ttk.Label(f1, text="The program finds every product code that starts with MML in the PDF and saves the "
                               "picture next to it as CODE.jpg. If a code appears more than once, the first picture is kept.",
                      style="Muted.TLabel", wraplength=W(880), justify="left").pack(anchor="w")

            f2 = Card(body, text=" 2.  Save the pictures in this folder ", padding=12)
            f2.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.v_dest = tk.StringVar()
            ttk.Entry(f2, textvariable=self.v_dest).grid(row=0, column=0, sticky="ew", padx=(0, 6))
            ttk.Button(f2, text="Browse...", command=self.browse).grid(row=0, column=1)
            f2.columnconfigure(0, weight=1)
            ttk.Label(f2, text="Existing files are never overwritten. Next, use the Image Renaming page on this folder to "
                               "give the pictures their NuORDER names.", style="Muted.TLabel", wraplength=W(880),
                      justify="left").grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

            f3 = Card(body, text=" 3.  Options ", padding=12)
            f3.pack(fill="x", padx=W(8), pady=(W(14), 0))
            self.v_resize = tk.BooleanVar(value=False)
            self.v_resize_px = tk.StringVar(value=str(RESIZE_DEFAULT_PX))
            resize_row(f3, self.v_resize, self.v_resize_px).pack(anchor="w")
            ttk.Label(f3, text="Off by default here: the pictures are saved in the PDF's own quality (JPG, quality 95) and the "
                               "Image Renaming page resizes them afterwards.", style="Muted.TLabel", wraplength=W(880),
                      justify="left").pack(anchor="w", pady=(6, 0))
            if not pymupdf_available():
                ttk.Label(f3, text="The PDF library (PyMuPDF) is missing: run Build_NuORDER_Client.bat again.",
                          style="Warn.TLabel").pack(anchor="w", pady=(6, 0))
            self.build_run_area(body)

        def badge_text(self):
            n = self.lb_files.size()
            return f"{n} PDF{'s' if n != 1 else ''}" if n else ""

        def add_files(self):
            have = set(self.lb_files.get(0, "end"))
            for p_ in filedialog.askopenfilenames(title="Choose the PDF catalog(s)",
                                                  filetypes=[("PDF documents", "*.pdf"), ("All files", "*.*")]):
                if p_ not in have:
                    self.lb_files.insert("end", p_)
                    if not self.v_dest.get().strip():
                        self.v_dest.set(os.path.join(os.path.dirname(p_), "MaxMara Images"))

        def remove_files(self):
            for i in reversed(self.lb_files.curselection()):
                self.lb_files.delete(i)

        def move(self, delta):
            sel = self.lb_files.curselection()
            if len(sel) != 1:
                return
            old = sel[0]
            new = old + delta
            if 0 <= new < self.lb_files.size():
                item = self.lb_files.get(old)
                self.lb_files.delete(old)
                self.lb_files.insert(new, item)
                self.lb_files.selection_set(new)

        def browse(self):
            p_ = filedialog.askdirectory(title="Folder where the pictures will be saved")
            if p_:
                self.v_dest.set(p_)

        def collect_cfg(self):
            files = list(self.lb_files.get(0, "end"))
            if not files:
                messagebox.showwarning(APP_NAME, "Please add at least one PDF (section 1).")
                return None
            dest = self.v_dest.get().strip()
            if not dest:
                messagebox.showwarning(APP_NAME, "Please choose the folder for the pictures (section 2).")
                return None
            px = resize_value(self.v_resize, self.v_resize_px)
            if px is False:
                return None
            return dict(pdfs=files, dest=dest, resize_px=px)

        def work(self, cfg, log, progress, stop):
            return run_pdf_images(cfg, log, progress, stop)

        def reset(self):
            self.lb_files.delete(0, "end")
            self.v_dest.set("")
            self.v_resize.set(False)
            self.v_resize_px.set(str(RESIZE_DEFAULT_PX))

    # ------------------------------------------------------------------
    #  SETTINGS - program update
    # ------------------------------------------------------------------
    class SettingsTab(ttk.Frame):
        def __init__(self, parent):
            super().__init__(parent, style="Page.TFrame", padding=(W(28), W(22), W(28), W(22)))
            self.q = queue.Queue()
            self.release = None
            self.newer = False
            self.busy = False
            self.is_busy = lambda: False          # set by the window: True while a run is in progress
            self.on_state = None                  # called with True/False when an update is (not) available
            self.stop_event = threading.Event()
            card = Card(self, text=" Program update ", padding=14)
            card.pack(fill="x")
            ttk.Label(card, text="New versions are compiled by GitHub. One click downloads the newest one, replaces this "
                                 "program and restarts it. Your History and settings are kept.",
                      style="Muted.TLabel", wraplength=W(640), justify="left").pack(anchor="w")
            info = ttk.Frame(card)
            info.pack(fill="x", pady=(12, 4))
            ttk.Label(info, text="Installed version:").grid(row=0, column=0, sticky="w")
            ttk.Label(info, text=f"v{APP_VERSION}", style="Info.TLabel").grid(row=0, column=1, sticky="w", padx=8)
            ttk.Label(info, text="Newest on GitHub:").grid(row=1, column=0, sticky="w", pady=(4, 0))
            self.v_latest = tk.StringVar(value="not checked yet")
            ttk.Label(info, textvariable=self.v_latest, style="Info.TLabel").grid(row=1, column=1, sticky="w", padx=8, pady=(4, 0))
            bf = ttk.Frame(card)
            bf.pack(fill="x", pady=(10, 4))
            self.btn_check = ttk.Button(bf, text="Check for updates", command=lambda: self.check(False))
            self.btn_check.pack(side="left")
            self.btn_update = ttk.Button(bf, text="Update now", style="Accent.TButton", command=self.update_now, state="disabled")
            self.btn_update.pack(side="left", padx=8)
            ttk.Button(bf, text="Open the releases page", command=lambda: open_url(f"https://github.com/{GITHUB_REPO}/releases")
                       ).pack(side="left")
            self.bar = ttk.Progressbar(card, mode="determinate", maximum=100)
            self.bar.pack(fill="x", pady=(8, 4))
            self.v_status = tk.StringVar(value="")
            ttk.Label(card, textvariable=self.v_status, style="Status.TLabel", wraplength=W(640), justify="left").pack(anchor="w")
            self.v_auto = tk.BooleanVar(value=bool(load_settings().get("check_updates", True)))
            ttk.Checkbutton(card, text="Check for updates when the program starts", variable=self.v_auto,
                            command=lambda: save_settings({"check_updates": bool(self.v_auto.get())})).pack(anchor="w", pady=(8, 0))
            how = ("Running as the compiled program (.exe): the new .exe replaces this one." if getattr(sys, "frozen", False)
                   else "Running from the Python file: nuorder_client.py is replaced (a .bak copy of the old one is kept).")
            ttk.Label(card, text=how, style="Muted.TLabel", wraplength=W(640), justify="left").pack(anchor="w", pady=(4, 0))
            notes = Card(self, text=" What is new in the newest version ", padding=12)
            notes.pack(fill="both", expand=True, pady=(W(14), 0))
            nf = ttk.Frame(notes)
            nf.pack(fill="both", expand=True)
            self.txt = tk.Text(nf, height=12, wrap="word", state="disabled", font=(UI["mono"], 9))
            sb = ttk.Scrollbar(nf, orient="vertical", command=self.txt.yview)
            self.txt.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            self.txt.pack(side="left", fill="both", expand=True)
            paint_classic(self)
            self.after(150, self._poll)

        def badge_text(self):
            return f"Update available: v{self.release['version']}" if self.newer and self.release else f"v{APP_VERSION}"

        def _set_notes(self, text):
            self.txt.configure(state="normal")
            self.txt.delete("1.0", "end")
            self.txt.insert("end", text)
            self.txt.configure(state="disabled")

        # ---- checking ----
        def check(self, silent):
            if self.busy:
                return
            self.busy = True
            self.btn_check.configure(state="disabled")
            if not silent:
                self.v_status.set("Checking GitHub ...")

            def work():
                try:
                    self.q.put(("checked", fetch_latest_release(), silent))
                except UserError as e:
                    self.q.put(("check_failed", str(e), silent))
                except Exception as e:  # noqa
                    self.q.put(("check_failed", f"Unexpected error: {e}", silent))
            threading.Thread(target=work, daemon=True).start()

        def update_now(self):
            if not self.release or not self.newer:
                return
            if self.is_busy():
                messagebox.showinfo(APP_NAME, "Please wait until the current run has finished, then update.")
                return
            frozen = getattr(sys, "frozen", False)
            asset = self.release["exe"] if frozen else self.release["py"]
            if not asset:
                messagebox.showinfo(APP_NAME, "GitHub is still compiling this version (it takes about 5 minutes). "
                                              "Please try again in a few minutes.")
                return
            if not messagebox.askyesno(APP_NAME, f"Update to v{self.release['version']}?\n\nThe program will download the new "
                                       "version, replace this one and restart.", parent=self.winfo_toplevel()):
                return
            target = sys.executable if frozen else os.path.abspath(sys.argv[0])
            dest = os.path.join(os.path.dirname(target), ("NuORDER Client.new.exe" if frozen else "nuorder_client.new.py"))
            self.busy = True
            self.btn_check.configure(state="disabled")
            self.btn_update.configure(state="disabled")
            self.bar["value"] = 0
            self.v_status.set("Downloading ...")

            def work():
                try:
                    def prog(got, total):
                        self.q.put(("progress", got, total))
                    download_asset(asset, dest, prog, self.stop_event)
                    self.q.put(("downloaded", dest, target, frozen))
                except UserError as e:
                    self.q.put(("update_failed", str(e)))
                except Exception as e:  # noqa
                    self.q.put(("update_failed", f"Unexpected error: {e}"))
            threading.Thread(target=work, daemon=True).start()

        def _poll(self):
            try:
                while True:
                    msg = self.q.get_nowait()
                    kind = msg[0]
                    if kind == "progress":
                        self.bar["maximum"] = msg[2] or 100
                        self.bar["value"] = msg[1]
                        if msg[2]:
                            self.v_status.set(f"Downloading ...  {msg[1] / 1048576:.1f} of {msg[2] / 1048576:.1f} MB")
                    elif kind == "checked":
                        self.busy = False
                        self.btn_check.configure(state="normal")
                        rel, silent = msg[1], msg[2]
                        self.release = rel
                        self.newer = is_newer(rel["version"], APP_VERSION)
                        self.v_latest.set(f"v{rel['version']}" + ("   (a newer version is available)" if self.newer
                                                                  else "   (you have the newest version)"))
                        self._set_notes(rel["notes"] or "(no description)")
                        self.btn_update.configure(state="normal" if self.newer else "disabled")
                        if not silent:
                            self.v_status.set("A new version is available - press 'Update now'." if self.newer
                                              else "You already have the newest version.")
                        if self.on_state:
                            self.on_state(self.newer)
                    elif kind == "check_failed":
                        self.busy = False
                        self.btn_check.configure(state="normal")
                        if not msg[2]:
                            self.v_status.set(msg[1])
                    elif kind == "update_failed":
                        self.busy = False
                        self.btn_check.configure(state="normal")
                        self.btn_update.configure(state="normal")
                        self.v_status.set(msg[1])
                        messagebox.showwarning(APP_NAME, msg[1] + f"\n\nYou can also download the new version by hand:\n"
                                               f"https://github.com/{GITHUB_REPO}/releases/latest")
                    elif kind == "downloaded":
                        dest, target, frozen = msg[1], msg[2], msg[3]
                        try:
                            (install_exe_update(dest, target) if frozen else install_py_update(dest, target))
                        except (UserError, OSError, SyntaxError) as e:
                            self.busy = False
                            self.btn_check.configure(state="normal")
                            self.btn_update.configure(state="normal")
                            self.v_status.set(f"The update could not be installed: {e}")
                            messagebox.showwarning(APP_NAME, f"The update could not be installed:\n{e}\n\n"
                                                   f"Download the new version by hand:\nhttps://github.com/{GITHUB_REPO}/releases/latest")
                            continue
                        self.v_status.set("Updated. Restarting ...")
                        messagebox.showinfo(APP_NAME, f"Updated to v{self.release['version']}. The program restarts now.",
                                            parent=self.winfo_toplevel())
                        relaunch_program()
                        self.winfo_toplevel().destroy()
                        return
            except queue.Empty:
                pass
            self.after(150, self._poll)

    # ------------------------------------------------------------------
    #  TAB 4 - HISTORY
    # ------------------------------------------------------------------
    class HistoryTab(ttk.Frame):
        COLS = (("time", "Date", 120), ("brand", "Brand", 150), ("files", "Files", 330), ("products", "Products", 70),
                ("linesheets", "Linesheets", 75), ("images", "Images", 60), ("status", "Status", 150))

        def __init__(self, parent):
            super().__init__(parent, style="Page.TFrame", padding=(W(28), W(22), W(28), W(22)))
            card = Card(self, text=" Past runs   —   search, reopen saved data files, export ")
            card.pack(fill="both", expand=True)
            top = ttk.Frame(card, padding=(0, 0, 0, 8))
            top.pack(fill="x")
            ttk.Label(top, text="Search:").pack(side="left")
            self.v_search = tk.StringVar()
            e = ttk.Entry(top, textvariable=self.v_search, width=40)
            e.pack(side="left", padx=6)
            self.v_search.trace_add("write", lambda *a: self.refresh())
            ttk.Label(top, text="(brand, file name, date, season ...)", style="Muted.TLabel").pack(side="left")
            self.v_count = tk.StringVar()
            ttk.Label(top, textvariable=self.v_count).pack(side="right")
            mid = ttk.Frame(card)
            mid.pack(fill="both", expand=True)
            self.tree = ttk.Treeview(mid, columns=[c[0] for c in self.COLS], show="headings", selectmode="extended")
            for key, title, w in self.COLS:
                self.tree.heading(key, text=title, anchor="w")
                self.tree.column(key, width=w, anchor="w")
            self.tree.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
            sb.pack(side="left", fill="y")
            self.tree.configure(yscrollcommand=sb.set)
            self.tree.bind("<Double-1>", lambda e: self.open_folder_sel())
            self.tree.bind("<<TreeviewSelect>>", lambda e: self.show_detail())
            self.v_detail = tk.StringVar()
            ttk.Label(card, textvariable=self.v_detail, wraplength=W(900), justify="left", padding=(0, 10)).pack(fill="x")
            bot = ttk.Frame(card)
            bot.pack(fill="x")
            ttk.Button(bot, text="Open saved copy", command=self.open_copy_sel, style="Accent.TButton").pack(side="left", padx=(0, 6))
            for txt, cmd in (("Save copy as...", self.save_copy_as), ("Open output folder", self.open_folder_sel),
                             ("Open log", self.open_log_sel)):
                ttk.Button(bot, text=txt, command=cmd).pack(side="left", padx=(0, 6))
            bot2 = ttk.Frame(card, padding=(0, 8, 0, 0))
            bot2.pack(fill="x")
            for txt, cmd in (("Edit brand...", self.edit_brand), ("Export to Excel...", self.export)):
                ttk.Button(bot2, text=txt, command=cmd).pack(side="left", padx=(0, 6))
            ttk.Button(bot2, text="Clear all", command=self.clear_all).pack(side="right")
            ttk.Button(bot2, text="Delete selected", command=self.delete_sel).pack(side="right", padx=6)
            ttk.Label(card, text=f"History and a saved copy of every cleaned data file (no images) are kept in "
                                 f"{os.path.dirname(history_file())}.", style="Muted.TLabel", wraplength=W(880),
                      justify="left", padding=(0, 10, 0, 0)).pack(anchor="w")
            self.items = []
            self.refresh()

        def badge_text(self):
            n = len(self.items)
            return f"{n} run{'s' if n != 1 else ''} saved" if n else ""

        def refresh(self):
            self.items = load_history()
            q = self.v_search.get().strip().lower()
            for iid in self.tree.get_children():
                self.tree.delete(iid)
            shown = 0
            for it in reversed(self.items):
                files = ", ".join(f["name"] + (f" ({f['type']})" if f.get("type") else "") for f in it.get("files", []))
                text = " ".join([it.get("time", ""), it.get("brand", ""), files, " ".join(it.get("seasons", [])),
                                 it.get("status", ""), it.get("output", "")]).lower()
                if q and q not in text:
                    continue
                status = it.get("status", "")
                if it.get("folder") and not os.path.isdir(it["folder"]):
                    status += "  (folder not found)"
                self.tree.insert("", "end", iid=it["id"], values=(
                    it.get("time", ""), it.get("brand", ""), files, it.get("products", ""), it.get("linesheets", ""),
                    it.get("images", ""), status))
                shown += 1
            self.v_count.set(f"{shown} of {len(self.items)} run(s)")

        def _selected(self):
            ids = set(self.tree.selection())
            return [it for it in self.items if it["id"] in ids]

        def show_detail(self):
            sel = self._selected()
            if not sel:
                self.v_detail.set("")
                return
            it = sel[0]
            lsn = it.get("linesheet_names") or []
            copy_ = it.get("saved_copy")
            self.v_detail.set(
                f"Output: {it.get('output', '')}\nSeasons: {', '.join(it.get('seasons', [])) or '-'}"
                + (f"\nLinesheets: " + "   •   ".join(f"{n} ({c})" for n, c in lsn) if lsn else "")
                + (f"\nSaved copy: {os.path.basename(copy_)}" + ("" if os.path.isfile(copy_) else "  (missing)")
                   if copy_ else ("\nSaved copy: none (image renaming runs have no data file)"
                                  if it.get("kind") == "rename" else "\nSaved copy: none (run made before v3.1)"))
                + (f"\nNotes: {it['notes']}" if it.get("notes") else ""))

        def _copy_of_sel(self):
            sel = self._selected()
            if not sel:
                messagebox.showinfo(APP_NAME, "Select a run in the list first.")
                return None
            p_ = sel[0].get("saved_copy")
            if not p_ or not os.path.isfile(p_):
                messagebox.showinfo(APP_NAME, "This run has no saved copy of its data file"
                                    + (" (image renaming runs don't create one)." if sel[0].get("kind") == "rename"
                                       else " (runs made before version 3.1 were not copied, or the copy was deleted)."))
                return None
            return p_

        def open_copy_sel(self):
            p_ = self._copy_of_sel()
            if p_:
                open_folder(p_)

        def save_copy_as(self):
            p_ = self._copy_of_sel()
            if not p_:
                return
            name = re.sub(r"^\d{4}-\d{2}-\d{2}_\d{4}_\d{4}_", "", os.path.basename(p_))
            dest = filedialog.asksaveasfilename(title="Save a copy of the data file", defaultextension=".xlsx",
                                                initialfile=name, filetypes=[("Excel", "*.xlsx")])
            if not dest:
                return
            import shutil
            try:
                shutil.copy2(p_, dest)
                messagebox.showinfo(APP_NAME, f"Saved:\n{dest}")
            except OSError as e:
                messagebox.showwarning(APP_NAME, f"Could not save - is that file open in Excel?\n{e}")

        def open_folder_sel(self):
            for it in self._selected()[:1]:
                if it.get("folder") and os.path.isdir(it["folder"]):
                    open_folder(it["folder"])
                else:
                    messagebox.showinfo(APP_NAME, "That output folder no longer exists (it was moved or deleted).")

        def open_log_sel(self):
            for it in self._selected()[:1]:
                if it.get("log") and os.path.isfile(it["log"]):
                    open_folder(it["log"])
                else:
                    messagebox.showinfo(APP_NAME, "That log file no longer exists.")

        def edit_brand(self):
            sel = self._selected()
            if not sel:
                return
            from tkinter import simpledialog
            new = simpledialog.askstring(APP_NAME, "Brand for the selected run(s):", initialvalue=sel[0].get("brand", ""),
                                         parent=self.winfo_toplevel())
            if new is None:
                return
            ids = {it["id"] for it in sel}
            for it in self.items:
                if it["id"] in ids:
                    it["brand"] = new.strip()
            save_history(self.items)
            self.refresh()

        def delete_sel(self):
            sel = self._selected()
            n_copies = sum(1 for it in sel if it.get("saved_copy") and os.path.isfile(it["saved_copy"]))
            if sel and messagebox.askyesno(APP_NAME, f"Delete {len(sel)} run(s) from the history?"
                                           + (f"\n\nWARNING: their {n_copies} saved data file copy(ies) will be deleted too. "
                                              "Your output folders are not touched." if n_copies else ""),
                                           icon="warning" if n_copies else "question"):
                ids = {it["id"] for it in sel}
                delete_history_copies(sel)
                save_history([it for it in self.items if it["id"] not in ids])
                self.refresh()
                self.v_detail.set("")

        def clear_all(self):
            if self.items and messagebox.askyesno(
                    APP_NAME, "Delete the WHOLE history? This cannot be undone.\n\nWARNING: every saved data file copy "
                              "is deleted too. Your output folders are not touched.", icon="warning"):
                delete_history_copies(self.items)
                save_history([])
                self.refresh()
                self.v_detail.set("")

        def export(self):
            p = filedialog.asksaveasfilename(title="Export the history", defaultextension=".xlsx",
                                             initialfile="NuORDER Client history.xlsx", filetypes=[("Excel", "*.xlsx")])
            if not p:
                return
            wb = Workbook()
            ws = wb.active
            ws.title = "History"
            ws.append(["DATE", "BRAND", "FILES", "DATA TYPES", "SEASONS", "PRODUCTS", "LINESHEETS", "IMAGES",
                       "STATUS", "NOTES", "OUTPUT FILE", "LINESHEET NAMES", "SAVED COPY"])
            for c in ws[1]:
                c.font = FONT_HEADER
            for it in self.items:
                ws.append([it.get("time"), it.get("brand"), ", ".join(f["name"] for f in it.get("files", [])),
                           ", ".join(dict.fromkeys(f.get("type", "") for f in it.get("files", []))),
                           ", ".join(it.get("seasons", [])), it.get("products"), it.get("linesheets"), it.get("images"),
                           it.get("status"), it.get("notes"), it.get("output"),
                           ", ".join(f"{n} ({c})" for n, c in (it.get("linesheet_names") or [])), it.get("saved_copy")])
            auto_fit(ws, max_w=60)
            try:
                wb.save(p)
                messagebox.showinfo(APP_NAME, f"History exported to:\n{p}")
            except PermissionError:
                messagebox.showwarning(APP_NAME, "Could not save - is that file open in Excel?")

    # ------------------------------------------------------------------


    def main():
        try:                                    # sharper text on Windows
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        root = tk.Tk()
        root.title(f"{APP_NAME}  v{APP_VERSION}")
        root.geometry("1340x920")
        root.minsize(1080, 640)
        build_app(root)
        root.mainloop()

    PAGES = [
        ("clean", "▤", "Clean Data & Images", "Brand linesheets  →  NuORDER-ready data and images"),
        ("upd", "⟳", "Linesheet Updater", "Apply a brand's changes (prices, styles, UPCs ...) to a data file or NuORDER export"),
        ("rename", "▣", "Image Renaming", "Match a folder of images to the products and copy them renamed and resized"),
        ("pdf", "▦", "MaxMara PDF", "Pull the product pictures (MML codes) out of MaxMara PDF catalogs"),
        ("hist", "◷", "History", "Every run, with a saved copy of each cleaned data file"),
        ("settings", "⚙", "Settings", "Program update"),
    ]

    def build_app(root):
        _pick_fonts(root)
        sett = load_settings()
        choice = sett.get("theme", "system")
        if choice not in THEME_CHOICES.values():
            choice = "system"
        accent = sett.get("accent", DEFAULT_ACCENT)
        if accent not in ACCENTS:
            accent = DEFAULT_ACCENT
        UI["accent"] = accent
        UI["choice"] = choice
        apply_theme(root, resolve_theme(choice), accent)

        # ---------------- sidebar ----------------
        side = ttk.Frame(root, style="Side.TFrame", width=W(260))
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        ttk.Frame(root, style="Line.TFrame", width=1).pack(side="left", fill="y")
        logo = ttk.Frame(side, style="Side.TFrame", padding=(W(22), W(26), W(16), W(28)))
        logo.pack(fill="x")
        ttk.Label(logo, text="N", style="Logo.TLabel", width=2).pack(side="left")
        nm = ttk.Frame(logo, style="Side.TFrame")
        nm.pack(side="left", padx=12)
        ttk.Label(nm, text=APP_NAME, style="App.TLabel").pack(anchor="w")
        ttk.Label(nm, text=f"v{APP_VERSION}", style="SideMuted.TLabel").pack(anchor="w")
        ttk.Label(side, text="WORKSPACE", style="Section.TLabel", padding=(26, 0, 0, 8)).pack(anchor="w")
        nav = {}
        navbox = ttk.Frame(side, style="Side.TFrame", padding=(W(14), 0))
        navbox.pack(fill="x")

        bottom = ttk.Frame(side, style="Side.TFrame", padding=(W(22), 0, W(22), W(24)))
        bottom.pack(side="bottom", fill="x")
        ttk.Frame(side, style="Line.TFrame", height=1).pack(side="bottom", fill="x", padx=18, pady=(0, 16))
        ttk.Label(bottom, text="Accent color", style="SideMuted.TLabel").pack(anchor="w")
        swrow = ttk.Frame(bottom, style="Side.TFrame", padding=(0, 8, 0, 14))
        swrow.pack(anchor="w")
        swatches = {}

        def draw_swatch(name):
            cv = swatches[name]
            c = UI["c"]
            cv.configure(background=c["side"])
            cv.delete("all")
            lt = ACCENTS[name][UI["mode"]]
            a, b, o = W(5), W(25), W(29)
            if name == UI["accent"]:
                cv.create_oval(1, 1, o, o, outline=c["fg"], width=max(2, W(2)))
            if lt["a2"] != lt["accent"]:
                cv.create_arc(a, a, b, b, start=45, extent=180, fill=lt["accent"], outline=lt["accent"])
                cv.create_arc(a, a, b, b, start=225, extent=180, fill=lt["a2"], outline=lt["a2"])
            else:
                cv.create_oval(a, a, b, b, fill=lt["accent"], outline=lt["accent"])

        def pick_accent(name):
            save_settings({"accent": name})
            apply_theme(root, UI["mode"], name)

        for name in ACCENTS:
            cv = tk.Canvas(swrow, width=W(30), height=W(30), highlightthickness=0, borderwidth=0, cursor="hand2")
            cv._keep = True
            cv.pack(side="left", padx=(0, 6))
            cv.bind("<Button-1>", lambda e, n=name: pick_accent(n))
            swatches[name] = cv
            draw_swatch(name)
        UI["hooks"].append(lambda: [draw_swatch(n) for n in swatches])
        ttk.Label(bottom, text="Theme   (Auto = same as Windows)", style="SideMuted.TLabel").pack(anchor="w")
        seg = ttk.Frame(bottom, style="Seg.TFrame", padding=3)
        seg.pack(fill="x", pady=(8, 0))
        segbtn = {}

        def pick_theme(ch):
            UI["choice"] = ch
            save_settings({"theme": ch})
            apply_theme(root, resolve_theme(ch), UI["accent"])

        for label, ch in THEME_CHOICES.items():
            b_ = ttk.Button(seg, text=label, width=1, command=lambda c_=ch: pick_theme(c_))
            b_.pack(side="left", fill="x", expand=True)
            segbtn[ch] = b_
        UI["hooks"].append(lambda: [b_.configure(style="SegOn.TButton" if ch == UI["choice"] else "Seg.TButton")
                                    for ch, b_ in segbtn.items()])

        # ---------------- main area ----------------
        main_ = ttk.Frame(root, style="Page.TFrame")
        main_.pack(side="left", fill="both", expand=True)
        top = ttk.Frame(main_, style="Card.TFrame", padding=(W(32), W(20), W(32), W(18)))
        top.pack(fill="x")
        ttk.Frame(main_, style="Line.TFrame", height=1).pack(fill="x")
        v_title, v_sub, v_badge = tk.StringVar(), tk.StringVar(), tk.StringVar()
        tl = ttk.Frame(top, style="Card.TFrame")
        tl.pack(side="left")
        ttk.Label(tl, textvariable=v_title, style="PageTitle.TLabel").pack(anchor="w")
        ttk.Label(tl, textvariable=v_sub, style="PageSub.TLabel").pack(anchor="w", pady=(2, 0))
        badge = ttk.Label(top, textvariable=v_badge, style="Badge.TLabel")
        badge.pack(side="right")
        holder = ttk.Frame(main_, style="Page.TFrame")
        holder.pack(fill="both", expand=True)
        clean, upd, rename, hist = CleanTab(holder), UpdaterTab(holder), RenameTab(holder), HistoryTab(holder)
        pdf = PdfTab(holder)
        settings = SettingsTab(holder)
        settings.is_busy = lambda: any(t.running for t in (clean, upd, rename, pdf))
        pages = {"clean": clean, "upd": upd, "rename": rename, "pdf": pdf, "hist": hist, "settings": settings}
        cur = {"key": None}

        def show(key):
            if cur["key"] == key:
                return
            if cur["key"]:
                pages[cur["key"]].pack_forget()
            cur["key"] = key
            pages[key].pack(fill="both", expand=True)
            for k_, b_ in nav.items():
                b_.configure(style="NavOn.TButton" if k_ == key else "Nav.TButton")
            meta = next(p_ for p_ in PAGES if p_[0] == key)
            v_title.set(meta[2])
            v_sub.set(meta[3])
            if key == "hist":
                hist.refresh()
            update_badge()

        for key, icon, title, _sub in PAGES:
            b_ = ttk.Button(navbox, text=title, style="Nav.TButton", command=lambda k_=key: show(k_))
            b_.pack(fill="x", pady=2)
            nav[key] = b_

        def update_badge():
            try:
                txt = pages[cur["key"]].badge_text() if cur["key"] else ""
            except Exception:
                txt = ""
            if v_badge.get() != txt:
                v_badge.set(txt)
            if txt and not badge.winfo_ismapped():
                badge.pack(side="right")
            elif not txt and badge.winfo_ismapped():
                badge.pack_forget()

        def tick():
            update_badge()
            if UI["choice"] == "system":                # 'Windows': follow Windows when it switches
                want = resolve_theme("system")
                if want != UI["mode"]:
                    apply_theme(root, want, UI["accent"])
            root.after(1500, tick)

        def mark_update(newer):
            nav["settings"].configure(text="Settings   •  update" if newer else "Settings")
            update_badge()

        settings.on_state = mark_update
        cleanup_old_update()
        if load_settings().get("check_updates", True):
            root.after(2500, lambda: settings.check(True))
        clean.on_history_change = hist.refresh
        rename.on_history_change = hist.refresh
        show("clean")
        apply_theme(root, UI["mode"], UI["accent"])
        root.after(1500, tick)
        return dict(show=show, clean=clean, upd=upd, rename=rename, pdf=pdf, hist=hist, settings=settings, pick_theme=pick_theme,
                    pick_accent=pick_accent)

else:

    def main():
        print("This program needs Tkinter (it comes with the normal Python installer from python.org).")


if __name__ == "__main__":
    main()
