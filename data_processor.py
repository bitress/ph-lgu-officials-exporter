"""
Data processor for Philippine Local Government Officials List.
Parses Officials_List.xlsx and groups officials by (Region, Province, Town).
Extracts Mayor Name, Municipal Hall Address, Town, Province, Vice Mayor.

Features:
- Fast loading with openpyxl read_only=True and automatic JSON caching (.officials_cache.json)
- Resilient text encoding and mojibake recovery (Ñ/ñ, accented letters, em-dashes)
- Robust suffix cleaning (strips N/A variants, typos, honorifics, normalizes Roman numerals)
- Address normalization preserving Philippine LGU conventions (Brgy., Poblacion, Nat'l Hwy)
- Multi-format exporters: CSV (UTF-8 BOM), Excel (.xlsx styled workbook), JSON
- Optional Contact Number inclusion
- Rolling timestamped backups before workbook updates
"""

import csv
import io
import re
import os
import json
import shutil
import datetime
from typing import Dict, List, Optional, Tuple, Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Common invalid/placeholder tokens found in DILG records
INVALID_TOKENS_EXACT = {
    'N/A', 'NA', 'NONE', 'N.A.', 'N.A', 'N/A.', 'N/A/', 'N\\A', 'N/A`', "''N/A",
    'NULL', ',N/A', 'N/A', 'N', '-', '--', '---', '.', '/', '\\', '`', "''", '""',
    'NOT APPLICABLE'
}

INVALID_PATTERN = re.compile(r'^(n[/\\]?a|none|null|not\s*applicable|[-_`\'"\\/.]+)$', re.IGNORECASE)

# Mojibake mapping table for mis-decoded UTF-8 sequences commonly found in Windows/DILG data
MOJIBAKE_MAP = {
    'Ã±': 'ñ',
    'Ã‘': 'Ñ',
    'Ã©': 'é',
    'Ã‰': 'É',
    'Ã³': 'ó',
    'Ã“': 'Ó',
    'Ã¡': 'á',
    'Ã': 'Á',
    'â€“': '-',
    'â€”': '—',
    'â€˜': "'",
    'â€™': "'",
    'â€œ': '"',
    'â€': '"',
}

# Misplaced honorifics that sometimes appear in the suffix column
HONORIFIC_PREFIXES = {'MR', 'MR.', 'MRS', 'MRS.', 'MS', 'MS.', 'DR', 'DR.', 'MIR', 'ENG', 'ENGR', 'ATTY', 'ATTY.'}

# Roman numerals up to X
ROMAN_NUMERALS = {'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X'}


def clean_text_encoding(text: str) -> str:
    """Fix replacement characters (\\ufffd) and common UTF-8/CP1252 mojibake artifacts."""
    if not text:
        return ""
    
    # 1. Direct replacement of common mojibake sequences
    s = str(text)
    for bad, good in MOJIBAKE_MAP.items():
        if bad in s:
            s = s.replace(bad, good)

    # 2. Fix replacement character (\ufffd) for Ñ / ñ
    if '\ufffd' in s:
        result = []
        i = 0
        while i < len(s):
            if s[i] == '\ufffd':
                if i + 1 < len(s) and s[i + 1].islower():
                    result.append('ñ')
                else:
                    result.append('Ñ')
            else:
                result.append(s[i])
            i += 1
        s = "".join(result)

    # 3. Strip zero-width / non-breaking space artifacts
    s = s.replace('\ufeff', '').replace('\u200b', '').replace('\xa0', ' ')
    return s


def clean_token(val: Any) -> str:
    """Strip whitespace, fix encoding, and return cleaned string or empty if invalid placeholder."""
    if val is None:
        return ""
    s = str(val).strip()
    s = clean_text_encoding(s)
    
    # Strip wrapping single/double quotes or backticks
    while (s.startswith(('"', "'", '`')) and s.endswith(('"', "'", '`')) and len(s) >= 2):
        s = s[1:-1].strip()

    if not s:
        return ""

    if s.upper() in INVALID_TOKENS_EXACT or INVALID_PATTERN.match(s):
        return ""

    # Strip stray surrounding backticks or quotes if still present
    s = s.strip("`'\"").strip()
    if s.upper() in INVALID_TOKENS_EXACT or INVALID_PATTERN.match(s):
        return ""

    return s


def clean_suffix(val: Any) -> str:
    """
    Standardize and clean personal suffixes (Jr., Sr., III, etc.).
    Filters out invalid placeholders, typos, and misplaced honorifics.
    """
    token = clean_token(val)
    if not token:
        return ""

    token_upper = token.upper().strip(".,`'\"").strip()
    
    # Filter out honorifics placed in suffix column (e.g. MR, DR.)
    if token_upper in HONORIFIC_PREFIXES:
        return ""

    # Filter out standalone invalid abbreviations
    if token_upper in {'TAN', 'N/A', 'NA', 'NONE', 'NULL'}:
        return ""

    # Junior variants: JR, JR., JR, JR.I -> Jr.
    if token_upper in {'JR', 'JR.', 'JR,', 'JR.I'}:
        return "Jr."

    # Senior variants: SR, SR. -> Sr.
    if token_upper in {'SR', 'SR.'}:
        return "Sr."

    # Roman numerals: II, III, IV, etc.
    if token_upper in ROMAN_NUMERALS:
        return token_upper

    # Compound suffixes: JR IV, JR-IV, jr ii
    if re.match(r'^JR[\s\-_]+(II|III|IV|V)$', token_upper):
        num = re.sub(r'^JR[\s\-_]+', '', token_upper)
        return f"Jr. {num}"

    # Return standard capitalized suffix if clean
    clean_suf = token.strip(".,`'\" ")
    if len(clean_suf) <= 5:
        if clean_suf.upper() in ROMAN_NUMERALS:
            return clean_suf.upper()
        return clean_suf.capitalize() + "." if not clean_suf.endswith('.') else clean_suf.capitalize()

    return ""


def clean_address(val: Any) -> str:
    """Clean address string, return empty if invalid."""
    if val is None:
        return ""
    s = str(val).strip()
    s = clean_text_encoding(s)
    if not s or s.upper() in INVALID_TOKENS_EXACT or INVALID_PATTERN.match(s):
        return ""
    return s


def format_address_title_case(addr: Any) -> str:
    """
    Format municipal hall address into clean Title Case while preserving
    Philippine LGU acronyms and address conventions.
    """
    if not addr:
        return ""
    s = clean_address(addr)
    if not s:
        return ""

    # Normalize double commas and spacing (e.g. 'Townsite,,Maluso' -> 'Townsite, Maluso')
    s = re.sub(r',+', ',', s)
    s = re.sub(r',([^\s\d])', r', \1', s)
    s = " ".join(s.split())

    def cap_word(w: str) -> str:
        if not w:
            return ""
        # Handle compound words with hyphens (e.g. 'LGU-Malungon', 'Sitio Lam-Ag')
        if '-' in w:
            return "-".join(cap_word(p) for p in w.split('-'))
        # Handle slashes (e.g. 'LGU/Town Hall')
        if '/' in w:
            return "/".join(cap_word(p) for p in w.split('/'))
        # Handle leading parenthesis or quotes
        if (w.startswith(('(', '"', "'")) and len(w) > 1):
            return w[0] + cap_word(w[1:])

        clean_upper = w.upper().rstrip('.,:;')
        punct = w[len(clean_upper):]

        # Philippine Address Acronyms and Roman Numerals
        if clean_upper in {'LGU', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X'}:
            return clean_upper + punct

        # Address abbreviation preservation
        abbr_map = {
            'POB': 'Pob.',
            'POBLACION': 'Poblacion',
            'BRGY': 'Brgy.',
            'BARANGAY': 'Brgy.',
            'BLDG': 'Bldg.',
            'NATL': "Nat'l",
            'NAT\'L': "Nat'l",
            'HIGHWAY': 'Hwy.',
            'HWY': 'Hwy.',
            'ST': 'St.',
            'AVE': 'Ave.',
            'RD': 'Rd.',
        }
        if clean_upper in abbr_map:
            replacement = abbr_map[clean_upper]
            if replacement.endswith('.') and punct.startswith('.'):
                punct = punct.lstrip('.')
            return replacement + punct

        return w.capitalize()

    tokens = s.split(' ')
    result = " ".join(cap_word(t) for t in tokens)
    result = re.sub(r'\.{2,}', '.', result)
    return result


def title_case_text(text: str) -> str:
    """Format text into Title Case while preserving hyphens and slashes."""
    if not text:
        return ""
    words = text.split()
    converted = []
    for w in words:
        if '-' in w:
            converted.append("-".join(part.capitalize() for part in w.split('-')))
        elif '/' in w:
            converted.append("/".join(part.capitalize() for part in w.split('/')))
        else:
            converted.append(w.capitalize())
    return " ".join(converted)


def get_middle_initial(middle: Any) -> str:
    """Extract clean middle initial with period, e.g. 'Dawal' -> 'D.'"""
    mn = clean_token(middle)
    if not mn:
        return ""
    s = mn.strip().rstrip('.')
    if not s:
        return ""
    return f"{s[0].upper()}."


def format_official_name(
    first: Any,
    middle: Any,
    last: Any,
    suffix: Any,
    name_case: str = "title",
    middle_as_initial: bool = True
) -> str:
    """
    Format full name as 'First [M.] Last [Suffix]'.
    name_case: 'title' | 'upper' | 'original'
    """
    fn = clean_token(first)
    ln = clean_token(last)
    suf = clean_suffix(suffix)

    if middle_as_initial:
        mi = get_middle_initial(middle)
    else:
        mi = clean_token(middle)

    parts = []
    if fn:
        parts.append(fn)
    if mi:
        parts.append(mi)
    if ln:
        parts.append(ln)
    if suf:
        parts.append(suf)

    full_name = " ".join(parts).strip()
    full_name = " ".join(full_name.split())

    if not full_name:
        return ""

    if name_case == "upper":
        return full_name.upper()
    elif name_case == "title":
        words = full_name.split()
        converted = []
        for w in words:
            # Preserve middle initial like 'D.'
            if len(w) == 2 and w[1] == '.':
                converted.append(w[0].upper() + '.')
                continue
            upper_w = w.upper().rstrip('.')
            if upper_w in {'JR', 'SR'}:
                converted.append('Jr.' if upper_w == 'JR' else 'Sr.')
            elif upper_w in ROMAN_NUMERALS:
                converted.append(upper_w)
            elif upper_w.startswith("MC") and len(w) > 2:
                converted.append("Mc" + w[2:].capitalize())
            elif '-' in w:
                converted.append("-".join(part.capitalize() for part in w.split('-')))
            else:
                converted.append(w.capitalize())
        return " ".join(converted)

    return full_name


REGION_ALIASES: Dict[str, str] = {
    "NCR": "NATIONAL CAPITAL REGION",
    "METRO MANILA": "NATIONAL CAPITAL REGION",
    "CAR": "CORDILLERA ADMINISTRATIVE REGION",
    "BARMM": "BANGSAMORO AUTONOMOUS REGION IN MUSLIM MINDANAO",
    "ARMM": "BANGSAMORO AUTONOMOUS REGION IN MUSLIM MINDANAO",
    "NIR": "NEGROS ISLAND REGION",
    "MIMAROPA": "MIMAROPA",
    "REGION 1": "REGION I - ILOCOS REGION",
    "REGION I": "REGION I - ILOCOS REGION",
    "ILOCOS": "REGION I - ILOCOS REGION",
    "REGION 2": "REGION II - CAGAYAN VALLEY",
    "REGION II": "REGION II - CAGAYAN VALLEY",
    "CAGAYAN VALLEY": "REGION II - CAGAYAN VALLEY",
    "REGION 3": "REGION III - CENTRAL LUZON",
    "REGION III": "REGION III - CENTRAL LUZON",
    "CENTRAL LUZON": "REGION III - CENTRAL LUZON",
    "REGION 4A": "REGION IV-A - CALABARZON",
    "REGION 4-A": "REGION IV-A - CALABARZON",
    "REGION IV-A": "REGION IV-A - CALABARZON",
    "REGION IVA": "REGION IV-A - CALABARZON",
    "CALABARZON": "REGION IV-A - CALABARZON",
    "REGION 4B": "MIMAROPA",
    "REGION 4-B": "MIMAROPA",
    "REGION IV-B": "MIMAROPA",
    "REGION 5": "REGION V - BICOL REGION",
    "REGION V": "REGION V - BICOL REGION",
    "BICOL": "REGION V - BICOL REGION",
    "REGION 6": "REGION VI - WESTERN VISAYAS",
    "REGION VI": "REGION VI - WESTERN VISAYAS",
    "WESTERN VISAYAS": "REGION VI - WESTERN VISAYAS",
    "REGION 7": "REGION VII - CENTRAL VISAYAS",
    "REGION VII": "REGION VII - CENTRAL VISAYAS",
    "CENTRAL VISAYAS": "REGION VII - CENTRAL VISAYAS",
    "REGION 8": "REGION VIII - EASTERN VISAYAS",
    "REGION VIII": "REGION VIII - EASTERN VISAYAS",
    "EASTERN VISAYAS": "REGION VIII - EASTERN VISAYAS",
    "REGION 9": "REGION IX - ZAMBOANGA PENINSULA",
    "REGION IX": "REGION IX - ZAMBOANGA PENINSULA",
    "ZAMBOANGA": "REGION IX - ZAMBOANGA PENINSULA",
    "REGION 10": "REGION X - NORTHERN MINDANAO",
    "REGION X": "REGION X - NORTHERN MINDANAO",
    "NORTHERN MINDANAO": "REGION X - NORTHERN MINDANAO",
    "REGION 11": "REGION XI - DAVAO REGION",
    "REGION XI": "REGION XI - DAVAO REGION",
    "DAVAO": "REGION XI - DAVAO REGION",
    "REGION 12": "REGION XII - SOCCSKSARGEN",
    "REGION XII": "REGION XII - SOCCSKSARGEN",
    "SOCCSKSARGEN": "REGION XII - SOCCSKSARGEN",
    "REGION 13": "REGION XIII - CARAGA",
    "REGION XIII": "REGION XIII - CARAGA",
    "CARAGA": "REGION XIII - CARAGA",
}


class OfficialsData:
    DEFAULT_CSV_COLUMNS = [
        "Mayor Name",
        "Municipal Hall Address",
        "Town",
        "Province",
        "Vice Mayor"
    ]

    def __init__(self, data_path: Optional[str] = None, cache_path: Optional[str] = None, excel_path: Optional[str] = None):
        target = data_path or excel_path
        if not target:
            xlsx_cand = os.path.join(os.path.dirname(os.path.abspath(__file__)), "officials_2025_2028.xlsx")
            csv_cand = os.path.join(os.path.dirname(os.path.abspath(__file__)), "officials_2025_2028.csv")
            target = xlsx_cand if os.path.exists(xlsx_cand) else (csv_cand if os.path.exists(csv_cand) else "officials_2025_2028.xlsx")

        self.data_path = os.path.abspath(target)
        self.excel_path = self.data_path  # Backwards compatibility alias

        if cache_path:
            self.cache_path = os.path.abspath(cache_path)
        else:
            base_dir = os.path.dirname(self.data_path)
            stem = os.path.splitext(os.path.basename(self.data_path))[0]
            self.cache_path = os.path.join(base_dir, f".{stem}_cache.json")

        self.regions: List[str] = []
        self.provinces_by_region: Dict[str, List[str]] = {}
        self.records: List[Dict[str, Any]] = []
        self._loaded = False

    def _is_cache_valid(self) -> bool:
        """Check if cached JSON exists and matches the source file's modified time and size."""
        if not os.path.exists(self.cache_path) or not os.path.exists(self.data_path):
            return False
        try:
            stat = os.stat(self.data_path)
            with open(self.cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                meta = data.get("meta", {})
                if meta.get("mtime") == stat.st_mtime and meta.get("size") == stat.st_size:
                    self._cached_data = data
                    return True
        except Exception:
            return False
        return False

    def _load_from_cache(self) -> bool:
        """Attempt to load parsed records from disk cache."""
        try:
            data = getattr(self, "_cached_data", None)
            if not data:
                with open(self.cache_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            self.records = data.get("records", [])
            self.regions = data.get("regions", [])
            self.provinces_by_region = data.get("provinces_by_region", {})
            self._loaded = True
            self._cached_data = None
            return True
        except Exception:
            return False

    def _save_to_cache(self):
        """Persist parsed records to JSON cache atomically."""
        try:
            stat = os.stat(self.data_path)
            cache_payload = {
                "meta": {
                    "mtime": stat.st_mtime,
                    "size": stat.st_size,
                    "cached_at": datetime.datetime.now().isoformat()
                },
                "regions": self.regions,
                "provinces_by_region": self.provinces_by_region,
                "records": self.records
            }
            temp_cache = self.cache_path + ".tmp"
            with open(temp_cache, "w", encoding="utf-8") as f:
                json.dump(cache_payload, f, ensure_ascii=False)
            if os.path.exists(self.cache_path):
                try:
                    os.remove(self.cache_path)
                except Exception:
                    pass
            shutil.move(temp_cache, self.cache_path)
        except Exception:
            # Cache failure is non-fatal
            pass

    def load(self, force_reload: bool = False):
        """
        Load and parse the data file (CSV or XLSX) with sub-50ms JSON caching.
        """
        if self._loaded and not force_reload:
            return

        if not os.path.exists(self.data_path):
            raise FileNotFoundError(f"Data file not found at: {self.data_path}")

        # Try fast cache load first
        if not force_reload and self._is_cache_valid():
            if self._load_from_cache():
                return

        # Parse CSV or Excel
        if self.data_path.lower().endswith(".csv"):
            self._load_from_csv()
        else:
            self._load_from_excel()

    def _load_from_csv(self):
        """Parse officials dataset from CSV format."""
        with open(self.data_path, "r", encoding="utf-8", errors="replace") as f:
            rows = list(csv.reader(f))

        header_idx = None
        for idx, r in enumerate(rows[:10]):
            r_upper = [c.strip().upper() for c in r]
            if any("REGION" in c for c in r_upper) and any("POSITION" in c for c in r_upper):
                header_idx = idx
                break

        if header_idx is None:
            raise ValueError(f"Could not find valid column header in CSV: {self.data_path}")

        headers = [h.strip().upper() for h in rows[header_idx]]

        def find_col(possible_names: List[str], default_idx: int) -> int:
            for name in possible_names:
                for idx, h in enumerate(headers):
                    if name in h:
                        return idx
            return default_idx

        reg_idx = find_col(["REGION"], 0)
        prov_idx = find_col(["PROVINCE"], 1)
        town_idx = find_col(["CITY/MUNICIPALITY", "MUNICIPALITY", "TOWN", "CITY"], 2)
        pos_idx = find_col(["POSITION"], 3)
        fn_idx = find_col(["FIRST NAME", "FIRST_NAME", "FNAME"], 5)
        mn_idx = find_col(["MIDDLE NAME", "MIDDLE_NAME", "MNAME"], 6)
        ln_idx = find_col(["LAST NAME", "LAST_NAME", "LNAME"], 7)
        suf_idx = find_col(["SUFFIX"], 8)
        addr_idx = find_col(["OFFICE ADDRESS", "ADDRESS"], 9)
        contact_idx = find_col(["OFFICE CONTACT", "CONTACT"], 10)
        email_idx = None
        for idx, h in enumerate(headers):
            if "EMAIL" in h:
                email_idx = idx
                break

        town_map: Dict[Tuple[str, str, str], Dict[str, Any]] = {}

        for row_idx, row in enumerate(rows[header_idx + 1:], start=header_idx + 1):
            if not row or not any(row):
                continue
            if len(row) <= max(reg_idx, prov_idx, town_idx):
                continue

            reg = clean_token(row[reg_idx])
            prov = clean_token(row[prov_idx])
            town = clean_token(row[town_idx])

            if not reg or not town:
                continue

            pos = str(row[pos_idx]).strip() if len(row) > pos_idx and row[pos_idx] else ""
            fn = row[fn_idx] if len(row) > fn_idx else None
            mn = row[mn_idx] if len(row) > mn_idx else None
            ln = row[ln_idx] if len(row) > ln_idx else None
            suf = row[suf_idx] if len(row) > suf_idx else None
            addr = row[addr_idx] if len(row) > addr_idx else None
            contact = clean_token(row[contact_idx]) if len(row) > contact_idx else ""
            email = clean_token(row[email_idx]) if email_idx is not None and len(row) > email_idx else ""

            key = (reg, prov, town)
            if key not in town_map:
                town_map[key] = {
                    "region": reg,
                    "province": prov,
                    "town": town,
                    "mayor": None,
                    "vice_mayor": None,
                    "sb_addresses": [],
                    "lgu_type": "Municipality"
                }

            entry = town_map[key]
            pos_upper = pos.upper()
            if "CITY" in pos_upper:
                entry["lgu_type"] = "City"

            if "MAYOR" in pos_upper and "VICE" not in pos_upper:
                if entry["mayor"] is None or (not entry["mayor"]["address"] and clean_address(addr)):
                    entry["mayor"] = {
                        "first": clean_token(fn),
                        "middle": clean_token(mn),
                        "last": clean_token(ln),
                        "suffix": clean_suffix(suf),
                        "address": clean_address(addr),
                        "contact": contact,
                        "email": email,
                        "row_idx": row_idx
                    }
            elif "VICE-MAYOR" in pos_upper or "VICE MAYOR" in pos_upper:
                if entry["vice_mayor"] is None or (not entry["vice_mayor"]["address"] and clean_address(addr)):
                    entry["vice_mayor"] = {
                        "first": clean_token(fn),
                        "middle": clean_token(mn),
                        "last": clean_token(ln),
                        "suffix": clean_suffix(suf),
                        "address": clean_address(addr),
                        "contact": contact,
                        "email": email,
                        "row_idx": row_idx
                    }
            else:
                clean_addr = clean_address(addr)
                if clean_addr and clean_addr not in entry["sb_addresses"]:
                    entry["sb_addresses"].append(clean_addr)

        self._build_records_from_map(town_map)

    def _load_from_excel(self):
        """Parse officials dataset from Excel (.xlsx) format using read_only=True with dynamic header detection."""
        wb = openpyxl.load_workbook(self.data_path, read_only=True, data_only=True)
        sheet = wb.active

        header_idx = None
        headers = []
        all_rows = list(sheet.iter_rows(values_only=True))
        wb.close()

        for idx, r in enumerate(all_rows[:10]):
            r_upper = [str(c).strip().upper() for c in r if c is not None]
            if any("REGION" in c for c in r_upper) and any("POSITION" in c for c in r_upper):
                header_idx = idx
                headers = [str(c).strip().upper() if c is not None else "" for c in r]
                break

        if header_idx is None:
            header_idx = 0
            headers = ["REGION", "PROVINCE", "MUNICIPALITY", "FIRST NAME", "MIDDLE NAME", "LAST NAME", "SUFFIX", "POSITION", "", "CONTACT", "ADDRESS"]

        def find_col(possible_names: List[str], default_idx: int) -> int:
            for name in possible_names:
                for idx, h in enumerate(headers):
                    if name in h:
                        return idx
            return default_idx

        reg_idx = find_col(["REGION"], 0)
        prov_idx = find_col(["PROVINCE"], 1)
        town_idx = find_col(["CITY/MUNICIPALITY", "MUNICIPALITY", "TOWN", "CITY"], 2)
        pos_idx = find_col(["POSITION"], 3 if len(headers) > 3 and "POSITION" in headers[3] else 7)
        fn_idx = find_col(["FIRST NAME", "FIRST_NAME"], 5 if len(headers) > 5 and "FIRST" in headers[5] else 3)
        mn_idx = find_col(["MIDDLE NAME", "MIDDLE_NAME"], 6 if len(headers) > 6 and "MIDDLE" in headers[6] else 4)
        ln_idx = find_col(["LAST NAME", "LAST_NAME"], 7 if len(headers) > 7 and "LAST" in headers[7] else 5)
        suf_idx = find_col(["SUFFIX"], 8 if len(headers) > 8 and "SUFFIX" in headers[8] else 6)
        addr_idx = find_col(["OFFICE ADDRESS", "ADDRESS"], 9 if len(headers) > 9 and "ADDRESS" in headers[9] else 10)
        contact_idx = find_col(["OFFICE CONTACT", "CONTACT"], 10 if len(headers) > 10 and "CONTACT" in headers[10] else 9)
        email_idx = None
        for idx, h in enumerate(headers):
            if "EMAIL" in h:
                email_idx = idx
                break

        town_map: Dict[Tuple[str, str, str], Dict[str, Any]] = {}

        for row_idx, row in enumerate(all_rows[header_idx + 1:], start=header_idx + 1):
            if not row or not any(row):
                continue
            if len(row) <= max(reg_idx, prov_idx, town_idx):
                continue

            reg = clean_token(row[reg_idx])
            prov = clean_token(row[prov_idx])
            town = clean_token(row[town_idx])

            if not reg or not town:
                continue

            pos = str(row[pos_idx]).strip() if len(row) > pos_idx and row[pos_idx] else ""
            fn = row[fn_idx] if len(row) > fn_idx else None
            mn = row[mn_idx] if len(row) > mn_idx else None
            ln = row[ln_idx] if len(row) > ln_idx else None
            suf = row[suf_idx] if len(row) > suf_idx else None
            addr = row[addr_idx] if len(row) > addr_idx else None
            contact = clean_token(row[contact_idx]) if len(row) > contact_idx else ""
            email = clean_token(row[email_idx]) if email_idx is not None and len(row) > email_idx else ""

            key = (reg, prov, town)
            if key not in town_map:
                town_map[key] = {
                    "region": reg,
                    "province": prov,
                    "town": town,
                    "mayor": None,
                    "vice_mayor": None,
                    "sb_addresses": [],
                    "lgu_type": "Municipality"
                }

            entry = town_map[key]
            pos_upper = pos.upper()
            if "CITY" in pos_upper:
                entry["lgu_type"] = "City"

            if "MAYOR" in pos_upper and "VICE" not in pos_upper:
                if entry["mayor"] is None or (not entry["mayor"]["address"] and clean_address(addr)):
                    entry["mayor"] = {
                        "first": clean_token(fn),
                        "middle": clean_token(mn),
                        "last": clean_token(ln),
                        "suffix": clean_suffix(suf),
                        "address": clean_address(addr),
                        "contact": contact,
                        "email": email,
                        "row_idx": row_idx
                    }
            elif "VICE-MAYOR" in pos_upper or "VICE MAYOR" in pos_upper:
                if entry["vice_mayor"] is None or (not entry["vice_mayor"]["address"] and clean_address(addr)):
                    entry["vice_mayor"] = {
                        "first": clean_token(fn),
                        "middle": clean_token(mn),
                        "last": clean_token(ln),
                        "suffix": clean_suffix(suf),
                        "address": clean_address(addr),
                        "contact": contact,
                        "email": email,
                        "row_idx": row_idx
                    }
            else:
                clean_addr = clean_address(addr)
                if clean_addr and clean_addr not in entry["sb_addresses"]:
                    entry["sb_addresses"].append(clean_addr)

        self._build_records_from_map(town_map)

    def _build_records_from_map(self, town_map: Dict[Tuple[str, str, str], Dict[str, Any]]):
        """Construct normalized records, sorted regions, and cache."""
        self.records = []
        reg_set = set()
        prov_map: Dict[str, set] = {}

        # Sort keys naturally: Region, Province, Town
        sorted_keys = sorted(town_map.keys(), key=lambda k: (k[0], k[1], k[2]))

        for key in sorted_keys:
            data = town_map[key]
            reg = data["region"]
            prov = data["province"]
            town = data["town"]

            reg_set.add(reg)
            if reg not in prov_map:
                prov_map[reg] = set()
            prov_map[reg].add(prov)

            mayor_data = data["mayor"]
            vm_data = data["vice_mayor"]

            # Address fallback logic
            address = ""
            if mayor_data and mayor_data.get("address"):
                address = mayor_data["address"]
            elif vm_data and vm_data.get("address"):
                address = vm_data["address"]
            elif data["sb_addresses"]:
                address = data["sb_addresses"][0]

            record = {
                "region": reg,
                "province": prov,
                "town": town,
                "lgu_type": data.get("lgu_type", "Municipality"),
                "raw_mayor": mayor_data,
                "raw_vm": vm_data,
                "address": address
            }
            self.records.append(record)

        self.regions = sorted(list(reg_set))
        self.provinces_by_region = {
            r: sorted(list(prov_map[r])) for r in self.regions
        }
        self._loaded = True
        self._save_to_cache()

    def get_provinces_for_region(self, region: str) -> List[str]:
        """Get sorted list of provinces for a region, or all provinces if region is empty or ALL."""
        if not region or region.upper() in {"ALL", "ALL REGIONS"}:
            all_provinces = set()
            for p_list in self.provinces_by_region.values():
                all_provinces.update(p_list)
            return sorted(list(all_provinces))
        target_reg = REGION_ALIASES.get(region.strip().upper(), region.strip().upper())
        if target_reg in self.provinces_by_region:
            return self.provinces_by_region[target_reg]
        for r_name, p_list in self.provinces_by_region.items():
            if target_reg == r_name or target_reg in r_name or region.strip().upper() in r_name:
                return p_list
        return []

    def filter_records(
        self,
        region: Optional[str] = None,
        province: Optional[str] = None,
        search_query: Optional[str] = None,
        name_case: str = "title",
        middle_as_initial: bool = True,
        include_contacts: bool = False,
        lgu_type: str = "all"
    ) -> List[Dict[str, str]]:
        """
        Filter records by region, province, search query, and LGU type ('all', 'municipality', 'city').
        Returns rows with formatted keys:
          'Mayor Name', 'Municipal Hall Address', 'Town', 'Province', 'Vice Mayor',
          'Region', 'LGU Type', and optionally 'Mayor Contact', 'Vice Mayor Contact', 'Mayor Email', 'Vice Mayor Email'.
        """
        if not self._loaded:
            self.load()

        is_all_reg = not region or region.strip().upper() in {"ALL", "ALL REGIONS"}
        is_all_prov = not province or province.strip().upper() in {"ALL", "ALL PROVINCES"}

        reg_filter = region.strip().upper() if not is_all_reg and region else None
        reg_target = REGION_ALIASES.get(reg_filter, reg_filter) if reg_filter else None
        prov_filter = province.strip().upper() if not is_all_prov and province else None
        query = search_query.strip().lower() if search_query else None
        lgu_filter = (lgu_type or "all").strip().lower()

        filtered = []
        for rec in self.records:
            rec_lgu_type = rec.get("lgu_type", "Municipality")
            if lgu_filter in {"city", "cities"} and rec_lgu_type != "City":
                continue
            if lgu_filter in {"municipality", "municipalities"} and rec_lgu_type != "Municipality":
                continue

            if reg_target:
                rec_reg_up = rec["region"].upper()
                if rec_reg_up != reg_target and reg_filter not in rec_reg_up and reg_target not in rec_reg_up:
                    continue
            if prov_filter:
                rec_prov_up = rec["province"].upper()
                if rec_prov_up != prov_filter and prov_filter not in rec_prov_up:
                    continue

            mayor_data = rec["raw_mayor"]
            vm_data = rec["raw_vm"]

            mayor_name = ""
            mayor_contact = ""
            if mayor_data:
                mayor_name = format_official_name(
                    mayor_data.get("first"), mayor_data.get("middle"),
                    mayor_data.get("last"), mayor_data.get("suffix"),
                    name_case=name_case,
                    middle_as_initial=middle_as_initial
                )
                mayor_contact = mayor_data.get("contact", "")

            vm_name = ""
            vm_contact = ""
            if vm_data:
                vm_name = format_official_name(
                    vm_data.get("first"), vm_data.get("middle"),
                    vm_data.get("last"), vm_data.get("suffix"),
                    name_case=name_case,
                    middle_as_initial=middle_as_initial
                )
                vm_contact = vm_data.get("contact", "")

            town = rec["town"]
            prov = rec["province"]
            reg = rec["region"]
            address = rec["address"]

            if name_case == "title":
                town = title_case_text(town)
                prov = title_case_text(prov)
                address = format_address_title_case(address)
            elif name_case == "upper":
                town = town.upper()
                prov = prov.upper()
                address = address.upper()
            else:
                address = format_address_title_case(address)

            if query:
                searchable = f"{town} {prov} {reg} {rec_lgu_type} {mayor_name} {vm_name} {address} {mayor_contact} {vm_contact}".lower()
                if query not in searchable:
                    continue

            item = {
                "Mayor Name": mayor_name,
                "Municipal Hall Address": address,
                "Town": town,
                "Province": prov,
                "Vice Mayor": vm_name,
                "Region": reg,
                "LGU Type": rec_lgu_type
            }
            if include_contacts:
                item["Mayor Contact"] = mayor_contact
                item["Vice Mayor Contact"] = vm_contact
                mayor_email = mayor_data.get("email", "") if mayor_data else ""
                vm_email = vm_data.get("email", "") if vm_data else ""
                if mayor_email:
                    item["Mayor Email"] = mayor_email
                if vm_email:
                    item["Vice Mayor Email"] = vm_email

            filtered.append(item)

        return filtered

    def get_statistics(self) -> Dict[str, Any]:
        """Generate comprehensive data coverage statistics."""
        if not self._loaded:
            self.load()

        total = len(self.records)
        all_provinces = set()
        for p_list in self.provinces_by_region.values():
            all_provinces.update(p_list)

        cities_count = sum(1 for r in self.records if r.get("lgu_type") == "City")
        munis_count = sum(1 for r in self.records if r.get("lgu_type") != "City")
        mayors_count = sum(1 for r in self.records if r.get("raw_mayor"))
        vm_count = sum(1 for r in self.records if r.get("raw_vm"))
        addr_count = sum(1 for r in self.records if r.get("address"))
        mayor_contact_count = sum(1 for r in self.records if r.get("raw_mayor") and r["raw_mayor"].get("contact"))

        missing_mayors = [
            {"town": r["town"], "province": r["province"], "region": r["region"], "lgu_type": r.get("lgu_type", "Municipality")}
            for r in self.records if not r.get("raw_mayor")
        ]
        missing_vms = [
            {"town": r["town"], "province": r["province"], "region": r["region"], "lgu_type": r.get("lgu_type", "Municipality")}
            for r in self.records if not r.get("raw_vm")
        ]

        return {
            "total_lgus": total,
            "total_municipalities": munis_count,
            "total_cities": cities_count,
            "total_regions": len(self.regions),
            "total_provinces": len(all_provinces),
            "mayors_count": mayors_count,
            "vice_mayors_count": vm_count,
            "addresses_count": addr_count,
            "mayor_contact_count": mayor_contact_count,
            "mayor_coverage_pct": round((mayors_count / total * 100), 1) if total else 0.0,
            "vm_coverage_pct": round((vm_count / total * 100), 1) if total else 0.0,
            "address_coverage_pct": round((addr_count / total * 100), 1) if total else 0.0,
            "missing_mayors": missing_mayors,
            "missing_vice_mayors": missing_vms
        }

    def _get_export_columns(
        self,
        include_region: bool = False,
        include_contacts: bool = False,
        include_lgu_type: bool = False,
        include_verified_col: bool = False,
        records: Optional[List[Dict[str, str]]] = None
    ) -> List[str]:
        """Construct the ordered list of export column headers."""
        cols = list(self.DEFAULT_CSV_COLUMNS)
        if include_lgu_type:
            if "Town" in cols:
                idx = cols.index("Town") + 1
                cols.insert(idx, "LGU Type")
            else:
                cols.append("LGU Type")
        if include_region:
            cols.append("Region")
        if include_contacts:
            cols.extend(["Mayor Contact", "Vice Mayor Contact"])
            if records and any("Mayor Email" in r or "Vice Mayor Email" in r for r in records):
                cols.extend(["Mayor Email", "Vice Mayor Email"])
        if include_verified_col:
            cols.append("Contact Verified")
        return cols

    def export_csv(
        self,
        filepath: str,
        records: List[Dict[str, str]],
        include_region: bool = False,
        include_contacts: bool = False,
        include_lgu_type: bool = False,
        include_verified_col: bool = False,
        verification_map: Optional[Dict[str, str]] = None
    ) -> int:
        """
        Export records to a CSV file in UTF-8 with BOM (utf-8-sig) for seamless Excel compatibility.
        """
        fieldnames = self._get_export_columns(
            include_region=include_region,
            include_contacts=include_contacts,
            include_lgu_type=include_lgu_type,
            include_verified_col=include_verified_col,
            records=records
        )
        os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)

        vmap = verification_map or {}

        with open(filepath, mode="w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=fieldnames,
                extrasaction="ignore",
                quoting=csv.QUOTE_MINIMAL
            )
            writer.writeheader()
            for r in records:
                row = dict(r)
                if include_verified_col:
                    key = f"{r.get('Province', '')}|{r.get('Town', '')}"
                    row["Contact Verified"] = "Yes" if vmap.get(key) else "No"
                writer.writerow(row)

        return len(records)

    def export_xlsx(
        self,
        filepath: str,
        records: List[Dict[str, str]],
        include_region: bool = False,
        include_contacts: bool = False,
        include_lgu_type: bool = False,
        include_verified_col: bool = False,
        verification_map: Optional[Dict[str, str]] = None
    ) -> int:
        """
        Export records to a professionally styled Excel workbook (.xlsx).
        Features: Dark blue header, bold white text, freeze pane, alternating rows, auto-fit columns.
        When include_verified_col=True: adds a 'Contact Verified' column and applies green highlights
        to rows whose contact info has been verified.
        """
        fieldnames = self._get_export_columns(
            include_region=include_region,
            include_contacts=include_contacts,
            include_lgu_type=include_lgu_type,
            include_verified_col=include_verified_col,
            records=records
        )
        if isinstance(filepath, str):
            os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)

        vmap = verification_map or {}

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Municipal Officials"

        # Styles
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        row_font = Font(name="Segoe UI", size=10)
        even_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        odd_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        verified_fill = PatternFill(start_color="D8F4D8", end_color="D8F4D8", fill_type="solid")
        verified_yes_font = Font(name="Segoe UI", size=10, bold=True, color="166534")
        thin_border = Border(
            left=Side(style='thin', color='E2E8F0'),
            right=Side(style='thin', color='E2E8F0'),
            top=Side(style='thin', color='E2E8F0'),
            bottom=Side(style='thin', color='E2E8F0')
        )

        # Write Header
        ws.row_dimensions[1].height = 28
        ws.append(fieldnames)
        for col_idx in range(1, len(fieldnames) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # Find verified_col_idx once
        verified_col_idx = None
        if include_verified_col and "Contact Verified" in fieldnames:
            verified_col_idx = fieldnames.index("Contact Verified") + 1

        # Write Rows
        for row_idx, r in enumerate(records, start=2):
            ws.row_dimensions[row_idx].height = 20

            # Determine verification status
            key = f"{r.get('Province', '')}|{r.get('Town', '')}"
            is_verified = bool(vmap.get(key)) if include_verified_col else False

            row = dict(r)
            if include_verified_col:
                row["Contact Verified"] = "Yes" if is_verified else "No"

            row_data = [row.get(col, "") for col in fieldnames]
            ws.append(row_data)

            row_fill = verified_fill if is_verified else (even_fill if row_idx % 2 == 0 else odd_fill)
            for col_idx in range(1, len(fieldnames) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.fill = row_fill
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
                # Bold green font for the 'Yes' cell in verified column
                if is_verified and col_idx == verified_col_idx:
                    cell.font = verified_yes_font
                else:
                    cell.font = row_font

        # Freeze Panes below header
        ws.freeze_panes = "A2"

        # Auto-fit Column Widths
        for col_idx, col_name in enumerate(fieldnames, start=1):
            max_len = len(col_name)
            for r in records:
                val_len = len(str(r.get(col_name, "")))
                if val_len > max_len:
                    max_len = val_len
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 45)

        if isinstance(filepath, str):
            os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)
            wb.save(filepath)
        else:
            wb.save(filepath)

        return len(records)

    def export_json(
        self,
        filepath: str,
        records: List[Dict[str, str]],
        include_region: bool = False,
        include_contacts: bool = False,
        include_lgu_type: bool = False,
        include_verified_col: bool = False,
        verification_map: Optional[Dict[str, str]] = None
    ) -> int:
        """Export records to formatted JSON."""
        fieldnames = self._get_export_columns(
            include_region=include_region,
            include_contacts=include_contacts,
            include_lgu_type=include_lgu_type,
            include_verified_col=include_verified_col,
            records=records
        )
        os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)

        vmap = verification_map or {}
        filtered_data = []
        for r in records:
            row = {col: r.get(col, "") for col in fieldnames}
            if include_verified_col:
                key = f"{r.get('Province', '')}|{r.get('Town', '')}"
                row["Contact Verified"] = "Yes" if vmap.get(key) else "No"
            filtered_data.append(row)

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(filtered_data, f, ensure_ascii=False, indent=2)

        return len(records)

    def export_all_regions(
        self,
        output_dir: str = "exports",
        file_format: str = "csv",
        include_region: bool = True,
        include_contacts: bool = False,
        include_lgu_type: bool = False,
        lgu_type: str = "all",
        name_case: str = "title"
    ) -> List[str]:
        """
        Batch-export separate official files for all 18 regions into output_dir.
        Returns a list of generated file paths.
        """
        if not self._loaded:
            self.load()

        os.makedirs(output_dir, exist_ok=True)
        generated_files = []

        for reg in self.regions:
            recs = self.filter_records(
                region=reg,
                name_case=name_case,
                include_contacts=include_contacts,
                lgu_type=lgu_type
            )
            reg_slug = re.sub(r'[^A-Za-z0-9]+', '_', reg).strip('_')
            filename = f"Officials_{reg_slug}.{file_format.lower()}"
            out_path = os.path.join(output_dir, filename)

            if file_format.lower() == "xlsx":
                self.export_xlsx(out_path, recs, include_region=include_region, include_contacts=include_contacts, include_lgu_type=include_lgu_type)
            elif file_format.lower() == "json":
                self.export_json(out_path, recs, include_region=include_region, include_contacts=include_contacts, include_lgu_type=include_lgu_type)
            else:
                self.export_csv(out_path, recs, include_region=include_region, include_contacts=include_contacts, include_lgu_type=include_lgu_type)

            generated_files.append(out_path)

        return generated_files

    def get_municipality_raw(self, town: str, province: str) -> Optional[Dict[str, Any]]:
        """Get raw municipality dictionary by town and province."""
        if not self._loaded:
            self.load()
        t_clean = town.strip().upper()
        p_clean = province.strip().upper()
        for rec in self.records:
            if rec["town"].strip().upper() == t_clean and rec["province"].strip().upper() == p_clean:
                return rec
        return None

    def update_municipality(
        self,
        town: str,
        province: str,
        mayor_first: str,
        mayor_middle: str,
        mayor_last: str,
        mayor_suffix: str,
        mayor_address: str,
        mayor_contact: str = "",
        vm_first: str = "",
        vm_middle: str = "",
        vm_last: str = "",
        vm_suffix: str = "",
        vm_address: str = "",
        vm_contact: str = "",
        mayor_email: str = "",
        vm_email: str = "",
        backup: bool = True
    ) -> bool:
        """
        Update Mayor and Vice Mayor info for a town, creating rolling timestamped backups.
        Updates CSV or Excel data source and refreshes memory/cache.
        """
        rec = self.get_municipality_raw(town, province)
        if not rec:
            return False

        has_mayor_data = any([mayor_first, mayor_middle, mayor_last, mayor_suffix, mayor_contact, mayor_address, mayor_email])
        has_vm_data = any([vm_first, vm_middle, vm_last, vm_suffix, vm_contact, vm_address, vm_email])

        is_csv = self.data_path.lower().endswith(".csv")

        # Rolling Backup System
        if backup and os.path.exists(self.data_path):
            base_dir = os.path.dirname(os.path.abspath(self.data_path))
            backup_dir = os.path.join(base_dir, "backups")
            os.makedirs(backup_dir, exist_ok=True)

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            stem = os.path.splitext(os.path.basename(self.data_path))[0]
            ext = ".csv" if is_csv else ".xlsx"
            backup_file = os.path.join(backup_dir, f"{stem}_{timestamp}{ext}")
            try:
                shutil.copy2(self.data_path, backup_file)
                shutil.copy2(self.data_path, self.data_path + f".backup{ext}")

                # Prune oldest backups if > 10
                all_backups = sorted([
                    os.path.join(backup_dir, f)
                    for f in os.listdir(backup_dir)
                    if f.startswith(f"{stem}_") and f.endswith(ext)
                ])
                if len(all_backups) > 10:
                    for old_b in all_backups[:-10]:
                        try:
                            os.remove(old_b)
                        except Exception:
                            pass
            except Exception:
                pass

        if is_csv:
            # Update CSV file
            with open(self.data_path, "r", encoding="utf-8", errors="replace") as f:
                rows = list(csv.reader(f))

            header_idx = None
            for idx, r in enumerate(rows[:10]):
                r_upper = [c.strip().upper() for c in r]
                if any("REGION" in c for c in r_upper) and any("POSITION" in c for c in r_upper):
                    header_idx = idx
                    break
            if header_idx is None:
                header_idx = 0

            headers = [h.strip().upper() for h in rows[header_idx]]

            def find_col(possible_names: List[str], default_idx: int) -> int:
                for name in possible_names:
                    for idx, h in enumerate(headers):
                        if name in h:
                            return idx
                return default_idx

            reg_idx = find_col(["REGION"], 0)
            prov_idx = find_col(["PROVINCE"], 1)
            town_idx = find_col(["CITY/MUNICIPALITY", "MUNICIPALITY", "TOWN", "CITY"], 2)
            pos_idx = find_col(["POSITION"], 3)
            year_idx = find_col(["ELECTION YEAR", "YEAR"], 4)
            fn_idx = find_col(["FIRST NAME", "FIRST_NAME"], 5)
            mn_idx = find_col(["MIDDLE NAME", "MIDDLE_NAME"], 6)
            ln_idx = find_col(["LAST NAME", "LAST_NAME"], 7)
            suf_idx = find_col(["SUFFIX"], 8)
            addr_idx = find_col(["OFFICE ADDRESS", "ADDRESS"], 9)
            contact_idx = find_col(["OFFICE CONTACT", "CONTACT"], 10)
            email_idx = None
            for idx, h in enumerate(headers):
                if "EMAIL" in h:
                    email_idx = idx
                    break

            # Update Mayor in CSV
            m_updated = False
            for r in rows[header_idx + 1:]:
                if len(r) > max(prov_idx, town_idx, pos_idx):
                    if r[town_idx].strip().upper() == town.strip().upper() and r[prov_idx].strip().upper() == province.strip().upper():
                        p = r[pos_idx].strip().upper()
                        if "MAYOR" in p and "VICE" not in p:
                            while len(r) <= max(fn_idx, mn_idx, ln_idx, suf_idx, addr_idx, contact_idx, email_idx or 0):
                                r.append("")
                            r[fn_idx] = mayor_first.strip()
                            r[mn_idx] = mayor_middle.strip()
                            r[ln_idx] = mayor_last.strip()
                            r[suf_idx] = mayor_suffix.strip()
                            r[addr_idx] = mayor_address.strip()
                            r[contact_idx] = mayor_contact.strip()
                            if email_idx is not None and mayor_email:
                                r[email_idx] = mayor_email.strip()
                            m_updated = True
                            break

            if not m_updated and has_mayor_data:
                new_row = [""] * max(len(headers), 12)
                new_row[reg_idx] = rec["region"]
                new_row[prov_idx] = rec["province"]
                new_row[town_idx] = rec["town"]
                new_row[pos_idx] = "MUNICIPAL MAYOR"
                if year_idx < len(new_row): new_row[year_idx] = "2025-2028"
                new_row[fn_idx] = mayor_first.strip()
                new_row[mn_idx] = mayor_middle.strip()
                new_row[ln_idx] = mayor_last.strip()
                new_row[suf_idx] = mayor_suffix.strip()
                new_row[addr_idx] = mayor_address.strip()
                new_row[contact_idx] = mayor_contact.strip()
                if email_idx is not None and email_idx < len(new_row):
                    new_row[email_idx] = mayor_email.strip()
                rows.append(new_row)

            # Update Vice Mayor in CSV
            vm_updated = False
            for r in rows[header_idx + 1:]:
                if len(r) > max(prov_idx, town_idx, pos_idx):
                    if r[town_idx].strip().upper() == town.strip().upper() and r[prov_idx].strip().upper() == province.strip().upper():
                        p = r[pos_idx].strip().upper()
                        if "VICE-MAYOR" in p or "VICE MAYOR" in p:
                            while len(r) <= max(fn_idx, mn_idx, ln_idx, suf_idx, addr_idx, contact_idx, email_idx or 0):
                                r.append("")
                            r[fn_idx] = vm_first.strip()
                            r[mn_idx] = vm_middle.strip()
                            r[ln_idx] = vm_last.strip()
                            r[suf_idx] = vm_suffix.strip()
                            r[addr_idx] = vm_address.strip()
                            r[contact_idx] = vm_contact.strip()
                            if email_idx is not None and vm_email:
                                r[email_idx] = vm_email.strip()
                            vm_updated = True
                            break

            if not vm_updated and has_vm_data:
                new_row = [""] * max(len(headers), 12)
                new_row[reg_idx] = rec["region"]
                new_row[prov_idx] = rec["province"]
                new_row[town_idx] = rec["town"]
                new_row[pos_idx] = "MUNICIPAL VICE-MAYOR"
                if year_idx < len(new_row): new_row[year_idx] = "2025-2028"
                new_row[fn_idx] = vm_first.strip()
                new_row[mn_idx] = vm_middle.strip()
                new_row[ln_idx] = vm_last.strip()
                new_row[suf_idx] = vm_suffix.strip()
                new_row[addr_idx] = vm_address.strip()
                new_row[contact_idx] = vm_contact.strip()
                if email_idx is not None and email_idx < len(new_row):
                    new_row[email_idx] = vm_email.strip()
                rows.append(new_row)

            tmp_path = self.data_path + ".tmp"
            with open(tmp_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(rows)
            if os.path.exists(self.data_path):
                try:
                    os.remove(self.data_path)
                except Exception:
                    pass
            shutil.move(tmp_path, self.data_path)

        else:
            # Update Excel (.xlsx) file
            wb = openpyxl.load_workbook(self.excel_path)
            sheet = wb.active

            mayor_info = rec.get("raw_mayor")
            if mayor_info and mayor_info.get("row_idx"):
                m_row = mayor_info["row_idx"]
                sheet.cell(row=m_row, column=4, value=mayor_first.strip() if mayor_first else None)
                sheet.cell(row=m_row, column=5, value=mayor_middle.strip() if mayor_middle else None)
                sheet.cell(row=m_row, column=6, value=mayor_last.strip() if mayor_last else None)
                sheet.cell(row=m_row, column=7, value=mayor_suffix.strip() if mayor_suffix else None)
                sheet.cell(row=m_row, column=10, value=mayor_contact.strip() if mayor_contact else None)
                sheet.cell(row=m_row, column=11, value=mayor_address.strip() if mayor_address else None)
            elif has_mayor_data:
                m_row = sheet.max_row + 1
                sheet.cell(row=m_row, column=1, value=rec["region"])
                sheet.cell(row=m_row, column=2, value=rec["province"])
                sheet.cell(row=m_row, column=3, value=rec["town"])
                sheet.cell(row=m_row, column=4, value=mayor_first.strip() if mayor_first else None)
                sheet.cell(row=m_row, column=5, value=mayor_middle.strip() if mayor_middle else None)
                sheet.cell(row=m_row, column=6, value=mayor_last.strip() if mayor_last else None)
                sheet.cell(row=m_row, column=7, value=mayor_suffix.strip() if mayor_suffix else None)
                sheet.cell(row=m_row, column=8, value="MUNICIPAL MAYOR")
                sheet.cell(row=m_row, column=10, value=mayor_contact.strip() if mayor_contact else None)
                sheet.cell(row=m_row, column=11, value=mayor_address.strip() if mayor_address else None)
                if not mayor_info:
                    rec["raw_mayor"] = {}
                rec["raw_mayor"]["row_idx"] = m_row

            vm_info = rec.get("raw_vm")
            if vm_info and vm_info.get("row_idx"):
                v_row = vm_info["row_idx"]
                sheet.cell(row=v_row, column=4, value=vm_first.strip() if vm_first else None)
                sheet.cell(row=v_row, column=5, value=vm_middle.strip() if vm_middle else None)
                sheet.cell(row=v_row, column=6, value=vm_last.strip() if vm_last else None)
                sheet.cell(row=v_row, column=7, value=vm_suffix.strip() if vm_suffix else None)
                sheet.cell(row=v_row, column=10, value=vm_contact.strip() if vm_contact else None)
                sheet.cell(row=v_row, column=11, value=vm_address.strip() if vm_address else None)
            elif has_vm_data:
                v_row = sheet.max_row + 1
                sheet.cell(row=v_row, column=1, value=rec["region"])
                sheet.cell(row=v_row, column=2, value=rec["province"])
                sheet.cell(row=v_row, column=3, value=rec["town"])
                sheet.cell(row=v_row, column=4, value=vm_first.strip() if vm_first else None)
                sheet.cell(row=v_row, column=5, value=vm_middle.strip() if vm_middle else None)
                sheet.cell(row=v_row, column=6, value=vm_last.strip() if vm_last else None)
                sheet.cell(row=v_row, column=7, value=vm_suffix.strip() if vm_suffix else None)
                sheet.cell(row=v_row, column=8, value="MUNICIPAL VICE-MAYOR")
                sheet.cell(row=v_row, column=10, value=vm_contact.strip() if vm_contact else None)
                sheet.cell(row=v_row, column=11, value=vm_address.strip() if vm_address else None)
                if not vm_info:
                    rec["raw_vm"] = {}
                rec["raw_vm"]["row_idx"] = v_row

            wb.save(self.excel_path)
            wb.close()

        # Update in-memory record
        if not rec.get("raw_mayor") and has_mayor_data:
            rec["raw_mayor"] = {}
        if rec.get("raw_mayor"):
            rec["raw_mayor"].update({
                "first": clean_token(mayor_first),
                "middle": clean_token(mayor_middle),
                "last": clean_token(mayor_last),
                "suffix": clean_suffix(mayor_suffix),
                "address": clean_address(mayor_address),
                "contact": clean_token(mayor_contact),
                "email": clean_token(mayor_email)
            })

        if not rec.get("raw_vm") and has_vm_data:
            rec["raw_vm"] = {}
        if rec.get("raw_vm"):
            rec["raw_vm"].update({
                "first": clean_token(vm_first),
                "middle": clean_token(vm_middle),
                "last": clean_token(vm_last),
                "suffix": clean_suffix(vm_suffix),
                "address": clean_address(vm_address),
                "contact": clean_token(vm_contact),
                "email": clean_token(vm_email)
            })

        rec["address"] = clean_address(mayor_address) or clean_address(vm_address) or rec.get("address", "")

        # Re-save updated cache
        self._save_to_cache()
        return True