"""
Data processor for Philippine Local Government Officials List.
Parses Officials_List.xlsx and groups officials by (Region, Province, Town).
Extracts Mayor Name, Municipal Hall Address, Town, Province, Vice Mayor.
"""

import csv
import io
import re
import os
import openpyxl
from typing import Dict, List, Optional, Tuple, Any

INVALID_TOKENS = {'N/A', 'NA', 'NONE', 'N.A.', '-', '"N/A"', "'N/A'", 'NULL', ',n/a', 'N/a', 'n'}

def clean_text_encoding(text: str) -> str:
    """Fix replacement characters (\\ufffd) commonly found in DILG exports for Ñ/ñ."""
    if not text or '\ufffd' not in text:
        return text
    result = []
    i = 0
    while i < len(text):
        if text[i] == '\ufffd':
            if i + 1 < len(text) and text[i + 1].islower():
                result.append('ñ')
            else:
                result.append('Ñ')
        else:
            result.append(text[i])
        i += 1
    return "".join(result)

def clean_token(val: Any) -> str:
    """Strip whitespace, fix encoding, and return cleaned string or empty if invalid."""
    if val is None:
        return ""
    s = str(val).strip()
    s = clean_text_encoding(s)
    if s.upper() in INVALID_TOKENS:
        return ""
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        s = s[1:-1].strip()
    return s

def clean_address(val: Any) -> str:
    """Clean address string, return empty if invalid."""
    if val is None:
        return ""
    s = str(val).strip()
    s = clean_text_encoding(s)
    if s.upper() in INVALID_TOKENS:
        return ""
    return s

def format_address_title_case(addr: Any) -> str:
    """Format municipal hall address into clean Title Case."""
    if not addr:
        return ""
    s = clean_address(addr)
    if not s:
        return ""
        
    # Ensure space after commas if missing (e.g., 'Townsite,Maluso' -> 'Townsite, Maluso')
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
        if (w.startswith('(') or w.startswith('"') or w.startswith("'")) and len(w) > 1:
            return w[0] + cap_word(w[1:])
            
        clean_upper = w.upper().rstrip('.,:;')
        if clean_upper in {'LGU', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII'}:
            punct = w[len(clean_upper):]
            return clean_upper + punct
            
        return w.capitalize()
        
    tokens = s.split(' ')
    return " ".join(cap_word(t) for t in tokens)

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
    suf = clean_token(suffix)
    
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
            elif upper_w in {'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII'}:
                converted.append(upper_w)
            elif upper_w.startswith("MC") and len(w) > 2:
                converted.append("Mc" + w[2:].capitalize())
            elif '-' in w:
                converted.append("-".join(part.capitalize() for part in w.split('-')))
            else:
                converted.append(w.capitalize())
        return " ".join(converted)
        
    return full_name

class OfficialsData:
    DEFAULT_CSV_COLUMNS = [
        "Mayor Name",
        "Municipal Hall Address",
        "Town",
        "Province",
        "Vice Mayor"
    ]

    def __init__(self, excel_path: str = "Officials_List.xlsx"):
        self.excel_path = excel_path
        self.regions: List[str] = []
        self.provinces_by_region: Dict[str, List[str]] = {}
        self.records: List[Dict[str, Any]] = []
        self._loaded = False
        
    def load(self, force_reload: bool = False):
        """Load and parse the Excel file."""
        if self._loaded and not force_reload:
            return
            
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"Excel file not found at: {self.excel_path}")
            
        wb = openpyxl.load_workbook(self.excel_path, data_only=True)
        sheet = wb.active
        
        town_map: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            
            reg = clean_token(row[0])
            prov = clean_token(row[1])
            town = clean_token(row[2])
            
            if not reg and not prov and not town:
                continue
                
            fn = row[3]
            mn = row[4]
            ln = row[5]
            suf = row[6]
            pos = str(row[7]).strip() if row[7] is not None else ""
            contact = clean_token(row[9])
            addr = row[10]
            
            key = (reg, prov, town)
            if key not in town_map:
                town_map[key] = {
                    "region": reg,
                    "province": prov,
                    "town": town,
                    "mayor": None,
                    "vice_mayor": None,
                    "sb_addresses": []
                }
                
            entry = town_map[key]
            
            pos_upper = pos.upper()
            if "MAYOR" in pos_upper and "VICE" not in pos_upper:
                if entry["mayor"] is None or (not entry["mayor"]["address"] and clean_address(addr)):
                    entry["mayor"] = {
                        "first": fn, "middle": mn, "last": ln, "suffix": suf,
                        "address": clean_address(addr), "contact": contact
                    }
            elif "VICE-MAYOR" in pos_upper or "VICE MAYOR" in pos_upper:
                if entry["vice_mayor"] is None or (not entry["vice_mayor"]["address"] and clean_address(addr)):
                    entry["vice_mayor"] = {
                        "first": fn, "middle": mn, "last": ln, "suffix": suf,
                        "address": clean_address(addr), "contact": contact
                    }
            else:
                clean_addr = clean_address(addr)
                if clean_addr and clean_addr not in entry["sb_addresses"]:
                    entry["sb_addresses"].append(clean_addr)
                    
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
            
            # Determine best address
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

    def get_provinces_for_region(self, region: str) -> List[str]:
        """Get sorted list of provinces for a region, or all provinces if region is empty or ALL."""
        if not region or region.upper() in {"ALL", "ALL REGIONS"}:
            all_provinces = set()
            for p_list in self.provinces_by_region.values():
                all_provinces.update(p_list)
            return sorted(list(all_provinces))
        return self.provinces_by_region.get(region, [])

    def filter_records(
        self,
        region: Optional[str] = None,
        province: Optional[str] = None,
        search_query: Optional[str] = None,
        name_case: str = "title",
        middle_as_initial: bool = True
    ) -> List[Dict[str, str]]:
        """
        Filter records by region, province, and optional search query.
        Returns rows with formatted keys:
          'Mayor Name', 'Municipal Hall Address', 'Town', 'Province', 'Vice Mayor', 'Region'
        """
        if not self._loaded:
            self.load()
            
        is_all_reg = not region or region.strip().upper() in {"ALL", "ALL REGIONS"}
        is_all_prov = not province or province.strip().upper() in {"ALL", "ALL PROVINCES"}
        
        reg_filter = region.strip().upper() if not is_all_reg and region else None
        prov_filter = province.strip().upper() if not is_all_prov and province else None
        query = search_query.strip().lower() if search_query else None
        
        filtered = []
        for rec in self.records:
            if reg_filter and rec["region"].upper() != reg_filter:
                continue
            if prov_filter and rec["province"].upper() != prov_filter:
                continue
                
            mayor_data = rec["raw_mayor"]
            vm_data = rec["raw_vm"]
            
            mayor_name = ""
            if mayor_data:
                mayor_name = format_official_name(
                    mayor_data["first"], mayor_data["middle"],
                    mayor_data["last"], mayor_data["suffix"],
                    name_case=name_case,
                    middle_as_initial=middle_as_initial
                )
                
            vm_name = ""
            if vm_data:
                vm_name = format_official_name(
                    vm_data["first"], vm_data["middle"],
                    vm_data["last"], vm_data["suffix"],
                    name_case=name_case,
                    middle_as_initial=middle_as_initial
                )
                
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
                searchable = f"{town} {prov} {reg} {mayor_name} {vm_name} {address}".lower()
                if query not in searchable:
                    continue
                    
            filtered.append({
                "Mayor Name": mayor_name,
                "Municipal Hall Address": address,
                "Town": town,
                "Province": prov,
                "Vice Mayor": vm_name,
                "Region": reg
            })
            
        return filtered

    def export_csv(
        self,
        filepath: str,
        records: List[Dict[str, str]],
        include_region: bool = False
    ) -> int:
        """
        Export records to a CSV file.
        Uses 'utf-8-sig' for seamless Microsoft Excel compatibility.
        Returns the number of rows exported.
        """
        fieldnames = list(self.DEFAULT_CSV_COLUMNS)
        if include_region:
            fieldnames.append("Region")
            
        with open(filepath, mode="w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=fieldnames,
                extrasaction="ignore",
                quoting=csv.QUOTE_MINIMAL
            )
            writer.writeheader()
            for r in records:
                writer.writerow(r)
                
        return len(records)
