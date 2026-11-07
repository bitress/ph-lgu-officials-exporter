"""
Philippine Municipal Officials Directory & CSV/Excel/JSON Exporter
A Python application to filter, view, edit, and export LGU Municipal Officials
(Mayor Name, Municipal Hall Address, Town, Province, Vice Mayor) by Region and Province.

Features:
- Sub-50ms instant startup via automatic JSON caching (.officials_cache.json)
- Multi-format exporters: CSV (Excel-compatible UTF-8 BOM), Excel (.xlsx styled workbook), JSON
- Optional Contact Number inclusion (--include-contact)
- Batch-export all 18 regions (--split-regions)
- Data health audit command (--stats)
- Modern interactive Desktop GUI (Tkinter) with responsive auto-fit columns & keyboard shortcuts
- Modern Web App (--web) built with Bootstrap 5, FontAwesome, pagination, sorting, metrics dashboard, and copy tools
- Direct Excel (.xlsx) editing with rolling timestamped backups
"""

import sys
import os
import io
import json
import datetime
import argparse
import webbrowser
from typing import Optional, List, Dict, Any

# Ensure UTF-8 console output for Philippine characters on Windows
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Local data processor
from lgu.data_processor import OfficialsData


def run_cli(args: argparse.Namespace, data: OfficialsData):
    """Run in command-line mode without GUI."""
    data.load()

    # 1. Print Statistics / Data Health Audit
    if getattr(args, "stats", False):
        stats = data.get_statistics()
        print("=" * 68)
        print("🏛️   PHILIPPINE LGU OFFICIALS DIRECTORY - DATA AUDIT")
        print("=" * 68)
        print(f"  Total LGUs             : {stats.get('total_lgus', stats['total_municipalities']):,}")
        print(f"  Total Municipalities   : {stats['total_municipalities']:,}")
        print(f"  Total Cities           : {stats.get('total_cities', 0):,}")
        print(f"  Total Regions          : {stats['total_regions']} Regions")
        print(f"  Total Provinces        : {stats['total_provinces']} Provinces")
        print(f"  Mayors                 : {stats['mayors_count']:,} ({stats['mayor_coverage_pct']}%)")
        print(f"  Vice Mayors            : {stats['vice_mayors_count']:,} ({stats['vm_coverage_pct']}%)")
        print(f"  Office Addresses       : {stats['addresses_count']:,} ({stats['address_coverage_pct']}%)")
        print(f"  Officials with Contact : {stats['mayor_contact_count']:,}")
        print("=" * 68)
        if stats['missing_mayors']:
            print(f"  LGUs with Unrecorded Mayor ({len(stats['missing_mayors'])}):")
            for m in stats['missing_mayors']:
                print(f"    - {m['region']} | {m['province']} | {m['town']} ({m.get('lgu_type', 'LGU')})")
        if stats['missing_vice_mayors']:
            print(f"\n  LGUs with Unrecorded Vice Mayor ({len(stats['missing_vice_mayors'])}):")
            for vm in stats['missing_vice_mayors']:
                print(f"    - {vm['region']} | {vm['province']} | {vm['town']} ({vm.get('lgu_type', 'LGU')})")
        print("=" * 68)
        return

    # 2. List Regions
    if args.list_regions:
        print("Available Regions (18):")
        for reg in data.regions:
            provs = data.provinces_by_region.get(reg, [])
            print(f"  - {reg} ({len(provs)} provinces)")
        return

    # 3. List Provinces
    if args.list_provinces:
        reg = args.list_provinces
        if reg.upper() in {"ALL", "ALL REGIONS"}:
            all_p = data.get_provinces_for_region("ALL")
            print(f"All Provinces ({len(all_p)}):")
            for p in all_p:
                print(f"  - {p}")
        else:
            provs = data.get_provinces_for_region(reg)
            if not provs:
                print(f"No provinces found for region: '{reg}'.")
                print("Use --list-regions to see valid region names.")
            else:
                print(f"Provinces in {reg} ({len(provs)}):")
                for p in provs:
                    print(f"  - {p}")
        return

    # 4. Batch export all 18 regions
    if getattr(args, "split_regions", False):
        out_dir = args.output or "exports"
        file_format = (getattr(args, "format", None) or "csv").lower()
        lgu_type = "all"
        if getattr(args, "cities_only", False) or getattr(args, "lgu_type", "") == "city":
            lgu_type = "city"
        elif getattr(args, "municipalities_only", False) or getattr(args, "lgu_type", "") == "municipality":
            lgu_type = "municipality"
        print(f"Batch exporting all 18 regions ({lgu_type}) to '{out_dir}/' in {file_format.upper()} format...")
        files = data.export_all_regions(
            output_dir=out_dir,
            file_format=file_format,
            include_region=args.include_region,
            include_contacts=getattr(args, "include_contact", False),
            include_lgu_type=getattr(args, "include_lgu_type", False),
            lgu_type=lgu_type,
            name_case=args.name_case or "title"
        )
        print(f"[SUCCESS] Exported {len(files)} regional files:")
        for f in files:
            print(f"  - {f}")
        return

    # 5. Filter records
    region = "ALL" if args.all else (args.region or "ALL")
    province = args.province or "ALL"
    name_case = args.name_case or "title"
    middle_as_initial = not getattr(args, "full_middle", False)
    include_contacts = getattr(args, "include_contact", False)
    include_lgu_type = getattr(args, "include_lgu_type", False)
    file_format = (getattr(args, "format", None) or "csv").lower()

    lgu_type = "all"
    if getattr(args, "cities_only", False) or getattr(args, "lgu_type", "") == "city":
        lgu_type = "city"
    elif getattr(args, "municipalities_only", False) or getattr(args, "lgu_type", "") == "municipality":
        lgu_type = "municipality"

    records = data.filter_records(
        region=region,
        province=province,
        search_query=args.search,
        name_case=name_case,
        middle_as_initial=middle_as_initial,
        include_contacts=include_contacts,
        lgu_type=lgu_type
    )

    print(f"Found {len(records)} records matching criteria:")
    print(f"  Region   : {region}")
    print(f"  Province : {province}")
    print(f"  LGU Type : {lgu_type.capitalize()}")
    if args.search:
        print(f"  Search   : '{args.search}'")
    print(f"  Casing   : {name_case}")
    print(f"  Format   : {file_format.upper()}")
    if include_contacts:
        print("  Contacts : Included")
    if include_lgu_type:
        print("  LGU Type Col : Included")

    output_path = args.output
    if not output_path:
        reg_clean = region.replace(" ", "_") if region != "ALL" else "ALL_REGIONS"
        prov_clean = f"_{province.replace(' ', '_')}" if province != "ALL" else ""
        lgu_slug = f"_{lgu_type.upper()}" if lgu_type != "all" else ""
        output_path = f"Officials_{reg_clean}{prov_clean}{lgu_slug}.{file_format}"

    if file_format == "xlsx":
        count = data.export_xlsx(
            filepath=output_path,
            records=records,
            include_region=args.include_region,
            include_contacts=include_contacts,
            include_lgu_type=include_lgu_type
        )
    elif file_format == "json":
        count = data.export_json(
            filepath=output_path,
            records=records,
            include_region=args.include_region,
            include_contacts=include_contacts,
            include_lgu_type=include_lgu_type
        )
    else:
        count = data.export_csv(
            filepath=output_path,
            records=records,
            include_region=args.include_region,
            include_contacts=include_contacts,
            include_lgu_type=include_lgu_type
        )

    print(f"\n[SUCCESS] Exported {count} records to: {os.path.abspath(output_path)}")


def run_web(port: int = 8000, data_path: Optional[str] = None, excel_path: Optional[str] = None):
    """Run an interactive web interface using Python's standard library."""
    import http.server
    import socketserver
    import urllib.parse

    target_path = data_path or excel_path
    data = OfficialsData(target_path)
    target_path = data.data_path
    data.load()

    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "index.html")
    verify_template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "verify.html")

    # Verification status file lives alongside the data source
    base_dir = os.path.dirname(os.path.abspath(data.data_path))
    verification_file = os.path.join(base_dir, "verification_status.json")

    def load_verification() -> dict:
        """Load verification status from JSON file. Returns empty dict if file missing."""
        if os.path.exists(verification_file):
            try:
                with open(verification_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_verification(vmap: dict):
        """Atomically save verification status to JSON file."""
        tmp_path = verification_file + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(vmap, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, verification_file)

    # Manually added LGU seals: image files in seals/ plus a JSON index keyed by "Province|Town"
    import base64
    import re
    seals_dir = os.path.join(base_dir, "seals")
    seals_index_file = os.path.join(seals_dir, "seals.json")
    seal_mime_ext = {
        "image/png": "png", "image/jpeg": "jpg", "image/jpg": "jpg",
        "image/gif": "gif", "image/webp": "webp", "image/svg+xml": "svg",
    }
    seal_ext_mime = {"png": "image/png", "jpg": "image/jpeg", "gif": "image/gif",
                     "webp": "image/webp", "svg": "image/svg+xml"}
    max_seal_bytes = 5 * 1024 * 1024

    def load_seals() -> dict:
        if os.path.exists(seals_index_file):
            try:
                with open(seals_index_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def save_seals(smap: dict):
        os.makedirs(seals_dir, exist_ok=True)
        tmp_path = seals_index_file + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(smap, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, seals_index_file)

    def seal_slug(province: str, town: str) -> str:
        raw = f"{province}__{town}".lower()
        return re.sub(r"[^a-z0-9ñ_]+", "-", raw).strip("-") or "seal"

    def remove_seal_file(entry: Optional[dict]):
        if entry and entry.get("file"):
            path = os.path.join(seals_dir, os.path.basename(entry["file"]))
            if os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass

    def fetch_seal_url(url: str):
        """Download an image from a URL. Returns (bytes, mime)."""
        import urllib.request
        if not url.lower().startswith(("http://", "https://")):
            raise ValueError("URL must start with http:// or https://")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (LGU Directory seal fetcher)"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            mime = (resp.headers.get_content_type() or "").lower()
            content = resp.read(max_seal_bytes + 1)
        if mime not in seal_mime_ext:
            raise ValueError(f"URL did not return a supported image (got {mime or 'unknown type'})")
        return content, mime

    # Custom Word template auto-fill
    from lgu.word_templates import WordTemplateStore, PLACEHOLDERS, MAX_TEMPLATE_BYTES
    word_store = WordTemplateStore(base_dir)
    docx_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    def seal_path_for(province: str, town: str, smap: Optional[dict] = None) -> Optional[str]:
        entry = (smap if smap is not None else load_seals()).get(f"{province}|{town}")
        if entry and entry.get("file"):
            return os.path.join(seals_dir, os.path.basename(entry["file"]))
        return None

    def parse_seal_width(query: dict) -> float:
        try:
            return min(max(float(query.get("seal_width", ["1"])[0]), 0.2), 8.0)
        except ValueError:
            return 1.0

    def content_disposition(filename: str) -> str:
        ascii_name = filename.encode("ascii", "replace").decode("ascii").replace("?", "_").replace('"', "")
        return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{urllib.parse.quote(filename)}"

    class WebHandler(http.server.BaseHTTPRequestHandler):
        def send_json(self, code: int, obj):
            body = json.dumps(obj).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def send_download(self, content: bytes, mime: str, filename: str):
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Disposition", content_disposition(filename))
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)

        def log_message(self, format, *args):
            # Suppress normal GET logging to keep console clean
            return

        def do_POST(self):
            parsed = urllib.parse.urlparse(self.path)
            if parsed.path == "/api/update":
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                try:
                    payload = json.loads(body.decode("utf-8"))
                    success = data.update_municipality(
                        town=payload.get("town", ""),
                        province=payload.get("province", ""),
                        mayor_first=payload.get("mayor_first", ""),
                        mayor_middle=payload.get("mayor_middle", ""),
                        mayor_last=payload.get("mayor_last", ""),
                        mayor_suffix=payload.get("mayor_suffix", ""),
                        mayor_address=payload.get("mayor_address", ""),
                        mayor_contact=payload.get("mayor_contact", ""),
                        vm_first=payload.get("vm_first", ""),
                        vm_middle=payload.get("vm_middle", ""),
                        vm_last=payload.get("vm_last", ""),
                        vm_suffix=payload.get("vm_suffix", ""),
                        vm_address=payload.get("vm_address", ""),
                        vm_contact=payload.get("vm_contact", ""),
                        mayor_email=payload.get("mayor_email", ""),
                        vm_email=payload.get("vm_email", "")
                    )
                    self.send_response(200 if success else 400)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": success}).encode("utf-8"))
                except Exception as ex:
                    self.send_response(500)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": False, "error": str(ex)}).encode("utf-8"))
            elif parsed.path == "/api/verification":
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                try:
                    payload = json.loads(body.decode("utf-8"))
                    town = payload.get("town", "").strip()
                    province = payload.get("province", "").strip()
                    verified = payload.get("verified", False)
                    if not town or not province:
                        self.send_response(400)
                        self.send_header("Content-Type", "application/json")
                        self.end_headers()
                        self.wfile.write(json.dumps({"success": False, "error": "town and province required"}).encode("utf-8"))
                        return
                    key = f"{province}|{town}"
                    vmap = load_verification()
                    timestamp = None
                    if verified:
                        timestamp = datetime.datetime.now().isoformat(timespec="seconds")
                        vmap[key] = timestamp
                    else:
                        vmap.pop(key, None)
                    save_verification(vmap)
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": True, "timestamp": timestamp}).encode("utf-8"))
                except Exception as ex:
                    self.send_response(500)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": False, "error": str(ex)}).encode("utf-8"))
            elif parsed.path in {"/api/word-templates", "/api/word-templates/delete"}:
                content_length = int(self.headers.get("Content-Length", 0))
                if content_length > MAX_TEMPLATE_BYTES * 2:
                    self.send_json(413, {"success": False, "error": "Template too large (max 15 MB)"})
                    return
                try:
                    payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
                    if parsed.path.endswith("/delete"):
                        word_store.delete(payload.get("name", ""))
                        self.send_json(200, {"success": True})
                        return
                    raw_data = payload.get("data", "")
                    if "," in raw_data and raw_data.startswith("data:"):
                        raw_data = raw_data.split(",", 1)[1]
                    info = word_store.save(payload.get("filename", ""), base64.b64decode(raw_data))
                    self.send_json(200, {"success": True, **info})
                except (ValueError, FileNotFoundError) as ex:
                    self.send_json(400, {"success": False, "error": str(ex)})
                except Exception as ex:
                    # Jinja syntax errors in the template land here
                    self.send_json(400, {"success": False, "error": f"Could not read template: {ex}"})
            elif parsed.path in {"/api/seal", "/api/seal/delete"}:
                def reply(code: int, obj: dict):
                    self.send_response(code)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps(obj).encode("utf-8"))

                content_length = int(self.headers.get("Content-Length", 0))
                if content_length > max_seal_bytes * 2:
                    reply(413, {"success": False, "error": "Image too large (max 5 MB)"})
                    return
                try:
                    payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
                    town = payload.get("town", "").strip()
                    province = payload.get("province", "").strip()
                    if not town or not province:
                        reply(400, {"success": False, "error": "town and province required"})
                        return
                    key = f"{province}|{town}"
                    smap = load_seals()

                    if parsed.path == "/api/seal/delete":
                        remove_seal_file(smap.pop(key, None))
                        save_seals(smap)
                        reply(200, {"success": True})
                        return

                    image = payload.get("image", "")
                    url = payload.get("url", "").strip()
                    if image:
                        m = re.match(r"^data:([\w/+.-]+);base64,(.*)$", image, re.S)
                        if not m:
                            reply(400, {"success": False, "error": "Invalid image data"})
                            return
                        mime = m.group(1).lower()
                        content = base64.b64decode(m.group(2))
                        source = str(payload.get("source") or "upload")[:500]
                    elif url:
                        content, mime = fetch_seal_url(url)
                        source = url
                    else:
                        reply(400, {"success": False, "error": "Provide an image file or an image URL"})
                        return

                    if mime not in seal_mime_ext:
                        reply(400, {"success": False, "error": f"Unsupported image type: {mime}"})
                        return
                    if len(content) > max_seal_bytes:
                        reply(413, {"success": False, "error": "Image too large (max 5 MB)"})
                        return

                    os.makedirs(seals_dir, exist_ok=True)
                    remove_seal_file(smap.get(key))
                    filename = f"{seal_slug(province, town)}.{seal_mime_ext[mime]}"
                    with open(os.path.join(seals_dir, filename), "wb") as f:
                        f.write(content)
                    updated = datetime.datetime.now().isoformat(timespec="seconds")
                    smap[key] = {"file": filename, "source": source, "updated": updated}
                    save_seals(smap)
                    reply(200, {"success": True, "file": filename, "updated": updated})
                except Exception as ex:
                    reply(500, {"success": False, "error": str(ex)})
            else:
                self.send_json(404, {"success": False, "error": f"Unknown endpoint: {self.path.split('?')[0]}"})

        def do_GET(self):
            parsed = urllib.parse.urlparse(self.path)
            path = parsed.path
            query = urllib.parse.parse_qs(parsed.query)

            if path in {"/", "/index.html"}:
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                if os.path.exists(template_path):
                    with open(template_path, "rb") as f:
                        self.wfile.write(f.read())
                else:
                    self.wfile.write(b"<h1>Error: templates/index.html not found.</h1>")

            elif path == "/verify":
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                if os.path.exists(verify_template_path):
                    with open(verify_template_path, "rb") as f:
                        self.wfile.write(f.read())
                else:
                    self.wfile.write(b"<h1>Error: templates/verify.html not found.</h1>")

            elif path == "/api/meta":
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                meta = {
                    "regions": data.regions,
                    "provinces_by_region": data.provinces_by_region
                }
                self.wfile.write(json.dumps(meta).encode("utf-8"))

            elif path == "/api/verification" and self.command == "GET":
                vmap = load_verification()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps(vmap).encode("utf-8"))

            elif path == "/api/word-templates":
                try:
                    templates = word_store.list()
                except Exception as ex:
                    self.send_json(500, {"error": str(ex)})
                    return
                self.send_json(200, {
                    "templates": templates,
                    "placeholders": [{"name": n, "description": d} for n, d in PLACEHOLDERS],
                })

            elif path == "/api/word-templates/sample":
                try:
                    self.send_download(word_store.sample_template(), docx_mime, "LGU Letter Template (sample).docx")
                except Exception as ex:
                    self.send_json(500, {"error": str(ex)})

            elif path == "/api/word-fill":
                tpl_name = query.get("template", [""])[0]
                town = query.get("town", [""])[0].strip()
                prov = query.get("province", [""])[0].strip()
                casing = query.get("casing", ["title"])[0]
                try:
                    matches = [
                        r for r in data.filter_records(province=prov, name_case=casing, include_contacts=True)
                        if r["Town"].strip().upper() == town.upper()
                    ]
                    if not matches:
                        self.send_json(404, {"error": f"LGU not found: {town}, {prov}"})
                        return
                    rec = matches[0]
                    content, filename = word_store.fill_one(
                        tpl_name, rec, data.get_municipality_raw(rec["Town"], rec["Province"]),
                        seal_path_for(rec["Province"], rec["Town"]), parse_seal_width(query)
                    )
                    self.send_download(content, docx_mime, filename)
                except FileNotFoundError as ex:
                    self.send_json(404, {"error": str(ex)})
                except Exception as ex:
                    self.send_json(400, {"error": f"Could not fill template: {ex}"})

            elif path == "/api/word-fill-batch":
                tpl_name = query.get("template", [""])[0]
                casing = query.get("casing", ["title"])[0]
                verified_only = query.get("verified_only", ["0"])[0] == "1"
                try:
                    records = data.filter_records(
                        region=query.get("region", ["ALL"])[0],
                        province=query.get("province", ["ALL"])[0],
                        search_query=query.get("q", [""])[0],
                        name_case=casing,
                        include_contacts=True,
                        lgu_type=query.get("lgu_type", ["all"])[0],
                    )
                    if verified_only:
                        vmap = load_verification()
                        records = [r for r in records if vmap.get(f"{r['Province']}|{r['Town']}")]
                    if not records:
                        self.send_json(404, {"error": "No LGUs match the current filters"})
                        return
                    smap = load_seals()
                    items = (
                        (r, data.get_municipality_raw(r["Town"], r["Province"]), seal_path_for(r["Province"], r["Town"], smap))
                        for r in records
                    )
                    reg = query.get("region", ["ALL"])[0]
                    content, filename, mime, _, _ = word_store.fill_batch(
                        tpl_name, items, parse_seal_width(query),
                        mode=query.get("mode", ["single"])[0],
                        label="All LGUs" if reg == "ALL" else reg.title(),
                    )
                    self.send_download(content, mime, filename)
                except FileNotFoundError as ex:
                    self.send_json(404, {"error": str(ex)})
                except Exception as ex:
                    self.send_json(400, {"error": f"Could not fill template: {ex}"})

            elif path == "/api/seals":
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps(load_seals()).encode("utf-8"))

            elif path == "/api/seal-proxy":
                # Fetch a remote image server-side so the browser can crop it without CORS canvas tainting
                try:
                    content, mime = fetch_seal_url(query.get("url", [""])[0].strip())
                    if len(content) > max_seal_bytes:
                        raise ValueError("Image too large (max 5 MB)")
                    self.send_response(200)
                    self.send_header("Content-Type", mime)
                    self.send_header("Content-Length", str(len(content)))
                    self.end_headers()
                    self.wfile.write(content)
                except Exception as ex:
                    self.send_response(400)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps({"success": False, "error": str(ex)}).encode("utf-8"))

            elif path.startswith("/seals/"):
                filename = os.path.basename(urllib.parse.unquote(path[len("/seals/"):]))
                file_path = os.path.join(seals_dir, filename)
                ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
                if filename and ext in seal_ext_mime and os.path.isfile(file_path):
                    with open(file_path, "rb") as f:
                        content = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", seal_ext_mime[ext])
                    self.send_header("Content-Length", str(len(content)))
                    self.send_header("Cache-Control", "no-cache")
                    self.end_headers()
                    self.wfile.write(content)
                else:
                    self.send_response(404)
                    self.end_headers()

            elif path == "/api/stats":
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                stats = data.get_statistics()
                self.wfile.write(json.dumps(stats).encode("utf-8"))

            elif path == "/api/records":
                reg = query.get("region", ["ALL"])[0]
                prov = query.get("province", ["ALL"])[0]
                q = query.get("q", [""])[0]
                casing = query.get("casing", ["title"])[0]
                lgu_t = query.get("lgu_type", ["all"])[0]
                inc_contacts = query.get("include_contacts", ["0"])[0] == "1"

                records = data.filter_records(
                    region=reg,
                    province=prov,
                    search_query=q,
                    name_case=casing,
                    include_contacts=inc_contacts,
                    lgu_type=lgu_t
                )
                for r in records:
                    raw = data.get_municipality_raw(r["Town"], r["Province"])
                    if raw:
                        r["raw_mayor"] = raw.get("raw_mayor") or {}
                        r["raw_vm"] = raw.get("raw_vm") or {}
                    else:
                        r["raw_mayor"] = {}
                        r["raw_vm"] = {}

                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps(records).encode("utf-8"))

            elif path == "/api/export":
                reg = query.get("region", ["ALL"])[0]
                prov = query.get("province", ["ALL"])[0]
                q = query.get("q", [""])[0]
                casing = query.get("casing", ["title"])[0]
                lgu_t = query.get("lgu_type", ["all"])[0]
                include_reg = query.get("include_region", ["0"])[0] == "1"
                include_contacts = query.get("include_contacts", ["0"])[0] == "1"
                include_lgu_type = query.get("include_lgu_type", ["0"])[0] == "1"
                include_verified_col = query.get("include_verified_col", ["0"])[0] == "1"
                verified_only = query.get("verified_only", ["0"])[0] == "1"
                export_fmt = query.get("format", ["csv"])[0].lower()

                # Load verification map once
                vmap = load_verification() if (include_verified_col or verified_only) else {}

                records = data.filter_records(
                    region=reg,
                    province=prov,
                    search_query=q,
                    name_case=casing,
                    include_contacts=include_contacts,
                    lgu_type=lgu_t
                )

                # Filter to only verified records if requested
                if verified_only:
                    records = [
                        r for r in records
                        if vmap.get(f"{r.get('Province', '')}|{r.get('Town', '')}")
                    ]

                reg_slug = reg.replace(" ", "_") if reg != "ALL" else "ALL_REGIONS"
                prov_slug = f"_{prov.replace(' ', '_')}" if prov != "ALL" else ""
                lgu_slug = f"_{lgu_t.upper()}" if lgu_t != "all" else ""
                verified_slug = "_VERIFIED" if verified_only else ""
                filename = f"Officials_{reg_slug}{prov_slug}{lgu_slug}{verified_slug}.{export_fmt}"

                if export_fmt == "xlsx":
                    out_buffer = io.BytesIO()
                    data.export_xlsx(
                        out_buffer, records,
                        include_region=include_reg,
                        include_contacts=include_contacts,
                        include_lgu_type=include_lgu_type,
                        include_verified_col=include_verified_col,
                        verification_map=vmap
                    )
                    content = out_buffer.getvalue()
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                elif export_fmt == "json":
                    cols = data._get_export_columns(
                        include_region=include_reg,
                        include_contacts=include_contacts,
                        include_lgu_type=include_lgu_type,
                        include_verified_col=include_verified_col
                    )
                    json_data = []
                    for r in records:
                        row = {c: r.get(c, "") for c in cols}
                        if include_verified_col:
                            key = f"{r.get('Province', '')}|{r.get('Town', '')}"
                            row["Contact Verified"] = "Yes" if vmap.get(key) else "No"
                        json_data.append(row)
                    content = json.dumps(json_data, ensure_ascii=False, indent=2).encode("utf-8")
                    mime_type = "application/json; charset=utf-8"
                else:
                    # CSV with UTF-8 BOM
                    csv_str_buf = io.StringIO()
                    fieldnames = data._get_export_columns(
                        include_region=include_reg,
                        include_contacts=include_contacts,
                        include_lgu_type=include_lgu_type,
                        include_verified_col=include_verified_col
                    )
                    import csv
                    writer = csv.DictWriter(csv_str_buf, fieldnames=fieldnames, extrasaction="ignore")
                    writer.writeheader()
                    for r in records:
                        row = dict(r)
                        if include_verified_col:
                            key = f"{r.get('Province', '')}|{r.get('Town', '')}"
                            row["Contact Verified"] = "Yes" if vmap.get(key) else "No"
                        writer.writerow(row)
                    content = ("\ufeff" + csv_str_buf.getvalue()).encode("utf-8")
                    mime_type = "text/csv; charset=utf-8"

                self.send_response(200)
                self.send_header("Content-Type", mime_type)
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)

            elif path == "/api/batch-export":
                lgu_t = query.get("lgu_type", ["all"])[0]
                inc_lgu = query.get("include_lgu_type", ["0"])[0] == "1"
                files = data.export_all_regions("exports", file_format="csv", lgu_type=lgu_t, include_lgu_type=inc_lgu)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True, "files_count": len(files)}).encode("utf-8"))

            else:
                self.send_json(404, {"success": False, "error": f"Unknown endpoint: {self.path.split('?')[0]}"})

    with socketserver.TCPServer(("", port), WebHandler) as httpd:
        url = f"http://localhost:{port}"
        print(f"\n[WEB APP] Running Philippine LGU Directory Web Server at: {url}")
        print("Press Ctrl+C to stop the server.\n")
        webbrowser.open(url)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down web server.")


def run_gui(data_path: Optional[str] = None, excel_path: Optional[str] = None):
    """Launch full native Windows Tkinter Desktop GUI application."""
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    target_path = data_path or excel_path
    data = OfficialsData(target_path)
    target_path = data.data_path
    try:
        data.load()
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error Loading Data", f"Could not load '{target_path}':\n{e}")
        return

    root = tk.Tk()
    root.title("Philippine LGU Officials Directory & Multi-Format Exporter (2025–2028)")
    root.geometry("1220x780")
    root.minsize(960, 620)

    # Set theme
    style = ttk.Style()
    available_themes = style.theme_names()
    if "vista" in available_themes:
        style.theme_use("vista")
    elif "clam" in available_themes:
        style.theme_use("clam")

    # Color palette
    COLOR_BG = "#f8fafc"
    COLOR_CARD = "#ffffff"
    COLOR_HEADER = "#0f172a"
    COLOR_PRIMARY = "#2563eb"
    COLOR_SUCCESS = "#059669"
    COLOR_MUTED = "#64748b"

    root.configure(bg=COLOR_BG)

    # State variables
    region_var = tk.StringVar(value="All Regions")
    province_var = tk.StringVar(value="All Provinces")
    lgu_type_var = tk.StringVar(value="All LGUs")
    search_var = tk.StringVar()
    casing_var = tk.StringVar(value="Title Case")
    include_lgu_type_var = tk.BooleanVar(value=False)
    include_region_var = tk.BooleanVar(value=False)
    include_contacts_var = tk.BooleanVar(value=False)
    status_var = tk.StringVar(value=f"Ready • {len(data.records):,} LGUs loaded (Cached)")
    count_summary_var = tk.StringVar(value=f"Showing {len(data.records):,} records")

    # Current filtered cache
    current_records: List[Dict[str, str]] = []
    sort_column = "Town"
    sort_descending = False

    # Main Frame
    main_frame = tk.Frame(root, bg=COLOR_BG, padx=16, pady=12)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 1. Header Frame
    header_frame = tk.Frame(main_frame, bg=COLOR_BG)
    header_frame.pack(fill=tk.X, pady=(0, 10))

    title_label = tk.Label(
        header_frame,
        text="🏛️ Philippine LGU Officials Directory",
        font=("Segoe UI", 16, "bold"),
        fg=COLOR_HEADER,
        bg=COLOR_BG
    )
    title_label.pack(anchor="w")

    stats = data.get_statistics()
    total_lgus = stats.get('total_lgus', len(data.records))
    sub_label = tk.Label(
        header_frame,
        text=f"Nationwide {total_lgus:,} LGUs ({stats['total_municipalities']:,} Municipalities & {stats.get('total_cities', 0):,} Cities) across 18 Regions • Mayors: {stats['mayor_coverage_pct']}% • Addresses: 100%",
        font=("Segoe UI", 9),
        fg=COLOR_MUTED,
        bg=COLOR_BG
    )
    sub_label.pack(anchor="w", pady=(2, 0))

    # 2. Control Card Frame
    control_card = tk.LabelFrame(
        main_frame,
        text=" Filter & Selection Controls ",
        font=("Segoe UI", 10, "bold"),
        fg="#334155",
        bg=COLOR_CARD,
        bd=1,
        relief=tk.SOLID,
        padx=14,
        pady=10
    )
    control_card.pack(fill=tk.X, pady=(0, 10))

    # Grid layout inside control card
    tk.Label(control_card, text="Region:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD, fg="#475569").grid(row=0, column=0, sticky="w", padx=(0, 4), pady=4)
    region_combo = ttk.Combobox(control_card, textvariable=region_var, state="readonly", width=18, font=("Segoe UI", 9))
    region_combo["values"] = ["All Regions"] + data.regions
    region_combo.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=4)

    tk.Label(control_card, text="Province:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD, fg="#475569").grid(row=0, column=2, sticky="w", padx=(0, 4), pady=4)
    province_combo = ttk.Combobox(control_card, textvariable=province_var, state="readonly", width=20, font=("Segoe UI", 9))
    province_combo.grid(row=0, column=3, sticky="w", padx=(0, 10), pady=4)

    tk.Label(control_card, text="LGU Type:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD, fg="#475569").grid(row=0, column=4, sticky="w", padx=(0, 4), pady=4)
    lgu_type_combo = ttk.Combobox(control_card, textvariable=lgu_type_var, state="readonly", width=16, font=("Segoe UI", 9))
    lgu_type_combo["values"] = ["All LGUs", "Municipalities Only", "Cities Only"]
    lgu_type_combo.grid(row=0, column=5, sticky="w", padx=(0, 10), pady=4)

    tk.Label(control_card, text="Search:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD, fg="#475569").grid(row=0, column=6, sticky="w", padx=(0, 4), pady=4)
    search_entry = ttk.Entry(control_card, textvariable=search_var, width=18, font=("Segoe UI", 9))
    search_entry.grid(row=0, column=7, sticky="w", padx=(0, 4), pady=4)

    btn_clear_search = ttk.Button(control_card, text="✕", width=3, command=lambda: (search_var.set(""), on_filter_change()))
    btn_clear_search.grid(row=0, column=8, sticky="w", pady=4)

    # Row 1: Options & Action Buttons
    options_frame = tk.Frame(control_card, bg=COLOR_CARD)
    options_frame.grid(row=1, column=0, columnspan=9, sticky="ew", pady=(8, 0))

    tk.Label(options_frame, text="Name Casing:", font=("Segoe UI", 9), bg=COLOR_CARD, fg="#475569").pack(side=tk.LEFT, padx=(0, 6))
    casing_combo = ttk.Combobox(
        options_frame,
        textvariable=casing_var,
        state="readonly",
        width=12,
        font=("Segoe UI", 9),
        values=["Title Case", "UPPERCASE", "Original"]
    )
    casing_combo.pack(side=tk.LEFT, padx=(0, 12))

    check_inc_lgu_type = ttk.Checkbutton(
        options_frame,
        text="Include LGU Type",
        variable=include_lgu_type_var,
        command=lambda: on_filter_change()
    )
    check_inc_lgu_type.pack(side=tk.LEFT, padx=(0, 10))

    check_inc_region = ttk.Checkbutton(
        options_frame,
        text="Include Region",
        variable=include_region_var,
        command=lambda: on_filter_change()
    )
    check_inc_region.pack(side=tk.LEFT, padx=(0, 10))

    check_inc_contact = ttk.Checkbutton(
        options_frame,
        text="Include Contacts",
        variable=include_contacts_var,
        command=lambda: on_filter_change()
    )
    check_inc_contact.pack(side=tk.LEFT, padx=(0, 12))

    btn_reset = ttk.Button(options_frame, text="Reset", command=lambda: reset_filters())
    btn_reset.pack(side=tk.LEFT, padx=(0, 12))

    # Action Buttons: Export CSV, Export Excel, Copy, Edit
    btn_export_xlsx = tk.Button(
        options_frame,
        text="📗 Excel (.xlsx)",
        bg="#1e40af",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        activebackground="#1d4ed8",
        activeforeground="white",
        relief=tk.FLAT,
        padx=10,
        pady=4,
        cursor="hand2",
        command=lambda: export_selected("xlsx")
    )
    btn_export_xlsx.pack(side=tk.RIGHT, padx=(4, 0))

    btn_export_csv = tk.Button(
        options_frame,
        text="⬇ CSV",
        bg=COLOR_SUCCESS,
        fg="white",
        font=("Segoe UI", 9, "bold"),
        activebackground="#047857",
        activeforeground="white",
        relief=tk.FLAT,
        padx=12,
        pady=4,
        cursor="hand2",
        command=lambda: export_selected("csv")
    )
    btn_export_csv.pack(side=tk.RIGHT, padx=(4, 0))

    btn_clipboard = tk.Button(
        options_frame,
        text="📋 Copy CSV",
        bg="#0284c7",
        fg="white",
        font=("Segoe UI", 9),
        activebackground="#0369a1",
        activeforeground="white",
        relief=tk.FLAT,
        padx=9,
        pady=4,
        cursor="hand2",
        command=lambda: copy_to_clipboard()
    )
    btn_clipboard.pack(side=tk.RIGHT, padx=(4, 0))

    btn_edit = tk.Button(
        options_frame,
        text="✏️ Edit Official",
        bg="#7c3aed",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        activebackground="#6d28d9",
        activeforeground="white",
        relief=tk.FLAT,
        padx=10,
        pady=4,
        cursor="hand2",
        command=lambda: open_edit_dialog()
    )
    btn_edit.pack(side=tk.RIGHT, padx=(0, 4))

    # 3. Stats & Summary Bar
    stats_bar = tk.Frame(main_frame, bg=COLOR_BG)
    stats_bar.pack(fill=tk.X, pady=(2, 6))

    lbl_count = tk.Label(stats_bar, textvariable=count_summary_var, font=("Segoe UI", 9, "bold"), fg="#1e293b", bg=COLOR_BG)
    lbl_count.pack(side=tk.LEFT)

    lbl_hint = tk.Label(stats_bar, text="💡 Tip: Ctrl+F to search • Ctrl+E to export CSV • Double-click row for details • Click header to sort", font=("Segoe UI", 8), fg="#64748b", bg=COLOR_BG)
    lbl_hint.pack(side=tk.RIGHT)

    # 4. Preview Table (Treeview)
    table_card = tk.Frame(main_frame, bg=COLOR_CARD, bd=1, relief=tk.SOLID)
    table_card.pack(fill=tk.BOTH, expand=True)

    col_widths = {
        "Town": 150,
        "LGU Type": 110,
        "Province": 150,
        "Region": 110,
        "Mayor Name": 190,
        "Mayor Contact": 130,
        "Vice Mayor": 190,
        "Vice Mayor Contact": 130,
        "Municipal Hall Address": 300
    }

    def get_active_columns():
        cols = ["Town"]
        if include_lgu_type_var.get():
            cols.append("LGU Type")
        cols.append("Province")
        if include_region_var.get():
            cols.append("Region")
        cols.append("Mayor Name")
        if include_contacts_var.get():
            cols.append("Mayor Contact")
        cols.append("Vice Mayor")
        if include_contacts_var.get():
            cols.append("Vice Mayor Contact")
        cols.append("Municipal Hall Address")
        return tuple(cols)

    tree = ttk.Treeview(table_card, columns=get_active_columns(), show="headings", selectmode="browse")

    vsb = ttk.Scrollbar(table_card, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(table_card, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    table_card.grid_rowconfigure(0, weight=1)
    table_card.grid_columnconfigure(0, weight=1)

    # Striped rows
    tree.tag_configure("evenrow", background="#ffffff")
    tree.tag_configure("oddrow", background="#f8fafc")

    # 5. Status Bar at bottom
    status_bar = tk.Frame(root, bg="#0f172a", height=24, padx=12, pady=4)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    status_label = tk.Label(status_bar, textvariable=status_var, font=("Segoe UI", 8), fg="#cbd5e1", bg="#0f172a")
    status_label.pack(side=tk.LEFT)

    source_display = os.path.basename(target_path) if target_path else "officials_2025_2028.xlsx"
    file_label = tk.Label(status_bar, text=f"Source: {source_display} • Auto-backup ON", font=("Segoe UI", 8), fg="#94a3b8", bg="#0f172a")
    file_label.pack(side=tk.RIGHT)

    # ------------------ EVENT HANDLERS & LOGIC ------------------

    def update_province_dropdown():
        selected_region = region_var.get()
        if selected_region == "All Regions":
            all_provinces = data.get_provinces_for_region("ALL")
            province_combo["values"] = ["All Provinces"] + all_provinces
            province_var.set("All Provinces")
        else:
            provs = data.get_provinces_for_region(selected_region)
            province_combo["values"] = [f"All Provinces ({selected_region})"] + provs
            province_var.set(province_combo["values"][0])

    def get_casing_mode() -> str:
        val = casing_var.get()
        if val == "Title Case":
            return "title"
        elif val == "UPPERCASE":
            return "upper"
        return "original"

    def on_filter_change(*args):
        nonlocal current_records
        reg = region_var.get()
        prov = province_var.get()

        reg_filter = None if reg == "All Regions" else reg
        prov_filter = None if "All Provinces" in prov else prov
        search = search_var.get().strip()
        casing = get_casing_mode()
        inc_contacts = include_contacts_var.get()

        lgu_mode = lgu_type_var.get()
        lgu_type = "all"
        if "City" in lgu_mode or "Cities" in lgu_mode:
            lgu_type = "city"
        elif "Municipality" in lgu_mode or "Municipalities" in lgu_mode:
            lgu_type = "municipality"

        current_records = data.filter_records(
            region=reg_filter,
            province=prov_filter,
            search_query=search,
            name_case=casing,
            include_contacts=inc_contacts,
            lgu_type=lgu_type
        )

        refresh_table()

    def refresh_table():
        active_cols = get_active_columns()

        if tuple(tree["columns"]) != active_cols:
            tree["columns"] = active_cols
            for col in active_cols:
                tree.column(col, width=col_widths.get(col, 150), minwidth=80, anchor="w", stretch=True)

        for col in active_cols:
            indicator = ""
            if col == sort_column:
                indicator = " ▼" if sort_descending else " ▲"
            tree.heading(col, text=col + indicator, command=lambda _c=col: sort_table_by_column(_c))

        for item in tree.get_children():
            tree.delete(item)

        sorted_records = list(current_records)
        if sort_column:
            def sort_key(rec):
                return str(rec.get(sort_column, "")).lower()
            sorted_records.sort(key=sort_key, reverse=sort_descending)

        for i, rec in enumerate(sorted_records):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            values = tuple(rec.get(col, "") for col in active_cols)
            tree.insert("", tk.END, values=values, tags=(tag,))

        reg_disp = region_var.get()
        prov_disp = province_var.get()
        filter_desc = f"{reg_disp}" if prov_disp.startswith("All") else f"{prov_disp}, {reg_disp}"
        count_summary_var.set(
            f"Showing {len(sorted_records)} municipalities ({filter_desc}) • Columns: {len(active_cols)}"
        )
        status_var.set(f"Loaded {len(sorted_records)} records • Ready to export")

    def sort_table_by_column(col: str):
        nonlocal sort_column, sort_descending
        if sort_column == col:
            sort_descending = not sort_descending
        else:
            sort_column = col
            sort_descending = False
        refresh_table()

    def reset_filters():
        region_var.set("All Regions")
        update_province_dropdown()
        search_var.set("")
        casing_var.set("Title Case")
        include_region_var.set(False)
        include_contacts_var.set(False)
        on_filter_change()

    search_var.trace_add("write", lambda *args: on_filter_change())
    region_combo.bind("<<ComboboxSelected>>", lambda e: (update_province_dropdown(), on_filter_change()))
    province_combo.bind("<<ComboboxSelected>>", lambda e: on_filter_change())
    casing_combo.bind("<<ComboboxSelected>>", lambda e: on_filter_change())

    # Keyboard shortcuts
    root.bind("<Control-f>", lambda e: search_entry.focus())
    root.bind("<Control-e>", lambda e: export_selected("csv"))

    def open_edit_dialog(target_rec=None):
        if not target_rec:
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showinfo("Select Municipality", "Please select a municipality from the table to edit.")
                return
            values = tree.item(selected_item[0], "values")
            if not values:
                return
            town_val = values[0]
            prov_val = values[1]
            target_rec = data.get_municipality_raw(town_val, prov_val)

        if not target_rec:
            messagebox.showerror("Error", "Could not locate municipality record.")
            return

        raw_mayor = target_rec.get("raw_mayor") or {}
        raw_vm = target_rec.get("raw_vm") or {}

        edit_win = tk.Toplevel(root)
        edit_win.title(f"Edit Officials - {target_rec['town']}, {target_rec['province']}")
        edit_win.geometry("640x660")
        edit_win.minsize(560, 580)
        edit_win.configure(bg=COLOR_BG)
        edit_win.transient(root)
        edit_win.grab_set()

        banner = tk.Frame(edit_win, bg="#0f172a", padx=16, pady=12)
        banner.pack(fill=tk.X)
        tk.Label(
            banner,
            text=f"🏛️ Edit Officials: {target_rec['town']}, {target_rec['province']}",
            font=("Segoe UI", 12, "bold"),
            fg="#f8fafc",
            bg="#0f172a"
        ).pack(anchor="w")
        tk.Label(
            banner,
            text=f"Region: {target_rec['region']} • Updates create timestamped backup in backups/",
            font=("Segoe UI", 8),
            fg="#94a3b8",
            bg="#0f172a"
        ).pack(anchor="w", pady=(2, 0))

        content_frame = tk.Frame(edit_win, bg=COLOR_BG, padx=16, pady=10)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Mayor Card
        m_card = tk.LabelFrame(content_frame, text=" Municipal Mayor ", font=("Segoe UI", 10, "bold"), fg="#1e40af", bg=COLOR_CARD, bd=1, relief=tk.SOLID, padx=12, pady=8)
        m_card.pack(fill=tk.X, pady=(0, 10))

        m_f1 = tk.Frame(m_card, bg=COLOR_CARD)
        m_f1.pack(fill=tk.X, pady=3)
        tk.Label(m_f1, text="First Name:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_m_first = ttk.Entry(m_f1, width=22)
        e_m_first.insert(0, raw_mayor.get("first") or "")
        e_m_first.pack(side=tk.LEFT, padx=(0, 16))

        tk.Label(m_f1, text="Middle Name:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_m_middle = ttk.Entry(m_f1, width=20)
        e_m_middle.insert(0, raw_mayor.get("middle") or "")
        e_m_middle.pack(side=tk.LEFT)

        m_f2 = tk.Frame(m_card, bg=COLOR_CARD)
        m_f2.pack(fill=tk.X, pady=3)
        tk.Label(m_f2, text="Last Name:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_m_last = ttk.Entry(m_f2, width=22)
        e_m_last.insert(0, raw_mayor.get("last") or "")
        e_m_last.pack(side=tk.LEFT, padx=(0, 16))

        tk.Label(m_f2, text="Suffix (Jr/III):", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_m_suf = ttk.Entry(m_f2, width=10)
        e_m_suf.insert(0, raw_mayor.get("suffix") or "")
        e_m_suf.pack(side=tk.LEFT)

        m_f3 = tk.Frame(m_card, bg=COLOR_CARD)
        m_f3.pack(fill=tk.X, pady=3)
        tk.Label(m_f3, text="Contact No:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_m_contact = ttk.Entry(m_f3, width=22)
        e_m_contact.insert(0, raw_mayor.get("contact") or "")
        e_m_contact.pack(side=tk.LEFT)

        m_f4 = tk.Frame(m_card, bg=COLOR_CARD)
        m_f4.pack(fill=tk.X, pady=3)
        tk.Label(m_f4, text="Hall Address:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_m_addr = ttk.Entry(m_f4)
        e_m_addr.insert(0, raw_mayor.get("address") or target_rec.get("address") or "")
        e_m_addr.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Vice Mayor Card
        v_card = tk.LabelFrame(content_frame, text=" Municipal Vice-Mayor ", font=("Segoe UI", 10, "bold"), fg="#065f46", bg=COLOR_CARD, bd=1, relief=tk.SOLID, padx=12, pady=8)
        v_card.pack(fill=tk.X, pady=(0, 10))

        v_f1 = tk.Frame(v_card, bg=COLOR_CARD)
        v_f1.pack(fill=tk.X, pady=3)
        tk.Label(v_f1, text="First Name:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_v_first = ttk.Entry(v_f1, width=22)
        e_v_first.insert(0, raw_vm.get("first") or "")
        e_v_first.pack(side=tk.LEFT, padx=(0, 16))

        tk.Label(v_f1, text="Middle Name:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_v_middle = ttk.Entry(v_f1, width=20)
        e_v_middle.insert(0, raw_vm.get("middle") or "")
        e_v_middle.pack(side=tk.LEFT)

        v_f2 = tk.Frame(v_card, bg=COLOR_CARD)
        v_f2.pack(fill=tk.X, pady=3)
        tk.Label(v_f2, text="Last Name:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_v_last = ttk.Entry(v_f2, width=22)
        e_v_last.insert(0, raw_vm.get("last") or "")
        e_v_last.pack(side=tk.LEFT, padx=(0, 16))

        tk.Label(v_f2, text="Suffix (Jr/III):", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_v_suf = ttk.Entry(v_f2, width=10)
        e_v_suf.insert(0, raw_vm.get("suffix") or "")
        e_v_suf.pack(side=tk.LEFT)

        v_f3 = tk.Frame(v_card, bg=COLOR_CARD)
        v_f3.pack(fill=tk.X, pady=3)
        tk.Label(v_f3, text="Contact No:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_v_contact = ttk.Entry(v_f3, width=22)
        e_v_contact.insert(0, raw_vm.get("contact") or "")
        e_v_contact.pack(side=tk.LEFT)

        v_f4 = tk.Frame(v_card, bg=COLOR_CARD)
        v_f4.pack(fill=tk.X, pady=3)
        tk.Label(v_f4, text="Office Address:", width=11, anchor="w", bg=COLOR_CARD, font=("Segoe UI", 9)).pack(side=tk.LEFT)
        e_v_addr = ttk.Entry(v_f4)
        e_v_addr.insert(0, raw_vm.get("address") or "")
        e_v_addr.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_frame = tk.Frame(content_frame, bg=COLOR_BG)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        def do_save():
            try:
                ok = data.update_municipality(
                    town=target_rec["town"],
                    province=target_rec["province"],
                    mayor_first=e_m_first.get().strip(),
                    mayor_middle=e_m_middle.get().strip(),
                    mayor_last=e_m_last.get().strip(),
                    mayor_suffix=e_m_suf.get().strip(),
                    mayor_address=e_m_addr.get().strip(),
                    mayor_contact=e_m_contact.get().strip(),
                    vm_first=e_v_first.get().strip(),
                    vm_middle=e_v_middle.get().strip(),
                    vm_last=e_v_last.get().strip(),
                    vm_suffix=e_v_suf.get().strip(),
                    vm_address=e_v_addr.get().strip(),
                    vm_contact=e_v_contact.get().strip()
                )
                if ok:
                    on_filter_change()
                    messagebox.showinfo("Saved", f"Updated {target_rec['town']}, {target_rec['province']}!\nBackup saved to backups/ folder.", parent=edit_win)
                    edit_win.destroy()
                else:
                    messagebox.showerror("Error", "Could not save municipality changes.", parent=edit_win)
            except Exception as ex:
                messagebox.showerror("Save Failed", f"An error occurred:\n{ex}", parent=edit_win)

        btn_save = tk.Button(btn_frame, text="💾 Save to Excel (.xlsx)", bg=COLOR_SUCCESS, fg="white", font=("Segoe UI", 9, "bold"), relief=tk.FLAT, padx=14, pady=6, cursor="hand2", command=do_save)
        btn_save.pack(side=tk.RIGHT, padx=(8, 0))
        btn_cancel = ttk.Button(btn_frame, text="Cancel", command=edit_win.destroy)
        btn_cancel.pack(side=tk.RIGHT)

    def show_row_details(event):
        selected_item = tree.selection()
        if not selected_item:
            return
        values = tree.item(selected_item[0], "values")
        if not values:
            return

        town_val = values[0]
        prov_val = values[1]
        rec = next((r for r in current_records if r["Town"] == town_val and r["Province"] == prov_val), None)
        if not rec:
            return

        raw = data.get_municipality_raw(rec["Town"], rec["Province"]) or {}
        raw_m = raw.get("raw_mayor") or {}
        raw_vm = raw.get("raw_vm") or {}

        detail_win = tk.Toplevel(root)
        detail_win.title(f"Details - {rec['Town']}, {rec['Province']}")
        detail_win.geometry("600x440")
        detail_win.minsize(500, 360)
        detail_win.configure(bg=COLOR_BG)
        detail_win.transient(root)

        card = tk.Frame(detail_win, bg=COLOR_CARD, padx=20, pady=20, bd=1, relief=tk.SOLID)
        card.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

        fields = [
            ("Town:", rec["Town"]),
            ("Province:", rec["Province"]),
            ("Region:", rec["Region"]),
            ("Mayor Name:", rec["Mayor Name"]),
            ("Mayor Contact:", raw_m.get("contact") or "(None)"),
            ("Vice Mayor:", rec["Vice Mayor"]),
            ("Vice Mayor Contact:", raw_vm.get("contact") or "(None)"),
            ("Municipal Hall Address:", rec["Municipal Hall Address"]),
        ]

        for i, (label_txt, val_txt) in enumerate(fields):
            tk.Label(card, text=label_txt, font=("Segoe UI", 9, "bold"), fg="#475569", bg=COLOR_CARD, anchor="w").grid(row=i, column=0, sticky="nw", pady=3, padx=(0, 10))
            tk.Label(card, text=val_txt or "(None)", font=("Segoe UI", 9), fg="#0f172a", bg=COLOR_CARD, anchor="w", wraplength=400, justify="left").grid(row=i, column=1, sticky="w", pady=3)

        btn_box = tk.Frame(card, bg=COLOR_CARD)
        btn_box.grid(row=len(fields), column=0, columnspan=2, pady=(16, 0), sticky="e")

        def copy_address():
            root.clipboard_clear()
            root.clipboard_append(rec["Municipal Hall Address"])
            messagebox.showinfo("Copied", "Address copied to clipboard!", parent=detail_win)

        ttk.Button(btn_box, text="✏️ Edit Officials", command=lambda: (detail_win.destroy(), open_edit_dialog(raw))).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_box, text="Copy Address", command=copy_address).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_box, text="Close", command=detail_win.destroy).pack(side=tk.LEFT)

    tree.bind("<Double-1>", show_row_details)

    # Right-click context menu
    context_menu = tk.Menu(root, tearoff=0)
    context_menu.add_command(label="✏️ Edit Officials...", command=lambda: open_edit_dialog())
    context_menu.add_command(label="👁️ View Details...", command=lambda: show_row_details(None))
    context_menu.add_separator()
    context_menu.add_command(label="📋 Copy CSV Row", command=lambda: copy_to_clipboard())

    def on_tree_right_click(event):
        row_id = tree.identify_row(event.y)
        if row_id:
            tree.selection_set(row_id)
            context_menu.post(event.x_root, event.y_root)

    tree.bind("<Button-3>", on_tree_right_click)

    def export_selected(file_format="csv"):
        if not current_records:
            messagebox.showwarning("No Records", "There are no records matching your current filter to export.")
            return

        reg = region_var.get()
        prov = province_var.get()
        reg_clean = reg.replace(" ", "_") if reg != "All Regions" else "ALL_REGIONS"
        prov_clean = f"_{prov.replace(' ', '_')}" if not prov.startswith("All Provinces") else ""
        default_filename = f"Officials_{reg_clean}{prov_clean}.{file_format}"

        ftypes = {
            "csv": [("CSV Files (*.csv)", "*.csv"), ("All Files (*.*)", "*.*")],
            "xlsx": [("Excel Files (*.xlsx)", "*.xlsx"), ("All Files (*.*)", "*.*")],
            "json": [("JSON Files (*.json)", "*.json"), ("All Files (*.*)", "*.*")]
        }

        filepath = filedialog.asksaveasfilename(
            parent=root,
            title=f"Save {file_format.upper()} File",
            initialfile=default_filename,
            defaultextension=f".{file_format}",
            filetypes=ftypes.get(file_format, [("All Files (*.*)", "*.*")])
        )
        if not filepath:
            return

        try:
            inc_reg = include_region_var.get()
            inc_con = include_contacts_var.get()
            inc_lgu = include_lgu_type_var.get()
            if file_format == "xlsx":
                count = data.export_xlsx(filepath=filepath, records=current_records, include_region=inc_reg, include_contacts=inc_con, include_lgu_type=inc_lgu)
            elif file_format == "json":
                count = data.export_json(filepath=filepath, records=current_records, include_region=inc_reg, include_contacts=inc_con, include_lgu_type=inc_lgu)
            else:
                count = data.export_csv(filepath=filepath, records=current_records, include_region=inc_reg, include_contacts=inc_con, include_lgu_type=inc_lgu)

            status_var.set(f"Successfully exported {count} records to: {os.path.basename(filepath)}")
            resp = messagebox.askyesno(
                "Export Complete",
                f"Successfully exported {count} records to:\n{filepath}\n\nWould you like to open this file now?"
            )
            if resp:
                try:
                    os.startfile(filepath)
                except Exception as ex:
                    messagebox.showerror("Could not open file", str(ex))
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred while saving:\n{e}")

    def copy_to_clipboard():
        if not current_records:
            messagebox.showwarning("No Records", "No records to copy.")
            return

        import csv
        output = io.StringIO()
        fieldnames = data._get_export_columns(include_region=include_region_var.get(), include_contacts=include_contacts_var.get(), include_lgu_type=include_lgu_type_var.get())
        writer = csv.DictWriter(output, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for r in current_records:
            writer.writerow(r)

        root.clipboard_clear()
        root.clipboard_append(output.getvalue())
        status_var.set(f"Copied {len(current_records)} records to clipboard as CSV!")
        messagebox.showinfo("Copied", f"Copied {len(current_records)} records to clipboard!\nYou can paste (Ctrl+V) directly into Excel or a text editor.")

    def reset_filters():
        region_var.set("All Regions")
        update_province_dropdown()
        lgu_type_var.set("All LGUs")
        search_var.set("")
        casing_var.set("Title Case")
        include_lgu_type_var.set(False)
        include_region_var.set(False)
        include_contacts_var.set(False)
        on_filter_change()

    update_province_dropdown()
    on_filter_change()

    # Center window
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    ws = root.winfo_screenwidth()
    hs = root.winfo_screenheight()
    x = max(0, (ws // 2) - (w // 2))
    y = max(0, (hs // 2) - (h // 2))
    root.geometry(f"{w}x{h}+{x}+{y}")

    root.mainloop()


def main():
    parser = argparse.ArgumentParser(
        description="Philippine LGU Officials Directory & CSV/Excel/JSON Exporter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Launch Desktop GUI:
  python app.py

  # Start Web UI:
  python app.py --web

  # Data health & coverage audit:
  python app.py --stats

  # Export Region 3, Pampanga to Excel:
  python app.py --region "REGION 3" --province "PAMPANGA" --format xlsx -o pampanga.xlsx

  # Export all cities nationwide to CSV:
  python app.py --cities-only -o cities.csv

  # Export all municipalities nationwide with contacts to CSV:
  python app.py --municipalities-only --include-contact -o municipalities.csv

  # Batch export all 18 regions to separate files in exports/ folder:
  python app.py --split-regions --format xlsx
        """
    )
    parser.add_argument("--data", default=None, help="Path to officials dataset (.xlsx or .csv, default: data/officials_2025_2028.xlsx)")
    parser.add_argument("--excel", default=None, help="Legacy alias for --data (path to Excel file)")
    parser.add_argument("--region", help="Region to filter (e.g. 'REGION 3', 'NCR', 'BARMM')")
    parser.add_argument("--province", help="Province to filter (e.g. 'PAMPANGA', 'TARLAC')")
    parser.add_argument("--lgu-type", choices=["all", "municipality", "city"], default="all", help="Filter by LGU type: all, municipality, or city (default: all)")
    parser.add_argument("--cities-only", action="store_true", help="Filter and export cities only")
    parser.add_argument("--municipalities-only", action="store_true", help="Filter and export municipalities only")
    parser.add_argument("--include-lgu-type", action="store_true", help="Include LGU Type (City / Municipality) column in export")
    parser.add_argument("--search", help="Search text filter for town or official name")
    parser.add_argument("--output", "-o", help="Output file path (CSV, XLSX, or JSON)")
    parser.add_argument("--all", action="store_true", help="Export all regions and provinces nationwide")
    parser.add_argument("--format", choices=["csv", "xlsx", "json"], default="csv", help="Export file format (default: csv)")
    parser.add_argument("--name-case", choices=["title", "upper", "original"], default="title", help="Name casing style (default: title)")
    parser.add_argument("--full-middle", action="store_true", help="Use full middle name instead of middle initial")
    parser.add_argument("--include-region", action="store_true", help="Include Region column in exported output")
    parser.add_argument("--include-contact", action="store_true", help="Include Mayor and Vice Mayor Office Contact Numbers & Emails")
    parser.add_argument("--stats", "--summary", action="store_true", help="Display data health audit & official coverage statistics")
    parser.add_argument("--split-regions", action="store_true", help="Batch export all 18 regions into separate files in exports/ folder")
    parser.add_argument("--list-regions", action="store_true", help="List all available regions and exit")
    parser.add_argument("--list-provinces", nargs="?", const="ALL", help="List provinces for a region (or all if not specified) and exit")
    parser.add_argument("--web", nargs="?", const=8000, type=int, help="Start web interface on specified port (default: 8000)")
    parser.add_argument("--cli", action="store_true", help="Force command line mode")

    args = parser.parse_args()

    # Resolve data path: preference given to --data or --excel, otherwise auto-detect
    data_file = args.data or args.excel

    # Determine execution mode
    is_cli_mode = (
        args.cli or
        args.stats or
        args.split_regions or
        args.list_regions or
        args.list_provinces is not None or
        args.region is not None or
        args.province is not None or
        args.all or
        args.cities_only or
        args.municipalities_only or
        args.lgu_type != "all" or
        args.output is not None
    )

    if args.web is not None:
        run_web(port=args.web, data_path=data_file)
    elif is_cli_mode:
        data = OfficialsData(data_file)
        run_cli(args, data)
    else:
        run_gui(data_path=data_file)


if __name__ == "__main__":
    main()