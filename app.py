"""
Philippine Municipal Officials Directory & CSV Exporter
A Python application to filter, view, and export LGU Municipal Officials
(Mayor Name, Municipal Hall Address, Town, Province, Vice Mayor) by Region and Province.

Features:
- Modern interactive Desktop GUI (Tkinter)
- Fast live filtering by Region and Province
- Real-time search by town, mayor, vice mayor, or province
- One-click CSV export with UTF-8 BOM (Excel compatible)
- CLI mode for batch automation and scripts
- Optional local web interface (--web)
"""

import sys
import os
import argparse
import webbrowser
from typing import Optional, List, Dict

# Local data processor
from data_processor import OfficialsData


def run_cli(args: argparse.Namespace, data: OfficialsData):
    """Run in command-line mode without GUI."""
    data.load()

    if args.list_regions:
        print("Available Regions (18):")
        for reg in data.regions:
            provs = data.provinces_by_region.get(reg, [])
            print(f"  - {reg} ({len(provs)} provinces)")
        return

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

    # Filter records
    region = "ALL" if args.all else (args.region or "ALL")
    province = args.province or "ALL"
    name_case = args.name_case or "title"

    records = data.filter_records(
        region=region,
        province=province,
        search_query=args.search,
        name_case=name_case
    )

    print(f"Found {len(records)} municipalities matching criteria:")
    print(f"  Region  : {region}")
    print(f"  Province: {province}")
    if args.search:
        print(f"  Search  : '{args.search}'")
    print(f"  Casing  : {name_case}")

    output_path = args.output
    if not output_path:
        # Default filename
        reg_clean = region.replace(" ", "_") if region != "ALL" else "ALL_REGIONS"
        prov_clean = f"_{province.replace(' ', '_')}" if province != "ALL" else ""
        output_path = f"Officials_{reg_clean}{prov_clean}.csv"

    count = data.export_csv(
        filepath=output_path,
        records=records,
        include_region=args.include_region
    )
    print(f"\n[SUCCESS] Exported {count} records to: {os.path.abspath(output_path)}")


def run_web(port: int = 8000, excel_path: str = "Officials_List.xlsx"):
    """Run an interactive web interface using Python's standard library."""
    import http.server
    import socketserver
    import json
    import urllib.parse

    data = OfficialsData(excel_path)
    data.load()

    class WebHandler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urllib.parse.urlparse(self.path)
            path = parsed.path
            query = urllib.parse.parse_qs(parsed.query)

            if path == "/" or path == "/index.html":
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                html = self.get_html()
                self.wfile.write(html.encode("utf-8"))
            elif path == "/api/meta":
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                meta = {
                    "regions": data.regions,
                    "provinces_by_region": data.provinces_by_region
                }
                self.wfile.write(json.dumps(meta).encode("utf-8"))
            elif path == "/api/records":
                reg = query.get("region", ["ALL"])[0]
                prov = query.get("province", ["ALL"])[0]
                q = query.get("q", [""])[0]
                casing = query.get("casing", ["title"])[0]

                records = data.filter_records(
                    region=reg,
                    province=prov,
                    search_query=q,
                    name_case=casing
                )
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps(records).encode("utf-8"))
            elif path == "/api/export":
                reg = query.get("region", ["ALL"])[0]
                prov = query.get("province", ["ALL"])[0]
                q = query.get("q", [""])[0]
                casing = query.get("casing", ["title"])[0]
                include_reg = query.get("include_region", ["0"])[0] == "1"

                records = data.filter_records(
                    region=reg,
                    province=prov,
                    search_query=q,
                    name_case=casing
                )

                import csv
                import io
                output = io.StringIO()
                fieldnames = list(data.DEFAULT_CSV_COLUMNS)
                if include_reg:
                    fieldnames.append("Region")

                writer = csv.DictWriter(output, fieldnames=fieldnames, extrasaction="ignore")
                writer.writeheader()
                for r in records:
                    writer.writerow(r)

                csv_bytes = ("\ufeff" + output.getvalue()).encode("utf-8")
                filename = f"Officials_{reg.replace(' ', '_')}_{prov.replace(' ', '_')}.csv"

                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(csv_bytes)))
                self.end_headers()
                self.wfile.write(csv_bytes)
            else:
                self.send_response(404)
                self.end_headers()

        def get_html(self):
            return """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Philippine Municipal Officials Directory & CSV Exporter</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --bg: #0f172a;
      --card: #1e293b;
      --card-border: #334155;
      --primary: #3b82f6;
      --primary-hover: #2563eb;
      --text: #f8fafc;
      --text-muted: #94a3b8;
      --accent: #10b981;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif; }
    body { background: var(--bg); color: var(--text); padding: 24px; min-height: 100vh; }
    .container { max-width: 1350px; margin: 0 auto; }
    header { margin-bottom: 24px; text-align: center; }
    h1 { font-size: 28px; font-weight: 700; color: #fff; margin-bottom: 8px; }
    p.sub { color: var(--text-muted); font-size: 14px; }
    .filters-card {
      background: var(--card); border: 1px solid var(--card-border);
      border-radius: 12px; padding: 20px; margin-bottom: 20px;
      display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)) auto;
      gap: 16px; align-items: end;
    }
    .field { display: flex; flex-direction: column; gap: 6px; }
    label { font-size: 13px; font-weight: 600; color: var(--text-muted); }
    select, input {
      background: #0f172a; border: 1px solid var(--card-border); color: #fff;
      padding: 10px 14px; border-radius: 8px; font-size: 14px; outline: none;
    }
    select:focus, input:focus { border-color: var(--primary); }
    .btn {
      background: var(--primary); color: #fff; border: none; padding: 10px 20px;
      border-radius: 8px; font-weight: 600; cursor: pointer; font-size: 14px;
      transition: background 0.2s; display: inline-flex; align-items: center; gap: 8px;
    }
    .btn:hover { background: var(--primary-hover); }
    .btn-success { background: var(--accent); }
    .btn-success:hover { background: #059669; }
    .stats-row {
      display: flex; justify-content: space-between; align-items: center;
      margin-bottom: 12px; color: var(--text-muted); font-size: 14px;
    }
    .table-container {
      background: var(--card); border: 1px solid var(--card-border);
      border-radius: 12px; overflow: hidden; max-height: 600px; overflow-y: auto;
    }
    table { width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }
    th {
      background: #111827; position: sticky; top: 0; padding: 12px 14px;
      font-weight: 600; color: #94a3b8; border-bottom: 1px solid var(--card-border);
    }
    td { padding: 12px 14px; border-bottom: 1px solid rgba(51, 65, 85, 0.4); }
    tr:hover td { background: rgba(59, 130, 246, 0.08); }
    .badge {
      display: inline-block; padding: 3px 8px; border-radius: 6px; font-size: 11px;
      background: rgba(59, 130, 246, 0.15); color: #60a5fa; font-weight: 600;
    }
  </style>
</head>
<body>
  <div class="container">
    <header>
      <h1>🏛️ Philippine Municipal Officials Directory</h1>
      <p class="sub">1,488 Municipalities • 18 Regions • Export Mayor Name, Municipal Hall Address, Town, Province, Vice Mayor</p>
    </header>

    <div class="filters-card">
      <div class="field">
        <label>Region</label>
        <select id="regSelect" onchange="onRegionChange()"></select>
      </div>
      <div class="field">
        <label>Province</label>
        <select id="provSelect" onchange="fetchRecords()"></select>
      </div>
      <div class="field">
        <label>Search Town / Official</label>
        <input type="text" id="searchInput" placeholder="Search..." oninput="fetchRecords()">
      </div>
      <div class="field">
        <label>Name Formatting</label>
        <select id="casingSelect" onchange="fetchRecords()">
          <option value="title">Title Case (e.g. Juan Dela Cruz)</option>
          <option value="upper">UPPERCASE</option>
          <option value="original">Original from File</option>
        </select>
      </div>
      <div style="display:flex; gap:8px;">
        <button class="btn btn-success" onclick="exportCsv()">⬇ Export CSV</button>
      </div>
    </div>

    <div class="stats-row">
      <div id="statusText">Loading...</div>
      <div>
        <label style="cursor:pointer; display:inline-flex; align-items:center; gap:6px;">
          <input type="checkbox" id="includeRegCheck" onchange="fetchRecords()"> Include Region column in CSV
        </label>
      </div>
    </div>

    <div class="table-container">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Town</th>
            <th>Province</th>
            <th id="thRegion" style="display:none;">Region</th>
            <th>Mayor Name</th>
            <th>Vice Mayor</th>
            <th>Municipal Hall Address</th>
          </tr>
        </thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>
  </div>

  <script>
    let meta = { regions: [], provinces_by_region: {} };

    async function init() {
      const res = await fetch('/api/meta');
      meta = await res.json();
      const regSelect = document.getElementById('regSelect');
      regSelect.innerHTML = '<option value="ALL">All Regions (18 Regions)</option>' +
        meta.regions.map(r => `<option value="${r}">${r}</option>`).join('');
      onRegionChange();
    }

    function onRegionChange() {
      const reg = document.getElementById('regSelect').value;
      const provSelect = document.getElementById('provSelect');
      let provs = [];
      if (reg === 'ALL') {
        const allSet = new Set();
        Object.values(meta.provinces_by_region).forEach(arr => arr.forEach(p => allSet.add(p)));
        provs = Array.from(allSet).sort();
        provSelect.innerHTML = '<option value="ALL">All Provinces</option>' +
          provs.map(p => `<option value="${p}">${p}</option>`).join('');
      } else {
        provs = meta.provinces_by_region[reg] || [];
        provSelect.innerHTML = `<option value="ALL">All Provinces (${reg})</option>` +
          provs.map(p => `<option value="${p}">${p}</option>`).join('');
      }
      fetchRecords();
    }

    async function fetchRecords() {
      const reg = document.getElementById('regSelect').value;
      const prov = document.getElementById('provSelect').value;
      const q = document.getElementById('searchInput').value;
      const casing = document.getElementById('casingSelect').value;
      const incReg = document.getElementById('includeRegCheck').checked;

      const thReg = document.getElementById('thRegion');
      if (thReg) thReg.style.display = incReg ? '' : 'none';

      const params = new URLSearchParams({ region: reg, province: prov, q: q, casing: casing });
      const res = await fetch('/api/records?' + params.toString());
      const data = await res.json();

      document.getElementById('statusText').innerText = `Showing ${data.length} municipal records • ${incReg ? '6 columns (incl. Region)' : '5 columns (Standard)'}`;

      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = data.map((r, i) => `
        <tr>
          <td><span class="badge">${i+1}</span></td>
          <td><strong>${r.Town}</strong></td>
          <td>${r.Province}</td>
          ${incReg ? `<td>${r.Region}</td>` : ''}
          <td style="color:#60a5fa; font-weight:600;">${r['Mayor Name'] || '<span style="color:#64748b">None</span>'}</td>
          <td style="color:#34d399; font-weight:600;">${r['Vice Mayor'] || '<span style="color:#64748b">None</span>'}</td>
          <td style="font-size:12px; color:#cbd5e1;">${r['Municipal Hall Address'] || '<span style="color:#64748b">None</span>'}</td>
        </tr>
      `).join('');
    }

    function exportCsv() {
      const reg = document.getElementById('regSelect').value;
      const prov = document.getElementById('provSelect').value;
      const q = document.getElementById('searchInput').value;
      const casing = document.getElementById('casingSelect').value;
      const incReg = document.getElementById('includeRegCheck').checked ? '1' : '0';

      const params = new URLSearchParams({
        region: reg, province: prov, q: q, casing: casing, include_region: incReg
      });
      window.location.href = '/api/export?' + params.toString();
    }

    init();
  </script>
</body>
</html>"""

    with socketserver.TCPServer(("", port), WebHandler) as httpd:
        url = f"http://localhost:{port}"
        print(f"\n[WEB APP] Running Philippine LGU Directory Web Server at: {url}")
        print("Press Ctrl+C to stop the server.\n")
        webbrowser.open(url)
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down web server.")


def run_gui(excel_path: str = "Officials_List.xlsx"):
    """Launch full native Windows Tkinter Desktop GUI application."""
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    data = OfficialsData(excel_path)
    try:
        data.load()
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error Loading Data", f"Could not load '{excel_path}':\n{e}")
        return

    root = tk.Tk()
    root.title("Philippine Municipal Officials Directory & CSV Exporter")
    root.geometry("1180x760")
    root.minsize(920, 600)

    # Set icon if available or configure theme
    style = ttk.Style()
    available_themes = style.theme_names()
    if "vista" in available_themes:
        style.theme_use("vista")
    elif "clam" in available_themes:
        style.theme_use("clam")

    # Color palette
    COLOR_BG = "#f8fafc"         # Slate 50
    COLOR_CARD = "#ffffff"       # White
    COLOR_BORDER = "#e2e8f0"     # Slate 200
    COLOR_HEADER = "#0f172a"     # Slate 900
    COLOR_PRIMARY = "#2563eb"    # Blue 600
    COLOR_SUCCESS = "#059669"    # Emerald 600

    root.configure(bg=COLOR_BG)

    # State variables
    region_var = tk.StringVar(value="All Regions")
    province_var = tk.StringVar(value="All Provinces")
    search_var = tk.StringVar()
    casing_var = tk.StringVar(value="Title Case")
    include_region_var = tk.BooleanVar(value=False)
    status_var = tk.StringVar(value="Ready • 1,488 Municipalities loaded")
    count_summary_var = tk.StringVar(value="Showing 1,488 records")

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
        text="🏛️ Philippine Municipal Officials Directory & CSV Exporter",
        font=("Segoe UI", 16, "bold"),
        fg=COLOR_HEADER,
        bg=COLOR_BG
    )
    title_label.pack(anchor="w")

    sub_label = tk.Label(
        header_frame,
        text="Filter by Region & Province to generate CSV with: Mayor Name, Municipal Hall Address, Town, Province, Vice Mayor",
        font=("Segoe UI", 9),
        fg="#64748b",
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
    # Row 0: Region, Province, Quick Search
    tk.Label(control_card, text="Region:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD, fg="#475569").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
    region_combo = ttk.Combobox(control_card, textvariable=region_var, state="readonly", width=22, font=("Segoe UI", 9))
    region_options = ["All Regions"] + data.regions
    region_combo["values"] = region_options
    region_combo.grid(row=0, column=1, sticky="w", padx=(0, 16), pady=4)

    tk.Label(control_card, text="Province:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD, fg="#475569").grid(row=0, column=2, sticky="w", padx=(0, 6), pady=4)
    province_combo = ttk.Combobox(control_card, textvariable=province_var, state="readonly", width=24, font=("Segoe UI", 9))
    province_combo.grid(row=0, column=3, sticky="w", padx=(0, 16), pady=4)

    tk.Label(control_card, text="Search:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD, fg="#475569").grid(row=0, column=4, sticky="w", padx=(0, 6), pady=4)
    search_entry = ttk.Entry(control_card, textvariable=search_var, width=22, font=("Segoe UI", 9))
    search_entry.grid(row=0, column=5, sticky="w", padx=(0, 6), pady=4)

    btn_clear_search = ttk.Button(control_card, text="✕", width=3, command=lambda: (search_var.set(""), on_filter_change()))
    btn_clear_search.grid(row=0, column=6, sticky="w", pady=4)

    # Row 1: Options & Action Buttons
    options_frame = tk.Frame(control_card, bg=COLOR_CARD)
    options_frame.grid(row=1, column=0, columnspan=7, sticky="ew", pady=(8, 0))

    tk.Label(options_frame, text="Name Casing:", font=("Segoe UI", 9), bg=COLOR_CARD, fg="#475569").pack(side=tk.LEFT, padx=(0, 6))
    casing_combo = ttk.Combobox(
        options_frame,
        textvariable=casing_var,
        state="readonly",
        width=14,
        font=("Segoe UI", 9),
        values=["Title Case", "UPPERCASE", "Original"]
    )
    casing_combo.pack(side=tk.LEFT, padx=(0, 16))

    check_inc_region = ttk.Checkbutton(
        options_frame,
        text="Include Region in CSV & Table",
        variable=include_region_var,
        command=lambda: on_filter_change()
    )
    check_inc_region.pack(side=tk.LEFT, padx=(0, 16))

    btn_reset = ttk.Button(options_frame, text="Reset Filters", command=lambda: reset_filters())
    btn_reset.pack(side=tk.LEFT, padx=(0, 16))

    # Right side action buttons
    btn_export = tk.Button(
        options_frame,
        text="⬇ Export to CSV",
        bg=COLOR_SUCCESS,
        fg="white",
        font=("Segoe UI", 9, "bold"),
        activebackground="#047857",
        activeforeground="white",
        relief=tk.FLAT,
        padx=14,
        pady=4,
        cursor="hand2",
        command=lambda: export_selected_csv()
    )
    btn_export.pack(side=tk.RIGHT, padx=(6, 0))

    btn_clipboard = tk.Button(
        options_frame,
        text="📋 Copy CSV",
        bg="#0284c7",
        fg="white",
        font=("Segoe UI", 9),
        activebackground="#0369a1",
        activeforeground="white",
        relief=tk.FLAT,
        padx=10,
        pady=4,
        cursor="hand2",
        command=lambda: copy_to_clipboard()
    )
    btn_clipboard.pack(side=tk.RIGHT, padx=(0, 6))

    # 3. Stats & Summary Bar
    stats_bar = tk.Frame(main_frame, bg=COLOR_BG)
    stats_bar.pack(fill=tk.X, pady=(2, 6))

    lbl_count = tk.Label(stats_bar, textvariable=count_summary_var, font=("Segoe UI", 9, "bold"), fg="#1e293b", bg=COLOR_BG)
    lbl_count.pack(side=tk.LEFT)

    lbl_hint = tk.Label(stats_bar, text="💡 Tip: Double-click any row to view full details • Click headers to sort", font=("Segoe UI", 8), fg="#64748b", bg=COLOR_BG)
    lbl_hint.pack(side=tk.RIGHT)

    # 4. Preview Table (Treeview)
    table_card = tk.Frame(main_frame, bg=COLOR_CARD, bd=1, relief=tk.SOLID)
    table_card.pack(fill=tk.BOTH, expand=True)

    col_widths = {
        "Town": 150,
        "Province": 150,
        "Region": 120,
        "Mayor Name": 210,
        "Vice Mayor": 210,
        "Municipal Hall Address": 320
    }

    def get_active_columns():
        if include_region_var.get():
            return ("Town", "Province", "Region", "Mayor Name", "Vice Mayor", "Municipal Hall Address")
        else:
            return ("Town", "Province", "Mayor Name", "Vice Mayor", "Municipal Hall Address")

    tree = ttk.Treeview(table_card, columns=get_active_columns(), show="headings", selectmode="browse")

    # Scrollbars
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

    file_label = tk.Label(status_bar, text=f"Source: {os.path.basename(excel_path)}", font=("Segoe UI", 8), fg="#94a3b8", bg="#0f172a")
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

        if reg == "All Regions":
            reg_filter = None
        else:
            reg_filter = reg

        if "All Provinces" in prov:
            prov_filter = None
        else:
            prov_filter = prov

        search = search_var.get().strip()
        casing = get_casing_mode()

        current_records = data.filter_records(
            region=reg_filter,
            province=prov_filter,
            search_query=search,
            name_case=casing
        )

        refresh_table()

    def refresh_table():
        active_cols = get_active_columns()

        # Reconfigure tree columns dynamically
        if tuple(tree["columns"]) != active_cols:
            tree["columns"] = active_cols
            for col in active_cols:
                tree.column(col, width=col_widths.get(col, 150), minwidth=80, anchor="w")

        # Set headings with sort indicators
        for col in active_cols:
            indicator = ""
            if col == sort_column:
                indicator = " ▼" if sort_descending else " ▲"
            tree.heading(col, text=col + indicator, command=lambda _c=col: sort_table_by_column(_c))

        # Clear tree
        for item in tree.get_children():
            tree.delete(item)

        # Sort if needed
        sorted_records = list(current_records)
        if sort_column:
            def sort_key(rec):
                val = rec.get(sort_column, "")
                return str(val).lower()
            sorted_records.sort(key=sort_key, reverse=sort_descending)

        # Insert rows
        inc_reg = include_region_var.get()
        mayors_count = 0
        vms_count = 0
        for i, rec in enumerate(sorted_records):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            if rec["Mayor Name"]:
                mayors_count += 1
            if rec["Vice Mayor"]:
                vms_count += 1

            if inc_reg:
                values = (
                    rec["Town"],
                    rec["Province"],
                    rec["Region"],
                    rec["Mayor Name"],
                    rec["Vice Mayor"],
                    rec["Municipal Hall Address"]
                )
            else:
                values = (
                    rec["Town"],
                    rec["Province"],
                    rec["Mayor Name"],
                    rec["Vice Mayor"],
                    rec["Municipal Hall Address"]
                )
            tree.insert("", tk.END, values=values, tags=(tag,))

        # Update summary text
        reg_disp = region_var.get()
        prov_disp = province_var.get()
        filter_desc = f"{reg_disp}" if prov_disp.startswith("All") else f"{prov_disp}, {reg_disp}"
        col_count_txt = "6 columns (incl. Region)" if inc_reg else "5 standard columns"
        count_summary_var.set(
            f"Showing {len(sorted_records)} municipalities ({filter_desc}) • CSV: {col_count_txt} • Mayors: {mayors_count} • Vice Mayors: {vms_count}"
        )
        status_var.set(f"Loaded {len(sorted_records)} records • Region column: {'ON' if inc_reg else 'OFF'} • Ready to export")

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
        on_filter_change()

    def on_region_selected(event):
        update_province_dropdown()
        on_filter_change()

    def on_province_selected(event):
        on_filter_change()

    def on_casing_selected(event):
        on_filter_change()

    # Search debounce or trace
    search_var.trace_add("write", lambda *args: on_filter_change())
    region_combo.bind("<<ComboboxSelected>>", on_region_selected)
    province_combo.bind("<<ComboboxSelected>>", on_province_selected)
    casing_combo.bind("<<ComboboxSelected>>", on_casing_selected)

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

        detail_win = tk.Toplevel(root)
        detail_win.title(f"Details - {rec['Town']}, {rec['Province']}")
        detail_win.geometry("560x390")
        detail_win.minsize(460, 320)
        detail_win.configure(bg=COLOR_BG)
        detail_win.transient(root)

        card = tk.Frame(detail_win, bg=COLOR_CARD, padx=20, pady=20, bd=1, relief=tk.SOLID)
        card.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

        fields = [
            ("Town:", rec["Town"]),
            ("Province:", rec["Province"]),
            ("Region:", rec["Region"]),
            ("Mayor Name:", rec["Mayor Name"]),
            ("Vice Mayor:", rec["Vice Mayor"]),
            ("Municipal Hall Address:", rec["Municipal Hall Address"]),
        ]

        for i, (label_txt, val_txt) in enumerate(fields):
            tk.Label(card, text=label_txt, font=("Segoe UI", 9, "bold"), fg="#475569", bg=COLOR_CARD, anchor="w").grid(row=i, column=0, sticky="nw", pady=4, padx=(0, 10))
            val_lbl = tk.Label(card, text=val_txt or "(None)", font=("Segoe UI", 9), fg="#0f172a", bg=COLOR_CARD, anchor="w", wraplength=360, justify="left")
            val_lbl.grid(row=i, column=1, sticky="w", pady=4)

        btn_box = tk.Frame(card, bg=COLOR_CARD)
        btn_box.grid(row=len(fields), column=0, columnspan=2, pady=(16, 0), sticky="e")

        def copy_address():
            root.clipboard_clear()
            root.clipboard_append(rec["Municipal Hall Address"])
            messagebox.showinfo("Copied", "Municipal Hall Address copied to clipboard!", parent=detail_win)

        ttk.Button(btn_box, text="Copy Address", command=copy_address).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_box, text="Close", command=detail_win.destroy).pack(side=tk.LEFT)

    tree.bind("<Double-1>", show_row_details)

    def export_selected_csv():
        if not current_records:
            messagebox.showwarning("No Records", "There are no records matching your current filter to export.")
            return

        reg = region_var.get()
        prov = province_var.get()
        reg_clean = reg.replace(" ", "_") if reg != "All Regions" else "ALL_REGIONS"
        prov_clean = f"_{prov.replace(' ', '_')}" if not prov.startswith("All Provinces") else ""
        default_filename = f"Officials_{reg_clean}{prov_clean}.csv"

        filepath = filedialog.asksaveasfilename(
            parent=root,
            title="Save CSV File",
            initialfile=default_filename,
            defaultextension=".csv",
            filetypes=[("CSV Files (*.csv)", "*.csv"), ("All Files (*.*)", "*.*")]
        )

        if not filepath:
            return

        try:
            count = data.export_csv(
                filepath=filepath,
                records=current_records,
                include_region=include_region_var.get()
            )
            status_var.set(f"Successfully exported {count} records to: {os.path.basename(filepath)}")

            # Ask user if they want to open the file or containing folder
            resp = messagebox.askyesno(
                "Export Complete",
                f"Successfully exported {count} municipal officials to:\n{filepath}\n\nWould you like to open this CSV file now?"
            )
            if resp:
                try:
                    os.startfile(filepath)
                except Exception as ex:
                    messagebox.showerror("Could not open file", str(ex))
        except Exception as e:
            messagebox.showerror("Export Failed", f"An error occurred while saving the CSV:\n{e}")

    def copy_to_clipboard():
        if not current_records:
            messagebox.showwarning("No Records", "No records to copy.")
            return

        import io
        import csv
        output = io.StringIO()
        fieldnames = list(data.DEFAULT_CSV_COLUMNS)
        if include_region_var.get():
            fieldnames.append("Region")

        writer = csv.DictWriter(output, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for r in current_records:
            writer.writerow(r)

        root.clipboard_clear()
        root.clipboard_append(output.getvalue())
        status_var.set(f"Copied {len(current_records)} records to clipboard as CSV!")
        messagebox.showinfo("Copied", f"Copied {len(current_records)} records to clipboard!\nYou can now paste (Ctrl+V) directly into Excel or a text editor.")

    # Initialize dropdown and table
    update_province_dropdown()
    on_filter_change()

    # Center window on screen
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
        description="Philippine Municipal Officials Directory & CSV Exporter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Launch the Desktop GUI:
  python app.py

  # Export Region 3, Pampanga:
  python app.py --region "REGION 3" --province "PAMPANGA" -o pampanga.csv

  # Export all municipalities nationwide:
  python app.py --all -o all_officials.csv

  # Start the web interface:
  python app.py --web
        """
    )
    parser.add_argument("--excel", default="Officials_List.xlsx", help="Path to Officials_List.xlsx (default: Officials_List.xlsx)")
    parser.add_argument("--region", help="Region to filter (e.g. 'REGION 3', 'NCR', 'BARMM')")
    parser.add_argument("--province", help="Province to filter (e.g. 'PAMPANGA', 'TARLAC')")
    parser.add_argument("--search", help="Search text filter for town or official name")
    parser.add_argument("--output", "-o", help="Output CSV file path")
    parser.add_argument("--all", action="store_true", help="Export all regions and provinces")
    parser.add_argument("--name-case", choices=["title", "upper", "original"], default="title", help="Name casing style (default: title)")
    parser.add_argument("--include-region", action="store_true", help="Include Region column in exported CSV")
    parser.add_argument("--list-regions", action="store_true", help="List all available regions and exit")
    parser.add_argument("--list-provinces", nargs="?", const="ALL", help="List provinces for a region (or all if not specified) and exit")
    parser.add_argument("--web", nargs="?", const=8000, type=int, help="Start web interface on specified port (default: 8000)")
    parser.add_argument("--cli", action="store_true", help="Force command line mode")

    args = parser.parse_args()

    # Determine execution mode
    is_cli_mode = (
        args.cli or
        args.list_regions or
        args.list_provinces is not None or
        args.region is not None or
        args.province is not None or
        args.all or
        args.output is not None
    )

    if args.web is not None:
        run_web(port=args.web, excel_path=args.excel)
    elif is_cli_mode:
        data = OfficialsData(args.excel)
        run_cli(args, data)
    else:
        # Default: Launch Tkinter GUI
        run_gui(excel_path=args.excel)


if __name__ == "__main__":
    main()
