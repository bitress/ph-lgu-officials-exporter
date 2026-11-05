# 🏛️ Philippine Local Government Officials Directory & Multi-Format Exporter

A high-performance Python application designed to process, search, edit, and export 2025–2028 elected officials from `officials_2025_2028.xlsx` (and `.csv`) into clean, Excel-compatible CSV, styled Excel (`.xlsx`), and JSON files.

---

## 📋 Key Features

- **Standard Required Columns**:
  - `Mayor Name`
  - `Municipal Hall Address`
  - `Town`
  - `Province`
  - `Vice Mayor`
  - *(Optional toggles to include `LGU Type` (City / Municipality), `Region`, `Mayor Contact`, and `Vice Mayor Contact / Email`)*
- **City & Municipality Filtering & Exporting**:
  - Nationwide **1,633 LGUs**: **1,484 Municipalities** and **149 Cities** across all **18 Regions** (including NIR, BARMM, and NCR).
  - Flexible export options: Export **All LGUs**, **Municipalities Only**, or **Cities Only**.
- **Ultra-Fast Startup & Caching**:
  - Uses `openpyxl` in `read_only=True` mode alongside a hash/mtime-validated JSON cache (`.officials_2025_2028_cache.json`).
  - Launches in **< 0.08 seconds** on subsequent runs.
- **Smart Data Cleaning & Address Fallback**:
  - **100% Address Coverage**: If an address is missing from the Mayor's row, it automatically draws the valid hall address from the Vice Mayor or Sangguniang Bayan/Panlungsod records across all 1,633 LGUs nationwide.
  - Cleans and formats official names, stripping extraneous `'N/A'`, `'N\A'`, `'-'`, `'MIR'` artifacts while preserving legitimate titles like `Jr.`, `Sr.`, `II`, `III`, `IV`.
  - Fixes Spanish en-eñe (`ñ`/`Ñ`) and common UTF-8 mojibake decoding artifacts (`Peñarrubia`, `Dasmariñas`, `Parañaque`).
- **Multi-Format Exporters**:
  - **CSV**: Encoded in **UTF-8 with BOM (`utf-8-sig`)** for seamless opening in Microsoft Excel on Windows without character corruption.
  - **Excel (`.xlsx`)**: Formatted workbook with dark navy headers, bold white text, zebra rows, freeze panes, and auto-adjusted column widths.
  - **JSON**: Formatted UTF-8 JSON for developers and APIs.
- **Safety First**:
  - **Rolling Timestamped Backups**: Every update creates an automatic snapshot in `backups/officials_2025_2028_YYYYMMDD_HHMMSS.xlsx` before modifying the spreadsheet.
- **Four Ways to Run**:
  1. **Native Desktop GUI** (Tkinter with LGU Type dropdown, live search, auto-fit columns, keyboard shortcuts, and one-click export).
  2. **Double-click Launcher** (`run_app.bat` with auto-dependency check).
  3. **Modern Web App** (`python app.py --web` built with Bootstrap 5, FontAwesome, LGU Type filter, pagination, metrics cards, and column sorting).
  4. **Command Line (CLI)** for automated scripting and data auditing.

---

## 🚀 How to Run

### Method 1: Double-Click Desktop Launcher
Simply double-click:
```
run_app.bat
```
This automatically verifies Python and `openpyxl`, then opens the desktop GUI immediately.

---

### Method 2: Python Command (Desktop GUI)
```bash
python app.py
```

**Desktop Keyboard Shortcuts**:
- `Ctrl + F`: Jump to Search input
- `Ctrl + E`: Export current table view to CSV
- `Double-click row`: View full LGU details, emails, and contact numbers
- `Right-click row`: Quick actions context menu (Edit, View, Copy)

---

### Method 3: Local Web Interface (Browser Mode)
```bash
python app.py --web
```
Starts a lightweight, zero-dependency local web app at `http://localhost:8000` and automatically opens your browser.

**Web App Features**:
- **Real-Time Metrics Cards**: Shows live counts for Total LGUs (1,633: 1,484 Municipalities & 149 Cities), Active Mayors (99.4%), Active Vice Mayors (99.7%), and Address Coverage (100%).
- **LGU Type Filter**: Easily filter by All, Municipalities Only, or Cities Only with a dedicated dropdown and column toggle.
- **Interactive Sorting**: Click any column header to sort ascending or descending.
- **Pagination Controls**: View 25, 50, 100, or All records per page with smooth navigation.
- **One-Click Exports**: Direct download buttons for CSV, Excel (`.xlsx`), and JSON.
- **Batch Export All**: Export all 18 regions to individual files with one button.

---

### Method 4: Command Line (CLI) Mode

#### 1. Data Health & Coverage Audit:
```bash
python app.py --stats
```

#### 2. Export Cities Only:
```bash
python app.py --cities-only --include-lgu-type -o philippine_cities.csv
```

#### 3. Export Municipalities Only:
```bash
python app.py --municipalities-only --format xlsx -o philippine_municipalities.xlsx
```

#### 4. Export a specific Region & Province to Excel (`.xlsx`):
```bash
python app.py --region "REGION 3" --province "PAMPANGA" --format xlsx -o pampanga.xlsx
```

#### 5. Export an entire Region with Contact Numbers & Emails:
```bash
python app.py --region "NCR" --include-contact -o ncr_with_contacts.csv
```

#### 6. Export all 1,633 LGUs nationwide:
```bash
python app.py --all --include-lgu-type -o all_officials.csv
```

#### 7. Batch-Export all 18 Regions into separate files:
```bash
python app.py --split-regions --format xlsx
```
*(Files will be saved into the `exports/` folder)*

#### 8. Change Name Casing (`title`, `upper`, `original`) or use Full Middle Name:
```bash
python app.py --region "REGION 4A" --province "CAVITE" --name-case upper -o cavite_upper.csv
python app.py --region "REGION 4A" --full-middle -o cavite_full_middle.csv
```

#### 9. List available Regions and Provinces:
```bash
python app.py --list-regions
python app.py --list-provinces "REGION 3"
```

---

## 🧪 Running Automated Unit Tests

A comprehensive unit test suite is included to ensure parsing accuracy, caching, address fallback, and export formatting:

```bash
python -m unittest discover tests
```

---

## 📁 File Structure

| File / Folder | Purpose |
|---|---|
| [`app.py`](app.py) | Main application entry point (Desktop GUI, CLI, & Web server) |
| [`data_processor.py`](data_processor.py) | Data parsing engine, caching, address fallback, and multi-format exporters |
| [`templates/index.html`](templates/index.html) | Bootstrap 5 & FontAwesome web UI with metrics dashboard, pagination, and sorting |
| [`run_app.bat`](run_app.bat) | Windows 1-click launcher with dependency verification |
| [`officials_2025_2028.xlsx`](officials_2025_2028.xlsx) | Primary dataset containing Philippine 2025–2028 elected officials (1,633 LGUs: 1,484 municipalities + 149 cities) |
| [`officials_2025_2028.csv`](officials_2025_2028.csv) | Raw CSV dataset (17,778 rows) across all 18 regions |
| [`requirements.txt`](requirements.txt) | Python dependencies (`openpyxl`) |
| `tests/` | Automated unit tests and web integration tests |
| `backups/` | Automated rolling timestamped backups of `officials_2025_2028.xlsx` |
| `exports/` | Destination folder for regional batch exports |
| `.officials_2025_2028_cache.json` | Local high-speed startup cache (auto-generated) |
