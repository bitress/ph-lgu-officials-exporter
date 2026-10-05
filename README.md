# 🏛️ Philippine Municipal Officials Directory & CSV Exporter

A Python application designed to process and export data from `Officials_List.xlsx` into clean, Excel-compatible CSV files with dynamic **Region** and **Province** selectors.

---

## 📋 Features

- **Exact Required Columns**:
  - `Mayor Name`
  - `Municipal Hall Address`
  - `Town`
  - `Province`
  - `Vice Mayor`
  *(Optional checkbox/flag to also include `Region` if needed)*
- **Dual Selectors**:
  - **Region Selector**: Select any of the 18 Philippine regions (or "All Regions").
  - **Dynamic Province Selector**: Automatically populates with provinces belonging to the selected Region (or "All Provinces").
- **Smart Data Processing**:
  - Automatically matches and pairs each municipality's **Mayor** and **Vice Mayor**.
  - **Municipal Hall Address Fallback**: If an address is missing from the Mayor's row, it automatically draws the valid municipal hall address from the Vice Mayor or Sangguniang Bayan records (achieving 100% address coverage across all 1,488 municipalities).
  - Cleans and formats official names, stripping extraneous `'N/A'` suffixes while preserving legitimate titles like `Jr.`, `Sr.`, `II`, `III`, `IV`.
  - Fixes Spanish en-eñe (`ñ`/`Ñ`) character decoding artifacts (e.g. `Peñarrubia`, `Dasmariñas`).
  - Saves in **UTF-8 with BOM (`utf-8-sig`)**, ensuring accented characters open cleanly in Microsoft Excel on Windows without garbling.
- **Three Ways to Run**:
  1. **Native Desktop GUI** (Windows Tkinter with live preview, search, sortable columns, and one-click export).
  2. **Double-click Launcher** (`run_app.bat`).
  3. **Command Line (CLI)** for automated scripting.
  4. **Local Web App Mode** (`python app.py --web`).

---

## 🚀 How to Run

### Method 1: Double-Click Desktop Launcher
Simply double-click:
```
run_app.bat
```
This opens the desktop GUI immediately.

---

### Method 2: Python Command (Desktop GUI)
Open your terminal in this folder and run:
```bash
python app.py
```

---

### Method 3: Command Line (CLI) Mode
You can also generate CSV files directly from the terminal without opening the GUI:

#### 1. Export a specific Region & Province:
```bash
python app.py --region "REGION 3" --province "PAMPANGA" -o pampanga_officials.csv
```

#### 2. Export an entire Region:
```bash
python app.py --region "REGION 1" -o region_1_officials.csv
```

#### 3. Export all 1,488 Municipalities nationwide:
```bash
python app.py --all -o all_officials.csv
```

#### 4. List all available Regions:
```bash
python app.py --list-regions
```

#### 5. List all Provinces in a Region:
```bash
python app.py --list-provinces "REGION 3"
```

#### 6. Change Name Casing (`title`, `upper`, `original`):
```bash
python app.py --region "REGION 4A" --province "CAVITE" --name-case upper -o cavite_upper.csv
```

---

### Method 4: Local Web Interface
If you prefer using a browser interface:
```bash
python app.py --web
```
This starts a lightweight web app at `http://localhost:8000` and automatically opens your browser.

---

## 📁 File Structure

| File | Purpose |
|---|---|
| [`app.py`](file:///c:/Users/bitress/Desktop/lgu%20official/app.py) | Main application entry point (Desktop GUI, CLI, & Web server) |
| [`data_processor.py`](file:///c:/Users/bitress/Desktop/lgu%20official/data_processor.py) | Data parsing, pairing logic, address fallback, and CSV exporter engine |
| [`run_app.bat`](file:///c:/Users/bitress/Desktop/lgu%20official/run_app.bat) | Windows 1-click launcher script |
| [`Officials_List.xlsx`](file:///c:/Users/bitress/Desktop/lgu%20official/Officials_List.xlsx) | Source dataset containing all Philippine LGU officials |
| [`all_municipalities.csv`](file:///c:/Users/bitress/Desktop/lgu%20official/all_municipalities.csv) | Pre-generated nationwide export of all 1,488 municipalities |
| [`pampanga_officials.csv`](file:///c:/Users/bitress/Desktop/lgu%20official/pampanga_officials.csv) | Sample generated export for Region 3 (Pampanga) |
