# Smart Meter JSON File Editor (Excel-Driven & Before/After Comparison)

A high-performance, structure-preserving JSON editor tailored specifically for DLMS/COSEM Smart Electricity Meter Load Survey data.

---

## 🎯 Key Features Added

1. **📊 4-Column Excel / CSV Input**:
   - Provide an Excel (`.xlsx`, `.xls`) or `.csv` file with 4 columns:
     1. **Meter Number** (`MeterNumber`)
     2. **Date & Time** (`0.0.1.0.0.255`)
     3. **Import kWh** (`1.0.1.29.0.255`)
     4. **Export kWh** (`1.0.2.29.0.255`)
   - The editor matches each row by `(Meter Number, Date & Time)` in the loaded JSON file(s) and updates the corresponding kWh values automatically.
   - Includes sample templates: [`sample_excel_input.xlsx`](sample_excel_input.xlsx) and [`sample_excel_input.csv`](sample_excel_input.csv).

2. **👁️ Before & After Data Viewer**:
   - The spreadsheet table displays both the original and modified values side-by-side:
     - `Meter Number`
     - `Date & Time`
     - `Import kWh [BEFORE]` vs `Import kWh [AFTER / EDITABLE]` (highlighted in amber when modified)
     - `Export kWh [BEFORE]` vs `Export kWh [AFTER / EDITABLE]` (highlighted in amber when modified)
     - `Status` badge (`Original` / `Modified`)
   - Filter by: `All Records`, `Modified Only`, or `Original Only`.

3. **💾 Automatic `_Modified` Output Naming**:
   - Output files are automatically saved with the `_Modified` suffix (e.g., `sample_load_survey__Modified.json`).
   - 100% byte-exact JSON structure preservation: all ~50+ other electrical parameters remain completely untouched.

4. **📥 Converted Data Exports & 9-Column Workflow**:
   - **Download 9-Col Excel Template**: Pre-populated template based on selected JSON file with Before & After columns.
   - **Export Source Data (CSV)**: Original 4-parameter tabular data.
   - **Export Output Data (CSV)**: Final updated 4-parameter tabular data.
   - **Export Before & After Comparison (CSV)**: Side-by-side comparison table with Status for reporting and audit verification.

---

## 🚀 How to Run

### Option 1: Standalone Windows Executable
Double-click:
```
SmartMeterJsonEditor.exe
```

### Option 2: Launcher Script
Double-click:
```
run_editor.bat
```

### Option 3: Command Line (CLI / Automation)
```powershell
# Apply Excel updates, export comparison CSV, and save _Modified.json:
python web_editor.py --file "samples/sample_load_survey.json" --excel "sample_excel_input.xlsx" --export-csv comparison --out "output"
```
