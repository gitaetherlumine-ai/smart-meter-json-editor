#!/usr/bin/env python3
"""
Smart Meter JSON Editor - Standalone Executable & Web Application
Enhanced with Excel 4-Column Batch Input, Before/After Comparison, and Multi-Format Exports.
Parameters:
  1. Meter Number (MeterNumber)
  2. Date & Time (OBIS 0.0.1.0.0.255)
  3. Import kWh (OBIS 1.0.1.29.0.255)
  4. Export kWh (OBIS 1.0.2.29.0.255)
Strictly preserves 100% of all other ~50+ electrical parameters and schema formatting.
"""

import os
import sys
import glob
import json
import base64
import argparse
import webbrowser
import threading
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from editor_core import JsonMeterEditor

# Windows GUI & Console safety
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.kernel32.AttachConsole(-1)
    except Exception:
        pass

if sys.stdout is None:
    try:
        sys.stdout = open(os.devnull, "w")
    except Exception:
        pass
else:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

if sys.stderr is None:
    try:
        sys.stderr = open(os.devnull, "w")
    except Exception:
        pass
else:
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Default directories
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.abspath(os.path.dirname(sys.executable))
else:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))

DEFAULT_SOURCE_DIR = BASE_DIR
DEFAULT_DEST_DIR = os.path.join(BASE_DIR, "output")

# Global editor instance
editor = JsonMeterEditor()
source_dir = DEFAULT_SOURCE_DIR
dest_dir = DEFAULT_DEST_DIR

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Smart Meter JSON Editor</title>
  <style>
    :root {
      --bg: #0f172a;
      --card: #1e293b;
      --border: #334155;
      --primary: #38bdf8;
      --primary-hover: #0284c7;
      --accent: #22c55e;
      --accent-hover: #16a34a;
      --danger: #ef4444;
      --text: #f8fafc;
      --text-muted: #94a3b8;
      --row-hover: #273549;
      --amber: #f59e0b;
      --amber-bg: rgba(245, 158, 11, 0.15);
      --green-bg: rgba(34, 197, 94, 0.15);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    body { background: var(--bg); color: var(--text); padding: 16px; font-size: 13px; }
    .container { max-width: 1560px; margin: 0 auto; display: flex; flex-direction: column; gap: 14px; }

    /* Cards & Header */
    .card { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 14px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.25); }
    .header-title { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
    .header-title h1 { font-size: 19px; font-weight: 700; color: var(--primary); display: flex; align-items: center; gap: 8px; }
    
    .badge { padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; text-transform: uppercase; }
    .badge-info { background: #0369a1; color: #e0f2fe; }
    .badge-kwh { background: #0369a1; color: #bae6fd; font-family: monospace; font-size: 12px; }
    .badge-warn { background: #b45309; color: #fef3c7; }
    .badge-success { background: #15803d; color: #dcfce7; }

    /* Top Picker Grid */
    .picker-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .input-group { display: flex; flex-direction: column; gap: 4px; }
    .input-group label { font-size: 11px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }
    .input-row { display: flex; gap: 8px; align-items: center; }
    input[type="text"], select { background: #0b1120; border: 1px solid var(--border); color: var(--text); padding: 7px 10px; border-radius: 6px; font-size: 13px; width: 100%; outline: none; }
    input[type="text"]:focus, select:focus { border-color: var(--primary); }

    /* Buttons */
    button { background: var(--primary); color: #0b1120; border: none; padding: 7px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; transition: 0.15s; display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; }
    button:hover { background: var(--primary-hover); color: white; }
    button.btn-accent { background: var(--accent); color: white; }
    button.btn-accent:hover { background: var(--accent-hover); }
    button.btn-secondary { background: #334155; color: var(--text); }
    button.btn-secondary:hover { background: #475569; }
    button.btn-warn { background: var(--amber); color: #0b1120; }
    button.btn-warn:hover { background: #d97706; color: white; }
    button.btn-excel { background: #107c41; color: white; }
    button.btn-excel:hover { background: #0b5a2f; }

    /* Multi-File Queue Strip */
    .queue-card { display: flex; flex-direction: column; gap: 8px; }
    .queue-header { display: flex; align-items: center; justify-content: space-between; }
    .queue-strip { display: flex; gap: 8px; overflow-x: auto; padding-bottom: 4px; align-items: center; }
    .file-chip { display: flex; align-items: center; gap: 6px; background: #0b1120; border: 1px solid var(--border); padding: 6px 12px; border-radius: 6px; cursor: pointer; transition: 0.15s; font-size: 12px; }
    .file-chip:hover { border-color: var(--primary); }
    .file-chip.active { background: #13243a; border-color: var(--primary); outline: 2px solid rgba(56, 189, 248, 0.4); font-weight: 600; }
    .file-chip .status-dot { width: 8px; height: 8px; border-radius: 50%; background: #94a3b8; }
    .file-chip.active .status-dot { background: var(--primary); }
    .file-chip .status-dot.mod { background: var(--amber); }
    .file-chip .status-dot.saved { background: var(--accent); }

    /* Controls Bar */
    .controls-bar { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
    .filter-group { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
    .action-group { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }

    /* Table Container - BEFORE & AFTER COMPARISON */
    .table-container { max-height: 540px; overflow-y: auto; border: 1px solid var(--border); border-radius: 6px; background: #0b1120; }
    table { width: 100%; border-collapse: collapse; text-align: left; }
    th { position: sticky; top: 0; background: #1e293b; color: var(--text-muted); font-size: 11px; font-weight: 700; padding: 9px 10px; border-bottom: 2px solid var(--border); z-index: 10; text-transform: uppercase; }
    td { padding: 6px 10px; border-bottom: 1px solid #1e293b; font-size: 13px; }
    tr:hover { background: var(--row-hover); }
    
    .editable-cell { cursor: pointer; border-radius: 4px; padding: 4px 6px; transition: 0.15s; font-family: monospace; display: inline-block; width: 100%; }
    .editable-cell:hover { background: #334155; outline: 1px dashed var(--primary); }
    .editable-cell.active-edit { padding: 0; }
    .editable-cell input { width: 100%; background: #0369a1; color: #ffffff; border: 1px solid var(--primary); padding: 4px 6px; border-radius: 4px; font-family: monospace; font-size: 13px; }
    
    .before-val { color: #64748b; font-family: monospace; font-size: 12px; }
    .after-val { font-family: monospace; font-weight: 600; color: #38bdf8; }
    .modified-cell { background: var(--amber-bg); color: #fbbf24 !important; font-weight: 700; border-left: 2px solid var(--amber); }

    /* Modal */
    .modal-backdrop { position: fixed; inset: 0; background: rgba(0,0,0,0.7); display: none; align-items: center; justify-content: center; z-index: 1000; }
    .modal-backdrop.open { display: flex; }
    .modal { background: var(--card); border: 1px solid var(--border); border-radius: 8px; width: 550px; max-width: 95%; padding: 20px; display: flex; flex-direction: column; gap: 14px; }
    .modal-header { display: flex; justify-content: space-between; align-items: center; }
    .modal-header h3 { color: var(--primary); font-size: 15px; }
    .close-btn { background: transparent; color: var(--text-muted); font-size: 18px; padding: 0; }
    .close-btn:hover { background: transparent; color: var(--text); }
    textarea { width: 100%; height: 160px; background: #0b1120; border: 1px solid var(--border); color: var(--text); padding: 8px; border-radius: 6px; font-family: monospace; font-size: 12px; resize: vertical; }

    /* Dropdown Menu */
    .dropdown { position: relative; display: inline-block; }
    .dropdown-content { display: none; position: absolute; right: 0; top: 100%; background: #1e293b; border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 8px 16px rgba(0,0,0,0.4); z-index: 100; min-width: 220px; margin-top: 4px; }
    .dropdown-content a { color: var(--text); padding: 9px 14px; text-decoration: none; display: block; font-size: 12px; font-weight: 600; cursor: pointer; }
    .dropdown-content a:hover { background: #334155; color: var(--primary); }
    .dropdown.show .dropdown-content { display: block; }

    /* Toast Notification */
    .toast { position: fixed; bottom: 20px; right: 20px; background: var(--accent); color: white; padding: 12px 18px; border-radius: 6px; font-weight: 600; display: none; box-shadow: 0 4px 12px rgba(0,0,0,0.4); z-index: 2000; }
  </style>
</head>
<body>
  <div class="container">
    <!-- Top Header -->
    <div class="card">
      <div class="header-title">
        <h1>
          <span>[*] Smart Meter JSON File Editor</span>
          <span class="badge badge-info">Zero Schema Drift</span>
        </h1>
        <div style="display:flex; gap:8px; align-items:center;">
          <div id="unsavedBadge" class="badge badge-warn" style="display:none;">0 Unsaved Edits</div>

          <!-- 9-Column Excel Template & Upload Buttons -->
          <button id="btnDownloadTemplate" class="btn-secondary" style="background:#0284c7; color:white; font-weight:700;" onclick="downloadChangeTemplate()">📥 Download 9-Col Excel Template (.xlsx)</button>
          <button id="btnUploadExcel" class="btn-excel" style="font-weight:700;" onclick="openExcelUploadModal()">📊 Upload Excel Changes</button>

          <!-- Export Dropdown -->
          <div class="dropdown" id="exportDropdown">
            <button class="btn-secondary" onclick="toggleExportMenu()">📥 Export Data ▾</button>
            <div class="dropdown-content">
              <a onclick="exportData('source')">📄 Export Source Data (CSV)</a>
              <a onclick="exportData('output')">📄 Export Output Data (CSV)</a>
              <a onclick="exportData('comparison')">📊 Export Before & After Comparison (CSV)</a>
            </div>
          </div>

          <!-- Save Buttons -->
          <button class="btn-accent" onclick="saveActiveFile()">💾 Save Output (*__Modified.json)</button>
          <button class="btn-accent" style="background:#059669;" onclick="saveAllFiles()">💾 Save All Files</button>
        </div>
      </div>

      <!-- Hidden Native HTML5 File Pickers -->
      <input type="file" id="browserFileInput" accept=".json" style="display:none;" onchange="handleBrowserFile(event)">
      <input type="file" id="browserMultiInput" accept=".json" multiple style="display:none;" onchange="handleBrowserMulti(event)">
      <input type="file" id="browserFolderInput" webkitdirectory directory style="display:none;" onchange="handleBrowserFolder(event)">
      <input type="file" id="browserExcelInput" accept=".xlsx,.xls,.csv" style="display:none;" onchange="handleBrowserExcel(event)">

      <!-- Quick File Selector Dropdown -->
      <div style="background: linear-gradient(135deg, #0b1523 0%, #10243e 100%); border: 1.5px solid #0284c7; border-radius: 8px; padding: 12px 16px; margin: 12px 0; display: flex; align-items: center; justify-content: space-between; gap: 14px; box-shadow: 0 4px 14px rgba(2, 132, 199, 0.2);">
        <div style="display:flex; align-items:center; gap:12px; flex:1;">
          <span style="font-size:22px;">⚡</span>
          <div style="display:flex; flex-direction:column; gap:3px; flex:1;">
            <div style="display:flex; align-items:center; gap:8px;">
              <span style="font-size:12px; font-weight:800; color:#38bdf8; text-transform:uppercase; letter-spacing:0.5px;">
                ACTIVE JSON FILE SELECTION:
              </span>
              <span class="badge badge-info" id="availableCountBadge">Scanning...</span>
            </div>
            <select id="quickSelectDropdown" style="background:#0b1120; border:1px solid #0284c7; color:#f8fafc; font-weight:600; padding:8px 12px; border-radius:6px; font-size:13px; width:100%; cursor:pointer;" onchange="handleQuickSelect(this.value)">
              <option value="">-- Choose a JSON File to View & Edit Intervals --</option>
            </select>
          </div>
        </div>
        <div style="display:flex; gap:8px; align-items:flex-end;">
          <button class="btn-accent" onclick="loadAvailableFilesList(true)" style="padding:8px 14px; font-size:12px; font-weight:700;">🔄 Refresh List</button>
        </div>
      </div>

      <!-- File & Folder Browser Pickers -->
      <div class="picker-grid">
        <!-- Source Selection -->
        <div class="input-group">
          <label>Source JSON Files / Folder:</label>
          <div class="input-row">
            <input type="text" id="srcPath" placeholder="Path to file or folder..." onkeydown="if(event.key==='Enter') loadManualPath(this.value)">
            <button class="btn-secondary" onclick="triggerBrowserFile()">📄 Browse File</button>
            <button class="btn-secondary" onclick="triggerBrowserMulti()">📚 Multi-Files</button>
            <button class="btn-secondary" onclick="triggerBrowserFolder()">📂 Folder</button>
          </div>
        </div>

        <!-- Destination Selection -->
        <div class="input-group">
          <label>Destination Folder (Output JSONs):</label>
          <div class="input-row">
            <input type="text" id="destPath" placeholder="Output destination directory..." onkeydown="if(event.key==='Enter') updateDestPath(this.value)">
            <button class="btn-secondary" onclick="browseDestFolder()">📂 Browse Output</button>
          </div>
        </div>
      </div>
    </div>

    <!-- Multi-File Queue Strip -->
    <div class="card queue-card" id="queueCard" style="display:none;">
      <div class="queue-header">
        <span style="font-size:12px; font-weight:700; color:var(--text-muted); text-transform:uppercase;">
          Loaded Files Queue (<span id="queueCount">0</span>)
        </span>
        <button class="btn-secondary" style="font-size:11px; padding:4px 8px;" onclick="openBulkActionModal()">🌐 Bulk Actions Across All Files</button>
      </div>
      <div class="queue-strip" id="queueStrip">
        <!-- File chips will be generated dynamically -->
      </div>
    </div>

    <!-- Active File Summary Bar -->
    <div id="summaryBar" class="card" style="display:none; padding:10px 14px; background:#131d2e; display:flex; justify-content:space-between; align-items:center;">
      <div style="display:flex; gap:20px; font-size:12px; color:var(--text-muted);">
        <div>File: <b id="summaryFilename" style="color:var(--text);">-</b></div>
        <div>Intervals: <b id="summaryTotalRecords" style="color:var(--text);">0</b></div>
        <div>Meters: <b id="summaryTotalMeters" style="color:var(--text);">0</b></div>
        <div>Output Name: <b id="summaryOutputName" style="color:#38bdf8;">-</b></div>
        <div>Status: <span id="summaryStatus" class="badge badge-info">Loaded</span></div>
      </div>
      <div style="display:flex; align-items:center; gap:12px;">
        <button class="btn-secondary" style="background:#0284c7; color:white; font-size:11px; padding:4px 10px;" onclick="downloadChangeTemplate()">📥 Download 9-Col Template (.xlsx)</button>
        <div style="font-size:11px; color:#38bdf8;">
          Tracking 4 Parameters: <b>Meter # | RTC Date/Time | Import kWh | Export kWh</b>
        </div>
      </div>
    </div>

    <!-- Filter & Action Controls Bar -->
    <div id="controlsBar" class="card controls-bar" style="display:none;">
      <div class="filter-group">
        <select id="meterFilter" style="width:170px;" onchange="applyFilter()">
          <option value="ALL">All Meters</option>
        </select>
        <select id="statusFilter" style="width:150px;" onchange="applyFilter()">
          <option value="ALL">All Records</option>
          <option value="MODIFIED">Modified Only</option>
          <option value="ORIGINAL">Original Only</option>
        </select>
        <input type="text" id="searchInput" placeholder="Search Meter / Time / kWh..." style="width:200px;" oninput="applyFilter()">
      </div>

      <div class="action-group">
        <button class="btn-secondary" style="background:#0284c7; color:white; font-weight:700;" onclick="downloadChangeTemplate()">📥 Download 9-Col Template (.xlsx)</button>
        <button class="btn-excel" onclick="openExcelUploadModal()">📊 Upload Excel Changes</button>
        <button class="btn-secondary" onclick="openPasteModal()">📋 Paste from Excel</button>
        <button class="btn-secondary" onclick="openBatchKwhModal()">⚡ Scale kWh</button>
        <button class="btn-secondary" onclick="openShiftTimeModal()">🕒 Shift Time</button>
        <button class="btn-secondary" onclick="openRenameModal()">🏷️ Rename Meter</button>
      </div>
    </div>

    <!-- Spreadsheet Table - BEFORE & AFTER DATA VIEWER -->
    <div class="table-container" id="tableContainer">
      <table>
        <thead>
          <tr>
            <th style="width: 50px;">#</th>
            <th style="width: 150px;">Meter Number</th>
            <th style="width: 190px;">Date & Time<br><small style="color:#38bdf8; font-weight:normal;">[0.0.1.0.0.255]</small></th>
            <th style="width: 150px;">Import kWh<br><small style="color:#64748b; font-weight:normal;">[BEFORE]</small></th>
            <th style="width: 160px;">Import kWh<br><small style="color:#22c55e; font-weight:normal;">[AFTER / EDITABLE]</small></th>
            <th style="width: 150px;">Export kWh<br><small style="color:#64748b; font-weight:normal;">[BEFORE]</small></th>
            <th style="width: 160px;">Export kWh<br><small style="color:#94a3b8; font-weight:normal;">[AFTER / EDITABLE]</small></th>
            <th style="width: 90px;">Status</th>
          </tr>
        </thead>
        <tbody id="tableBody">
          <tr><td colspan="8" style="text-align: center; color: var(--text-muted); padding: 40px;">Browse or select a JSON file above to view intervals</td></tr>
        </tbody>
      </table>
    </div>

    <!-- Pagination Footer -->
    <div id="paginationBar" class="card" style="display:none; justify-content: space-between; align-items: center; padding: 8px 14px;">
      <div style="color: var(--text-muted); font-size:12px;">
        Showing <span id="pageRange">0-0</span> of <span id="totalFilteredCount">0</span> records
      </div>
      <div style="display: flex; gap: 8px; align-items: center;">
        <button class="btn-secondary" onclick="changePage(-1)">◀ Prev</button>
        <span id="pageIndicator" style="font-weight: 600; font-size:12px;">Page 1</span>
        <button class="btn-secondary" onclick="changePage(1)">Next ▶</button>
      </div>
    </div>
  </div>

  <!-- Excel Upload Modal -->
  <div id="excelModal" class="modal-backdrop">
    <div class="modal" style="width:640px;">
      <div class="modal-header">
        <h3>📊 Apply Excel Changes (9-Column or 4-Column Format)</h3>
        <button class="close-btn" onclick="closeModal('excelModal')">✕</button>
      </div>

      <!-- Step 1: Download Template -->
      <div style="background:#13243a; border:1px solid #0369a1; border-radius:6px; padding:12px; display:flex; justify-content:space-between; align-items:center;">
        <div>
          <div style="font-weight:700; color:#38bdf8; font-size:12px;">📥 Step 1: Download 9-Column Change Template</div>
          <div style="color:#94a3b8; font-size:11px; margin-top:2px;">
            Pre-fills <b>JSON File Name</b> + 4 <b>[Before]</b> parameters from JSON. You only edit the 4 <b>[After]</b> columns!
          </div>
        </div>
        <button class="btn-secondary" onclick="downloadChangeTemplate()" style="padding:6px 12px; font-size:11px;">📥 Download (.xlsx)</button>
      </div>

      <!-- Step 2: Upload or Paste Changes -->
      <div style="font-weight:700; color:var(--text); font-size:12px; margin-top:2px;">
        📤 Step 2: Upload Edited Excel or Paste Rows
      </div>

      <div style="background: rgba(56, 189, 248, 0.08); border: 1px solid rgba(56, 189, 248, 0.25); border-radius: 6px; padding: 8px 12px; font-size: 11px; color: #bae6fd;">
        💡 <b>Supports 9 Columns or 4 Columns:</b>
        <br>• <b>9 Columns:</b> <code>JSON File Name | Meter Before | Time Before | Imp Before | Exp Before | Meter After | Time After | Imp After | Exp After</code>
        <br>• <b>4 Columns:</b> <code>Meter Number | Date & Time | Import kWh | Export kWh</code>
      </div>

      <div class="input-group">
        <label>Target Meter in JSON to Replace / Update (Optional for 4-Col):</label>
        <select id="excelTargetMeter">
          <option value="AUTO">⚡ Auto-Match (or replace active filtered meter)</option>
        </select>
      </div>
      
      <div style="display:flex; flex-direction:column; gap:10px;">
        <div style="display:flex; gap:8px;">
          <input type="text" id="excelFilePath" placeholder="Select your edited .xlsx or .csv file...">
          <button class="btn-secondary" onclick="triggerBrowserExcel()">📄 Browse Excel</button>
        </div>
        <div style="text-align:center; color:var(--text-muted); font-size:11px;">— OR PASTE ROWS DIRECTLY BELOW —</div>
        <textarea id="excelPasteArea" placeholder="Paste 9 columns or 4 columns here..."></textarea>
      </div>

      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:4px;">
        <button class="btn-secondary" onclick="closeModal('excelModal')">Cancel</button>
        <button class="btn-excel" onclick="applyExcelUpdates()">Apply Excel Changes</button>
      </div>
    </div>
  </div>

  <!-- Excel Single-Column Paste Modal (Legacy helper) -->
  <div id="pasteModal" class="modal-backdrop">
    <div class="modal">
      <div class="modal-header">
        <h3>📋 Paste Column from Excel (Import kWh)</h3>
        <button class="close-btn" onclick="closeModal('pasteModal')">✕</button>
      </div>
      <p style="color:var(--text-muted); font-size:12px;">Paste a single column of kWh numbers to sequentially fill down filtered intervals.</p>
      <textarea id="pasteTextarea" placeholder="0.05200&#10;0.05400&#10;0.05100&#10;..."></textarea>
      <div style="display:flex; justify-content:flex-end; gap:8px;">
        <button class="btn-secondary" onclick="closeModal('pasteModal')">Cancel</button>
        <button class="btn-accent" onclick="applySingleColumnPaste()">Apply Values</button>
      </div>
    </div>
  </div>

  <!-- Batch Edit kWh Modal -->
  <div id="batchKwhModal" class="modal-backdrop">
    <div class="modal">
      <div class="modal-header">
        <h3>⚡ Batch Scale / Set Import kWh</h3>
        <button class="close-btn" onclick="closeModal('batchKwhModal')">✕</button>
      </div>
      <div class="input-group">
        <label>Multiply by Factor (e.g. 1.05 for +5%):</label>
        <input type="text" id="batchMultiplyKwh" placeholder="e.g. 1.05">
      </div>
      <div style="text-align:center; color:var(--text-muted); font-size:11px;">— OR —</div>
      <div class="input-group">
        <label>Set Exact Fixed Value (kWh):</label>
        <input type="text" id="batchFixedKwh" placeholder="e.g. 0.04500">
      </div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:8px;">
        <button class="btn-secondary" onclick="closeModal('batchKwhModal')">Cancel</button>
        <button class="btn-accent" onclick="applyBatchKwh()">Apply to Filtered</button>
      </div>
    </div>
  </div>

  <!-- Shift Time Modal -->
  <div id="shiftTimeModal" class="modal-backdrop">
    <div class="modal">
      <div class="modal-header">
        <h3>🕒 Shift Interval Date & Time</h3>
        <button class="close-btn" onclick="closeModal('shiftTimeModal')">✕</button>
      </div>
      <div class="input-group">
        <label>Shift Minutes (+ or -):</label>
        <input type="number" id="shiftMinutes" value="0">
      </div>
      <div class="input-group">
        <label>Shift Days (+ or -):</label>
        <input type="number" id="shiftDays" value="0">
      </div>
      <label style="display:flex; align-items:center; gap:8px; font-size:12px; color:var(--text-muted);">
        <input type="checkbox" id="syncHesCheck" checked> Synchronize HES reception time by same offset
      </label>
      <div style="display:flex; justify-content:flex-end; gap:8px;">
        <button class="btn-secondary" onclick="closeModal('shiftTimeModal')">Cancel</button>
        <button class="btn-accent" onclick="applyShiftTime()">Shift Timestamps</button>
      </div>
    </div>
  </div>

  <!-- Rename Meter Modal -->
  <div id="renameModal" class="modal-backdrop">
    <div class="modal">
      <div class="modal-header">
        <h3>🏷️ Rename Meter ID</h3>
        <button class="close-btn" onclick="closeModal('renameModal')">✕</button>
      </div>
      <div class="input-group">
        <label>Target Meter Number:</label>
        <input type="text" id="renameOldMeter" readonly style="opacity:0.8;">
      </div>
      <div class="input-group">
        <label>New Meter Number:</label>
        <input type="text" id="renameNewMeter" placeholder="e.g. MTR_00001_NEW">
      </div>
      <div style="display:flex; justify-content:flex-end; gap:8px;">
        <button class="btn-secondary" onclick="closeModal('renameModal')">Cancel</button>
        <button class="btn-accent" onclick="applyRenameMeter()">Rename Across File</button>
      </div>
    </div>
  </div>

  <!-- Bulk Action Across All Files Modal -->
  <div id="bulkActionModal" class="modal-backdrop">
    <div class="modal">
      <div class="modal-header">
        <h3>🌐 Bulk Actions Across ALL Queued Files</h3>
        <button class="close-btn" onclick="closeModal('bulkActionModal')">✕</button>
      </div>
      <div class="input-group">
        <label>Action Type:</label>
        <select id="bulkActionType" onchange="toggleBulkActionInputs()">
          <option value="scale_kwh">Multiply Import kWh by Factor</option>
          <option value="set_kwh">Set Exact Fixed Import kWh</option>
          <option value="shift_time">Shift Timestamps (Days / Minutes)</option>
          <option value="rename_meter">Rename Meter ID</option>
        </select>
      </div>
      <div id="bulkInputContainer">
        <div class="input-group" id="bulkParam1Group">
          <label id="bulkParam1Label">Scaling Factor (e.g. 1.10 for +10%):</label>
          <input type="text" id="bulkParam1" placeholder="e.g. 1.05">
        </div>
      </div>
      <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:8px;">
        <button class="btn-secondary" onclick="closeModal('bulkActionModal')">Cancel</button>
        <button class="btn-warn" onclick="applyBulkActionAll()">Execute on All Files</button>
      </div>
    </div>
  </div>

  <!-- Toast -->
  <div id="toast" class="toast">Action completed successfully</div>

  <script>
    let currentPage = 1;
    let pageSize = 100;
    let currentMeters = [];
    let queue = [];
    let activeFile = null;

    let pendingExcelBase64 = null;
    let pendingExcelFilename = null;

    window.onload = async () => {
      try {
        const cfg = await apiCall('/api/config');
        if (cfg && cfg.source_dir) {
          document.getElementById('srcPath').value = cfg.source_dir;
        }
        if (cfg && cfg.dest_dir) {
          document.getElementById('destPath').value = cfg.dest_dir;
        }
      } catch (e) {
        console.warn('Config fetch warning:', e);
      }

      // Auto-scan and pre-load available JSON files
      await loadAvailableFilesList();

      // Close dropdown when clicking outside
      document.addEventListener('click', (e) => {
        if (!e.target.closest('#exportDropdown')) {
          document.getElementById('exportDropdown').classList.remove('show');
        }
      });

      // Drag and Drop support
      window.addEventListener('dragover', (e) => { e.preventDefault(); });
      window.addEventListener('drop', async (e) => {
        e.preventDefault();
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.toLowerCase().endsWith('.json'));
        if (files.length > 0) {
          showToast(`Uploading ${files.length} dropped JSON files...`);
          for (const f of files) {
            try {
              const content = await f.text();
              await apiCall('/api/upload_json_file', { filename: f.name, content: content });
            } catch (err) {
              console.error(err);
            }
          }
          await loadAvailableFilesList();
          await refreshQueue();
          if (queue.length > 0) {
            await switchFile(queue[0].path);
          }
        }
      });
    };

    function triggerBrowserFile() {
      document.getElementById('browserFileInput').click();
    }

    function triggerBrowserMulti() {
      document.getElementById('browserMultiInput').click();
    }

    function triggerBrowserFolder() {
      document.getElementById('browserFolderInput').click();
    }

    function triggerBrowserExcel() {
      document.getElementById('browserExcelInput').click();
    }

    async function handleBrowserFile(event) {
      const file = event.target.files[0];
      if (!file) return;
      showToast(`Loading ${file.name}...`);
      try {
        const content = await file.text();
        const res = await apiCall('/api/upload_json_file', {
          filename: file.name,
          content: content
        });
        if (res && res.file_name) {
          await loadAvailableFilesList();
          await refreshQueue();
          displayFile(res);
          showToast(`Loaded ${res.file_name} successfully!`);
        }
      } catch (err) {
        alert(`Error loading file: ${err.message}`);
      }
      event.target.value = '';
    }

    async function handleBrowserMulti(event) {
      const files = Array.from(event.target.files);
      if (files.length === 0) return;
      showToast(`Loading ${files.length} files...`);
      for (const f of files) {
        try {
          const content = await f.text();
          await apiCall('/api/upload_json_file', { filename: f.name, content: content });
        } catch (e) {
          console.error(e);
        }
      }
      await loadAvailableFilesList();
      await refreshQueue();
      if (queue.length > 0) {
        await switchFile(queue[0].path);
      }
      showToast(`Loaded ${files.length} files into queue!`);
      event.target.value = '';
    }

    async function handleBrowserFolder(event) {
      const files = Array.from(event.target.files).filter(f => f.name.toLowerCase().endsWith('.json'));
      if (files.length === 0) {
        alert('No .json files found in selected folder!');
        return;
      }
      showToast(`Found ${files.length} JSON files in folder. Loading...`);
      for (const f of files) {
        try {
          const content = await f.text();
          await apiCall('/api/upload_json_file', { filename: f.name, content: content });
        } catch (e) {
          console.error(e);
        }
      }
      await loadAvailableFilesList();
      await refreshQueue();
      if (queue.length > 0) {
        await switchFile(queue[0].path);
      }
      showToast(`Loaded ${files.length} JSON files from folder!`);
      event.target.value = '';
    }

    function handleBrowserExcel(event) {
      const file = event.target.files[0];
      if (!file) return;
      document.getElementById('excelFilePath').value = file.name;
      pendingExcelFilename = file.name;
      const reader = new FileReader();
      reader.onload = (e) => {
        pendingExcelBase64 = e.target.result.split(',')[1];
        showToast(`Ready to apply Excel changes from: ${file.name}`);
      };
      reader.readAsDataURL(file);
    }

    async function handleQuickSelect(filePath) {
      if (!filePath) return;
      showToast('Loading selected JSON file...');
      try {
        const summary = await apiCall('/api/switch_file', { file_path: filePath });
        await refreshQueue();
        displayFile(summary);
        showToast(`Loaded: ${summary.file_name}`);
      } catch (err) {
        alert(`Error loading file: ${err.message}`);
      }
    }

    async function loadAvailableFilesList(userTriggered = false) {
      try {
        const res = await apiCall('/api/available_files');
        const sel = document.getElementById('quickSelectDropdown');
        sel.innerHTML = '<option value="">-- Choose a JSON File to View & Edit Intervals --</option>';
        const badge = document.getElementById('availableCountBadge');
        if (res.files && res.files.length > 0) {
          badge.innerText = `${res.files.length} Files Detected`;
          res.files.forEach(f => {
            const opt = document.createElement('option');
            opt.value = f.path;
            opt.innerText = `${f.name} (${f.size_mb} MB) [${f.folder}]`;
            sel.appendChild(opt);
          });
          if (res.source_dir && !document.getElementById('srcPath').value) {
            document.getElementById('srcPath').value = res.source_dir;
          }
          if (res.dest_dir && !document.getElementById('destPath').value) {
            document.getElementById('destPath').value = res.dest_dir;
          }
          // Automatically load the first file on initial page load
          if (!activeFile && !userTriggered) {
            const firstPath = res.files[0].path;
            const summary = await apiCall('/api/switch_file', { file_path: firstPath });
            await refreshQueue();
            displayFile(summary);
            sel.value = firstPath;
          }
        } else {
          badge.innerText = '0 Files Detected';
        }
      } catch (err) {
        console.error('Failed to load available files:', err);
      }
    }

    async function loadManualPath(pathVal) {
      if (!pathVal) return;
      showToast(`Loading: ${pathVal}...`);
      try {
        const res = await apiCall('/api/scan_folder', { folder: pathVal });
        if (res.files && res.files.length > 0) {
          await loadMultipleFilesList(res.files.map(f => f.path));
        } else {
          await loadSingleFile(pathVal);
        }
      } catch (e) {
        alert(`Error loading path: ${e.message}`);
      }
    }

    function updateDestPath(val) {
      if (val) {
        showToast(`Destination folder set to: ${val}`);
      }
    }

    function showToast(msg) {
      const t = document.getElementById('toast');
      t.innerText = msg;
      t.style.display = 'block';
      setTimeout(() => { t.style.display = 'none'; }, 3500);
    }

    function toggleExportMenu() {
      document.getElementById('exportDropdown').classList.toggle('show');
    }

    async function apiCall(endpoint, data = null) {
      const opts = { method: data ? 'POST' : 'GET' };
      if (data) {
        opts.headers = { 'Content-Type': 'application/json' };
        opts.body = JSON.stringify(data);
      }
      const res = await fetch(endpoint, opts);
      return await res.json();
    }

    async function browseSingleFile() {
      const res = await apiCall('/api/native_browse_file');
      if (res.file_path) {
        document.getElementById('srcPath').value = res.file_path;
        await loadSingleFile(res.file_path);
      }
    }

    async function browseMultipleFiles() {
      const res = await apiCall('/api/native_browse_files');
      if (res.file_paths && res.file_paths.length > 0) {
        await loadMultipleFilesList(res.file_paths);
      }
    }

    async function browseFolder() {
      const res = await apiCall('/api/native_browse_folder');
      if (res.folder_path) {
        document.getElementById('srcPath').value = res.folder_path;
        if (res.files && res.files.length > 0) {
          await loadMultipleFilesList(res.files.map(f => f.path));
        }
      }
    }

    async function browseDestFolder() {
      const res = await apiCall('/api/native_browse_dest_folder');
      if (res.folder_path) {
        document.getElementById('destPath').value = res.folder_path;
        showToast('Destination folder updated');
      }
    }

    async function browseExcelFile() {
      const res = await apiCall('/api/native_browse_excel');
      if (res.file_path) {
        document.getElementById('excelFilePath').value = res.file_path;
      }
    }

    async function scanDefaultFolder() {
      const src = document.getElementById('srcPath').value;
      const res = await apiCall('/api/scan_folder', { folder: src });
      if (res.files && res.files.length > 0) {
        await loadMultipleFilesList(res.files.map(f => f.path));
      }
    }

    async function loadSingleFile(filePath) {
      const summary = await apiCall('/api/load_file', { file_path: filePath });
      await refreshQueue();
      displayFile(summary);
      showToast(`Loaded ${summary.file_name}`);
    }

    async function loadMultipleFilesList(filePaths) {
      const res = await apiCall('/api/load_multiple_files', { file_paths: filePaths });
      await refreshQueue();
      if (res.summaries && res.summaries.length > 0) {
        displayFile(res.summaries[0]);
        showToast(`Loaded ${res.summaries.length} files into queue`);
      }
    }

    async function refreshQueue() {
      const res = await apiCall('/api/get_queue');
      queue = res.queue || [];
      const strip = document.getElementById('queueStrip');
      strip.innerHTML = '';

      if (queue.length > 0) {
        document.getElementById('queueCard').style.display = 'flex';
        document.getElementById('queueCount').innerText = queue.length;

        queue.forEach(f => {
          const chip = document.createElement('div');
          chip.className = `file-chip ${f.is_active ? 'active' : ''}`;
          
          let dotClass = '';
          if (f.status === 'Modified') dotClass = 'mod';
          else if (f.status === 'Saved') dotClass = 'saved';

          chip.innerHTML = `
            <div class="status-dot ${dotClass}"></div>
            <span>${f.name}</span>
            <small style="color:var(--text-muted);">(${f.total_records})</small>
            ${f.modified_count > 0 ? `<span class="badge badge-warn" style="font-size:9px;">${f.modified_count}</span>` : ''}
          `;
          chip.onclick = () => switchFile(f.path);
          strip.appendChild(chip);
        });
      } else {
        document.getElementById('queueCard').style.display = 'none';
      }
    }

    async function switchFile(filePath) {
      const summary = await apiCall('/api/switch_file', { file_path: filePath });
      await refreshQueue();
      displayFile(summary);
    }

    function displayFile(summary) {
      activeFile = summary;
      document.getElementById('summaryFilename').innerText = summary.file_name;
      document.getElementById('summaryTotalRecords').innerText = summary.total_records.toLocaleString();
      document.getElementById('summaryTotalMeters').innerText = summary.unique_meters_count;
      
      const modName = summary.output_name || (summary.file_name.toLowerCase().endsWith('.json') ? summary.file_name.slice(0, -5) + '__Modified.json' : summary.file_name + '__Modified.json');
      document.getElementById('summaryOutputName').innerText = modName;

      const st = document.getElementById('summaryStatus');
      st.innerText = summary.status || 'Loaded';
      st.className = `badge ${summary.status === 'Modified' ? 'badge-warn' : (summary.status === 'Saved' ? 'badge-success' : 'badge-info')}`;

      document.getElementById('summaryBar').style.display = 'flex';
      document.getElementById('controlsBar').style.display = 'flex';
      document.getElementById('paginationBar').style.display = 'flex';

      // Show Excel template download and upload buttons now that JSON is selected
      const btnDl = document.getElementById('btnDownloadTemplate');
      if (btnDl) btnDl.style.display = 'inline-flex';
      const btnUl = document.getElementById('btnUploadExcel');
      if (btnUl) btnUl.style.display = 'inline-flex';

      // Populate meter filter
      currentMeters = summary.meters || [];
      const mf = document.getElementById('meterFilter');
      mf.innerHTML = `<option value="ALL">All Meters (${currentMeters.length})</option>`;
      currentMeters.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m;
        opt.innerText = m;
        mf.appendChild(opt);
      });

      currentPage = 1;
      updateUnsavedBadge();
      loadTableRows();
    }

    async function loadTableRows() {
      const meter = document.getElementById('meterFilter').value;
      const statusF = document.getElementById('statusFilter').value;
      const search = document.getElementById('searchInput').value;
      const res = await apiCall('/api/get_rows', {
        meter_filter: meter,
        status_filter: statusF,
        search: search,
        page: currentPage,
        page_size: pageSize
      });

      document.getElementById('totalFilteredCount').innerText = res.total.toLocaleString();
      const start = (currentPage - 1) * pageSize + 1;
      const end = Math.min(currentPage * pageSize, res.total);
      document.getElementById('pageRange').innerText = res.total > 0 ? `${start}-${end}` : '0-0';
      document.getElementById('pageIndicator').innerText = `Page ${currentPage} of ${Math.max(1, Math.ceil(res.total / pageSize))}`;

      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';

      if (res.rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; color: var(--text-muted); padding: 30px;">No matching records found</td></tr>';
        return;
      }

      // BEFORE & AFTER COMPARISON TABLE
      res.rows.forEach(r => {
        const tr = document.createElement('tr');
        const isMod = r.is_modified;
        const meterChanged = r.meter_number !== r.meter_number_before;
        const kwhImportChanged = r.kwh_import !== r.kwh_import_before;
        const kwhExportChanged = r.kwh_export !== r.kwh_export_before;

        tr.innerHTML = `
          <td style="color:var(--text-muted);">${r.index}</td>
          <td>
            <span class="editable-cell ${meterChanged ? 'modified-cell' : ''}" title="${meterChanged ? 'Original: ' + r.meter_number_before : ''}" onclick="startEdit(this, ${r.index}, 'meter_number')">${r.meter_number}</span>
          </td>
          <td>
            <span class="editable-cell" onclick="startEdit(this, ${r.index}, 'rtc_time')">${r.rtc_time}</span>
          </td>
          <td>
            <span class="before-val">${r.kwh_import_before}</span>
          </td>
          <td>
            <span class="editable-cell after-val ${kwhImportChanged ? 'modified-cell' : ''}" onclick="startEdit(this, ${r.index}, 'kwh_import')">${r.kwh_import}</span>
          </td>
          <td>
            <span class="before-val">${r.kwh_export_before || '0.00000'}</span>
          </td>
          <td>
            <span class="editable-cell after-val ${kwhExportChanged ? 'modified-cell' : ''}" onclick="startEdit(this, ${r.index}, 'kwh_export')">${r.kwh_export || '0.00000'}</span>
          </td>
          <td>
            ${isMod ? '<span class="badge badge-warn">Modified</span>' : '<span style="color:var(--text-muted); font-size:11px;">Original</span>'}
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function startEdit(cellSpan, recordIndex, field) {
      if (cellSpan.classList.contains('active-edit')) return;
      const curVal = cellSpan.innerText;
      cellSpan.classList.add('active-edit');
      cellSpan.innerHTML = `<input type="text" value="${curVal}">`;
      const input = cellSpan.querySelector('input');
      input.focus();
      input.select();

      async function finishEdit() {
        const newVal = input.value.trim();
        cellSpan.classList.remove('active-edit');
        cellSpan.innerText = newVal;
        if (newVal !== curVal) {
          await apiCall('/api/update_cell', {
            index: recordIndex,
            field: field,
            value: newVal
          });
          await refreshQueue();
          updateUnsavedBadge();
          loadTableRows();
        }
      }

      input.onblur = finishEdit;
      input.onkeydown = (e) => {
        if (e.key === 'Enter') finishEdit();
        if (e.key === 'Escape') {
          cellSpan.classList.remove('active-edit');
          cellSpan.innerText = curVal;
        }
      };
    }

    function updateUnsavedBadge() {
      const activeEntry = queue.find(f => f.is_active);
      const b = document.getElementById('unsavedBadge');
      if (activeEntry && activeEntry.modified_count > 0) {
        b.style.display = 'inline-block';
        b.innerText = `${activeEntry.modified_count} Unsaved Edits`;
      } else {
        b.style.display = 'none';
      }
    }

    function applyFilter() {
      currentPage = 1;
      loadTableRows();
    }

    function changePage(delta) {
      currentPage = Math.max(1, currentPage + delta);
      loadTableRows();
    }

    // Modal Helpers
    function openModal(id) { document.getElementById(id).classList.add('open'); }
    function closeModal(id) { document.getElementById(id).classList.remove('open'); }

    function downloadChangeTemplate() {
      if (!activeFile || !activeFile.file_path) {
        alert('Please select or load a source JSON file first!');
        return;
      }
      const meter = document.getElementById('meterFilter').value;
      window.location.href = `/api/download_template_excel?meter=${encodeURIComponent(meter)}`;
      showToast(`Downloading 9-column change template for ${activeFile.file_name}...`);
    }

    // Excel Batch Update Modal
    function openExcelUploadModal() {
      document.getElementById('excelFilePath').value = '';
      document.getElementById('excelPasteArea').value = '';

      const targetSel = document.getElementById('excelTargetMeter');
      targetSel.innerHTML = '';

      const curFilter = document.getElementById('meterFilter').value;
      const autoOpt = document.createElement('option');
      autoOpt.value = 'AUTO';
      autoOpt.innerText = '⚡ Auto-Match (or replace active filtered meter)';
      targetSel.appendChild(autoOpt);

      if (currentMeters && currentMeters.length > 0) {
        currentMeters.forEach(m => {
          const opt = document.createElement('option');
          opt.value = m;
          opt.innerText = `Replace Meter: ${m}`;
          if (curFilter !== 'ALL' && curFilter === m) {
            opt.selected = true;
          }
          targetSel.appendChild(opt);
        });
      }

      openModal('excelModal');
    }

    async function applyExcelUpdates() {
      const filePath = document.getElementById('excelFilePath').value.trim();
      const pasteText = document.getElementById('excelPasteArea').value.trim();
      const targetMeter = document.getElementById('excelTargetMeter').value;

      if (!filePath && !pasteText) {
        alert('Please browse an Excel file or paste rows below.');
        return;
      }

      let res;
      if (filePath) {
        if (pendingExcelBase64) {
          res = await apiCall('/api/upload_excel_file', {
            filename: pendingExcelFilename || filePath,
            base64_data: pendingExcelBase64,
            target_meter: targetMeter
          });
        } else {
          res = await apiCall('/api/upload_excel_file', { 
            file_path: filePath,
            target_meter: targetMeter
          });
        }
      } else {
        const lines = pasteText.split(String.fromCharCode(10)).map(l => l.replace(String.fromCharCode(13), '').trim()).filter(l => l.length > 0);
        const rows = lines.map(line => line.includes(String.fromCharCode(9)) ? line.split(String.fromCharCode(9)).map(c => c.trim()) : line.split(',').map(c => c.trim()));
        res = await apiCall('/api/paste_excel_rows', { 
          rows: rows,
          target_meter: targetMeter
        });
      }

      pendingExcelBase64 = null;
      pendingExcelFilename = null;

      closeModal('excelModal');
      await refreshQueue();
      updateUnsavedBadge();
      
      if (activeFile && activeFile.file_path) {
        const updatedSummary = await apiCall('/api/switch_file', { file_path: activeFile.file_path });
        displayFile(updatedSummary);
      } else {
        loadTableRows();
      }
      showToast(`Updated ${res.updated_intervals} intervals from Excel! (Unmatched: ${res.unmatched_rows})`);
    }

    function openPasteModal() {
      document.getElementById('pasteTextarea').value = '';
      openModal('pasteModal');
    }

    async function applySingleColumnPaste() {
      const raw = document.getElementById('pasteTextarea').value;
      const lines = raw.split(String.fromCharCode(10)).map(l => l.replace(String.fromCharCode(13), '').trim()).filter(l => l.length > 0);
      if (lines.length === 0) {
        alert('Please paste at least one value.');
        return;
      }
      const meter = document.getElementById('meterFilter').value;
      const res = await apiCall('/api/batch_paste', {
        start_index: (currentPage - 1) * pageSize,
        meter_filter: meter,
        values: lines
      });
      closeModal('pasteModal');
      await refreshQueue();
      updateUnsavedBadge();
      loadTableRows();
      showToast(`Pasted ${res.updated_count} values!`);
    }

    function openBatchKwhModal() {
      document.getElementById('batchFixedKwh').value = '';
      document.getElementById('batchMultiplyKwh').value = '';
      openModal('batchKwhModal');
    }

    async function applyBatchKwh() {
      const fixed = document.getElementById('batchFixedKwh').value.trim();
      const mult = document.getElementById('batchMultiplyKwh').value.trim();
      const meter = document.getElementById('meterFilter').value;

      if (!fixed && !mult) {
        alert('Enter a fixed value or multiply factor.');
        return;
      }

      const res = await apiCall('/api/batch_set_kwh', {
        meter_filter: meter,
        new_kwh: fixed ? fixed : null,
        multiply_factor: mult ? mult : null
      });

      closeModal('batchKwhModal');
      await refreshQueue();
      updateUnsavedBadge();
      loadTableRows();
      showToast(`Updated kWh for ${res.count} intervals!`);
    }

    function openShiftTimeModal() { openModal('shiftTimeModal'); }

    async function applyShiftTime() {
      const mins = parseInt(document.getElementById('shiftMinutes').value, 10) || 0;
      const days = parseInt(document.getElementById('shiftDays').value, 10) || 0;
      const syncHes = document.getElementById('syncHesCheck').checked;
      const meter = document.getElementById('meterFilter').value;

      const res = await apiCall('/api/batch_shift_time', {
        meter_filter: meter,
        shift_minutes: mins,
        shift_days: days,
        sync_hes: syncHes
      });

      closeModal('shiftTimeModal');
      await refreshQueue();
      updateUnsavedBadge();
      loadTableRows();
      showToast(`Shifted timestamps for ${res.count} records!`);
    }

    function openRenameModal() {
      const meter = document.getElementById('meterFilter').value;
      if (meter === 'ALL') {
        alert('Please filter to a specific meter first to rename it.');
        return;
      }
      document.getElementById('renameOldMeter').value = meter;
      document.getElementById('renameNewMeter').value = '';
      openModal('renameModal');
    }

    async function applyRenameMeter() {
      const oldM = document.getElementById('renameOldMeter').value;
      const newM = document.getElementById('renameNewMeter').value.trim();
      if (!newM) return;

      const res = await apiCall('/api/batch_rename_meter', {
        old_meter: oldM,
        new_meter: newM
      });

      closeModal('renameModal');
      await refreshQueue();
      updateUnsavedBadge();
      showToast(`Renamed ${oldM} to ${newM}!`);
      switchFile(activeFile.file_path);
    }

    function openBulkActionModal() {
      toggleBulkActionInputs();
      openModal('bulkActionModal');
    }

    function toggleBulkActionInputs() {
      const t = document.getElementById('bulkActionType').value;
      const c = document.getElementById('bulkInputContainer');
      if (t === 'scale_kwh') {
        c.innerHTML = `
          <div class="input-group">
            <label>Multiplier Factor (e.g. 1.10 for +10%):</label>
            <input type="text" id="bulkParam1" placeholder="e.g. 1.05">
          </div>`;
      } else if (t === 'set_kwh') {
        c.innerHTML = `
          <div class="input-group">
            <label>Fixed Import kWh Value:</label>
            <input type="text" id="bulkParam1" placeholder="e.g. 0.05000">
          </div>`;
      } else if (t === 'shift_time') {
        c.innerHTML = `
          <div class="input-group">
            <label>Shift Minutes (+ or -):</label>
            <input type="number" id="bulkParam1" value="0">
          </div>
          <div class="input-group" style="margin-top:6px;">
            <label>Shift Days (+ or -):</label>
            <input type="number" id="bulkParam2" value="0">
          </div>`;
      } else if (t === 'rename_meter') {
        c.innerHTML = `
          <div class="input-group">
            <label>Existing Meter ID:</label>
            <input type="text" id="bulkParam1" placeholder="e.g. MTR_00001">
          </div>
          <div class="input-group" style="margin-top:6px;">
            <label>New Meter ID:</label>
            <input type="text" id="bulkParam2" placeholder="e.g. MTR_00001_NEW">
          </div>`;
      }
    }

    async function applyBulkActionAll() {
      const act = document.getElementById('bulkActionType').value;
      let params = {};

      if (act === 'scale_kwh') {
        params.multiply_factor = parseFloat(document.getElementById('bulkParam1').value) || 1.0;
      } else if (act === 'set_kwh') {
        params.new_kwh = document.getElementById('bulkParam1').value.trim();
      } else if (act === 'shift_time') {
        params.shift_minutes = parseInt(document.getElementById('bulkParam1').value, 10) || 0;
        params.shift_days = parseInt(document.getElementById('bulkParam2').value, 10) || 0;
        params.sync_hes = true;
      } else if (act === 'rename_meter') {
        params.old_meter = document.getElementById('bulkParam1').value.trim();
        params.new_meter = document.getElementById('bulkParam2').value.trim();
      }

      const res = await apiCall('/api/bulk_action_all', {
        action_type: act,
        params: params
      });

      closeModal('bulkActionModal');
      await refreshQueue();
      updateUnsavedBadge();
      if (activeFile) {
        switchFile(activeFile.file_path);
      }
      showToast(`Bulk action executed across all files: ${res.total_affected} records modified!`);
    }

    // Export Data Handler
    function exportData(exportType) {
      document.getElementById('exportDropdown').classList.remove('show');
      window.location.href = `/api/export_csv?type=${exportType}`;
      showToast(`Exporting ${exportType} CSV data...`);
    }

    async function saveActiveFile() {
      const dest = document.getElementById('destPath').value;
      if (!dest) {
        alert('Please specify a destination folder.');
        return;
      }
      const res = await apiCall('/api/save_active', { destination_folder: dest });
      await refreshQueue();
      updateUnsavedBadge();
      loadTableRows();
      showToast(`Saved to ${res.filename}!`);
      alert(`Success! Updated file saved with exact schema to:\\n${res.saved_path}`);
    }

    async function saveAllFiles() {
      const dest = document.getElementById('destPath').value;
      if (!dest) {
        alert('Please specify a destination folder.');
        return;
      }
      const res = await apiCall('/api/save_all', { destination_folder: dest });
      await refreshQueue();
      updateUnsavedBadge();
      loadTableRows();
      showToast(`Saved all ${res.saved_files.length} modified files!`);
      alert(`Success! Saved ${res.saved_files.length} files with '_Modified' suffix to:\\n${dest}`);
    }
  </script>
</body>
</html>
"""

def open_native_file_dialog(filetypes=None, title="Select File", multiple=False):
    """Opens native Windows Explorer File Open dialog."""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    types = filetypes or [("JSON Files", "*.json"), ("All Files", "*.*")]
    if multiple:
        files = filedialog.askopenfilenames(title=title, filetypes=types)
        root.destroy()
        return list(files)
    else:
        file = filedialog.askopenfilename(title=title, filetypes=types)
        root.destroy()
        return file

def open_native_folder_dialog(title="Select Folder"):
    """Opens native Windows Explorer Directory picker."""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder = filedialog.askdirectory(title=title)
    root.destroy()
    return folder

class JsonEditorRequestHandler(BaseHTTPRequestHandler):
    def _send_json(self, data, status=200):
        body = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_body_json(self):
        length = int(self.headers.get("Content-Length", 0))
        if length > 0:
            raw = self.rfile.read(length).decode("utf-8")
            return json.loads(raw)
        return {}

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = HTML_PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        elif parsed.path == "/api/config":
            self._send_json({
                "source_dir": source_dir,
                "dest_dir": dest_dir
            })

        elif parsed.path == "/api/available_files":
            files_list = []
            if os.path.exists(source_dir):
                for fpath in sorted(glob.glob(os.path.join(source_dir, "*.json"))):
                    if "output" not in fpath.split(os.sep):
                        size_mb = f"{os.path.getsize(fpath) / (1024 * 1024):.2f}"
                        files_list.append({
                            "name": os.path.basename(fpath),
                            "path": os.path.abspath(fpath).replace("\\", "/"),
                            "size_mb": size_mb,
                            "folder": "Project Root"
                        })
                for item in sorted(os.listdir(source_dir)):
                    sub = os.path.join(source_dir, item)
                    if os.path.isdir(sub) and item not in ["output", "build", "dist", "runtime", ".git", ".gemini", "scratch", "__pycache__"]:
                        for fpath in sorted(glob.glob(os.path.join(sub, "*.json"))):
                            size_mb = f"{os.path.getsize(fpath) / (1024 * 1024):.2f}"
                            files_list.append({
                                "name": f"{item}/{os.path.basename(fpath)}",
                                "path": os.path.abspath(fpath).replace("\\", "/"),
                                "size_mb": size_mb,
                                "folder": item
                            })
            self._send_json({
                "files": files_list,
                "source_dir": source_dir,
                "dest_dir": dest_dir
            })

        elif parsed.path == "/api/get_queue":
            queue_data = editor.get_queue_status()
            self._send_json({"queue": queue_data})

        elif parsed.path == "/api/export_csv":
            query = parse_qs(parsed.query)
            exp_type = query.get("type", ["comparison"])[0]
            try:
                csv_str = editor.export_data_csv(export_type=exp_type)
                base = "data"
                if editor.active_file_path:
                    base = os.path.splitext(os.path.basename(editor.active_file_path))[0]
                filename = f"{base}_{exp_type}.csv"
                
                body = csv_str.encode("utf-8-sig")
                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)

        elif parsed.path == "/api/download_template_excel":
            query = parse_qs(parsed.query)
            m_filter = query.get("meter", ["ALL"])[0]
            try:
                xlsx_bytes = editor.generate_template_excel(meter_filter=m_filter)
                base = "meter_change_template"
                if editor.active_file_path:
                    base = os.path.splitext(os.path.basename(editor.active_file_path))[0]
                if m_filter and m_filter != "ALL":
                    filename = f"{base}_{m_filter}_Change_Template.xlsx"
                else:
                    filename = f"{base}_Change_Template.xlsx"

                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(xlsx_bytes)))
                self.end_headers()
                self.wfile.write(xlsx_bytes)
            except Exception as e:
                self._send_json({"error": str(e)}, status=500)

        elif parsed.path == "/api/native_browse_file":
            selected = open_native_file_dialog(title="Select Smart Meter JSON File")
            self._send_json({"file_path": selected})

        elif parsed.path == "/api/native_browse_files":
            selected = open_native_file_dialog(title="Select Smart Meter JSON Files", multiple=True)
            self._send_json({"file_paths": selected})

        elif parsed.path == "/api/native_browse_folder":
            folder = open_native_folder_dialog("Select Source Folder")
            files_list = []
            if folder and os.path.exists(folder):
                for fpath in glob.glob(os.path.join(folder, "*.json")):
                    if "output" not in fpath.split(os.sep):
                        files_list.append({"name": os.path.basename(fpath), "path": fpath})
            self._send_json({"folder_path": folder, "files": files_list})

        elif parsed.path == "/api/native_browse_dest_folder":
            folder = open_native_folder_dialog("Select Destination Folder")
            self._send_json({"folder_path": folder})

        elif parsed.path == "/api/native_browse_excel":
            types = [("Excel & CSV Files", "*.xlsx;*.xls;*.csv"), ("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
            selected = open_native_file_dialog(filetypes=types, title="Select Excel or CSV File")
            self._send_json({"file_path": selected})

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        global source_dir, dest_dir
        parsed = urlparse(self.path)
        payload = self._read_body_json()

        if parsed.path == "/api/scan_folder":
            folder = payload.get("folder", source_dir)
            source_dir = folder
            files_list = []
            if os.path.exists(folder):
                for fpath in glob.glob(os.path.join(folder, "*.json")):
                    if "output" not in fpath.split(os.sep):
                        size_mb = f"{os.path.getsize(fpath) / (1024 * 1024):.2f}"
                        files_list.append({"name": os.path.basename(fpath), "path": fpath, "size_mb": size_mb})
            self._send_json({"files": files_list})

        elif parsed.path == "/api/load_file":
            fpath = payload.get("file_path")
            summary = editor.load_file(fpath)
            self._send_json(summary)

        elif parsed.path == "/api/load_multiple_files":
            fpaths = payload.get("file_paths", [])
            summaries = editor.load_multiple_files(fpaths)
            self._send_json({"summaries": summaries})

        elif parsed.path == "/api/switch_file":
            fpath = payload.get("file_path")
            summary = editor.switch_active_file(fpath)
            self._send_json(summary)

        elif parsed.path == "/api/get_rows":
            m_filter = payload.get("meter_filter")
            st_filter = payload.get("status_filter", "ALL")
            q = payload.get("search", "")
            p = payload.get("page", 1)
            psize = payload.get("page_size", 100)
            rows_data = editor.get_rows(meter_filter=m_filter, status_filter=st_filter, search_query=q, page=p, page_size=psize)
            self._send_json(rows_data)

        elif parsed.path == "/api/update_cell":
            idx = payload.get("index")
            field = payload.get("field")
            val = payload.get("value")
            editor.update_single_cell(idx, field, val)
            self._send_json({"success": True})

        elif parsed.path == "/api/upload_json_file":
            filename = payload.get("filename", "uploaded.json")
            content = payload.get("content", "")
            safe_name = os.path.basename(filename)
            save_path = os.path.join(source_dir, safe_name)
            with open(save_path, "w", encoding="utf-8") as fp:
                fp.write(content)
            summary = editor.load_file(save_path)
            self._send_json(summary)

        elif parsed.path == "/api/upload_excel_file":
            target_m = payload.get("target_meter")
            if "base64_data" in payload and payload["base64_data"]:
                fname = payload.get("filename", "changes.xlsx")
                raw = base64.b64decode(payload["base64_data"])
                temp_path = os.path.join(source_dir, f"_temp_{os.path.basename(fname)}")
                with open(temp_path, "wb") as fp:
                    fp.write(raw)
                try:
                    res = editor.apply_excel_file(temp_path, target_meter=target_m)
                    self._send_json(res)
                finally:
                    if os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except Exception:
                            pass
            else:
                fpath = payload.get("file_path")
                res = editor.apply_excel_file(fpath, target_meter=target_m)
                self._send_json(res)

        elif parsed.path == "/api/paste_excel_rows":
            rows = payload.get("rows", [])
            target_m = payload.get("target_meter")
            res = editor.apply_excel_rows(rows, target_meter=target_m)
            self._send_json(res)

        elif parsed.path == "/api/batch_paste":
            start_idx = payload.get("start_index", 0)
            m_filter = payload.get("meter_filter")
            vals = payload.get("values", [])
            count = editor.batch_paste_kwh(start_idx, m_filter, vals)
            self._send_json({"updated_count": count})

        elif parsed.path == "/api/batch_set_kwh":
            m_filter = payload.get("meter_filter")
            fixed = payload.get("new_kwh")
            mult = payload.get("multiply_factor")
            count = editor.batch_set_kwh(m_filter, new_kwh=fixed, multiply_factor=mult)
            self._send_json({"count": count})

        elif parsed.path == "/api/batch_shift_time":
            m_filter = payload.get("meter_filter")
            mins = payload.get("shift_minutes", 0)
            days = payload.get("shift_days", 0)
            sync_hes = payload.get("sync_hes", True)
            count = editor.batch_shift_time(m_filter, shift_minutes=mins, shift_days=days, sync_hes=sync_hes)
            self._send_json({"count": count})

        elif parsed.path == "/api/batch_rename_meter":
            old_m = payload.get("old_meter")
            new_m = payload.get("new_meter")
            count = editor.batch_rename_meter(old_m, new_m)
            self._send_json({"count": count})

        elif parsed.path == "/api/bulk_action_all":
            act = payload.get("action_type")
            params = payload.get("params", {})
            res = editor.apply_bulk_action_all_files(act, params)
            self._send_json(res)

        elif parsed.path == "/api/save_active":
            destination = payload.get("destination_folder", dest_dir)
            dest_dir = destination
            result = editor.save_file(destination)
            self._send_json(result)

        elif parsed.path == "/api/save_all":
            destination = payload.get("destination_folder", dest_dir)
            dest_dir = destination
            saved_list = editor.save_all_files(destination)
            self._send_json({"saved_files": saved_list})

        else:
            self.send_response(404)
            self.end_headers()

def run_server(port=8080, open_browser=True):
    actual_port = port
    server = None
    for p in range(port, port + 20):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", p), JsonEditorRequestHandler)
            server.daemon_threads = True
            actual_port = p
            break
        except OSError:
            continue

    if not server:
        print("Could not find an available port to start server.")
        return

    url = f"http://127.0.0.1:{actual_port}"
    print(f"\n========================================================")
    print(f"[*] Smart Meter JSON Editor is running at: {url}")
    print(f"========================================================\n")
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        server.server_close()

def run_cli(args):
    """Headless CLI mode for testing, automation, and verification."""
    ed = JsonMeterEditor()
    s = ed.load_file(args.file)
    print(f"Loaded: {s['file_name']} ({s['total_records']} intervals, {s['unique_meters_count']} meters)")

    if args.excel:
        res = ed.apply_excel_file(args.excel)
        print(f"Applied Excel changes: {res}")

    if args.edit_index is not None and args.field and args.value is not None:
        ed.update_single_cell(args.edit_index, args.field, args.value)
        print(f"Updated record #{args.edit_index}: {args.field} -> {args.value}")

    if args.batch_scale is not None:
        c = ed.batch_set_kwh("ALL", multiply_factor=args.batch_scale)
        print(f"Scaled kWh by {args.batch_scale} across {c} records")

    if args.export_csv:
        csv_str = ed.export_data_csv(args.export_csv)
        out_csv = os.path.join(args.out or "output", f"{s['file_name']}_{args.export_csv}.csv")
        os.makedirs(os.path.dirname(out_csv), exist_ok=True)
        with open(out_csv, "w", encoding="utf-8-sig") as fp:
            fp.write(csv_str)
        print(f"Exported {args.export_csv} CSV to: {out_csv}")

    if args.out:
        res = ed.save_file(args.out)
        print(f"Saved to: {res['saved_path']} ({res['size_bytes']} bytes)")

def main():
    parser = argparse.ArgumentParser(description="Smart Meter JSON Editor")
    parser.add_argument("--file", "-f", help="Direct path to input JSON file for headless operation")
    parser.add_argument("--excel", "-e", help="Excel/CSV file with 4 columns to apply updates")
    parser.add_argument("--edit-index", type=int, help="Record index to edit")
    parser.add_argument("--field", choices=["meter_number", "rtc_time", "kwh_import", "kwh_export"], help="Field to edit")
    parser.add_argument("--value", help="New value for field")
    parser.add_argument("--batch-scale", type=float, help="Batch multiply kWh by factor")
    parser.add_argument("--export-csv", choices=["source", "output", "comparison"], help="Export data as CSV")
    parser.add_argument("--out", "-o", help="Output destination folder")
    parser.add_argument("--port", "-p", type=int, default=8080, help="Port for web server")
    parser.add_argument("--no-browser", action="store_true", help="Do not automatically open browser")

    args = parser.parse_args()

    if args.file:
        run_cli(args)
    else:
        run_server(port=args.port, open_browser=not args.no_browser)

if __name__ == "__main__":
    main()
