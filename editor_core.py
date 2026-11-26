"""
Smart Meter JSON Editor Core Engine
Provides exact structure-preserving reads, Excel-driven batch updates, Before/After comparison, and exact compact serialization.
Focuses strictly on the 4 critical parameters:
  1. Meter Number (MeterNumber)
  2. Date & Time (OBIS 0.0.1.0.0.255)
  3. Import kWh (OBIS 1.0.1.29.0.255)
  4. Export kWh (OBIS 1.0.2.29.0.255)
Preserves 100% of all other ~50+ electrical parameters, attribute IDs ("AID":"2"), and schema formatting without data loss.
"""

import os
import io
import csv
import json
from collections import OrderedDict
from datetime import datetime, timedelta

RTC_OBIS = "0.0.1.0.0.255"
KWH_IMPORT_OBIS = "1.0.1.29.0.255"
KWH_EXPORT_OBIS = "1.0.2.29.0.255"

def parse_flexible_datetime(val):
    """Normalizes Excel datetimes or strings to DD/MM/YYYY HH:MM:SS."""
    if val is None or str(val).strip() == "":
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y %H:%M:%S")
    
    s = str(val).strip()
    formats = [
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        except ValueError:
            pass
    return s

class JsonMeterEditor:
    def __init__(self):
        self.files_queue = OrderedDict()
        self.active_file_path = None

    @property
    def index_lookup(self):
        """Convenience property for active file's index lookup."""
        if self.active_file_path and self.active_file_path in self.files_queue:
            return self.files_queue[self.active_file_path]["index_lookup"]
        return []

    @property
    def raw_data(self):
        """Convenience property for active file's raw JSON data."""
        if self.active_file_path and self.active_file_path in self.files_queue:
            return self.files_queue[self.active_file_path]["raw_data"]
        return None

    def load_file(self, file_path):
        """Loads a single JSON file into the queue preserving key order and structure."""
        file_path = os.path.abspath(file_path)
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        with open(file_path, "r", encoding="utf-8") as fp:
            raw_data = json.load(fp, object_pairs_hook=OrderedDict)

        index_lookup = self._build_index(raw_data)
        meters = sorted(list(set(r["meter_number"] for r in index_lookup)))

        summary = {
            "file_name": os.path.basename(file_path),
            "file_path": file_path,
            "output_name": self.get_output_filename(file_path),
            "total_records": len(index_lookup),
            "unique_meters_count": len(meters),
            "meters": meters,
            "modified_count": 0,
            "status": "Loaded"
        }

        lookup_map, time_map = self._build_lookup_map(index_lookup)
        self.files_queue[file_path] = {
            "raw_data": raw_data,
            "index_lookup": index_lookup,
            "lookup_map": lookup_map,
            "time_map": time_map,
            "modified_count": 0,
            "status": "Loaded",
            "summary": summary
        }
        self.active_file_path = file_path
        return summary

    def load_multiple_files(self, file_paths):
        """Loads multiple JSON files into the queue."""
        summaries = []
        for p in file_paths:
            if os.path.exists(p) and p.lower().endswith(".json"):
                try:
                    s = self.load_file(p)
                    summaries.append(s)
                except Exception as e:
                    print(f"Error loading {p}: {e}")
        if file_paths and os.path.abspath(file_paths[0]) in self.files_queue:
            self.active_file_path = os.path.abspath(file_paths[0])
        return summaries

    def switch_active_file(self, file_path):
        """Switches the active file being viewed/edited."""
        abs_p = os.path.abspath(file_path)
        if abs_p in self.files_queue:
            self.active_file_path = abs_p
            return self.files_queue[abs_p]["summary"]
        return self.load_file(abs_p)

    def get_queue_status(self):
        """Returns the list of queued files and their statuses."""
        result = []
        for p, data in self.files_queue.items():
            result.append({
                "path": p,
                "name": os.path.basename(p),
                "output_name": self.get_output_filename(p),
                "total_records": len(data["index_lookup"]),
                "meters_count": data["summary"]["unique_meters_count"],
                "modified_count": data["modified_count"],
                "status": data["status"],
                "is_active": (p == self.active_file_path)
            })
        return result

    def _build_index(self, raw_data):
        """Builds a fast-access index storing Before and After states for the 4 critical parameters."""
        index_lookup = []
        for idx, record in enumerate(raw_data):
            meter_num = record.get("MeterNumber", "")
            hes_time = record.get("HesTime", "")
            mdata = record.get("MeterData", [])

            rtc_val = ""
            kwh_import_val = ""
            kwh_export_val = ""

            for item in mdata:
                oc = item.get("OC")
                if oc == RTC_OBIS:
                    rtc_val = item.get("V", "")
                elif oc == KWH_IMPORT_OBIS:
                    kwh_import_val = item.get("V", "")
                elif oc == KWH_EXPORT_OBIS:
                    kwh_export_val = item.get("V", "")

            index_lookup.append({
                "index": idx,
                "meter_number": meter_num,
                "meter_number_before": meter_num,
                "hes_time": hes_time,
                "hes_time_before": hes_time,
                "rtc_time": rtc_val,
                "rtc_time_before": rtc_val,
                "kwh_import": kwh_import_val,
                "kwh_import_before": kwh_import_val,
                "kwh_export": kwh_export_val,
                "kwh_export_before": kwh_export_val,
                "is_modified": False
            })
        return index_lookup

    def _build_lookup_map(self, index_lookup):
        """Builds (meter_number, rtc_time) -> index map and rtc_time -> [indices] for flexible matching."""
        mapping = {}
        time_map = {}
        for r in index_lookup:
            norm_time = parse_flexible_datetime(r["rtc_time"])
            raw_time = r["rtc_time"].strip()
            idx = r["index"]
            m_upper = r["meter_number"].strip().upper()

            mapping[(m_upper, norm_time)] = idx
            mapping[(m_upper, raw_time)] = idx

            time_map.setdefault(norm_time, []).append(idx)
            if raw_time != norm_time:
                time_map.setdefault(raw_time, []).append(idx)

        return mapping, time_map

    def get_active_summary(self):
        """Returns summary of the currently active file."""
        if not self.active_file_path or self.active_file_path not in self.files_queue:
            return {}
        return self.files_queue[self.active_file_path]["summary"]

    def get_rows(self, file_path=None, meter_filter=None, search_query="", status_filter="ALL", page=1, page_size=100):
        """Retrieves paginated and filtered rows with Before and After comparison values."""
        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            return {"total": 0, "page": page, "page_size": page_size, "rows": []}

        index_lookup = self.files_queue[target_path]["index_lookup"]
        filtered = index_lookup

        if meter_filter and meter_filter != "ALL":
            filtered = [r for r in filtered if r["meter_number"] == meter_filter]

        if status_filter == "MODIFIED":
            filtered = [r for r in filtered if r["is_modified"]]
        elif status_filter == "ORIGINAL":
            filtered = [r for r in filtered if not r["is_modified"]]

        if search_query:
            q = search_query.strip().lower()
            filtered = [
                r for r in filtered 
                if q in r["meter_number"].lower() 
                or q in r["rtc_time"].lower() 
                or q in r["kwh_import"].lower() 
                or q in r["kwh_export"].lower()
            ]

        total_filtered = len(filtered)
        start = (page - 1) * page_size
        end = start + page_size
        paged_rows = filtered[start:end]

        return {
            "total": total_filtered,
            "page": page,
            "page_size": page_size,
            "rows": paged_rows
        }

    def update_single_cell(self, record_index, field, new_value, file_path=None):
        """
        Updates a single focused field in the in-memory JSON structure.
        All other fields and characters remain completely untouched.
        """
        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            raise KeyError("No active file loaded")

        file_entry = self.files_queue[target_path]
        raw_data = file_entry["raw_data"]
        index_lookup = file_entry["index_lookup"]

        if record_index < 0 or record_index >= len(raw_data):
            raise IndexError("Record index out of range")

        record = raw_data[record_index]
        new_val_str = str(new_value).strip()

        if field == "meter_number":
            record["MeterNumber"] = new_val_str
            index_lookup[record_index]["meter_number"] = new_val_str

        elif field == "rtc_time":
            found = False
            for item in record.get("MeterData", []):
                if item.get("OC") == RTC_OBIS:
                    item["V"] = new_val_str
                    found = True
                    break
            if not found:
                record.setdefault("MeterData", []).append(OrderedDict([("OC", RTC_OBIS), ("V", new_val_str), ("AID", "2")]))
            index_lookup[record_index]["rtc_time"] = new_val_str

        elif field in ["kwh_import", "kwh_export"]:
            try:
                f_val = float(new_val_str)
                new_val_str = f"{f_val:.5f}"
            except ValueError:
                pass

            target_obis = KWH_IMPORT_OBIS if field == "kwh_import" else KWH_EXPORT_OBIS
            found = False
            for item in record.get("MeterData", []):
                if item.get("OC") == target_obis:
                    item["V"] = new_val_str
                    found = True
                    break
            if not found:
                record.setdefault("MeterData", []).append(OrderedDict([("OC", target_obis), ("V", new_val_str), ("AID", "2")]))
            index_lookup[record_index][field] = new_val_str

        # Update modification state
        is_mod = (
            index_lookup[record_index]["meter_number"] != index_lookup[record_index]["meter_number_before"] or
            index_lookup[record_index]["rtc_time"] != index_lookup[record_index]["rtc_time_before"] or
            index_lookup[record_index]["kwh_import"] != index_lookup[record_index]["kwh_import_before"] or
            index_lookup[record_index]["kwh_export"] != index_lookup[record_index]["kwh_export_before"]
        )
        index_lookup[record_index]["is_modified"] = is_mod

        # Recompute total modified count
        mod_count = sum(1 for r in index_lookup if r["is_modified"])
        file_entry["modified_count"] = mod_count
        file_entry["status"] = "Modified" if mod_count > 0 else "Loaded"
        file_entry["summary"]["modified_count"] = mod_count
        file_entry["summary"]["status"] = file_entry["status"]
        return True

    def generate_template_excel(self, file_path=None, meter_filter=None):
        """
        Generates a 9-column Excel (.xlsx) change template containing the data read from the JSON.
        Columns:
          1. JSON File Name
          2. Meter Number [Before]
          3. Date & Time [Before]
          4. Import kWh [Before]
          5. Export kWh [Before]
          6. Meter Number [After]  (pre-filled, user edits this!)
          7. Date & Time [After]   (pre-filled, user edits this!)
          8. Import kWh [After]    (pre-filled, user edits this!)
          9. Export kWh [After]    (pre-filled, user edits this!)
        Returns bytes of the .xlsx file.
        """
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from io import BytesIO

        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            raise KeyError("No active file loaded to generate template")

        file_name = os.path.basename(target_path)
        index_lookup = self.files_queue[target_path]["index_lookup"]

        filtered = index_lookup
        if meter_filter and meter_filter != "ALL":
            filtered = [r for r in filtered if r["meter_number_before"] == meter_filter or r["meter_number"] == meter_filter]

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Change_Template"

        headers = [
            "JSON File Name",
            "Meter Number [Before]",
            "Date & Time [Before]",
            "Import kWh [Before]",
            "Export kWh [Before]",
            "Meter Number [After]",
            "Date & Time [After]",
            "Import kWh [After]",
            "Export kWh [After]"
        ]
        ws.append(headers)

        # Style header row
        fill_file = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
        fill_before = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        fill_after = PatternFill(start_color="14532D", end_color="14532D", fill_type="solid")
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col_idx in range(1, 10):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = font_header
            cell.alignment = align_header
            if col_idx == 1:
                cell.fill = fill_file
            elif col_idx in [2, 3, 4, 5]:
                cell.fill = fill_before
            else:
                cell.fill = fill_after

        # Append data rows
        font_after_data = Font(name="Calibri", size=11, bold=True, color="166534")
        fill_after_cell = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")

        for r_idx, r in enumerate(filtered, start=2):
            row_data = [
                file_name,
                r["meter_number_before"],
                r["rtc_time_before"],
                r["kwh_import_before"],
                r["kwh_export_before"],
                r["meter_number"],
                r["rtc_time"],
                r["kwh_import"],
                r["kwh_export"]
            ]
            ws.append(row_data)

            # Style the editable 'After' cells
            for col_idx in range(6, 10):
                c = ws.cell(row=r_idx, column=col_idx)
                c.fill = fill_after_cell
                c.font = font_after_data

        # Set column widths
        col_widths = {
            1: 36, 2: 22, 3: 22, 4: 20, 5: 20, 6: 22, 7: 22, 8: 20, 9: 20
        }
        for col_idx, width in col_widths.items():
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        ws.freeze_panes = "A2"

        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def apply_excel_file(self, excel_path, target_file_path=None, target_meter=None):
        """
        Parses an Excel (.xlsx / .xls / .csv) file with 4 or 9 columns:
          9-Column Template:
            [JSON File Name, Meter Before, Time Before, Imp Before, Exp Before, Meter After, Time After, Imp After, Exp After]
          4-Column Format:
            [Meter Number, Date & Time, Import kWh, Export kWh]
        Matches records in JSON and updates Meter Number, Import kWh, and Export kWh.
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        rows = []
        ext = os.path.splitext(excel_path)[1].lower()

        if ext == ".csv":
            with open(excel_path, "r", encoding="utf-8-sig", errors="replace") as f:
                reader = csv.reader(f)
                for r in reader:
                    if r and any(cell.strip() for cell in r):
                        rows.append(r)
        else:
            import openpyxl
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            sheet = wb.active
            for row in sheet.iter_rows(values_only=True):
                if row and any(cell is not None and str(cell).strip() for cell in row):
                    rows.append(list(row))

        return self.apply_excel_rows(rows, target_file_path=target_file_path, target_meter=target_meter)

    def apply_excel_rows(self, rows, target_file_path=None, target_meter=None):
        """
        Applies a list of rows (from Excel or clipboard paste) to the loaded JSON file(s).
        Supports both:
          - 9-Column Template:
            [JSON File Name, Meter Before, Time Before, Imp Before, Exp Before, Meter After, Time After, Imp After, Exp After]
          - 4-Column Format:
            [Meter Number, Date & Time, Import kWh, Export kWh]
        """
        if not rows:
            return {"updated_intervals": 0, "total_excel_rows": 0, "unmatched_rows": 0}

        # Check if row 0 is header
        start_row = 0
        first_row_str = " ".join(str(c or "").lower() for c in rows[0])
        if any(keyword in first_row_str for keyword in ["meter", "date", "time", "kwh", "import", "export", "obis", "json", "file", "before", "after"]):
            start_row = 1

        target_files = [target_file_path] if target_file_path else (
            [self.active_file_path] if self.active_file_path else list(self.files_queue.keys())
        )

        total_updated = 0
        unmatched_count = 0
        processed_rows = 0

        for r_idx in range(start_row, len(rows)):
            row = rows[r_idx]
            if len(row) < 3:
                continue

            # Check if this row is 9-column (or 8+ column) template format
            if len(row) >= 8:
                json_name = str(row[0] or "").strip()
                before_meter = str(row[1] or "").strip()
                before_time_raw = row[2]
                before_time = parse_flexible_datetime(before_time_raw)
                raw_before_time = str(before_time_raw or "").strip()

                after_meter = str(row[5] or "").strip() if len(row) > 5 and row[5] is not None else ""
                after_time_raw = row[6] if len(row) > 6 and row[6] is not None else ""
                after_time = parse_flexible_datetime(after_time_raw) if after_time_raw else ""
                after_imp = str(row[7] or "").strip() if len(row) > 7 and row[7] is not None else ""
                after_exp = str(row[8] or "").strip() if len(row) > 8 and row[8] is not None else ""

                if not before_time and not raw_before_time:
                    continue

                processed_rows += 1
                matched = False

                row_targets = target_files
                if json_name:
                    matching_fpaths = [p for p in self.files_queue if os.path.basename(p).lower() == json_name.lower()]
                    if matching_fpaths:
                        row_targets = matching_fpaths

                for fpath in row_targets:
                    file_entry = self.files_queue.get(fpath)
                    if not file_entry:
                        continue

                    lookup_map = file_entry["lookup_map"]
                    time_map = file_entry.get("time_map", {})
                    index_lookup = file_entry["index_lookup"]

                    matched_idx = None

                    # Match by (before_meter, before_time)
                    if before_meter:
                        m_upper = before_meter.upper()
                        if (m_upper, before_time) in lookup_map:
                            matched_idx = lookup_map[(m_upper, before_time)]
                        elif (m_upper, raw_before_time) in lookup_map:
                            matched_idx = lookup_map[(m_upper, raw_before_time)]

                    # If not found by before_meter, search time_map
                    if matched_idx is None:
                        candidates = time_map.get(before_time) or time_map.get(raw_before_time) or []
                        if candidates:
                            if before_meter:
                                b_upper = before_meter.upper()
                                for c_idx in candidates:
                                    if index_lookup[c_idx]["meter_number_before"].strip().upper() == b_upper or index_lookup[c_idx]["meter_number"].strip().upper() == b_upper:
                                        matched_idx = c_idx
                                        break
                            if matched_idx is None:
                                matched_idx = candidates[0]

                    if matched_idx is not None:
                        if after_meter:
                            self.update_single_cell(matched_idx, "meter_number", after_meter, file_path=fpath)
                        if after_time and after_time != before_time:
                            self.update_single_cell(matched_idx, "rtc_time", after_time, file_path=fpath)
                        if after_imp != "":
                            self.update_single_cell(matched_idx, "kwh_import", after_imp, file_path=fpath)
                        if after_exp != "":
                            self.update_single_cell(matched_idx, "kwh_export", after_exp, file_path=fpath)

                        matched = True
                        total_updated += 1
                        break

                if not matched:
                    unmatched_count += 1

            else:
                # 4-Column Format
                excel_meter = str(row[0] or "").strip()
                date_time_raw = row[1]
                date_time = parse_flexible_datetime(date_time_raw)
                raw_time = str(date_time_raw or "").strip()
                kwh_import = str(row[2] or "").strip()
                kwh_export = str(row[3] or "").strip() if len(row) > 3 and row[3] is not None else ""

                if not date_time and not raw_time:
                    continue

                processed_rows += 1
                matched = False

                for fpath in target_files:
                    file_entry = self.files_queue.get(fpath)
                    if not file_entry:
                        continue

                    lookup_map = file_entry["lookup_map"]
                    time_map = file_entry.get("time_map", {})
                    index_lookup = file_entry["index_lookup"]

                    matched_idx = None

                    # 1. Attempt exact match by (excel_meter, date_time)
                    if excel_meter:
                        m_upper = excel_meter.upper()
                        if (m_upper, date_time) in lookup_map:
                            matched_idx = lookup_map[(m_upper, date_time)]
                        elif (m_upper, raw_time) in lookup_map:
                            matched_idx = lookup_map[(m_upper, raw_time)]

                    # 2. Attempt match by target_meter + date_time
                    if matched_idx is None and target_meter and target_meter not in ("AUTO", "ALL"):
                        t_upper = target_meter.strip().upper()
                        if (t_upper, date_time) in lookup_map:
                            matched_idx = lookup_map[(t_upper, date_time)]
                        elif (t_upper, raw_time) in lookup_map:
                            matched_idx = lookup_map[(t_upper, raw_time)]

                    # 3. Attempt match by date_time in time_map
                    if matched_idx is None:
                        candidates = time_map.get(date_time) or time_map.get(raw_time) or []
                        if candidates:
                            if target_meter and target_meter not in ("AUTO", "ALL"):
                                t_upper = target_meter.strip().upper()
                                for c_idx in candidates:
                                    if index_lookup[c_idx]["meter_number"].strip().upper() == t_upper:
                                        matched_idx = c_idx
                                        break
                            elif len(candidates) == 1:
                                matched_idx = candidates[0]
                            else:
                                unmod = [c for c in candidates if not index_lookup[c]["is_modified"]]
                                matched_idx = unmod[0] if unmod else candidates[0]

                    if matched_idx is not None:
                        if excel_meter:
                            self.update_single_cell(matched_idx, "meter_number", excel_meter, file_path=fpath)
                        if kwh_import != "":
                            self.update_single_cell(matched_idx, "kwh_import", kwh_import, file_path=fpath)
                        if kwh_export != "":
                            self.update_single_cell(matched_idx, "kwh_export", kwh_export, file_path=fpath)

                        matched = True
                        total_updated += 1
                        break

                if not matched:
                    unmatched_count += 1

        # Refresh file summary meters list in case meter numbers were updated
        for fpath in target_files:
            file_entry = self.files_queue.get(fpath)
            if file_entry:
                idx_lookup = file_entry["index_lookup"]
                all_m = sorted(list(set(r["meter_number"] for r in idx_lookup)))
                file_entry["summary"]["meters"] = all_m
                file_entry["summary"]["unique_meters_count"] = len(all_m)
                new_lmap, new_tmap = self._build_lookup_map(idx_lookup)
                file_entry["lookup_map"] = new_lmap
                file_entry["time_map"] = new_tmap

        return {
            "total_excel_rows": processed_rows,
            "updated_intervals": total_updated,
            "unmatched_rows": unmatched_count
        }

    def batch_paste_kwh(self, start_index, meter_filter, kwh_values_list, file_path=None):
        """Pastes a column of numbers sequentially into consecutive intervals."""
        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            return 0

        index_lookup = self.files_queue[target_path]["index_lookup"]
        matching_indices = [
            r["index"] for r in index_lookup
            if (not meter_filter or meter_filter == "ALL" or r["meter_number"] == meter_filter)
        ]
        
        try:
            start_pos = matching_indices.index(start_index)
        except ValueError:
            start_pos = 0

        updated_count = 0
        for i, val in enumerate(kwh_values_list):
            pos = start_pos + i
            if pos >= len(matching_indices):
                break
            target_idx = matching_indices[pos]
            self.update_single_cell(target_idx, "kwh_import", val, file_path=target_path)
            updated_count += 1

        return updated_count

    def batch_set_kwh(self, meter_filter, new_kwh=None, multiply_factor=None, precision=5, file_path=None):
        """Batch set or multiply kWh values for active file."""
        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            return 0

        index_lookup = self.files_queue[target_path]["index_lookup"]
        matching_indices = [
            r["index"] for r in index_lookup
            if (not meter_filter or meter_filter == "ALL" or r["meter_number"] == meter_filter)
        ]

        count = 0
        for idx in matching_indices:
            current_v = index_lookup[idx]["kwh_import"]
            if new_kwh is not None:
                new_v = f"{float(new_kwh):.{precision}f}"
            else:
                try:
                    num = float(current_v)
                    if multiply_factor is not None:
                        num = num * float(multiply_factor)
                    new_v = f"{num:.{precision}f}"
                except ValueError:
                    continue
            self.update_single_cell(idx, "kwh_import", new_v, file_path=target_path)
            count += 1
        return count

    def batch_shift_time(self, meter_filter, shift_minutes=0, shift_days=0, sync_hes=True, file_path=None):
        """Batch shifts timestamps while preserving exact DD/MM/YYYY HH:MM:SS format."""
        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            return 0

        file_entry = self.files_queue[target_path]
        index_lookup = file_entry["index_lookup"]
        raw_data = file_entry["raw_data"]

        matching_indices = [
            r["index"] for r in index_lookup
            if (not meter_filter or meter_filter == "ALL" or r["meter_number"] == meter_filter)
        ]

        delta = timedelta(days=shift_days, minutes=shift_minutes)
        count = 0
        for idx in matching_indices:
            rtc_str = index_lookup[idx]["rtc_time"]
            hes_str = index_lookup[idx]["hes_time"]

            try:
                rtc_dt = datetime.strptime(rtc_str, "%d/%m/%Y %H:%M:%S")
                new_rtc = (rtc_dt + delta).strftime("%d/%m/%Y %H:%M:%S")
                self.update_single_cell(idx, "rtc_time", new_rtc, file_path=target_path)

                if sync_hes and hes_str:
                    hes_dt = datetime.strptime(hes_str, "%d/%m/%Y %H:%M:%S")
                    new_hes = (hes_dt + delta).strftime("%d/%m/%Y %H:%M:%S")
                    raw_data[idx]["HesTime"] = new_hes
                    index_lookup[idx]["hes_time"] = new_hes

                count += 1
            except Exception:
                continue

        lmap, tmap = self._build_lookup_map(index_lookup)
        file_entry["lookup_map"] = lmap
        file_entry["time_map"] = tmap
        return count

    def batch_rename_meter(self, old_meter, new_meter, file_path=None):
        """Batch renames a meter ID across all records in active file."""
        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            return 0

        file_entry = self.files_queue[target_path]
        index_lookup = file_entry["index_lookup"]
        count = 0
        for idx in range(len(index_lookup)):
            if index_lookup[idx]["meter_number"] == old_meter:
                self.update_single_cell(idx, "meter_number", new_meter, file_path=target_path)
                count += 1

        lmap, tmap = self._build_lookup_map(index_lookup)
        file_entry["lookup_map"] = lmap
        file_entry["time_map"] = tmap
        return count

    def export_data_csv(self, export_type="comparison", file_path=None):
        """
        Exports tabular data as CSV string.
        Options:
          - 'source': Original source 4 parameters
          - 'output': Modified output 4 parameters
          - 'comparison': Before & After comparison with status
        """
        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            raise KeyError("No active file loaded to export")

        index_lookup = self.files_queue[target_path]["index_lookup"]
        output = io.StringIO()
        writer = csv.writer(output)

        if export_type == "source":
            writer.writerow(["Meter Number", "Date & Time (0.0.1.0.0.255)", "Import kWh (1.0.1.29.0.255)", "Export kWh (1.0.2.29.0.255)"])
            for r in index_lookup:
                writer.writerow([r["meter_number_before"], r["rtc_time_before"], r["kwh_import_before"], r["kwh_export_before"]])

        elif export_type == "output":
            writer.writerow(["Meter Number", "Date & Time (0.0.1.0.0.255)", "Import kWh (1.0.1.29.0.255)", "Export kWh (1.0.2.29.0.255)"])
            for r in index_lookup:
                writer.writerow([r["meter_number"], r["rtc_time"], r["kwh_import"], r["kwh_export"]])

        elif export_type == "comparison":
            writer.writerow([
                "Meter Number [Before]",
                "Meter Number [After]",
                "Date & Time", 
                "Import kWh [Before]", 
                "Import kWh [After]", 
                "Export kWh [Before]", 
                "Export kWh [After]", 
                "Status"
            ])
            for r in index_lookup:
                writer.writerow([
                    r["meter_number_before"],
                    r["meter_number"],
                    r["rtc_time"],
                    r["kwh_import_before"],
                    r["kwh_import"],
                    r["kwh_export_before"],
                    r["kwh_export"],
                    "Modified" if r["is_modified"] else "Original"
                ])

        return output.getvalue()

    def get_output_filename(self, file_path):
        """Generates destination filename with __Modified suffix (e.g. LP_...__Modified.json)."""
        base_name = os.path.basename(file_path)
        name, ext = os.path.splitext(base_name)
        if not name.endswith("__Modified") and not name.endswith("_Modified"):
            name = f"{name}__Modified"
        return f"{name}{ext}"

    def save_file(self, destination_folder, file_path=None, custom_filename=None):
        """
        Saves the modified JSON to the destination folder with __Modified suffix.
        Uses separators=(',', ':') to preserve the exact compact structure
        without changing any punctuation, whitespace, commas, or order.
        """
        target_path = os.path.abspath(file_path) if file_path else self.active_file_path
        if not target_path or target_path not in self.files_queue:
            raise ValueError("No file loaded to save")

        file_entry = self.files_queue[target_path]
        raw_data = file_entry["raw_data"]

        os.makedirs(destination_folder, exist_ok=True)
        filename = custom_filename or self.get_output_filename(target_path)
        out_path = os.path.join(destination_folder, filename)

        with open(out_path, "w", encoding="utf-8") as fp:
            json.dump(raw_data, fp, separators=(',', ':'), ensure_ascii=False)

        file_entry["status"] = "Saved"
        file_entry["summary"]["status"] = "Saved"

        return {
            "saved_path": out_path,
            "filename": filename,
            "size_bytes": os.path.getsize(out_path),
            "total_records": len(raw_data)
        }

    def save_all_files(self, destination_folder):
        """Saves ALL files in the queue to destination_folder/<filename>__Modified.json."""
        os.makedirs(destination_folder, exist_ok=True)
        saved_list = []
        for fpath in list(self.files_queue.keys()):
            res = self.save_file(destination_folder, file_path=fpath)
            saved_list.append(res)
        return saved_list
