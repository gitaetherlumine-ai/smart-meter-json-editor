#!/usr/bin/env python3
"""
Smart Meter JSON Editor - Command Line Interface (CLI)
Provides fast headless operations for batch edits, replacements, and transformations.
Preserves 100% exact JSON structure, commas, and formatting.
"""

import os
import sys
import argparse
from editor_core import JsonMeterEditor

def main():
    parser = argparse.ArgumentParser(description="Smart Meter JSON Editor CLI")
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # Command: list
    list_p = subparsers.add_parser("info", help="Inspect and summarize a JSON file")
    list_p.add_argument("--file", "-f", required=True, help="Path to JSON file")

    # Command: set-kwh
    kwh_p = subparsers.add_parser("set-kwh", help="Set or scale kWh import values")
    kwh_p.add_argument("--file", "-f", required=True, help="Path to input JSON file")
    kwh_p.add_argument("--meter", "-m", default="ALL", help="Target meter number (or ALL)")
    kwh_p.add_argument("--val", "-v", type=float, help="Exact fixed kWh value")
    kwh_p.add_argument("--factor", type=float, help="Multiply factor (e.g. 1.05)")
    kwh_p.add_argument("--out", "-o", required=True, help="Output destination folder")

    # Command: shift-time
    time_p = subparsers.add_parser("shift-time", help="Shift interval date/times")
    time_p.add_argument("--file", "-f", required=True, help="Path to input JSON file")
    time_p.add_argument("--meter", "-m", default="ALL", help="Target meter number (or ALL)")
    time_p.add_argument("--mins", type=int, default=0, help="Shift minutes (+ or -)")
    time_p.add_argument("--days", type=int, default=0, help="Shift days (+ or -)")
    time_p.add_argument("--no-sync-hes", action="store_true", help="Do not shift HES time")
    time_p.add_argument("--out", "-o", required=True, help="Output destination folder")

    # Command: rename-meter
    ren_p = subparsers.add_parser("rename-meter", help="Rename a meter ID across all records")
    ren_p.add_argument("--file", "-f", required=True, help="Path to input JSON file")
    ren_p.add_argument("--old", required=True, help="Existing meter number")
    ren_p.add_argument("--new", required=True, help="New meter number")
    ren_p.add_argument("--out", "-o", required=True, help="Output destination folder")

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)

    editor = JsonMeterEditor()

    if args.command == "info":
        s = editor.load_file(args.file)
        print(f"File: {s['file_name']}")
        print(f"Total Intervals: {s['total_records']}")
        print(f"Unique Meters: {s['unique_meters_count']}")
        print(f"Sample meters: {s['meters'][:10]}")

    elif args.command == "set-kwh":
        editor.load_file(args.file)
        count = editor.batch_set_kwh(args.meter, new_kwh=args.val, multiply_factor=args.factor)
        res = editor.save_file(args.out)
        print(f"Updated {count} records. Saved to: {res['saved_path']}")

    elif args.command == "shift-time":
        editor.load_file(args.file)
        count = editor.batch_shift_time(args.meter, shift_minutes=args.mins, shift_days=args.days, sync_hes=not args.no_sync_hes)
        res = editor.save_file(args.out)
        print(f"Shifted timestamps for {count} records. Saved to: {res['saved_path']}")

    elif args.command == "rename-meter":
        editor.load_file(args.file)
        count = editor.batch_rename_meter(args.old, args.new)
        res = editor.save_file(args.out)
        print(f"Renamed meter in {count} records. Saved to: {res['saved_path']}")

if __name__ == "__main__":
    main()
