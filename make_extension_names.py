"""Rebuild extension_names.xlsx from a Yeastar user export (CSV or XLSX).

    python make_extension_names.py Ext_Name_Mapping.csv

Only extension numbers and names are read; passwords, PINs and e-mail addresses in the export are ignored and are
not written to the result.
"""
import sys

from transfer_report.ext_names import DEFAULT_NAMES_PATH, load_extension_names, save_extension_names

if len(sys.argv) != 2:
    sys.exit("Usage: python make_extension_names.py <Yeastar user export .csv/.xlsx>")
names = load_extension_names(sys.argv[1])
print(f"{len(names)} extensions -> {save_extension_names(names, DEFAULT_NAMES_PATH)}")
