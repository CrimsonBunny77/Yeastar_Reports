"""Regenerate the sample department_mapping.xlsx (overwrites the existing file)."""
from transfer_report.config import DEFAULT_MAPPING_PATH
from transfer_report.mapping import create_template

print("Written:", create_template(DEFAULT_MAPPING_PATH))
