"""Shared RAW CSV reader so analysis and Excel export use identical normalization.

- Encoding: UTF-8 (with/without BOM) first, then CP949 (Korean Excel "CSV" save).
- Fully blank rows (e.g. ",,,,," left by Excel) are skipped.
- Leading/trailing whitespace is stripped from every field.
- instrument_id is upper-cased (device_a -> DEVICE_A).
"""
from pathlib import Path
import csv

ENCODINGS = ('utf-8-sig', 'cp949')


def read_raw_rows(path):
    path = Path(path)
    last_error = None
    for enc in ENCODINGS:
        try:
            with path.open('r', encoding=enc, newline='') as f:
                reader = csv.DictReader(f)
                fieldnames = [str(c).strip() for c in (reader.fieldnames or [])]
                rows = []
                for line_no, r in enumerate(reader, 2):
                    if None in r:
                        raise ValueError(f'RAW row {line_no}: extra CSV fields')
                    values = [(v or '').strip() for v in r.values()]
                    if not any(values):
                        continue  # blank row
                    rec = {k: v for k, v in zip(fieldnames, values)}
                    if 'instrument_id' in rec:
                        rec['instrument_id'] = rec['instrument_id'].upper()
                    rec['_line'] = line_no
                    rows.append(rec)
            return fieldnames, rows, enc
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ValueError(f'RAW CSV encoding not supported (tried {ENCODINGS}): {last_error}')


def natural_key(text):
    """S2 < S10 ordering for sample IDs without zero padding."""
    import re
    return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', str(text))]
