# CLAUDE.md

Project notes for Claude. See [README.md](README.md) for the script plugin interface
(`SCRIPT_NAME`, `REQUIRED_FILE_COUNT`, `ALLOW_MULTIPLE_FILES`, `process_data`).

## PO data assumptions (confirmed by owner — do not "fix" these)

- **SKUCODE + ETD is unique** within a PO file. SKUCODE alone may repeat, but the
  SKUCODE + ETD pair never does. The exact-match lookup in `scripts/compare_po.py`
  relies on this and is working as intended.
- **ETD is always filled** on every PO row. Blank/NaN ETD handling is not needed.
- **Header row position:**
  - `scripts/merge_po_files.py` must accept the PO header on **row 1 or row 2**,
    because source files come from other departments with inconsistent layouts.
  - `scripts/compare_po.py` only reads files **produced by the merge script**, which
    always writes the header on row 1. Row-2 header support is not needed there.

## Known improvement ideas (not yet done)

- `merge_po_files.py` collects warnings (skipped sheets, unreadable workbooks) but only
  `print`s them to the console, so GUI users never see them. Candidate improvement:
  surface them in the GUI or write them to a sheet in the output workbook.
- Output files are saved to the current working directory. The owner may change this
  later if needed. Leave it as is unless asked.

## Other

- `scripts/test.py` is a demo script and is intentionally kept. More personal-use
  scripts will be added to `scripts/` over time.
