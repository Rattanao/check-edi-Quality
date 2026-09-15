# check-edi-Quality

`build_edi.py` compares an ocean-freight carrier's **MANIFEST** (`.xls`/`.xlsx`, free-form cargo
report) against a customer's **ENTER** document (`.pdf`, per-B/L confirmation/amendment sheet,
often with hand-typed red annotations for STATUS/DG/REEFER TEMP/TRANSIT) before filing a Thailand
customs import declaration (ใบขนสินค้าขาเข้า). It produces a single color-coded `EDI.xlsx` report
comparing every sub-B/L on CONSIGNEE, CONTAINER NO., STATUS, TOTAL PACKAGE, PACKAGING, GROSS
WEIGHT, MEASUREMENT, MARKS, DESCRIPTION, REEFER TEMP, DG (CLASS/UN), and TRANSIT/TRANSHIPMENT,
plus a TOTAL row cross-checking summed values against each document's own stated totals, and a
critical row for any B/L that exists in ENTER but is entirely missing from MANIFEST.

## Usage

```bash
pip install -r requirements.txt
python build_edi.py --manifest MANIFEST.xls --enter ENTER.pdf --outdir .
```

Or drop `MANIFEST.xls`/`.xlsx` and `ENTER.pdf` into an `input/` folder next to the script and run
`python build_edi.py` with no arguments — it auto-detects the files.

## Report conventions

- Red = confident mismatch, orange = needs human confirmation, yellow = no ENTER doc to compare
  against, green = pass (summary column only).
- MARKS/DESCRIPTION comparisons ignore spacing/punctuation differences (PDF/Excel line-wrap
  artifacts) but flag any real wording difference.
- DG, REEFER TEMP, and TRANSIT/TRANSHIPMENT are flagged whenever they appear in either document,
  since they always require customs follow-up.

See the inline comments in `build_edi.py` for the exact column-index assumptions per document
layout, and the two ENTER PDF layouts (and a B/L-range FreeText-annotation convention) it
auto-detects.

## Note on input data

`input/` and `output/` are git-ignored — this repo tracks the program only, not the customer
shipping documents it processes.
