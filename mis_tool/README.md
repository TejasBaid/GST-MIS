# GSTR-3B → Excel Auto-Mapper

A Streamlit app that reads GSTR-3B PDFs and auto-populates your GST Monthly Analysis Excel file — preserving all formulas and formatting.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Usage

1. Open the app in your browser (http://localhost:8501)
2. Upload your **GST Monthly Analysis Excel** (e.g. `GST_Monthly_Analysis_-_FY_2026-27.xlsx`)
3. Upload **one or more GSTR-3B PDFs** (any number of GSTINs / months at once)
4. Review the extracted data preview
5. Click **Map Data & Download Excel**
6. Open the downloaded file — all data is populated, formulas intact

## What gets mapped

### Per-GSTIN sheets (e.g. `01AABCG4768K1ZL`, `02AABCG4768K1ZJ`, ...)

| Section | Fields written |
|---------|----------------|
| 3.1(a) Outward taxable | Taxable Value, IGST, CGST, Cess |
| 3.1(b) Zero rated | Taxable Value, IGST |
| 3.1(c) Nil/Exempt | Taxable Value |
| 3.1(d) RCM inward | Taxable Value, IGST, CGST |
| 3.1(e) Non-GST | Value |
| 3.2 Inter-state | URD/Composition/UIN taxable + IGST |
| 4A ITC Available | IGST (all other + ISD), CGST, RCM ITC |
| 4B ITC Reversed | Rules 38/42/43 + Others (combined) |
| 5 Exempt/Non-GST | Inter/Intra state supplies |
| 5.1 Interest & Late Fee | All tax types |
| 6.1A Payment | ITC paid through IGST/CGST/SGST |
| 6.1B RCM Payment | Cash paid for reverse charge |

### Turnover sheet

New rows are appended for each GSTIN × supply type × month combination (matching the existing pivot-table source data format).

## Formula preservation

- Cells with formulas like `=EOMONTH(A5,1)`, `=SUM(B5:B16)`, `=D7` are **never overwritten**
- SGST columns that use `=CGST formula` are preserved unless SGST ≠ CGST
- Tax-payable cells in section 6.1 that reference section 3.1 are left as formulas

## Month row mapping

The template uses rows 5–16 for April–March:

| Month | Row |
|-------|-----|
| April | 5 |
| May | 6 |
| June | 7 |
| July | 8 |
| August | 9 |
| September | 10 |
| October | 11 |
| November | 12 |
| December | 13 |
| January | 14 |
| February | 15 |
| March | 16 |

(Same offset applies to ITC section rows 22–33, Section 5 rows 40–51, and Payment rows 58–69 / 76–87 / 94–105.)
