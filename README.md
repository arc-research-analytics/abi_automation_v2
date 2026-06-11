# ABI Excel Cleaning Robot (T-1000)

A Streamlit web application that automates the cleaning and processing of ABI construction/contracting invoice Excel files. Access the live app [here](https://abi-t1000.streamlit.app/).

---

## What It Does

The app accepts one or more raw ABI invoice Excel files and transforms them into clean, standardized tabular spreadsheets ready for analysis or reporting. Key steps in the workflow:

1. **Upload** — Users drag and drop one or more `.xlsx` invoice files into the browser interface.
2. **Parse** — Each file is read into a Pandas DataFrame. Invoice metadata is extracted from fixed cell positions in the header block (rows 1–11, column B), and vendor/subcontractor line items are pulled from the table body (rows 14–32, columns A–H).
3. **Clean & Standardize** — The raw layout is restructured into a flat, one-row-per-vendor table with 21 named columns. Financial values are formatted as currency strings. Dropdown placeholder values (`"Select"`) and blanks are normalized to `"N/A"`. Each vendor row is classified as `"Prime"` or `"Sub"` by comparing the vendor name against the prime contractor field.
4. **Rename** — Each output file is named using a structured convention: `{sender}_{date}_{PrimeVendorName}_{InvoiceNumber}.xlsx`. The sender and date are parsed from the original filename; the vendor name is sanitized (alphanumeric only, capped at 20 characters).
5. **Package & Download** — All cleaned files are bundled into a timestamped `.zip` archive (Eastern time) and offered as a single download button.

Files that fail to process are skipped with a detailed inline error message; the rest of the batch continues uninterrupted.

---

## Output Columns

Each cleaned file contains the following columns:

| Column | Source |
|---|---|
| Invoice Date | Row 1, Col B |
| Vendor Invoice # | Row 2, Col B |
| ABI Contract Name | Row 3, Col B |
| ABI Contract # | Row 4, Col B |
| Project or Work Order Name | Row 5, Col B |
| ABI Project # | Row 6, Col B |
| ABI Work Order # | Row 7, Col B (N/A if blank) |
| Total Contract/Work Order Amt | Row 8, Col B |
| Current Invoice Amount | Row 9, Col B |
| ABI Cost Code # | Row 10, Col B |
| Prime Contractor/Vendor | Row 11, Col B |
| Vendor/Subcontractor | Rows 14–32, Col A |
| Prime/Sub | Derived: matches vendor against prime |
| Certification | Rows 14–32, Col B |
| Race/Ethnicity | Rows 14–32, Col C |
| Additional DBE Types | Rows 14–32, Col D |
| Net Invoice Amount ($) | Rows 14–32, Col E |
| Net Contracted Amount ($) | Rows 14–32, Col F |
| Total Invoiced to Date ($) | Rows 14–32, Col G |
| Newly Added? | Rows 14–32, Col H |
| Invoice Total | Row 34, Col E |

---

## Tech Stack

| Technology | Role |
|---|---|
| [Streamlit](https://streamlit.io/) `1.27.2` | Web app framework and UI |
| [Pandas](https://pandas.pydata.org/) `2.1.4` | DataFrame parsing, transformation, and output |
| [openpyxl](https://openpyxl.readthedocs.io/) `3.1.2` | Reading `.xlsx` input files |
| [XlsxWriter](https://xlsxwriter.readthedocs.io/) `3.1.9` | Writing output `.xlsx` files with auto-fit columns |
| [Pillow](https://pillow.readthedocs.io/) | Rendering the ABI logo image |
| [pytz](https://pythonhosted.org/pytz/) `2024.2` | Eastern timezone formatting for zip timestamps |

---

## Running Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app is deployed on [Streamlit Community Cloud](https://streamlit.io/cloud).

---

## Important Constraint

The extraction logic relies on a **fixed Excel template structure**. If the source template changes, the cell references in `clean_dataframe()` must be updated to match.
