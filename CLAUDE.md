# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Streamlit web application nicknamed "T-1000" that automates the cleaning and processing of ABI (construction/contracting) invoice Excel files. The app is deployed at https://abi-t1000.streamlit.app/.

## Application Architecture

**Single-file application**: The entire application logic is contained in `app.py` (250 lines), making this a straightforward single-module Streamlit app.

**Core workflow**:
1. Users upload multiple Excel invoice files via Streamlit file uploader
2. Each file is processed through `clean_dataframe()` which extracts invoice metadata and vendor/subcontractor details from specific cell locations
3. Cleaned data is restructured into a standardized tabular format
4. Output files are renamed following the pattern: `{sendername}_{vendorname}_{invoicenumber}_clean.xlsx`
5. All cleaned files are packaged into a timestamped zip file for download

**Data extraction logic** (`clean_dataframe()` function):
- Invoice metadata is extracted from hardcoded cell positions (rows 0-10, column 1)
- Vendor/subcontractor data is extracted from rows 13-32 across columns 0-7
- The function assumes a rigid Excel template structure - cell positions are not dynamic
- Financial values are formatted with currency symbols and thousand separators
- "Prime/Sub" classification is determined by matching vendor names against the prime contractor field

**File naming convention**: Output files are constructed from three components split by underscores:
- Sender name: extracted from original filename before '@' symbol (`filename.split('@')[0]`)
- Vendor name: from first row's vendor field, with spaces/periods/commas removed
- Invoice number: from the vendor invoice metadata field

## Development Commands

**Run locally**:
```bash
streamlit run app.py
```

**Install dependencies**:
```bash
pip install -r requirements.txt
```

**Deploy**: The app is configured for Streamlit Cloud deployment (see `.streamlit/config.toml` for theme settings).

## Key Technical Details

**Dependencies**:
- `streamlit==1.27.2` - Web app framework
- `pandas==2.1.4` - Data processing
- `openpyxl==3.1.2` - Reading Excel files
- `xlsxwriter==3.1.9` - Writing Excel files with autofit columns
- `pytz==2024.2` - Eastern timezone timestamps for zip filenames

**Custom styling**: CSS is injected via `st.markdown()` to hide default Streamlit UI elements (footer, menu, sidebar toggle) and apply custom color scheme matching ABI branding (blue: `#005cb9`).

**Error handling**: The app catches exceptions during file processing and displays user-friendly error messages without breaking the entire batch operation. Files that error are skipped with a red warning message.

**Assets**: The `Content/` directory contains the ABI logo (`abi_2.png`) displayed in the bottom-right corner.

**Timezone handling**: All zip file timestamps use America/New_York timezone, formatted as `MM-DD-YYYY_HH.MMam/pm`.

## Important Constraints

When modifying the data extraction logic, note that the Excel template has a **fixed structure**:
- Invoice header metadata: Rows 0-10, Column B (index 1)
- Vendor table: Rows 13-32 (max 19 vendors), Columns A-H (indices 0-7)
- Invoice total: Row 33, Column E (index 4)

These hardcoded positions are critical to the cleaning logic and should not be changed without corresponding updates to the source Excel template.
