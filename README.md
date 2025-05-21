# gm---bulk-specs-converter---streamlit--v1

## Description
The `gm---bulk-specs-converter---streamlit--v1` script is a Python-based application designed to convert product specifications from Excel format into beautifully formatted HTML. It supports features such as single product SKUs, package SKUs with multiple products displayed in tabs, collapsible sections (using "Start"/"End" markers) within any SKU or tab, dedicated Care Instruction sections, and notes within specification or care sections.

## Features
- Conversion of product specifications from Excel to HTML format.
- Support for single product SKUs and package SKUs with multiple products in tabs.
- Collapsible sections (using "Start"/"End" markers) within any SKU or tab.
- Dedicated Care Instruction sections.
- Support for notes within specification or care sections.

## Prerequisites/Dependencies
The script requires the following Python libraries:
- `pandas`
- `math`
- `traceback`
- `bs4` (BeautifulSoup)
- `datetime`
- `io`
- `streamlit`

## How to Use/Run
1. Launch the Streamlit application.
2. Click "Browse files" to choose your Excel file (.xlsx).
3. (Optional) Enter a custom width for the first column (e.g., "180px") or leave "Auto width" checked.
4. Click "Convert to HTML" to start the process.
5. Progress will be shown, and a message will appear upon completion or error.
6. The output is saved as a new Excel file that you can download.

## Input Format
The input Excel file should be structured according to the instructions provided in the "Preparing Your Input (Tabs & Details)" section of the instructions HTML.

## License
License to be determined.