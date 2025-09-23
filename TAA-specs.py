# -*- coding: utf-8 -*-
import sys
import os
import pandas as pd
import math  # Import math for isnan check
import traceback
from bs4 import BeautifulSoup
from datetime import datetime
import io  # For BytesIO

import streamlit as st

# --- Instructions HTML ---
def get_instructions_html():
    return """
    <h1>Specs HTML Converter User Guide</h1>

    <h2>Table of Contents</h2>
    <ol>
        <li><a href="#introduction">Introduction</a></li>
        <li><a href="#getting-started">Getting Started</a></li>
        <li><a href="#using-the-application">Using the Application</a></li>
        <li><a href="#preparing-your-input">Preparing Your Input (Tabs & Details)</a></li>
        <li><a href="#understanding-the-output">Understanding the Output</a></li>
        <li><a href="#troubleshooting">Troubleshooting</a></li>
    </ol>

    <h2 id="introduction">1. Introduction</h2>
    <p>The Specs HTML Converter is designed to convert product specifications from Excel format into beautifully formatted HTML. It supports:</p>
    <ul>
        <li>Single product SKUs.</li>
        <li>Package SKUs with multiple products displayed in <strong>Tabs</strong>.</li>
        <li><strong>Collapsible sections</strong> (using "Start"/"End" markers) within any SKU or tab.</li>
        <li>Dedicated Care Instruction sections.</li>
        <li>Notes within specification or care sections.</li>
    </ul>


    <h2 id="getting-started">2. Getting Started</h2>
    <p>To use the Specs HTML Converter:</p>
    <ol>
        <li>Ensure your Excel file is properly formatted (see Preparing Your Input).</li>
        <li>Launch the application.</li>
        <li>Select your input Excel file.</li>
        <li>Configure the width settings if needed (Auto width is recommended).</li>
        <li>Click "Convert to HTML" to generate the output Excel file.</li>
    </ol>

    <h2 id="using-the-application">3. Using the Application</h2>
    <ol>
        <li>Click "Browse files" to choose your Excel file (.xlsx).</li>
        <li>(Optional) Enter a custom width for the first column (e.g., "180px") or leave "Auto width" checked.</li>
        <li>Click "Convert to HTML" to start the process.</li>
        <li>Progress will be shown, and a message will appear upon completion or error.</li>
        <li>The output is saved as a new Excel file that you can download.</li>
    </ol>

    <h2 id="preparing-your-input">4. Preparing Your Input (Tabs & Details)</h2>
    <p>Prepare an Excel (.xlsx) file. The structure depends on whether you have a single product or a package with tabs:</p>

    <h3>A. General Structure:</h3>
    <ul>
        <li><b>Column A:</b> Contains the SKU for the product/package OR a numeric marker (1, 2, 3...) to indicate the start of a new Tab within a package. Leave blank for continuation rows, "Start", or "End" markers. (The SKU row itself is for identification and will not be rendered as a spec line in the HTML table).</li>
        <li><b>Column B:</b>
            <ul>
                <li>If Column A is a <strong>numeric Tab marker</strong>: This column MUST contain the <strong>Tab Title</strong> (e.g., "Product 1 Details", "Canopy Specs").</li>
                <li>If Column A is a <strong>SKU</strong> (on the SKU identification row) or <strong>blank</strong> (on a data row): This is the first column for US/CA data (e.g., Specification Title, Care Header, Note). Ignore rows where Column A contains only 'US' or 'UK'.</li>
                <li>If the row is for a <strong>Collapsible Section Title</strong>: Put the title here (e.g., "Dimensions & Weights").</li>
                <li>If the row marks the <strong>Start</strong> of a collapsible section: This column contains the <em>first header</em> for the nested table (e.g., "Dimension").</li>
                 <li>If the row marks the <strong>End</strong> of a collapsible section: This column contains the <em>first data cell</em> for the last row of the nested table.</li>
           </ul>
        </li>
        <li><b>Column C & D (for US/CA data):</b> Specification values, care details, notes content, or subsequent headers/data for collapsible sections.</li>
         <li><b>Column E onwards (for UK/AU/NZ data):</b> The corresponding data for these regions. Column E should align with Column B's purpose (Spec Title, Care Header, etc. for UK), Column F with Column C (Spec Value, Care Detail), and so on.</li>
         <li><b>Care Instructions:</b> Use headers like "Graphic Care Instructions", "Washing Instructions", etc., in the appropriate column (B for US, E for UK) to start a care section within a SKU or Tab. Subsequent rows list instructions or notes.</li>
         <li><b>Notes:</b> Start a cell in the relevant 'title' column (B for US, E for UK) with "Note:" (case-insensitive) to create a formatted note block.</li>
    </ul>

    <h3>B. Creating Tabs (for Package SKUs):</h3>
    <ol>
        <li>Enter the main Package SKU in Column A of the first row for that package. (This row identifies the package; its other columns are not rendered as specs).</li>
        <li>For the first tab, add a row below the SKU row:
            <ul><li>Enter '1' in Column A.</li><li>Enter the desired <strong>Tab Title</strong> (e.g., "Frame Details") in Column B.</li></ul>
        </li>
        <li>Add the specification rows, care instructions, notes, and any collapsible sections for this first tab below the Tab Title row. Data for US/CA goes in columns B, C, D and data for UK/AU/NZ goes in columns E, F, G ... Rows starting with just 'US' or 'UK' in Column A are ignored.</li>
        <li>For the second tab, add a new row:
             <ul><li>Enter '2' in Column A.</li><li>Enter the <strong>Tab Title</strong> for the second tab (e.g., "Graphic Specs") in Column B.</li></ul>
        </li>
         <li>Add the data rows for the second tab below its title row.</li>
         <li>Repeat for any subsequent tabs (using '3', '4', etc. in Column A).</li>
         <li>The data rows following a tab marker belong to that tab until the next tab marker or a new SKU is encountered.</li>
    </ol>
     <p><b>Example Structure (Package SKU with Tabs):</b></p>
    <pre>
    | Col A    | Col B                 | Col C            | Col D | Col E (UK Title)     | Col F (UK Value) | Col G (UK Value2) |
    |----------|-----------------------|------------------|-------|----------------------|------------------|-------------------|
    | PKGSKU01 | Description...        | URL...           |       | Description UK...    | URL UK...        |                   | &lt;- SKU row, not rendered as spec
    | US       |                       |                  |       |                      |                  |                   | &lt;- This row ignored
    | 1        | Frame Specifications  |                  |       | Frame Specs UK       |                  |                   | &lt;- Tab Title row
    |          | Material              | Aluminum         |       | Material             | Aluminium        |                   |
    |          | Weight                | 5 kg             |       | Weight               | 5 kg             |                   |
    | 2        | Canopy Details        |                  |       | Canopy Details UK    |                  |                   |
    |          | Fabric Type           | Polyester        |       | Fabric Type          | Polyester        |                   |
    |          | Graphic Care...       | Wipe clean only  |       | Graphic Care...      | Wipe clean only  |                   |
    |          | Note:                 | Handle with care |       | Note:                | Handle with care |                   |
    </pre>

    <h3>C. Creating Collapsible Sections (within a SKU or Tab):</h3>
    <p>These work the same way whether inside a tab or a regular SKU's data.</p>
    <ol>
        <li>Add a row with the main title for the section (e.g., "Dimensions & Weights") in the appropriate title column (B for US, E for UK).</li>
        <li>Immediately below this title row, add a row with:
            <ul>
                <li>"Start" (case-insensitive) in Column A.</li>
                <li>The <b>header row</b> data for the nested table starting from the appropriate title column (B for US headers in B,C,D; E for UK headers in E,F,G).</li>
            </ul>
        </li>
        <li>Add the subsequent data rows for the nested table (Column A empty, data starts in appropriate columns B,C,D for US or E,F,G for UK).</li>
        <li>Immediately after the last data row for the nested table, add a row with:
             <ul>
                <li>"End" (case-insensitive) in Column A.</li>
                <li>The data for the *last row* of the nested table starting from the appropriate columns.</li>
            </ul>
        </li>
    </ol>
    <p><b>Example (Collapsible Section within US data):</b></p>
    <pre>
    | Column A | Column B             | Column C           | Column D          | Column E... |
    |----------|----------------------|--------------------|-------------------|-------------|
    |          | Flag Size (W x H)... |                    |                   | (UK Data...) |
    | Start    | 1.5' x 5.5'          | 8.86' (3 poles)    |                   |             |
    |          | 2' x 7.5'            | 11.48' (4 poles)   |                   |             |
    | ...      | ...                  | ...                | ...               |             |
    | End      | 2.5' x 15.5'         | 18.04' (5 poles)   |                   |             |
    |          | Total Package Weight | 44 lbs             |                   |             |
    </pre>


    <h2 id="understanding-the-output">5. Understanding the Output</h2>
    <ul>
        <li>The application produces a new Excel file named like `YourInputFile_output_YYYYMMDD_HHMMSS.xlsx`.</li>
        <li>It contains columns: SKU, Region, HTML.</li>
        <li>Each input SKU will have multiple rows in the output, one for each target region (default, canada, unitedkingdom, australia, newzealand).</li>
        <li>The HTML column contains the fully formatted HTML, including styles, tabs (if applicable), collapsible sections, tables, lists, and notes, ready to be used.</li>
    </ul>

    <h2 id="troubleshooting">6. Troubleshooting</h2>
    <ul>
        <li><strong>File Read Error:</strong> Ensure the Excel file is closed in other applications (like Excel itself). Verify it's a valid .xlsx file.</li>
        <li><strong>Percentages (e.g., 50%) showing as decimals (0.5):</strong> This tool is designed to read all data as text to prevent this. If it still occurs, ensure the cells in Excel are formatted as 'Text' before entering the percentage value.</li>
        <li><strong>No Output / Incorrect HTML / `No valid SKU data...` Error:</strong>
            <ul>
                <li>Double-check the Excel structure against the "Preparing Your Input" guide. Pay close attention to column usage for SKUs (Col A, must be first row for the product, this row itself is not output as a spec), Tab Markers (Col A, numeric), Tab Titles (Col B, on same row as Tab Marker), US data (B, C, D), UK data (E, F, G...), "Start"/"End" markers (Col A), and Care Headers/Notes (Col B/E).</li>
                 <li>Ensure the *very first row* for each product/package has the SKU in Column A.</li>
                 <li>Ensure Tab Markers (1, 2...) are numeric and in Column A, with Tab Titles immediately following in Column B.</li>
                <li>Verify "Start" is on the same row as the *headers* of the collapsible table, and "End" is on the same row as the *last data row* of that table.</li>
                <li>Check for hidden characters or extra spaces, especially in marker cells (A, B, E).</li>
                <li>Make sure specification data exists in the correct columns for US and UK regions on rows *after* the SKU identification row.</li>
                <li>Rows with just 'US' or 'UK' in Column A should exist but are ignored for data processing.</li>
            </ul>
        </li>
         <li><strong>Auto Width Issues:</strong> If auto-width seems too wide or narrow, try setting a manual width (e.g., "180px"). Very long headers in collapsible sections can sometimes affect auto-width calculation.</li>
        <li><strong>Other Errors:</strong> Check the message in the application's text box for specific error details or consult the console output if running from source.</li>
    </ul>
    """

# --- Helper Functions ---
def is_number(s):
    if s is None: retur
