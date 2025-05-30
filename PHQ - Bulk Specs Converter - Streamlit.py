# -*- coding: utf-8 -*-
import sys
import os
import pandas as pd
import math # Import math for isnan check
import traceback
from bs4 import BeautifulSoup
from datetime import datetime
import io # For BytesIO

import streamlit as st

# --- Instructions HTML (Copied from PyQt App) ---
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
        <li><b>Column A:</b> Contains the SKU for the product/package OR a numeric marker (1, 2, 3...) to indicate the start of a new Tab within a package. Leave blank for continuation rows, "Start", or "End" markers.</li>
        <li><b>Column B:</b>
            <ul>
                <li>If Column A is a <strong>numeric Tab marker</strong>: This column MUST contain the <strong>Tab Title</strong> (e.g., "Product 1 Details", "Canopy Specs").</li>
                <li>If Column A is a <strong>SKU</strong> or <strong>blank</strong> (and not a "Start"/"End" row): This is the first column for US/CA data (e.g., Specification Title, Care Header, Note). Ignore rows where Column A contains only 'US' or 'UK'.</li>
                <li>If the row is for a <strong>Collapsible Section Title</strong>: Put the title here (e.g., "Dimensions & Weights").</li>
                <li>If the row marks the <strong>Start</strong> of a collapsible section: This column contains the <em>first header</em> for the nested table (e.g., "Dimension").</li>
                 <li>If the row marks the <strong>End</strong> of a collapsible section: This column contains the <em>first data cell</em> for the last row of the nested table.</li>
           </ul>
        </li>
        <li><b>Column C onwards (for US/CA data):</b> Specification values, care details, notes content, or subsequent headers/data for collapsible sections.</li>
         <li><b>Column E onwards (for UK/AU/NZ data):</b> The corresponding data for these regions. Column E should align with Column B's purpose (Spec Title, Care Header, etc. for UK), Column F with Column C (Spec Value, Care Detail), and so on.</li>
         <li><b>Care Instructions:</b> Use headers like "Graphic Care Instructions", "Washing Instructions", etc., in the appropriate column (B for US, E for UK) to start a care section within a SKU or Tab. Subsequent rows list instructions or notes.</li>
         <li><b>Notes:</b> Start a cell in the relevant 'title' column (B for US, E for UK) with "Note:" (case-insensitive) to create a formatted note block.</li>
    </ul>

    <h3>B. Creating Tabs (for Package SKUs):</h3>
    <ol>
        <li>Enter the main Package SKU in Column A of the first row for that package. The content in Columns B, C... on this SKU row will be part of the first tab's data.</li>
        <li>For the first tab (if you want an explicit title different from "Details"), add a row below the SKU row:
            <ul><li>Enter '1' in Column A.</li><li>Enter the desired <strong>Tab Title</strong> (e.g., "Frame Details") in Column B.</li></ul>
             If you don't add a '1' marker, all data following the SKU row (until a '2' or new SKU) will be under an implicit "Details" tab or the first named tab if a '1' marker is used for the *very first set of data*.
        </li>
        <li>Add the specification rows, care instructions, notes, and any collapsible sections for this first tab below the Tab Title row (if used) or SKU row. Data for US/CA goes in columns B, C, ... and data for UK/AU/NZ goes in columns E, F, ... Rows starting with just 'US' or 'UK' in Column A are ignored.</li>
        <li>For the second tab, add a new row:
             <ul><li>Enter '2' in Column A.</li><li>Enter the <strong>Tab Title</strong> for the second tab (e.g., "Graphic Specs") in Column B.</li></ul>
        </li>
         <li>Add the data rows for the second tab below its title row.</li>
         <li>Repeat for any subsequent tabs (using '3', '4', etc. in Column A).</li>
         <li>The data rows following a tab marker belong to that tab until the next tab marker or a new SKU is encountered.</li>
    </ol>
     <p><b>Example Structure (Package SKU with Tabs):</b></p>
    <pre>
    | Col A    | Col B                 | Col C            | Col D | Col E (UK Title)     | Col F (UK Value) |
    |----------|-----------------------|------------------|-------|----------------------|------------------|
    | PKGSKU01 | Package Description   | Main Image URL   |       | Package Description  | Main Image URL   |
    | US       |                       |                  |       |                      |                  | <= This row ignored
    | 1        | Frame Specifications  |                  |       | Frame Specifications |                  |
    |          | Material              | Aluminum         |       | Material             | Aluminium        |
    |          | Weight                | 5 kg             |       | Weight               | 5 kg             |
    | 2        | Canopy Details        |                  |       | Canopy Details       |                  |
    |          | Fabric Type           | Polyester        |       | Fabric Type          | Polyester        |
    |          | Graphic Care...       | Wipe clean only  |       | Graphic Care...      | Wipe clean only  |
    |          | Note:                 | Handle with care |       | Note:                | Handle with care |
    </pre>
    <p><b>Example Structure (Single SKU, no explicit tabs - content from SKU row used):</b></p>
    <pre>
    | Col A    | Col B                 | Col C            | Col D | Col E (UK Title)     | Col F (UK Value) |
    |----------|-----------------------|------------------|-------|----------------------|------------------|
    | SKUABC   | Product Name Here     |                  |       | Product Name Here UK |                  |
    |          | Feature 1             | Value 1          |       | Feature 1 UK         | Value 1 UK       |
    |          | Feature 2             | Value 2          |       | Feature 2 UK         | Value 2 UK       |
    </pre>


    <h3>C. Creating Collapsible Sections (within a SKU or Tab):</h3>
    <p>These work the same way whether inside a tab or a regular SKU's data.</p>
    <ol>
        <li>Add a row with the main title for the section (e.g., "Dimensions & Weights") in the appropriate title column (B for US, E for UK).</li>
        <li>Immediately below this title row, add a row with:
            <ul>
                <li>"Start" (case-insensitive) in Column A.</li>
                <li>The <b>header row</b> data for the nested table starting from the appropriate title column (B for US, E for UK).</li>
            </ul>
        </li>
        <li>Add the subsequent data rows for the nested table (Column A empty, data starts in appropriate columns).</li>
        <li>Immediately after the last data row for the nested table, add a row with:
             <ul>
                <li>"End" (case-insensitive) in Column A.</li>
                <li>The data for the *last row* of the nested table starting from the appropriate title column.</li>
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
        <li><strong>No Output / Incorrect HTML / `No valid SKU data...` Error:</strong>
            <ul>
                <li>Double-check the Excel structure against the "Preparing Your Input" guide. Pay close attention to column usage for SKUs (Col A, must be first row for the product), Tab Markers (Col A, numeric), Tab Titles (Col B, on same row as Tab Marker), US data (B, C...), UK data (E, F...), "Start"/"End" markers (Col A), and Care Headers/Notes (Col B/E).</li>
                 <li>Ensure the *very first row* for each product/package has the SKU in Column A. Content in columns B, C... on this SKU row is considered part of the first tab's data.</li>
                 <li>Ensure Tab Markers (1, 2...) are numeric and in Column A, with Tab Titles immediately following in Column B.</li>
                <li>Verify "Start" is on the same row as the *headers* of the collapsible table, and "End" is on the same row as the *last data row* of that table.</li>
                <li>Check for hidden characters or extra spaces, especially in marker cells (A, B, E).</li>
                <li>Make sure specification data exists in the correct columns for US and UK regions.</li>
                <li>Rows with just 'US' or 'UK' in Column A should exist but are ignored for data processing.</li>
            </ul>
        </li>
         <li><strong>`NameError` in HTML Generation:</strong> This often points to an issue in the CSS generation within the Python code (like the 'content' error). Report this specific error.</li>
         <li><strong>Auto Width Issues:</strong> If auto-width seems too wide or narrow, try setting a manual width (e.g., "180px"). Very long headers in collapsible sections can sometimes affect auto-width calculation.</li>
        <li><strong>Other Errors:</strong> Check the message in the application's text box for specific error details or consult the console output if running from source.</li>
    </ul>
    """

# --- Helper Functions (from ConversionWorker, now standalone) ---
def is_number(s):
    """Checks if a value is numeric (int or float), handling strings."""
    if s is None: return False
    if isinstance(s, (int, float)): return not math.isnan(s)
    s_str = str(s).strip()
    if not s_str: return False
    try:
        val = float(s_str)
        return not math.isnan(val)
    except (ValueError, TypeError):
        return False

def process_cell(content, replace_newlines=True):
    """Processes cell content: converts to string, strips, handles newlines."""
    content_str = str(content).strip() if content is not None else ""
    if not content_str:
        return ""
    if not replace_newlines:
        return content_str
    lines = [line.strip() for line in content_str.split('\n') if line.strip()]
    return '<br>'.join(lines) if len(lines) > 1 else content_str

# --- Core HTML Generation Logic (from ConversionWorker, now standalone functions) ---
def generate_formatted_html_for_tab(raw_data_rows, region):
    """
    Generates HTML for specs, care, notes, and details for a SINGLE tab's data block.
    Args:
        raw_data_rows: List of lists, where each inner list is a row's cell values (strings).
        region: 'us' or 'uk'.
    Returns:
        Dictionary: {'specs_html': str, 'care_html': str, 'header_lengths': list}
    """
    if not raw_data_rows:
        return {'specs_html': '', 'care_html': '', 'header_lengths': []}

    title_col_idx = 1 if region == 'us' else 4
    value_cols_start_idx = 2 if region == 'us' else 5

    processed_block = []
    i = 0
    while i < len(raw_data_rows):
        row = raw_data_rows[i]
        first_cell_raw = row[0] if len(row) > 0 else ""
        first_cell_lower = str(first_cell_raw).strip().lower()

        if first_cell_lower == 'start' and processed_block:
            potential_trigger_row = processed_block.pop()
            if (isinstance(potential_trigger_row, list) and
                    len(potential_trigger_row) > title_col_idx and
                    str(potential_trigger_row[title_col_idx]).strip()):

                details_title = process_cell(potential_trigger_row[title_col_idx], False) 
                summary_text = f"Click to view" 

                details_header_row_raw = [row[idx] for idx in range(title_col_idx, len(row))]
                details_header_row = [process_cell(c, False) for c in details_header_row_raw if str(c).strip()] 

                details_data_rows = []
                data_row_idx = i + 1

                while data_row_idx < len(raw_data_rows):
                    current_data_row_list = raw_data_rows[data_row_idx]
                    marker_cell_raw = current_data_row_list[0] if len(current_data_row_list) > 0 else ""
                    marker_cell_str = str(marker_cell_raw).strip()

                    data_cells_raw = [current_data_row_list[idx] for idx in range(title_col_idx, len(current_data_row_list))]

                    if marker_cell_str.lower() == 'end':
                        details_data_rows.append([process_cell(c, True) for c in data_cells_raw]) 
                        i = data_row_idx 
                        break
                    if any(str(cell).strip() for cell in data_cells_raw):
                        details_data_rows.append([process_cell(c, True) for c in data_cells_raw]) 
                    data_row_idx += 1
                else: 
                    print(f"Warning: 'Start' found for '{details_title}' but no matching 'End' marker.")
                    processed_block.append(potential_trigger_row) 
                    i += 1
                    continue

                processed_block.append({
                    'type': 'details', 'label': details_title, 'summary': summary_text,
                    'header': details_header_row, 'data': details_data_rows
                })
                i += 1 
            else: 
                print(f"Warning: Found 'Start' marker at index {i} without a valid preceding title row for region '{region}'.")
                if potential_trigger_row: processed_block.append(potential_trigger_row)
                i += 1
        else: 
            processed_block.append(row)
            i += 1

    spec_sections = []
    current_spec_section_rows = []
    current_section_title = None
    care_instructions_html_parts = []
    list_open = False
    header_lengths = []
    care_instructions_started = False
    last_header = None
    current_td_contents = []
    section_notes = []

    for item in processed_block:
        if isinstance(item, dict) and item.get('type') == 'details':
            if last_header: 
                spec_row = f'<tr>\n<th class="th150" style="text-align: left;">{last_header}</th>\n<td>{("<br>".join(current_td_contents))}</td>\n</tr>'
                current_spec_section_rows.append(spec_row)
                last_header = None; current_td_contents = []

            details_html = '<details>\n'
            details_html += f'<summary>{item["summary"]}</summary>\n'
            has_nested_content = False

            if item['header']: 
                details_html += '<table>\n<thead>\n<tr>\n'
                header_classes = ["th-nested-1", "th-nested-2", "th-nested-3", "th-nested-4", "th-nested-5"]
                col_count = len(item['header']) 
                for idx, header_text in enumerate(item['header']):
                    css_class = header_classes[idx] if idx < len(header_classes) else ""
                    details_html += f'<th class="{css_class}">{header_text}</th>\n'
                details_html += '</tr>\n</thead>\n<tbody>\n'
                has_nested_content = True

                for data_row_cells in item['data']:
                     if any(str(cell).strip() for cell in data_row_cells):
                        details_html += '<tr>\n'
                        for cell_idx in range(col_count): 
                            cell_text = data_row_cells[cell_idx] if cell_idx < len(data_row_cells) else ""
                            details_html += f'<td>{cell_text}</td>\n'
                        details_html += '</tr>\n'
                details_html += '</tbody>\n</table>\n'
            elif item['data']: 
                details_html += '<table>\n<tbody>\n'
                has_nested_content = True
                for data_row_cells in item['data']:
                    if any(str(cell).strip() for cell in data_row_cells):
                        details_html += '<tr>\n'
                        for cell_text in data_row_cells:
                            details_html += f'<td>{cell_text}</td>\n' 
                        details_html += '</tr>\n'
                details_html += '</tbody>\n</table>\n'
            if not has_nested_content: 
                details_html += '<p style="margin-left: 20px; margin-top: 10px;">No details available.</p>\n'
            details_html += '</details>'

            details_label = item["label"]
            spec_row = f'<tr>\n<th class="th150" style="text-align: left;">{details_label}</th>\n<td>{details_html}</td>\n</tr>'
            current_spec_section_rows.append(spec_row)
            header_lengths.append(len(str(details_label)))
            last_header = None; current_td_contents = []
            continue
        elif isinstance(item, list):
            row_data = item # Renamed from 'row' to avoid conflict with outer scope 'row' in run_conversion_logic
            if not any(cell for cell in row_data): continue

            cell_title = process_cell(row_data[title_col_idx], False) if len(row_data) > title_col_idx else "" 
            cell_title_lower = cell_title.lower()
            cell_values_raw = row_data[value_cols_start_idx:] if len(row_data) > value_cols_start_idx else []
            has_value_content = any(str(v).strip() for v in cell_values_raw)

            care_headers = ["graphic care instructions", "washing instructions", "washing options",
                            "drying options", "removing wrinkles","care essentials","maintenance"]
            if cell_title_lower in care_headers or care_instructions_started:
                if last_header: 
                    spec_row_html = f'<tr>\n<th class="th150" style="text-align: left;">{last_header}</th>\n<td>{("<br>".join(current_td_contents))}</td>\n</tr>'
                    current_spec_section_rows.append(spec_row_html)
                    last_header = None; current_td_contents = []
                if not care_instructions_started: care_instructions_started = True

                if cell_title_lower in care_headers:
                    if list_open: care_instructions_html_parts.append("</ul>"); list_open = False
                    care_instructions_html_parts.append(f"<h3>{cell_title}</h3>")
                    list_open = True; care_instructions_html_parts.append("<ul>")
                    for val in cell_values_raw:
                        processed_val = process_cell(val, True) 
                        if processed_val:
                            for line in processed_val.split('<br>'):
                                if line: care_instructions_html_parts.append(f"<li>{line}</li>")
                elif cell_title_lower.startswith("note:"):
                    if list_open: care_instructions_html_parts.append("</ul>"); list_open = False
                    note_text = cell_title + " " + " ".join(filter(None, [str(v).strip() for v in cell_values_raw]))
                    if note_text.lower().startswith("note:"): note_text = note_text[5:].strip()
                    care_instructions_html_parts.append(f'<p class="note"><strong>Note:</strong> {process_cell(note_text)}</p>') 
                else: 
                    if not list_open: care_instructions_html_parts.append("<ul>"); list_open = True
                    full_instruction_text = cell_title + " " + " ".join(filter(None, [str(v).strip() for v in cell_values_raw]))
                    processed_instruction = process_cell(full_instruction_text, True) 
                    if processed_instruction:
                         for line in processed_instruction.split('<br>'):
                             if line: care_instructions_html_parts.append(f"<li>{line}</li>")
                continue
            else:
                if cell_title_lower.startswith("note:"):
                    if last_header: 
                        spec_row_html = f'<tr>\n<th class="th150" style="text-align: left;">{last_header}</th>\n<td>{("<br>".join(current_td_contents))}</td>\n</tr>'
                        current_spec_section_rows.append(spec_row_html)
                        last_header = None; current_td_contents = []
                    note_text = cell_title + " " + " ".join(filter(None, [str(v).strip() for v in cell_values_raw]))
                    if note_text.lower().startswith("note:"): note_text = note_text[5:].strip()
                    section_notes.append(f'<p class="note"><strong>Note:</strong> {process_cell(note_text)}</p>') 
                    continue

                is_section_title = bool(cell_title) and not has_value_content
                if is_section_title:
                    if last_header: 
                        spec_row_html = f'<tr>\n<th class="th150" style="text-align: left;">{last_header}</th>\n<td>{("<br>".join(current_td_contents))}</td>\n</tr>'
                        current_spec_section_rows.append(spec_row_html)
                    if current_spec_section_rows or current_section_title is not None or section_notes: 
                        spec_sections.append({'title': current_section_title, 'rows': current_spec_section_rows, 'notes': section_notes})
                        current_spec_section_rows = []; section_notes = []
                    current_section_title = cell_title
                    last_header = None; current_td_contents = []
                else: 
                    if cell_title: 
                        if last_header: 
                            spec_row_html = f'<tr>\n<th class="th150" style="text-align: left;">{last_header}</th>\n<td>{("<br>".join(current_td_contents))}</td>\n</tr>'
                            current_spec_section_rows.append(spec_row_html)
                        last_header = cell_title
                        header_lengths.append(len(last_header))
                        current_td_contents = [process_cell(v) for v in cell_values_raw if str(v).strip()] 
                    elif last_header: 
                        continuation_contents = [process_cell(v) for v in cell_values_raw if str(v).strip()] 
                        current_td_contents.extend(continuation_contents)
        else:
             print(f"Warning: Unexpected item type in processed_block: {type(item)}")

    if last_header:
        spec_row_html = f'<tr>\n<th class="th150" style="text-align: left;">{last_header}</th>\n<td>{("<br>".join(current_td_contents))}</td>\n</tr>'
        current_spec_section_rows.append(spec_row_html)
    if current_spec_section_rows or current_section_title is not None or section_notes:
         spec_sections.append({'title': current_section_title, 'rows': current_spec_section_rows, 'notes': section_notes})
    if list_open: care_instructions_html_parts.append("</ul>")

    specs_tab_html = ""
    if any(s.get('title') or s.get('rows') or s.get('notes') for s in spec_sections):
        specs_box_content = '<div class="productDetails">\n'
        for section in spec_sections:
             if section.get('title'): specs_box_content += f'<h3>{section["title"]}</h3>\n'
             if section.get('rows'):
                 specs_box_content += '<table class="productDetailsSection">\n<tbody>\n'
                 specs_box_content += '\n'.join(section['rows'])
                 specs_box_content += '\n</tbody>\n</table>\n'
             if section.get('notes'): specs_box_content += '\n'.join(section['notes']) + '\n'
        specs_box_content += '</div>' 
        specs_tab_html = f'<div class="newSpecificationBox specs-box">\n{specs_box_content}\n</div>'

    care_tab_html = ""
    filtered_care_parts = [part for part in care_instructions_html_parts if part.strip()]
    if filtered_care_parts:
        care_box_content = '<div class="productDetails">\n' + '\n'.join(filtered_care_parts) + '\n</div>'
        care_tab_html = f'<div class="newSpecificationBox care-box">\n{care_box_content}\n</div>'

    return {'specs_html': specs_tab_html, 'care_html': care_tab_html, 'header_lengths': header_lengths}

def generate_tabbed_html(tabs_data, region, auto_width_enabled, th150_width_input_value):
    """ Generates the complete HTML structure for tabs """
    if not tabs_data: return ""

    all_header_lengths = []
    tab_contents_html = []
    radio_buttons_html = []
    labels_html = []
    active_tab_ids = []

    for i, tab_info in enumerate(tabs_data):
        tab_id = f"tab{region}{i+1}"
        data_block = tab_info.get('data_rows', [])
        # Call the standalone generate_formatted_html_for_tab
        tab_result = generate_formatted_html_for_tab(data_block, region)

        if tab_result['specs_html'] or tab_result['care_html']:
            all_header_lengths.extend(tab_result['header_lengths'])
            active_tab_ids.append(tab_id)

            is_first_visible_tab = not radio_buttons_html
            radio_buttons_html.append(f'<input type="radio" id="{tab_id}" name="tabs{region}"{" checked" if is_first_visible_tab else ""}>')
            # Call standalone process_cell
            labels_html.append(f'<label for="{tab_id}">{process_cell(tab_info.get("title", f"Tab {i+1}"))}</label>')

            content_id = f"content{region}{i+1}"
            tab_content = f'<div class="tab-content" id="{content_id}">\n'
            tab_content += tab_result['specs_html'] + '\n' if tab_result['specs_html'] else ''
            tab_content += tab_result['care_html'] + '\n' if tab_result['care_html'] else ''
            tab_content += '</div>'
            tab_contents_html.append(tab_content)

    if not radio_buttons_html: return "<p>No specification data available for this product in this region.</p>"

    # Determine Width
    final_th150_width = '180px' # Default
    if auto_width_enabled:
         if all_header_lengths:
             try:
                 max_len = max(all_header_lengths)
                 min_width_px = 150; avg_char_px = 7.5; padding_allowance_px = 30
                 calculated_width = max(min_width_px, (max_len * avg_char_px) + padding_allowance_px)
                 final_th150_width = f'{int(round(calculated_width / 10.0)) * 10}px'
             except ValueError: final_th150_width = '200px' # Fallback if max fails (e.g. empty list)
         else: final_th150_width = '180px' # Fallback if no headers
    elif th150_width_input_value: # Use manual input if provided and auto_width is off
        final_th150_width = th150_width_input_value
        if not (final_th150_width.endswith('px') or final_th150_width.endswith('%')):
             print(f"Warning: Manual width '{final_th150_width}' might not be valid CSS. Using it anyway.")
    
    # CSS (th150_width replaced by final_th150_width)
    # Single tab content wrapper style
    single_tab_style = f"""<style>
    :root {{
        --primary-green: #2C3413;
        --secondary-neutral: #F2EFE4;
        --secondary-orange: #DE6E26;
        --secondary-sage: #707A5C;
        --light-sage: #dde2d4;
        --lighter-neutral: #f9f7f1;
        --border-sage: #b8bfab;
        --text-on-sage: #FFFFFF;
        --shadow-color: rgba(44, 52, 19, 0.15);
    }}

    * {{ font-family: nunitoregular, sans-serif; font-size: 14px; box-sizing: border-box; margin: 0; padding: 0; }}
    .content-wrapper {{ 
        border: 1px solid var(--border-sage);
        border-radius: 5px;
        padding: 25px 20px;
        background: var(--secondary-neutral);
        position: relative;
        width: 100%;
        clear: both;
    }}
    .newSpecificationBox {{
        box-shadow: 0 1px 3px var(--shadow-color);
        background-color: transparent; 
        margin-bottom: 25px;
        padding: 0;
        border: none;
        border-radius: 0;
    }}
    .newSpecificationBox.care-box {{
        padding: 0 5px; 
    }}
    .newSpecificationBox:last-child {{ margin-bottom: 0; border-bottom: 2px solid var(--light-sage); }}
    .productDetails {{ width: 100%; }}
    .productDetailsSection {{
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 20px;
    }}
    .productDetails > h3 + .productDetailsSection {{ margin-top: 0px; }}
    .productDetails > *:last-child {{ margin-bottom: 0 !important; }}
    .productDetailsSection tr:nth-child(odd) td, .productDetailsSection tr:nth-child(odd) th {{
        background-color: var(--lighter-neutral);
    }}
    .productDetailsSection tr:nth-child(even) td, .productDetailsSection tr:nth-child(even) th {{
        background-color: var(--secondary-neutral);
    }}
    .productDetailsSection th, .productDetailsSection td {{
        padding: 14px 18px;
        border-bottom: 1px solid var(--light-sage);
        text-align: left;
        vertical-align: top;
    }}
    .productDetailsSection tr:last-child th, .productDetailsSection tr:last-child td {{
        border-bottom: none;
    }}
    .th150 {{
        width: {final_th150_width}; 
        padding-right: 25px;
        font-weight: normal;
        vertical-align: top;
    }}
    h3 {{
        font-size: 14px;
        font-weight: bold;
        margin-top: 30px;
        margin-bottom: 12px;
        color: var(--primary-green);
    }}
    .productDetails > h3:first-child {{ margin-top: 0; }}
    .care-box h3 {{ margin-top: 0; }}
    ul {{ margin: 0 0 15px 0; padding-left: 25px; list-style: disc; }}
    li {{ margin-bottom: 6px; line-height: 1.5; }}
    p.note {{
        margin-top: 15px; margin-bottom: 15px;
        font-style: italic;
        color: var(--primary-green);
        background-color: var(--lighter-neutral);
        padding: 12px 15px;
        border-left: 4px solid var(--secondary-sage);
    }}
    .productDetails > p.note:first-child {{ margin-top: 0; }}
    summary {{
        cursor: pointer;
        display: inline-block;
        padding: 5px 10px;
        border-radius: 4px;
        background-color: var(--light-sage);
        border: 1px solid var(--border-sage);
        margin-bottom: 8px;
        margin-top: -8px;
        font-weight: normal;
        transition: background-color 0.2s ease, color 0.2s ease;
        color: var(--primary-green);
    }}
    summary:hover {{
        background-color: var(--secondary-sage);
        color: var(--text-on-sage);
    }}
    summary::marker {{ display: none; content: ""; }}
    details {{ margin-top: 5px; }}
    details[open] > summary {{ margin-bottom: 10px; }}
    details > table {{
        margin-top: 10px;
        width: 98%;
        max-width: 700px;
        border-collapse: collapse;
        margin-left: 5px;
        border: 1px solid var(--border-sage);
        font-size: 13px;
    }}
    details > table th, details > table td {{
        border: 1px solid var(--light-sage);
        padding: 8px 10px;
        text-align: left;
        vertical-align: middle;
        background-color: var(--secondary-neutral);
    }}
    details > table th {{
        background-color: var(--light-sage);
        color: var(--primary-green);
        font-weight: bold;
        border-bottom: 2px solid var(--border-sage);
    }}
    details > table tbody tr:nth-child(even) td {{
        background-color: var(--lighter-neutral);
    }}
    a {{
        color: var(--secondary-orange);
        text-decoration: none;
    }}
    a:hover {{
        color: var(--primary-green);
        text-decoration: underline;
    }}
</style>"""

    if len(active_tab_ids) == 1: # If only one tab has content, use simpler wrapper
        # For single tab, tab_contents_html[0] already includes <div class="tab-content"...
        # We need to wrap it with content-wrapper and the overall style.
        # The single_tab_style is already defined. We just need to make sure the content is placed correctly.
        # tab_contents_html[0] contains: <div class="tab-content" id="contentregion1"> ... </div>
        # We want: <div class="content-wrapper"> <div class="newSpecificationBox..."> ... </div> <div class="newSpecificationBox care-box..."> ... </div> </div>
        # The current structure of tab_contents_html[0] IS:
        # <div class="tab-content" id="contentregion1">
        #    <div class="newSpecificationBox specs-box">...</div>  <-- From generate_formatted_html_for_tab
        #    <div class="newSpecificationBox care-box">...</div>   <-- From generate_formatted_html_for_tab
        # </div>
        # So, we need to extract the *inner content* of tab_contents_html[0] if we use .content-wrapper as the sole container.
        # Or, we can just embed tab_contents_html[0] directly if .content-wrapper styles are compatible with .tab-content existing there.
        # Let's keep it simple: for a single active tab, we use the .content-wrapper and place the combined specs/care HTML into it.
        
        single_tab_inner_html = ""
        # Reconstruct the inner content similar to how it's done for multi-tabs but without the .tab-content div
        tab_info_single = tabs_data[0] # Assuming the first tab in tabs_data is the one that's active
        # Find the first tab_info that actually has content (matching active_tab_ids logic)
        first_active_tab_index = -1
        for i, t_info in enumerate(tabs_data):
            temp_tab_id = f"tab{region}{i+1}"
            if temp_tab_id == active_tab_ids[0]:
                first_active_tab_index = i
                break
        
        if first_active_tab_index != -1:
            tab_info_single = tabs_data[first_active_tab_index]
            data_block_single = tab_info_single.get('data_rows', [])
            tab_result_single = generate_formatted_html_for_tab(data_block_single, region) # regenerate for clarity
            if tab_result_single['specs_html']:
                single_tab_inner_html += tab_result_single['specs_html'] + '\n'
            if tab_result_single['care_html']:
                single_tab_inner_html += tab_result_single['care_html'] + '\n'
        
        return single_tab_style + '\n\n<div class="content-wrapper">\n' + single_tab_inner_html.strip() + '\n</div>'


    tab_content_selectors = []
    tab_label_selectors = []
    for tab_id in active_tab_ids:
         content_id = tab_id.replace('tab', 'content')
         tab_content_selectors.append(f'#{tab_id}:checked ~ #{content_id}')
         tab_label_selectors.append(f'#{tab_id}:checked ~ label[for="{tab_id}"]')

    multi_tab_style = f"""<style>
    :root {{
        --primary-green: #2C3413;
        --secondary-neutral: #F2EFE4;
        --secondary-orange: #DE6E26;
        --secondary-sage: #707A5C;
        --light-sage: #dde2d4; 
        --lighter-neutral: #f9f7f1; 
        --border-sage: #b8bfab; 
        --text-on-sage: #FFFFFF; 
        --shadow-color: rgba(44, 52, 19, 0.15); 
    }}

    * {{ font-family: nunitoregular, sans-serif; font-size: 14px; box-sizing: border-box; margin: 0; padding: 0; }}
    .tabs {{ width: 100%; margin-bottom: 20px; position: relative; clear: both; }}
    .tabs input[type="radio"] {{ display: none; }}
    .tabs label {{
        display: inline-block;
        padding: 10px 18px;
        background: var(--light-sage);
        border: 1px solid var(--border-sage);
        color: var(--primary-green);
        border-bottom: none;
        border-radius: 5px 5px 0 0;
        margin-top: 10px;
        margin-right: 3px;
        margin-left: 3px;
        margin-bottom: -1px;
        cursor: pointer;
        font-weight: bold;
        position: relative;
        z-index: 1;
        transition: background-color 0.2s ease, color 0.2s ease;
    }}
    .tabs label:hover {{
        background-color: var(--secondary-sage);
        color: var(--text-on-sage);
    }}
    .tabs .tab-content {{
        display: none;
        border: 1px solid var(--border-sage);
        border-radius: 0 5px 5px 5px;
        padding: 25px 20px;
        background: var(--secondary-neutral);
        position: relative;
        width: 100%;
        clear: both;
        margin-top: 0;
    }}
    {', '.join(tab_content_selectors)} {{ display: block; }}
    {', '.join(tab_label_selectors)} {{
        background: var(--secondary-neutral);
        border-bottom: 1px solid var(--secondary-neutral);
        border-top: 2px solid var(--secondary-orange);
        padding-top: 9px; 
        color: var(--secondary-orange);
        z-index: 2;
    }}
    .newSpecificationBox {{
        box-shadow: 0 1px 3px var(--shadow-color);
        background-color: transparent; 
        margin-bottom: 25px;
        padding: 0;
        border: none;
        border-radius: 0;
    }}
    .newSpecificationBox.care-box {{
        padding: 0 5px;
    }}
    .newSpecificationBox:last-child {{ margin-bottom: 0; border-bottom: 2px solid var(--light-sage); }}
    .productDetails {{ width: 100%; }}
    .productDetailsSection {{
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 20px;
    }}
    .productDetails > h3 + .productDetailsSection {{ margin-top: 0px; }}
    .productDetails > *:last-child {{ margin-bottom: 0 !important; }}
    .productDetailsSection tr:nth-child(odd) td, .productDetailsSection tr:nth-child(odd) th {{
        background-color: var(--lighter-neutral);
    }}
    .productDetailsSection tr:nth-child(even) td, .productDetailsSection tr:nth-child(even) th {{
        background-color: var(--secondary-neutral);
    }}
    .productDetailsSection th, .productDetailsSection td {{
        padding: 14px 18px;
        border-bottom: 1px solid var(--light-sage);
        text-align: left;
        vertical-align: top;
    }}
     .productDetailsSection tr:last-child th, .productDetailsSection tr:last-child td {{
        border-bottom: none;
    }}
    .th150 {{
        width: {final_th150_width};
        padding-right: 25px;
        font-weight: normal;
        vertical-align: top;
    }}
    h3 {{
        font-size: 14px;
        font-weight: bold;
        margin-top: 30px;
        margin-bottom: 12px;
        color: var(--primary-green);
    }}
    .productDetails > h3:first-child {{ margin-top: 0; }}
    .care-box h3 {{ margin-top: 0; }}
    ul {{ margin: 0 0 15px 0; padding-left: 25px; list-style: disc; }}
    li {{ margin-bottom: 6px; line-height: 1.5; }}
    p.note {{
        margin-top: 15px; margin-bottom: 15px;
        font-style: italic;
        color: var(--primary-green);
        background-color: var(--lighter-neutral);
        padding: 12px 15px;
        border-left: 4px solid var(--secondary-sage);
    }}
    .productDetails > p.note:first-child {{ margin-top: 0; }}
    summary {{
        cursor: pointer;
        display: inline-block;
        padding: 5px 10px;
        border-radius: 4px;
        background-color: var(--light-sage);
        border: 1px solid var(--border-sage);
        margin-bottom: 8px;
        margin-top: -8px;
        font-weight: normal;
        transition: background-color 0.2s ease, color 0.2s ease;
        color: var(--primary-green);
    }}
    summary:hover {{
        background-color: var(--secondary-sage);
        color: var(--text-on-sage);
    }}
    summary::marker {{ display: none; content: ""; }}
    details {{ margin-top: 5px; }}
    details[open] > summary {{ margin-bottom: 10px; }}
    details > table {{
        margin-top: 10px;
        width: 98%;
        max-width: 700px;
        border-collapse: collapse;
        margin-left: 5px;
        border: 1px solid var(--border-sage);
        font-size: 13px;
    }}
    details > table th, details > table td {{
        border: 1px solid var(--light-sage);
        padding: 8px 10px;
        text-align: left;
        vertical-align: middle;
        background-color: var(--secondary-neutral);
    }}
    details > table th {{
        background-color: var(--light-sage);
        color: var(--primary-green);
        font-weight: bold;
        border-bottom: 2px solid var(--border-sage);
    }}
    details > table tbody tr:nth-child(even) td {{
         background-color: var(--lighter-neutral);
    }}
    a {{
        color: var(--secondary-orange);
        text-decoration: none;
    }}
    a:hover {{
        color: var(--primary-green);
        text-decoration: underline;
    }}
</style>"""

    html_output = multi_tab_style + '\n\n'
    html_output += '<div class="tabs">\n'
    html_output += '    <!-- Tab Radio Buttons (Hidden) -->\n'
    html_output += '    ' + '\n    '.join(radio_buttons_html) + '\n\n'
    html_output += '    <!-- Tab Labels -->\n'
    html_output += '    ' + '\n    '.join(labels_html) + '\n\n'
    html_output += '    <!-- Tab Content Panes -->\n'
    html_output += '    ' + '\n    '.join(tab_contents_html) + '\n'
    html_output += '</div> <!-- end tabs -->\n'

    try:
        soup = BeautifulSoup(html_output, 'html.parser')
        pretty_html = soup.prettify(formatter="minimal")
        pretty_html = '\n'.join(line for line in pretty_html.split('\n') if line.strip())
        return pretty_html
    except Exception as e:
        print(f"HTML parsing/prettifying error: {e}. Returning raw HTML.")
        return html_output


def run_conversion_logic(input_file_buffer, input_filename_for_output, th150_width_manual, auto_width_enabled, progress_bar, status_area):
    """
    Core conversion logic.
    Returns a tuple (output_dataframe, error_message_string)
    """
    try:
        df = pd.read_excel(input_file_buffer, header=None, na_filter=False)
        df = df.applymap(lambda x: str(x).strip())
    except Exception as e:
        err_msg = f"Error reading Excel file: {str(e)}. Ensure it's closed and not corrupted."
        status_area.error(err_msg)
        return None, err_msg

    output_rows = []
    total_rows = len(df)
    progress_bar.progress(0)

    current_sku = None
    current_sku_tabs_data = []  # List of dicts: [{'title': str, 'data_rows': list_of_lists}, ...]
    current_tab_rows = []       # Rows for the *current* tab being processed

    for index, row_series in df.iterrows():
        if index % 10 == 0 or index == total_rows - 1:
            progress_bar.progress(int((index + 1) / total_rows * 100))

        # Current row being processed
        current_processing_row = [str(cell).strip() for cell in row_series.tolist()]
        first_cell_value = current_processing_row[0] if len(current_processing_row) > 0 else ""

        if first_cell_value.upper() in ['US', 'UK'] and not any(current_processing_row[1:]):
            continue # Skip placeholder US/UK rows

        is_tab_marker_str = str(first_cell_value)
        is_tab_marker = is_number(is_tab_marker_str)
        is_start_end_marker = first_cell_value.lower() in ['start', 'end']
        
        is_potential_new_sku_row = bool(first_cell_value) and not is_tab_marker and not is_start_end_marker
        
        is_new_sku = False
        if is_potential_new_sku_row:
            if current_sku is None: # First SKU in the file
                is_new_sku = True
            elif first_cell_value != current_sku: # New SKU encountered
                is_new_sku = True

        if is_new_sku:
            # 1. Process and finalize the PREVIOUS SKU's data (if exists)
            if current_sku is not None:
                if current_tab_rows: # Add any remaining rows to the last tab of the previous SKU
                    if not current_sku_tabs_data: # Should ideally not happen if logic is right, implies data without tab
                        current_sku_tabs_data.append({'title': 'Details', 'data_rows': current_tab_rows})
                    else:
                        current_sku_tabs_data[-1]['data_rows'].extend(current_tab_rows)
                
                if current_sku_tabs_data: # If previous SKU had any tab data
                    try:
                        us_html = generate_tabbed_html(current_sku_tabs_data, 'us', auto_width_enabled, th150_width_manual)
                        uk_html = generate_tabbed_html(current_sku_tabs_data, 'uk', auto_width_enabled, th150_width_manual)
                        output_rows.extend([
                            [current_sku, 'default', us_html]
                        ])
                    except Exception as e:
                        error_details = traceback.format_exc()
                        err_msg = f"Error generating HTML for SKU '{current_sku}': {str(e)}\nDetails:\n{error_details}"
                        status_area.error(err_msg); print(err_msg) # Log and show error
                else:
                    print(f"Info: Previous SKU '{current_sku}' had no processable tab data to finalize.")
            
            # 2. Initialize for the NEW SKU
            current_sku = first_cell_value
            current_sku_tabs_data = [] # Reset tabs for the new SKU
            current_tab_rows = []      # Reset rows for the first tab of the new SKU
            
            # 3. The NEW SKU row itself is skipped - we don't add its content to current_tab_rows
            # This matches v1's behavior where SKU row content is not included in the tab data
            continue 

        # --- This part executes if it's NOT a new SKU row, but we ARE under an active SKU ---
        if current_sku is not None: 
            if is_tab_marker:
                # A new tab is starting.
                # 1. Finalize data for the *previous* tab section
                if current_tab_rows: # If there were rows collected before this tab marker
                    if not current_sku_tabs_data: 
                        # This is the first *explicit* tab marker for the current SKU.
                        # Create a default "Details" tab for any rows collected before this marker
                        current_sku_tabs_data.append({'title': 'Details', 'data_rows': current_tab_rows})
                    else: 
                        # `current_tab_rows` belong to the previously defined tab in `current_sku_tabs_data`.
                        current_sku_tabs_data[-1]['data_rows'].extend(current_tab_rows)
                        # Now, create the NEW tab entry from the current marker row.
                        tab_title_val = current_processing_row[1] if len(current_processing_row) > 1 and current_processing_row[1] else f"Tab {int(float(first_cell_value))}"
                        current_sku_tabs_data.append({'title': tab_title_val, 'data_rows': []}) # New tab starts empty
                else: # No `current_tab_rows` before this marker; just start a new tab.
                      # This happens if a SKU row is immediately followed by a tab marker,
                      # or if a tab marker immediately follows another tab marker.
                    tab_title_val = current_processing_row[1] if len(current_processing_row) > 1 and current_processing_row[1] else f"Tab {int(float(first_cell_value))}"
                    current_sku_tabs_data.append({'title': tab_title_val, 'data_rows': []}) 
                
                current_tab_rows = [] # Reset `current_tab_rows` to collect data for this newly defined tab.
                # The tab marker row itself (e.g., "1 | Tab Title") does not contribute spec data, so `current_processing_row` is not added here.
            else: # This row is regular data (spec, Start/End, note, care, or continuation) for the current tab.
                is_meaningful_data_row = any(cell_content for cell_content in current_processing_row[1:]) or \
                                         first_cell_value.lower() in ['start', 'end'] or \
                                         (bool(first_cell_value) and not is_tab_marker and not is_start_end_marker) # handles blank Col A but with data in B,C...
                                                                                                                     # or Col A non-empty and not a marker
                if is_meaningful_data_row or (not first_cell_value and any(current_processing_row[1:])): # Add if Col A has content OR (Col A is blank AND Col B+ has content)
                    current_tab_rows.append(current_processing_row)
        # else: current_sku is None. This row must be the very first SKU row (or malformed data before any SKU).
        #       The `is_new_sku` block should handle the first SKU row correctly by setting `current_sku`.


    # After loop, process the very last SKU's collected data
    if current_sku is not None:
         if current_tab_rows: # If there are any remaining rows for the last tab/SKU
             if not current_sku_tabs_data: # No explicit tabs were defined for this SKU
                 current_sku_tabs_data.append({'title': 'Details', 'data_rows': current_tab_rows})
             else: # Add to the last defined tab of the current SKU
                 current_sku_tabs_data[-1]['data_rows'].extend(current_tab_rows)
         
         if current_sku_tabs_data: # If there's any tab data to process for the SKU
            try:
                us_html = generate_tabbed_html(current_sku_tabs_data, 'us', auto_width_enabled, th150_width_manual)
                uk_html = generate_tabbed_html(current_sku_tabs_data, 'uk', auto_width_enabled, th150_width_manual)
                output_rows.extend([
                    [current_sku, 'default', us_html],
                    [current_sku, 'canada', us_html],
                    [current_sku, 'unitedkingdom', uk_html],
                    [current_sku, 'australia', uk_html],
                    [current_sku, 'newzealand', uk_html]
                ])
            except Exception as e:
                error_details = traceback.format_exc()
                err_msg = f"Error generating HTML for last SKU '{current_sku}': {str(e)}\nDetails:\n{error_details}"
                status_area.error(err_msg); print(err_msg)
         else:
             print(f"Info: Last SKU '{current_sku}' had no processable tab data upon loop completion.")


    if not output_rows:
         err_msg = ("Conversion finished, but NO valid SKU data resulted in HTML output.\n"
                    "Please check:\n"
                    "- Did the input file contain SKUs in Column A for each product group?\n"
                    "- Was data present in the expected US/UK columns (B/C+, E/F+)?\n"
                    "- Was the Excel sheet structure correct per instructions (SKU first, then optional tab markers)?")
         status_area.warning(err_msg)
         return None, err_msg 

    output_df = pd.DataFrame(output_rows, columns=['SKU', 'Region', 'HTML'])
    progress_bar.progress(100)
    return output_df, None


# --- Streamlit Application UI ---
def main():
    st.set_page_config(page_title="Specs HTML Converter", layout="wide")

    # Logo - assuming VO-Logo.png is in the same directory as the script
    logo_path = "VO-Logo.png"
    if os.path.exists(logo_path):
        # Use columns to place logo to the right
        col1, col2 = st.columns([4,1])
        with col1:
            st.title("PatioHQ Specs HTML - Tabs & Dropdowns")
        with col2:
            st.image(logo_path, width=113)
    else:
        st.title("PatioHQ Specs HTML Converter (Tabs & Dropdowns)")
        st.caption("Logo (VO-Logo.png) not found.")


    with st.expander("Help / Instructions", expanded=False):
        st.markdown(get_instructions_html(), unsafe_allow_html=True)

    st.subheader("1. Upload Excel File")
    uploaded_file = st.file_uploader("Choose an .xlsx file", type="xlsx")

    st.subheader("2. Configure Settings")
    col1, col2 = st.columns(2)
    with col1:
        auto_width_checkbox = st.checkbox("Auto width for Spec Header", value=True,
                                          help="Automatically adjust first column width based on content (Recommended).")
    with col2:
        th150_width_input = st.text_input("Manual Spec Header Width (if Auto unchecked)",
                                          placeholder="e.g., 180px",
                                          help="Enter manual width for the first column (e.g., '180px'). Overridden by 'Auto width'.",
                                          disabled=auto_width_checkbox)

    st.subheader("3. Convert")
    convert_button = st.button("Convert to HTML")

    status_area = st.empty() # For messages like errors or warnings
    progress_bar = st.progress(0)

    if convert_button:
        if uploaded_file is not None:
            input_filename = uploaded_file.name
            status_area.info(f"Starting conversion for: {input_filename}...")
            progress_bar.progress(0) # Reset progress bar

            # Get width settings
            manual_width_val = th150_width_input if not auto_width_checkbox else ""
            if not auto_width_checkbox and manual_width_val:
                 if not (manual_width_val.endswith('px') or manual_width_val.endswith('%')):
                     st.warning(f"Manual width '{manual_width_val}' does not end with 'px' or '%'. The converter will attempt to use it as is.")
            
            try:
                # Renamed 'row' in the loop within run_conversion_logic to current_processing_row
                output_df, error_msg = run_conversion_logic(
                    uploaded_file,
                    input_filename,
                    manual_width_val,
                    auto_width_checkbox,
                    progress_bar,
                    status_area 
                )

                if error_msg and output_df is None : 
                    st.error(f"Conversion failed. See details above or console log.")
                elif error_msg and output_df is not None: # Errors occurred but some output might exist
                    st.warning(f"Conversion completed with some issues. Please review output and messages above.")
                    # Proceed to show download and preview if output_df exists
                    # (This part is the same as the success case below, just with a warning first)
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        output_df.to_excel(writer, index=False)
                    output_buffer.seek(0)
                    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename_base = os.path.splitext(input_filename)[0]
                    download_filename = f"{output_filename_base}_output_{current_time}.xlsx"
                    st.download_button(
                        label="Download Output Excel File (with potential issues)",
                        data=output_buffer,
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.markdown("---")
                    st.markdown("### Preview of Generated HTML (first 5 rows - review carefully due to warnings):")
                    preview_df = output_df[['SKU', 'Region']].copy()
                    preview_df['HTML_Preview (Scrollable)'] = output_df['HTML'].apply(
                        lambda x: f'<div style="max-height: 200px; overflow-y: auto; border: 1px solid #eee; padding: 5px; background-color: #f9f9f9;">{x[:2000]}{"..." if len(x)>2000 else ""}</div>'
                    )
                    st.markdown(preview_df.head().to_html(escape=False, index=False), unsafe_allow_html=True)
                elif output_df is not None: # Success
                    status_area.success("Conversion complete!")
                    
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        output_df.to_excel(writer, index=False)
                    output_buffer.seek(0)

                    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename_base = os.path.splitext(input_filename)[0]
                    download_filename = f"{output_filename_base}_output_{current_time}.xlsx"

                    st.download_button(
                        label="Download Output Excel File",
                        data=output_buffer,
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.markdown("---")
                    st.markdown("### Preview of Generated HTML (first 5 rows):")
                    
                    preview_df = output_df[['SKU', 'Region']].copy()
                    preview_df['HTML_Preview (Scrollable)'] = output_df['HTML'].apply(
                        lambda x: f'<div style="max-height: 200px; overflow-y: auto; border: 1px solid #eee; padding: 5px; background-color: #f9f9f9;">{x[:2000]}{"..." if len(x)>2000 else ""}</div>'
                    )
                    st.markdown(preview_df.head().to_html(escape=False, index=False), unsafe_allow_html=True)

                else: 
                    status_area.error("An unexpected issue occurred. No output generated and no specific error message.")


            except Exception as e:
                error_details = traceback.format_exc()
                status_area.error(f"A critical error occurred: {str(e)}\n\nTraceback:\n{error_details}")
                progress_bar.progress(0)

        else:
            status_area.warning("Please upload an Excel file first.")
            progress_bar.progress(0)
            
    st.markdown("---")
    st.markdown("<p style='text-align: center; color: gray;'>Developed by Mohit Dhaker © 2025</p>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
