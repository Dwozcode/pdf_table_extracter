import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import argparse

def detect_bordered_tables(page):
    """Detect tables using lines (borders)"""
    tables = []
    # Get horizontal and vertical lines
    h_lines = [line for line in page.lines if line["height"] < 1]
    v_lines = [line for line in page.lines if line["width"] < 1]
    
    if not h_lines or not v_lines:
        return []
    
    # Find bounding boxes formed by lines
    x_coords = sorted({line["x0"] for line in v_lines} | {line["x1"] for line in v_lines})
    y_coords = sorted({line["top"] for line in h_lines} | {line["bottom"] for line in h_lines})
    
    # Create cell boundaries
    cells = []
    for i in range(len(y_coords)-1):
        row = []
        for j in range(len(x_coords)-1):
            left = x_coords[j]
            right = x_coords[j+1]
            top = y_coords[i]
            bottom = y_coords[i+1]
            row.append((left, top, right, bottom))
        cells.append(row)
    
    # Map words to cells
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    table_data = []
    for row in cells:
        row_data = []
        for cell in row:
            cell_text = ""
            for word in words:
                if (cell[0] <= word["x0"] <= cell[2] and
                    cell[1] <= word["top"] <= cell[3]):
                    cell_text += f"{word['text']} "
            row_data.append(cell_text.strip())
        table_data.append(row_data)
    
    if table_data:
        tables.append(table_data)
    return tables

def detect_borderless_tables(page):
    """Detect tables using text alignment (no borders)"""
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    if not words:
        return []
    
    # Cluster words into rows based on y-coordinate
    rows = {}
    for word in words:
        y = round(word["top"], 1)  # Round to group nearby lines
        if y not in rows:
            rows[y] = []
        rows[y].append(word)
    
    # Sort rows by y-position
    sorted_y = sorted(rows.keys())
    table_data = []
    column_positions = []  # Track typical x-positions of columns
    
    for y in sorted_y:
        row_words = sorted(rows[y], key=lambda x: x["x0"])
        row_text = []
        current_col = 0
        
        for word in row_words:
            # If we don't have column positions yet, initialize them
            if not column_positions:
                column_positions.append((word["x0"], word["x1"]))
                row_text.append(word["text"])
                continue
                
            # Find which column this word belongs to
            matched = False
            for i, (col_start, col_end) in enumerate(column_positions):
                if abs(word["x0"] - col_start) < 10:  # 10px tolerance
                    # If we're skipping columns, fill blanks
                    while current_col < i:
                        row_text.append("")
                        current_col += 1
                    row_text.append(word["text"])
                    current_col += 1
                    matched = True
                    break
                    
            if not matched:
                # New column found
                column_positions.append((word["x0"], word["x1"]))
                row_text.append(word["text"])
                current_col += 1
                
        table_data.append(row_text)
    
    return [table_data] if table_data else []

def extract_tables(pdf_path):
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            # Try bordered tables first
            tables = detect_bordered_tables(page)
            if not tables:
                tables = detect_borderless_tables(page)
            
            for table in tables:
                all_tables.append({
                    "page": page_num+1,
                    "data": table
                })
    return all_tables

def save_to_excel(tables, output_file):
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for idx, table in enumerate(tables):
            df = pd.DataFrame(table["data"])
            sheet_name = f"Page_{table['page']}_Table_{idx+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF Table Extractor")
    parser.add_argument("pdf_file", help="Path to input PDF file")
    parser.add_argument("output_file", help="Path to output Excel file")
    args = parser.parse_args()

    tables = extract_tables(args.pdf_file)
    save_to_excel(tables, args.output_file)
    print(f"Extracted {len(tables)} tables to {args.output_file}")