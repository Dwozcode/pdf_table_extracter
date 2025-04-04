# PDF Table Extractor

## Project Overview

This tool detects and extracts tables from system-generated PDFs and saves them to Excel files. The solution handles various table types including bordered tables, borderless tables, and irregularly shaped tables without using pre-built table extraction tool.
## Features

- **Multi-method table detection**:
  - Bordered table detection using line analysis
  - Borderless table detection using text alignment patterns
  - Handles edge cases like merged cells and multi-line content
  
- **Accurate data extraction**:
  - Preserves table structure (rows and columns)
  - Maintains cell content integrity
  
- **Excel output**:
  - Each table is saved to a separate worksheet
  - Worksheets are named based on page number and table index
  - Basic formatting is applied for readability

## Installation

### Prerequisites

- Python 3.6+
- pip (Python package installer)

### Required Libraries

```
pdfplumber
pandas
openpyxl
```

### Installation Steps

1. Clone or download this repository
2. Install required libraries:

```bash
pip install pdfplumber pandas openpyxl
```

## Usage

Run the script from the command line with the following syntax:

```bash
python pdf_table_extractor.py [input_pdf_file] [output_excel_file]
```

Example:

```bash
python pdf_table_extractor.py sample.pdf extracted_tables.xlsx
```

## How It Works

The tool uses a multi-layered approach to identify and extract tables:

1. **PDF Processing**: Uses `pdfplumber` to extract text, layout, and line information from the PDF.

2. **Table Detection**:
   - First attempts to detect tables with borders by analyzing horizontal and vertical lines
   - If no bordered tables are found, uses text alignment patterns to detect borderless tables
   - Creates cell boundaries based on the detected structure

3. **Content Extraction**:
   - Maps text elements to their appropriate cells in the table structure
   - Handles special cases like text spanning across multiple cells

4. **Excel Creation**:
   - Converts extracted data to pandas DataFrames
   - Writes each table to a separate sheet in the Excel file
   - Names sheets based on page number and table index

## Algorithm Details

### Bordered Table Detection

- Identifies horizontal and vertical lines in the PDF
- Creates a grid based on line intersections
- Maps text elements to the resulting cells

### Borderless Table Detection

- Groups text elements based on vertical positioning (rows)
- Analyzes horizontal positioning patterns to determine columns
- Creates a virtual grid and maps text to appropriate cells

### Edge Case Handling

- Word clustering with tolerance parameters to handle text alignment variations
- Cell boundary analysis to account for merged cells
- Post-processing to ensure consistent column counts across rows

## Limitations

- Very complex table layouts with nested tables may not be detected accurately
- Tables spanning multiple pages are treated as separate tables
- PDF documents with non-standard fonts or heavily formatted text may yield imperfect results

## Future Improvements

- Add support for tables spanning multiple pages
- Implement more advanced merged cell detection
- Add a user interface for easier operation
- Improve handling of complex nested tables
- Add options for customizing output formatting

## Project Structure

```
pdf-table-extractor/
├── pdf_table_extractor.py     # Main script
├── README.md                  # Project documentation
├── sample/                    # Sample files
│   ├── sample1.pdf            # Sample PDF with bordered tables
│   ├── sample2.pdf            # Sample PDF with borderless tables
│   └── sample3.pdf            # Sample PDF with irregular tables
└── output/                    # Example outputs
    ├── sample1_tables.xlsx    # Extracted tables from sample1.pdf
    ├── sample2_tables.xlsx    # Extracted tables from sample2.pdf
    └── sample3_tables.xlsx    # Extracted tables from sample3.pdf
```

## Testing Methodology

The tool was tested on a variety of PDFs including:
1. Simple documents with standard bordered tables
2. Financial reports with borderless tables
3. Technical documents with irregularly shaped tables
4. Documents with mixed table formats

Accuracy was validated by comparing extracted data with the original PDF tables.

## Performance Considerations

- Processing time primarily depends on PDF complexity and size
- For large PDFs with many tables, processing may take several minutes
- Memory usage scales with PDF size and number of tables

## Hackathon Solution Approach

This project was developed to address the challenge of extracting tables from PDFs without using common tools like Tabula or Camelot. The key innovations in our approach include:

1. **Multi-method detection strategy** - Using different algorithms depending on table type
2. **Content-aware cell mapping** - Intelligent assignment of text to table cells
3. **Robust post-processing** - Cleaning and normalizing extracted data for consistency

Our solution demonstrates that effective table extraction is possible without specialized table extraction libraries or image conversion techniques.

## Presentation Key Points

- Demonstration with different table types (bordered, borderless, irregular)
- Comparison of extraction accuracy against manual extraction
- Discussion of technical challenges and solutions
- Performance metrics on various PDF types
- Potential real-world applications

---

This project was developed for the PDF Table Extraction=.
