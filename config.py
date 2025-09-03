"""
Configuration file for Excel to Word Table Extractor
Modify these settings to specify which data to extract from your Excel file.
"""

# Input Excel file configuration
EXCEL_FILE = "excel-data/data.xlsx"   # Path to your Excel file
SHEET_NAME = None                     # Sheet name (None = active sheet, or specify like "Sheet1")

# Data extraction range
START_ROW = 1                         # Starting row number
END_ROW = 10                          # Ending row number (inclusive)
START_COL = "A"                       # Starting column letter (A, B, C, etc. or AA, AB, etc.)
END_COL = "E"                         # Ending column letter (inclusive)

# Output configuration
OUTPUT_FILE = "docx-output/extracted_data.docx"  # Output Word document path
DOCUMENT_TITLE = "Extracted Excel Data"  # Title that appears in the Word document

# Table formatting options
TABLE_STYLE = "Table Grid"            # Word table style
AUTO_FIT = True                       # Auto-fit table width to content
CENTER_TABLE = True                   # Center the table in the document

# Optional: Header row formatting
FIRST_ROW_IS_HEADER = True            # Treat first row as header (will be bold)