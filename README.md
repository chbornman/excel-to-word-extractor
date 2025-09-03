# Excel to Word Table Extractor

A Python utility that extracts data from Excel files and creates formatted Word documents with tables. Includes both manual processing and automatic file watching capabilities.

## Directory Structure

```
excel-docx/
├── excel-data/           # Place your Excel files here
│   └── processed/        # Processed files moved here (watcher mode)
├── docx-output/          # Generated Word documents saved here
├── config.py             # Configuration settings
├── main.py               # Manual conversion script
├── watch_main.py         # File watcher script
├── requirements.txt      # Python package dependencies
└── create_test_data.py   # Generate test Excel files
```

## Features

- **Configurable Data Extraction**: Specify exact rows and columns to extract via `config.py`
- **Flexible Sheet Selection**: Work with specific sheets or active sheet
- **Formatted Word Output**: Creates professional Word documents with styled tables
- **File Watcher Mode**: Automatically processes new Excel files as they appear
- **Batch Processing**: Process existing files in a directory
- **Error Handling**: Robust error handling with clear user feedback

## Table of Contents

1. [Installation](#installation)
   - [Installing Python on Windows](#installing-python-on-windows)
   - [Installing Dependencies](#installing-dependencies)
2. [Configuration](#configuration)
3. [Usage](#usage)
   - [Manual Processing](#manual-processing)
   - [File Watcher Mode](#file-watcher-mode)
4. [Setting Up Auto-Start Service](#setting-up-auto-start-service)
   - [Windows Task Scheduler](#windows-task-scheduler)
5. [Examples](#examples)
6. [Troubleshooting](#troubleshooting)

## Installation

### Installing Python on Windows

If you don't have Python installed on your Windows machine, follow these steps:

#### Method 1: Official Python Installer (Recommended)

1. **Download Python:**
   - Visit [python.org/downloads](https://www.python.org/downloads/)
   - Click "Download Python 3.x.x" (latest version)
   - Choose the 64-bit installer for modern systems

2. **Install Python:**
   - Run the downloaded installer
   - **IMPORTANT:** Check ✅ "Add Python to PATH" at the bottom of the installer
   - Click "Install Now"
   - Wait for installation to complete

3. **Verify Installation:**
   - Open Command Prompt (Win+R, type `cmd`, press Enter)
   - Type `python --version` and press Enter
   - You should see: `Python 3.x.x`
   - Type `pip --version` to verify pip is installed

#### Method 2: Microsoft Store (Windows 10/11)

1. Open Microsoft Store
2. Search for "Python"
3. Select "Python 3.x" by Python Software Foundation
4. Click "Get" or "Install"
5. Open Command Prompt and verify with `python --version`

#### Method 3: Using winget (Windows Package Manager)

If you have Windows 11 or updated Windows 10:
```cmd
winget install Python.Python.3
```

### Installing Dependencies

Once Python is installed, install the required packages:

```bash
# Using requirements.txt (recommended)
pip install -r requirements.txt
```

Or install packages individually:
```bash
# Open Command Prompt or PowerShell as Administrator
pip install openpyxl python-docx
```

For the file watcher functionality:
```bash
pip install watchdog
```

For development/testing, you might also want:
```bash
pip install pandas  # For creating test Excel files
```

## Configuration

Edit `config.py` to specify your extraction settings:

```python
# Input Excel file configuration
EXCEL_FILE = "excel-data/data.xlsx"   # Path to your Excel file
SHEET_NAME = None                     # None for active sheet, or "Sheet1", "Sheet2", etc.

# Data extraction range
START_ROW = 1                         # First row to extract
END_ROW = 10                          # Last row to extract
START_COL = "A"                       # First column letter
END_COL = "E"                         # Last column letter

# Output configuration
OUTPUT_FILE = "docx-output/extracted_data.docx"  # Output path
DOCUMENT_TITLE = "Extracted Excel Data"

# Formatting options
FIRST_ROW_IS_HEADER = True           # Bold first row as header
```

**Note:** The current selection mechanism is intentionally basic, extracting a simple rectangular range of cells. This can be improved in future versions to support:
- Multiple non-contiguous ranges
- Named ranges from Excel
- Dynamic range detection based on data presence
- Column/row filtering based on criteria
- Cell formatting preservation

## Usage

### Manual Processing

1. Place your Excel file in the `excel-data/` folder
2. Update `config.py` with the filename and extraction range
3. Run the conversion:

```bash
python main.py
```

This will:
- Read the Excel file from `excel-data/`
- Extract the specified range of data
- Create a formatted Word document
- Save to `docx-output/`

### File Watcher Mode

Monitor the `excel-data/` folder for new Excel files and process them automatically:

```bash
python watch_main.py
```

Or specify a custom directory to watch:

```bash
python watch_main.py /path/to/watch/directory
```

**Features:**
- Watches `excel-data/` folder for new Excel files
- Processes files automatically using config.py settings
- Moves processed files to `excel-data/processed/`
- Saves Word documents to `docx-output/`
- Can process existing files on startup

**Default Watcher Configuration** (in `watch_main.py`):
```python
WATCH_DIRECTORY = "./excel-data"      # Watch for Excel files here
OUTPUT_DIRECTORY = "./docx-output"    # Save Word documents here
PROCESSED_DIRECTORY = "./excel-data/processed"  # Move processed files here
AUTO_PROCESS = True                   # False for manual confirmation
FILE_PATTERNS = ['*.xlsx', '*.xls']   # File types to watch
```

## Setting Up Auto-Start Service

To have the watcher script run automatically when your machine starts, you need to set up a system service.

### Windows Task Scheduler

#### Method 1: Using Task Scheduler GUI

1. **Open Task Scheduler:**
   - Press Win+R, type `taskschd.msc`, press Enter
   - Or search "Task Scheduler" in Start Menu

2. **Create Basic Task:**
   - Click "Create Basic Task..." in the Actions panel
   - Name: "Excel to Word Watcher"
   - Description: "Monitors for new Excel files and converts to Word"
   - Click Next

3. **Set Trigger:**
   - Choose "When the computer starts"
   - Click Next

4. **Set Action:**
   - Choose "Start a program"
   - Program: `C:\Path\To\Python\python.exe`
   - Arguments: `C:\Path\To\Your\Project\watch_main.py`
   - Start in: `C:\Path\To\Your\Project`
   - Click Next, then Finish

5. **Configure Advanced Settings:**
   - Find your task in the Task Scheduler Library
   - Right-click → Properties
   - General tab: Check "Run with highest privileges"
   - Conditions tab: Uncheck "Start only if on AC power"
   - Settings tab: Check "Allow task to be run on demand"

#### Method 2: Using PowerShell

Create a scheduled task via PowerShell (run as Administrator):

```powershell
$action = New-ScheduledTaskAction -Execute "python.exe" -Argument "C:\Path\To\watch_main.py" -WorkingDirectory "C:\Path\To\Project"
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName "ExcelToWordWatcher" -Action $action -Trigger $trigger -Principal $principal -Settings $settings
```

#### Method 3: Creating a Windows Service

For a more robust solution, create a Windows service using `nssm`:

1. Download [NSSM](https://nssm.cc/download)
2. Extract to a folder (e.g., `C:\nssm`)
3. Open Command Prompt as Administrator:

```cmd
cd C:\nssm\win64
nssm install ExcelToWordWatcher
```

4. In the GUI that opens:
   - Path: `C:\Path\To\Python\python.exe`
   - Arguments: `C:\Path\To\watch_main.py`
   - Startup directory: `C:\Path\To\Project`

5. Start the service:
```cmd
nssm start ExcelToWordWatcher
```

## Examples

### Example 1: Extract Specific Data Range

1. Place `sales_report.xlsx` in the `excel-data/` folder
2. Configure `config.py`:
```python
EXCEL_FILE = "excel-data/sales_report.xlsx"
SHEET_NAME = "Q4 Sales"
START_ROW = 5
END_ROW = 25
START_COL = "B"  # Column B
END_COL = "G"    # Column G
OUTPUT_FILE = "docx-output/q4_sales_summary.docx"
```

Run:
```bash
python main.py
```

### Example 2: Watch for Daily Reports

1. Drop Excel files into the `excel-data/` folder
2. Run the watcher:
```bash
python watch_main.py
```

To watch a different folder (e.g., network drive):
```bash
python watch_main.py "\\\\NetworkDrive\\DailyReports"
```

### Example 3: Process Multiple Files

Create a batch script (`process_all.bat` for Windows):
```batch
@echo off
for %%f in (excel-data\*.xlsx) do (
    echo Processing %%f
    copy "%%f" "excel-data\data.xlsx"
    python main.py
    move "docx-output\extracted_data.docx" "docx-output\%%~nf.docx"
)
```

## Creating Test Data

A sample Excel file `test_data.xlsx` is included. To create your own test file:

```python
import pandas as pd
import numpy as np

# Create sample data
data = {
    'Product': ['Widget A', 'Widget B', 'Gadget X', 'Gadget Y', 'Tool Z'],
    'Q1 Sales': np.random.randint(100, 1000, 5),
    'Q2 Sales': np.random.randint(100, 1000, 5),
    'Q3 Sales': np.random.randint(100, 1000, 5),
    'Q4 Sales': np.random.randint(100, 1000, 5),
    'Total': np.random.randint(1000, 5000, 5)
}

df = pd.DataFrame(data)
df.to_excel('excel-data/test_data.xlsx', index=False, sheet_name='Sales Data')
print("Test file created: test_data.xlsx")
```

### Example 4: Extract from Multi-Column Ranges

For wider ranges spanning many columns:
```python
# In config.py
START_COL = "A"
END_COL = "Z"     # Extracts columns A through Z

# Or for columns beyond Z:
START_COL = "AA"
END_COL = "AF"    # Extracts columns AA through AF
```

### Example 5: Reprocessing Previously Output Files

If you need to edit data from a previously processed Excel file:

1. **Locate the file** in `excel-data/processed/`
2. **Edit the Excel file** with your changes
3. **Move it back** to the main `excel-data/` folder
4. **Reprocess the file:**
   - For manual processing: Update `config.py` with the filename and run `python main.py`
   - For automatic processing: If the watcher is running, it will detect and process automatically

**Note:** The file will be moved back to `processed/` after reprocessing if using the watcher. To edit multiple times, either:
- Work on a copy of the file
- Temporarily stop the watcher while editing
- Use manual processing mode instead

## Troubleshooting

### Common Issues

1. **"Module not found" error:**
   ```bash
   pip install --upgrade openpyxl python-docx watchdog
   ```

2. **"Permission denied" error:**
   - Ensure Excel file is not open in another program
   - Run script with appropriate permissions
   - Check file/directory permissions

3. **"Excel file not found":**
   - Use absolute paths in `config.py`
   - Verify file exists and path is correct

4. **Watcher not detecting files:**
   - Check file patterns match your Excel files
   - Ensure watch directory is correct
   - Verify no antivirus blocking file system events

5. **Service not starting on Windows:**
   - Check Python is in system PATH
   - Use absolute paths in service configuration
   - Check Event Viewer for error messages
   - Ensure all dependencies are installed for the service user

### Debug Mode

For debugging, modify the scripts to add verbose output:

```python
# In main.py or watch_main.py
import logging
logging.basicConfig(level=logging.DEBUG, 
                   format='%(asctime)s - %(levelname)s - %(message)s')
```

## Requirements

- Python 3.7+
- openpyxl (for Excel file reading)
- python-docx (for Word document creation)
- watchdog (for file system monitoring)

Install via:
```bash
pip install -r requirements.txt
```

## License

MIT License - Feel free to modify and distribute as needed.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.