#!/usr/bin/env python3
"""
Excel to Word Table Extractor - File Watcher
Monitors a directory for new Excel files and automatically processes them.
"""

import sys
import os
import time
import shutil
from pathlib import Path
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import openpyxl
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Import configuration and main processing functions
try:
    import config
    from main import extract_excel_data, create_word_table, validate_config, parse_xml
except ImportError as e:
    print(f"Error: {e}")
    print("Please ensure config.py and main.py exist in the same directory.")
    sys.exit(1)


class ExcelFileHandler(FileSystemEventHandler):
    """Handler for monitoring and processing Excel files."""
    
    def __init__(self, watch_directory, output_directory, processed_directory=None, 
                 auto_process=True, file_patterns=None):
        """
        Initialize the Excel file handler.
        
        Args:
            watch_directory (str): Directory to monitor for new Excel files
            output_directory (str): Directory where Word documents will be saved
            processed_directory (str): Directory to move processed Excel files (optional)
            auto_process (bool): Automatically process new files
            file_patterns (list): List of file patterns to watch (e.g., ['*.xlsx', '*.xls'])
        """
        self.watch_directory = Path(watch_directory)
        self.output_directory = Path(output_directory)
        self.processed_directory = Path(processed_directory) if processed_directory else None
        self.auto_process = auto_process
        self.file_patterns = file_patterns or ['*.xlsx', '*.xls', '*.xlsm']
        self.processing_files = set()  # Track files being processed to avoid duplicates
        
        # Create directories if they don't exist
        self.output_directory.mkdir(parents=True, exist_ok=True)
        if self.processed_directory:
            self.processed_directory.mkdir(parents=True, exist_ok=True)
        
        print(f"Watching directory: {self.watch_directory}")
        print(f"Output directory: {self.output_directory}")
        if self.processed_directory:
            print(f"Processed files directory: {self.processed_directory}")
        print(f"File patterns: {', '.join(self.file_patterns)}")
        print("-" * 50)
    
    def on_created(self, event):
        """Handle file creation events."""
        if not event.is_directory:
            self.process_file(event.src_path)
    
    def on_modified(self, event):
        """Handle file modification events."""
        if not event.is_directory:
            # Only process if file hasn't been processed recently
            file_path = Path(event.src_path)
            if file_path not in self.processing_files:
                # Wait a moment to ensure file write is complete
                time.sleep(0.5)
                self.process_file(event.src_path)
    
    def on_moved(self, event):
        """Handle file move events."""
        if not event.is_directory:
            self.process_file(event.dest_path)
    
    def is_valid_excel_file(self, file_path):
        """
        Check if the file is a valid Excel file matching our patterns.
        
        Args:
            file_path (str): Path to the file
            
        Returns:
            bool: True if valid Excel file, False otherwise
        """
        file_path = Path(file_path)
        
        # Check if file matches patterns
        for pattern in self.file_patterns:
            if file_path.match(pattern):
                # Ignore temporary Excel files
                if file_path.name.startswith('~$'):
                    return False
                return True
        return False
    
    def process_file(self, file_path):
        """
        Process an Excel file and convert it to Word document.
        
        Args:
            file_path (str): Path to the Excel file to process
        """
        file_path = Path(file_path)
        
        # Check if it's a valid Excel file
        if not self.is_valid_excel_file(file_path):
            return
        
        # Skip if file is already being processed
        if file_path in self.processing_files:
            return
        
        # Add to processing set
        self.processing_files.add(file_path)
        
        try:
            # Wait for file to be fully written (in case it's still being saved)
            time.sleep(1)
            
            # Check if file is accessible
            if not file_path.exists():
                print(f"File no longer exists: {file_path}")
                return
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] New Excel file detected: {file_path.name}")
            
            if not self.auto_process:
                response = input("Process this file? (y/n): ").lower().strip()
                if response != 'y':
                    print("Skipping file.")
                    return
            
            # Generate output filename
            output_filename = f"{file_path.stem}_{timestamp}.docx"
            output_path = self.output_directory / output_filename
            
            print(f"Processing: {file_path.name} -> {output_filename}")
            
            # Convert column letters to numbers and extract data
            from main import column_letter_to_number
            start_col_num = column_letter_to_number(config.START_COL)
            end_col_num = column_letter_to_number(config.END_COL)
            
            data = extract_excel_data(
                str(file_path),
                config.SHEET_NAME,
                config.START_ROW,
                config.END_ROW,
                start_col_num,
                end_col_num
            )
            
            if data is None:
                print(f"✗ Failed to extract data from {file_path.name}")
                return
            
            print(f"✓ Extracted {len(data)} rows with {len(data[0]) if data else 0} columns")
            
            # Create Word document
            title = f"{config.DOCUMENT_TITLE} - {file_path.stem}"
            success = create_word_table(data, str(output_path), title, config)
            
            if success:
                print(f"✓ Created Word document: {output_filename}")
                
                # Move processed file if configured
                if self.processed_directory:
                    processed_path = self.processed_directory / f"{file_path.stem}_{timestamp}{file_path.suffix}"
                    try:
                        shutil.move(str(file_path), str(processed_path))
                        print(f"✓ Moved processed file to: {processed_path.name}")
                    except Exception as e:
                        print(f"Warning: Could not move file: {e}")
            else:
                print(f"✗ Failed to create Word document for {file_path.name}")
        
        except PermissionError:
            print(f"✗ Permission denied: {file_path.name} (file may be open)")
        except Exception as e:
            print(f"✗ Error processing {file_path.name}: {str(e)}")
        finally:
            # Remove from processing set
            self.processing_files.discard(file_path)


def scan_existing_files(handler):
    """
    Scan and optionally process existing Excel files in the watch directory.
    
    Args:
        handler (ExcelFileHandler): The file handler to use for processing
    """
    excel_files = []
    for pattern in handler.file_patterns:
        excel_files.extend(handler.watch_directory.glob(pattern))
    
    # Filter out temporary files
    excel_files = [f for f in excel_files if not f.name.startswith('~$')]
    
    if excel_files:
        print(f"\nFound {len(excel_files)} existing Excel file(s) in watch directory:")
        for i, file in enumerate(excel_files, 1):
            print(f"  {i}. {file.name}")
        
        response = input("\nProcess existing files? (y/n/q to quit): ").lower().strip()
        if response == 'q':
            return False
        elif response == 'y':
            for file in excel_files:
                handler.process_file(str(file))
            print("\n" + "=" * 50)
    
    return True


def main():
    """Main function to set up and run the file watcher."""
    print("=" * 50)
    print("Excel to Word Table Extractor - File Watcher")
    print("=" * 50)
    
    # Configuration for the watcher
    WATCH_DIRECTORY = "./excel-data"  # Watch the excel-data folder
    OUTPUT_DIRECTORY = "./docx-output"  # Output to docx-output folder
    PROCESSED_DIRECTORY = "./excel-data/processed"  # Subfolder for processed files
    AUTO_PROCESS = True  # Set to False for manual confirmation
    FILE_PATTERNS = ['*.xlsx', '*.xls', '*.xlsm']
    
    # Allow command-line override of watch directory
    if len(sys.argv) > 1:
        WATCH_DIRECTORY = sys.argv[1]
        print(f"Watch directory overridden: {WATCH_DIRECTORY}")
    
    # Validate main configuration
    print("\nValidating extraction configuration...")
    if not validate_config():
        print("\n✗ Please fix configuration errors in config.py and try again.")
        print("Note: The watcher uses the same config.py settings for data extraction.")
        return 1
    print("✓ Configuration valid")
    
    # Create event handler
    event_handler = ExcelFileHandler(
        watch_directory=WATCH_DIRECTORY,
        output_directory=OUTPUT_DIRECTORY,
        processed_directory=PROCESSED_DIRECTORY,
        auto_process=AUTO_PROCESS,
        file_patterns=FILE_PATTERNS
    )
    
    # Scan for existing files
    if not scan_existing_files(event_handler):
        print("Exiting...")
        return 0
    
    # Set up file system observer
    observer = Observer()
    observer.schedule(event_handler, str(event_handler.watch_directory), recursive=False)
    
    # Start watching
    observer.start()
    print(f"Watching for new Excel files... (Press Ctrl+C to stop)")
    print("=" * 50)
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\nStopping file watcher...")
        observer.stop()
    
    observer.join()
    print("File watcher stopped.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(f"\nUnexpected error: {str(e)}")
        sys.exit(1)