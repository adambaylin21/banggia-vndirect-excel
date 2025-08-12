# HTML to Excel Converter

This Python script converts HTML table data from the file 'Untitled-1.html' to an Excel file.

## Files Created

1. **html_to_excel.py** - Basic version for HTML to Excel conversion
2. **html_to_excel_improved.py** - Enhanced version with better data cleaning and column naming
3. **banggia_phaisinh.xlsx** - Output Excel file with converted data

## Usage

### Prerequisites
Install required Python packages:
```bash
pip install pandas beautifulsoup4 openpyxl
```

### Run the Script
```bash
python html_to_excel_improved.py
```

## Output Details

- **Input file**: `Untitled-1.html`
- **Output file**: `banggia_phaisinh.xlsx`
- **Data rows**: 246 rows exported
- **Columns**: 28 columns including:
  - Symbol (stock code)
  - Exchange (trading venue)
  - Expiry_Date (contract expiration)
  - Price columns (various price types)
  - Volume data
  - Margin information
  - Strike price and ratio
  - And other financial data

## Features

- Automatically parses HTML table structure
- Cleans and formats numeric data
- Handles different data types (numbers, dates, text)
- Provides detailed conversion statistics
- Shows sample of converted data

## Script Structure

The improved script includes:
- HTML parsing using BeautifulSoup
- Data cleaning and formatting
- Pandas DataFrame creation
- Excel export with proper column headers
- Error handling and progress reporting