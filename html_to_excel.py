import pandas as pd
from bs4 import BeautifulSoup
import re

def clean_numeric_value(text):
    """Clean and convert numeric values"""
    if not text or text == '':
        return None
    
    # Remove commas and clean the text
    cleaned = str(text).replace(',', '').strip()
    
    # Handle different number formats
    try:
        if '.' in cleaned and cleaned.count('.') == 1:
            # Handle decimal numbers (e.g., "4.27", "32.05")
            return float(cleaned)
        elif cleaned.replace('-', '').isdigit():
            # Handle integers
            return int(cleaned)
        else:
            # Handle special cases like ratios "2:1", dates, etc.
            return cleaned
    except:
        return cleaned

def extract_table_data(html_content):
    """Extract data from HTML table"""
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Find the table
    table = soup.find('table', {'id': 'banggia-phaisinh'})
    if not table:
        print("Table with id 'banggia-phaisinh' not found!")
        return None
    
    # Define column headers based on the table structure
    headers = [
        'Symbol', 'Exchange', 'Expiry_Date', 'Price_TC', 'Price_Tran', 'Price_San',
        'Volume', 'Price_1', 'Volume_1', 'Price_2', 'Volume_2', 'Price_3', 'Volume_3',
        'Price_Close', 'Volume_Close', 'Change', 'Price_Buy_1', 'Volume_Buy_1',
        'Price_Buy_2', 'Volume_Buy_2', 'Price_Buy_3', 'Volume_Buy_3', 'Price_Sell_1',
        'Margin', 'Strike_Price', 'Ratio', 'Theo_chờ', 'Volume_Theo_chờ'
    ]
    
    # Extract table rows
    rows_data = []
    tbody = table.find('tbody')
    if tbody:
        rows = tbody.find_all('tr')
        
        for row in rows:
            cells = row.find_all('td')
            row_data = []
            
            for cell in cells:
                # Extract text from span tags or direct text
                span = cell.find('span')
                if span:
                    text = span.get_text(strip=True)
                else:
                    text = cell.get_text(strip=True)
                
                # Clean and format the data
                cleaned_value = clean_numeric_value(text)
                row_data.append(cleaned_value)
            
            if row_data:  # Only add non-empty rows
                rows_data.append(row_data)
    
    return rows_data, headers

def convert_html_to_excel(html_file_path, excel_file_path):
    """
    Convert HTML table data to Excel file
    """
    try:
        # Read the HTML file
        with open(html_file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()
        
        # Extract table data
        rows_data, headers = extract_table_data(html_content)
        
        if rows_data is None:
            return
        
        # Create DataFrame
        df = pd.DataFrame(rows_data)
        
        # Set column names
        df.columns = headers[:len(df.columns)]  # Match headers to actual columns
        
        # Save to Excel
        df.to_excel(excel_file_path, index=False)
        print(f"✓ Data successfully exported to {excel_file_path}")
        print(f"✓ Total rows exported: {len(df)}")
        print(f"✓ Total columns: {len(df.columns)}")
        
        # Show first few rows
        print("\nFirst 5 rows of data:")
        print(df.head())
        
    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    # File paths
    html_file = "filehtml.html"
    excel_file = "banggia_phaisinh.xlsx"
    
    print("Converting HTML table to Excel...")
    print(f"Input file: {html_file}")
    print(f"Output file: {excel_file}")
    print("-" * 50)
    
    # Convert HTML to Excel
    convert_html_to_excel(html_file, excel_file)
    
    print("-" * 50)
    print("Conversion completed!")