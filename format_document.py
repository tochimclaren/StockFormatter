import csv
import re
import random
import string
import argparse
import os
import sys
import pandas as pd

class FormatDocument:
    header = [
        "id",
        "name",
        "brand",
        "color",
        "code",
        "quantity",
        "price",
        "branch",
        "branch_id",
        "description",
    ]

    def __init__(self, excel_path, sheet_name=0):
        """
        Initialize with the path to an Excel file.
        sheet_name can be the name of a specific sheet or an index (default is 0 for first sheet)
        """
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.items = []
        self.parse_excel()

    def parse_excel(self):
        """Parse the Excel file and extract item data"""
        try:
            # Read the Excel file
            df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)
            
            # Find the row index where the actual data starts (after headers)
            start_row = 0
            for i, col_value in enumerate(df.iloc[:, 1]):  # Check the second column which should have "ITEM DESCRIPTION"
                if isinstance(col_value, str) and "ITEM DESCRIPTION" in col_value:
                    start_row = i + 1
                    break
            
            # Extract the data rows
            data_df = df.iloc[start_row:].reset_index(drop=True)
            
            # Process each row
            for idx, row in data_df.iterrows():
                try:
                    # Skip rows without an item number in the first column
                    if not isinstance(row.iloc[0], (int, float)) or pd.isna(row.iloc[0]):
                        continue
                    
                    # Extract data from the row
                    item_id = str(int(row.iloc[0]))  # Convert to int first to remove decimal points
                    item_name = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ""
                    
                    # Skip empty item names
                    if not item_name.strip():
                        continue
                    
                    # Extract quantity if available (column index 2)
                    quantity = "1"  # Default
                    if len(row) > 2 and not pd.isna(row.iloc[2]) and str(row.iloc[2]).isdigit():
                        quantity = str(int(row.iloc[2]))
                    
                    # Extract brand from name
                    brand = self.extract_brand(item_name)
                    
                    # Create a random code
                    code = self.generate_code(item_name)
                    
                    # Extract other properties
                    color = self.extract_color(item_name)
                    price = "20.00"  # Default price
                    branch = "ojodu"  # Default branch
                    branch_id = "3"   # Default branch ID
                    description = f"{item_name}"
                    
                    # Add to items list
                    self.items.append({
                        "id": item_id,
                        "name": item_name.lower(),
                        "brand": brand,
                        "color": color,
                        "code": code,
                        "quantity": quantity,
                        "price": price,
                        "branch": branch,
                        "branch_id": branch_id,
                        "description": description
                    })
                except Exception as e:
                    print(f"Error processing row {idx + start_row}: {e}")
                    continue
                    
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            raise

    def extract_brand(self, name):
        """Extract brand from item name"""
        # Common brand keywords
        brand_keywords = ["LINSAN", "BREVILLE", "PRIMA", "TESCO", "COLEMAN", "SAINSBURY'S", 
                        "PYREX", "PHILIPS", "ELPINE", "OZARK", "SNAPWARE", "RUBBERMAID",
                        "CORNINGWARE", "KIRKLAND", "HAIER", "INDESIT", "MIKASA", "LUMINARC"]
        
        # Check if any brand keyword is in the name
        for brand in brand_keywords:
            if brand in name.upper():
                return brand.title()
        
        # If no brand found, return first word as brand
        words = name.split()
        return words[0].title() if words else "Unknown"

    def extract_color(self, name):
        """Extract color from item name if present"""
        colors = ["RED", "BLUE", "GREEN", "BLACK", "WHITE", "GREY", "PINK", "SILVER", "NAVY"]
        for color in colors:
            if color in name.upper():
                return color.lower()
        return "silver"  # Default color

    def generate_code(self, name):
        """Generate a unique code for the item"""
        # Use first 3 letters of name + random numbers
        prefix = re.sub(r'[^a-zA-Z]', '', name)[:3].upper()
        random_num = ''.join(random.choices(string.digits, k=9))
        return f"{prefix}{random_num}"

    def export_to_csv(self, filename="formatted_items.csv"):
        """Export items to CSV file"""
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=self.header)
            writer.writeheader()
            writer.writerows(self.items)
        return filename

    @classmethod
    def format(cls, excel_path, sheet_name=0, output_filename="formatted_items.csv"):
        """Static method to create and process document in one step"""
        formatter = cls(excel_path, sheet_name)
        return formatter.export_to_csv(output_filename)


def list_sheets(excel_path):
    """List all sheet names in the Excel file"""
    try:
        xl = pd.ExcelFile(excel_path)
        print(f"Available sheets in {excel_path}:")
        for i, sheet in enumerate(xl.sheet_names):
            print(f"  {i}: {sheet}")
        return xl.sheet_names
    except Exception as e:
        print(f"Error reading sheets: {e}")
        return []


def print_usage():
    """Print usage information with examples"""
    usage_text = """
Excel-to-CSV Inventory Formatter

This script converts Excel inventory files (.xlsx or .xls) to CSV format 
compatible with Django models.

Usage Examples:
  # Basic usage - process the first sheet and output to formatted_items.csv
  python format_document.py inventory.xlsx
  
  # Specify an output file
  python format_document.py inventory.xlsx --output my_inventory.csv
  
  # Process a specific sheet by index (starting from 0)
  python format_document.py inventory.xlsx --sheet 2
  
  # Process a specific sheet by name
  python format_document.py inventory.xlsx --sheet "Middle Container"
  
  # List all available sheets in an Excel file
  python format_document.py inventory.xlsx --list-sheets
  
  # Set branch information
  python format_document.py inventory.xlsx --branch ikeja --branch-id 5

Required Dependencies:
  pip install pandas openpyxl xlrd
  
For help on specific options:
  python format_document.py --help
"""
    print(usage_text)


def main():
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(
        description='Format inventory Excel document to CSV',
        formatter_class=argparse.RawDescriptionHelpFormatter,  # Use raw formatting for the epilog
        epilog="""
Examples:
  python format_document.py inventory.xlsx
  python format_document.py inventory.xlsx --output my_inventory.csv
  python format_document.py inventory.xlsx --sheet "Sheet2" --branch ikeja
  python format_document.py inventory.xlsx --list-sheets
"""
    )
    
    parser.add_argument('excel_path', nargs='?', help='Path to the Excel file to be processed')
    parser.add_argument('--output', '-o', default='formatted_items.csv',
                      help='Output CSV file path (default: formatted_items.csv)')
    parser.add_argument('--sheet', '-s', default=0,
                      help='Sheet name or index to process (default: 0)')
    parser.add_argument('--branch', default='ojodu',
                      help='Branch name to assign to items (default: ojodu)')
    parser.add_argument('--branch-id', default='3',
                      help='Branch ID to assign to items (default: 3)')
    parser.add_argument('--list-sheets', '-l', action='store_true',
                      help='List all sheets in the Excel file and exit')
    
    args = parser.parse_args()
    
    # Print usage if no arguments provided
    if len(sys.argv) == 1:
        print_usage()
        sys.exit(0)
    
    # Check if Excel path is provided (when not using --help)
    if not args.excel_path:
        parser.print_help()
        print("\nError: Excel file path is required")
        sys.exit(1)
    
    # Validate input file
    if not os.path.isfile(args.excel_path):
        print(f"Error: File not found - {args.excel_path}")
        sys.exit(1)
    
    # List sheets if requested
    if args.list_sheets:
        list_sheets(args.excel_path)
        sys.exit(0)
    
    try:
        # Convert sheet argument to int if it's a number
        sheet = args.sheet
        try:
            sheet = int(sheet)
        except ValueError:
            # Keep it as string (sheet name)
            pass
        
        # Format the document
        formatter = FormatDocument(args.excel_path, sheet)
        
        # Update branch info if provided
        if args.branch != 'ojodu' or args.branch_id != '3':
            for item in formatter.items:
                item['branch'] = args.branch
                item['branch_id'] = args.branch_id
        
        # Export to CSV
        csv_file = formatter.export_to_csv(args.output)
        print(f"Success! Formatted document saved to {csv_file}")
        print(f"Processed {len(formatter.items)} items")
        
    except Exception as e:
        print(f"Error processing file: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()