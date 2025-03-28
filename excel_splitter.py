import pandas as pd
import os
import argparse

def split_excel_sheets(input_file, output_dir=None):
    """
    Split an Excel file into multiple files, one for each sheet.
    Each new file will be named after its sheet.
    """
    if output_dir is None:
        output_dir = os.path.dirname(input_file) or '.'
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Read the Excel file
    print(f"Reading Excel file: {input_file}")
    try:
        excel = pd.ExcelFile(input_file)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # Get all sheet names
    sheet_names = excel.sheet_names
    print(f"Found {len(sheet_names)} sheets: {', '.join(sheet_names)}")
    
    # Process each sheet
    for sheet_name in sheet_names:
        # Create a safe filename (replace invalid characters)
        safe_name = sheet_name.replace('/', '_').replace('\\', '_').replace(':', '_')
        
        # Create output path
        output_file = os.path.join(output_dir, f"{safe_name}.xlsx")
        
        print(f"Processing sheet: {sheet_name} -> {output_file}")
        
        try:
            # Read the sheet
            df = pd.read_excel(excel, sheet_name=sheet_name)
            
            # Write to new Excel file
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Format the worksheet with headers and column widths
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                
                # Format headers
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'bg_color': '#D7E4BC',
                    'border': 1
                })
                
                for col_num, value in enumerate(df.columns):
                    worksheet.write(0, col_num, str(value), header_format)
                
                # Auto-filter for headers
                last_col = len(df.columns) - 1
                if last_col >= 0:
                    worksheet.autofilter(0, 0, 0, last_col)
                
                # Set freeze panes to keep headers visible
                worksheet.freeze_panes(1, 0)
                
                # Auto-fit columns
                for i, col in enumerate(df.columns):
                    # Set a minimum width
                    width = max(15, len(str(col)) + 2)
                    # Increase width for certain column types
                    if 'email' in str(col).lower() or 'url' in str(col).lower() or any(s in str(col).lower() for s in ['facebook', 'linkedin', 'twitter', 'instagram']):
                        width = max(width, 40)
                    worksheet.set_column(i, i, width)
            
            print(f"Successfully created: {output_file}")
        
        except Exception as e:
            print(f"Error processing sheet '{sheet_name}': {e}")

def main():
    parser = argparse.ArgumentParser(description='Split Excel sheets into separate files')
    parser.add_argument('--input', type=str, required=True,
                        help='Input Excel file to split')
    parser.add_argument('--output-dir', type=str, default=None,
                        help='Output directory for the split files (default: same as input)')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' not found.")
        return
    
    split_excel_sheets(args.input, args.output_dir)
    print("Processing complete!")

if __name__ == "__main__":
    main() 