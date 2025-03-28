# Excel Contact Info Processor

This repository contains Python scripts to organize and process contact information from CSV files into various Excel formats.

## Features

The processor offers the following organization options:

1. **Domain Summary**: Counts of contact information (emails, phones, social media) for each domain
2. **Original Data**: The original CSV data with cleaned column names
3. **Domain Consolidated**: All contact info for each domain in a consolidated view 
4. **Social Media**: Separate sheets for each social media platform (Facebook, Twitter, LinkedIn, etc.)
5. **Contact Types**: Contact information organized by type (emails, phones, uncertain phones)

- **Deduplication**: Automatically removes duplicate contact information
- **Clean formatting**: Professionally formatted Excel sheets with proper column widths and headers
- **Consolidated view**: All contact information for each domain in a single row
- **Filtered views**: Specific sheets for each type of contact information
- **Flexible output options**: Choose which organization formats to include in the output

## Requirements

- Python 3.6+
- Required packages: pandas, xlsxwriter

To install the required packages:

```bash
pip install pandas xlsxwriter
```

## Usage

### Basic Script

The basic script (`excel_processor.py`) processes the CSV file and generates all formats:

```bash
python excel_processor.py
```

### Enhanced Script

The enhanced script (`excel_processor_enhanced.py`) provides additional options:

```bash
python excel_processor_enhanced.py --input your_input_file.csv --output your_output_file.xlsx --formats all
```

### Command-line Arguments

- `--input`: Path to the input CSV file (default: dataset_contact-info-scraper_2025-03-22_10-29-36-105.csv)
- `--output`: Path to the output Excel file (default: contact_info_organized.xlsx)
- `--formats`: Formats to include in the output Excel file. Options:
  - `all`: Include all formats (default)
  - `original`: Include the original data
  - `domain`: Include the domain-consolidated view
  - `social`: Include social media sheets
  - `contact`: Include contact type sheets

Example for generating only domain summary and social media sheets:

```bash
python excel_processor_enhanced.py --formats domain social
```

## Output File Structure

The output Excel file contains multiple sheets:

1. **Domain_Summary**: Overview of all domains with contact counts
2. **Original_Data** (if requested): Original data with cleaned column names
3. **Domain_Consolidated** (if requested): All contact info by domain
4. **Social_[Platform]** (if requested): Each social media platform has its own sheet
5. **Contact_[Type]** (if requested): Contact information organized by type

## Examples

### Basic Example

```bash
python excel_processor.py
```

This will read the default CSV file and create a comprehensive Excel file with all organization formats.

### Custom Example

```bash
python excel_processor_enhanced.py --input my_data.csv --output organized_contacts.xlsx --formats domain social
```

This will:
1. Read `my_data.csv`
2. Generate an Excel file named `organized_contacts.xlsx`
3. Include only the domain summary and social media sheets 

## Additional Utilities

### Excel Sheet Splitter

The repository also includes a utility script to split an Excel workbook into multiple files, one for each sheet.

#### Usage

```bash
python excel_splitter.py --input <path_to_excel_file> --output-dir <optional_output_directory>
```

Example:
```bash
python excel_splitter.py --input contact_info_organized.xlsx
```

This will create separate Excel files for each sheet in the input file, with filenames matching the sheet names:
- Domain_Consolidated.xlsx
- Contact_Emails.xlsx
- Contact_Phones.xlsx
- etc.

Each output file maintains the formatting of the original sheet, including headers, column widths, and data. 