import pandas as pd
import re
from urllib.parse import urlparse
import os

def extract_domain(url):
    """Extract domain from URL, handling various formats."""
    if pd.isna(url):
        return None
    try:
        # Remove any whitespace and convert to string
        url = str(url).strip()
        # Add http:// if no protocol is specified
        if not url.startswith(('http://', 'https://')):
            url = 'http://' + url
        # Parse the URL and get the domain
        parsed = urlparse(url)
        domain = parsed.netloc.lower()
        # Remove 'www.' if present
        domain = re.sub(r'^www\.', '', domain)
        return domain
    except:
        return None

def process_file_pair(source_file, contact_file, output_file):
    """Process a single pair of files."""
    print(f"\nProcessing {source_file} -> {contact_file}")
    
    # Read source file and extract domains
    source_df = pd.read_excel(source_file)
    print("Source file columns:", source_df.columns.tolist())
    
    # Extract domains from source file's website column
    source_df['domain'] = source_df['Website'].apply(extract_domain)
    
    # Create a mapping dictionary from domain to company name and bio
    domain_mapping = {}
    for _, row in source_df.iterrows():
        if pd.notna(row['domain']):
            domain_mapping[row['domain']] = {
                'company_name': row['Company Name'],
                'bio': row['Bio']
            }
    
    # Get all sheet names from contact file
    excel_file = pd.ExcelFile(contact_file)
    sheet_names = excel_file.sheet_names
    print(f"\nProcessing {len(sheet_names)} sheets from contact file:", sheet_names)
    
    # Create a new Excel writer
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Process each sheet
        for sheet_name in sheet_names:
            print(f"\nProcessing sheet: {sheet_name}")
            
            # Read the sheet
            contact_df = pd.read_excel(contact_file, sheet_name=sheet_name)
            print(f"Columns in sheet: {contact_df.columns.tolist()}")
            
            # Add new columns to contact_df if they don't exist
            if 'company_name' not in contact_df.columns:
                contact_df['company_name'] = None
            if 'bio' not in contact_df.columns:
                contact_df['bio'] = None
            
            # Update company name and bio based on domain matching
            matches = 0
            for idx, row in contact_df.iterrows():
                domain = row['domain']
                if domain in domain_mapping:
                    contact_df.at[idx, 'company_name'] = domain_mapping[domain]['company_name']
                    contact_df.at[idx, 'bio'] = domain_mapping[domain]['bio']
                    matches += 1
            
            print(f"Found {matches} matches out of {len(contact_df)} total records in sheet {sheet_name}")
            
            # Save the sheet
            contact_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"\nAll sheets processed and saved to {output_file}")

def update_contact_info():
    # Define all file pairs
    file_pairs = [
        {
            'source': 'HealthOptimisation.xlsx',
            'contact': 'output/HealthOptimization_contact_info_organized.xlsx',
            'output': 'output1/HealthOptimization_contact_info_updated.xlsx'
        },
        {
            'source': 'Logevity Med.xlsx',
            'contact': 'output/LongevityMed_contact_info_organized.xlsx',
            'output': 'output1/LongevityMed_contact_info_updated.xlsx'
        },
        {
            'source': 'Longevity.xlsx',
            'contact': 'output/Longevity_contact_info_organized.xlsx',
            'output': 'output1/Longevity_contact_info_updated.xlsx'
        },
        {
            'source': 'NextMed.xlsx',
            'contact': 'output/NextMed_contact_info_organized.xlsx',
            'output': 'output1/NextMed_contact_info_updated.xlsx'
        }
    ]
    
    # Process each pair of files
    for pair in file_pairs:
        try:
            process_file_pair(pair['source'], pair['contact'], pair['output'])
            print(f"\nSuccessfully processed {pair['source']}")
        except Exception as e:
            print(f"\nError processing {pair['source']}: {str(e)}")

if __name__ == "__main__":
    update_contact_info() 