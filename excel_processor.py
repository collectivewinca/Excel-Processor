import pandas as pd
import os
import re
import argparse
from collections import defaultdict

def clean_column_names(df):
    """Clean column names by replacing '/' with '_' and converting to lowercase."""
    df.columns = [col.lower().replace('/', '_') for col in df.columns]
    return df

def split_long_urls(url_str, max_length=2000):
    """Split long URLs into chunks that fit Excel's limit."""
    if pd.isna(url_str) or not url_str:
        return []
    
    # Split URLs if they're comma-separated
    urls = [u.strip() for u in str(url_str).split(',')]
    
    # Handle each URL
    processed_urls = []
    current_chunk = []
    current_length = 0
    
    for url in urls:
        url = url.strip()
        if not url:
            continue
            
        url_length = len(url)
        
        # If single URL is longer than max_length, truncate it
        if url_length > max_length:
            processed_urls.append(url[:max_length])
            continue
            
        # If adding this URL would exceed max_length, start new chunk
        if current_length + url_length + 2 > max_length:  # +2 for comma and space
            if current_chunk:
                processed_urls.append(', '.join(current_chunk))
            current_chunk = [url]
            current_length = url_length
        else:
            current_chunk.append(url)
            current_length += url_length + 2  # +2 for comma and space
    
    # Add remaining chunk
    if current_chunk:
        processed_urls.append(', '.join(current_chunk))
    
    return processed_urls

def extract_domains(df):
    """Extract unique domains from the dataset."""
    return df['domain'].unique().tolist()

def get_contact_counts(df):
    """Count number of contacts (emails, phones, social media) for each domain."""
    contact_data = []
    
    for domain in df['domain'].unique():
        domain_rows = df[df['domain'] == domain]
        
        # Count non-empty values for each contact type
        email_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                          if col.startswith('emails_') and pd.notna(row.get(col)) and row.get(col) != '')
        
        phone_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                          if col.startswith('phones_') and pd.notna(row.get(col)) and row.get(col) != '')
        
        uncertain_phone_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                                   if col.startswith('phonesuncertain_') and pd.notna(row.get(col)) and row.get(col) != '')
        
        facebook_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                            if 'facebook' in col and pd.notna(row.get(col)) and row.get(col) != '')
        
        linkedin_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                            if 'linkedin' in col and pd.notna(row.get(col)) and row.get(col) != '')
        
        twitter_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                           if ('twitter' in col or 'x.com' in col) and pd.notna(row.get(col)) and row.get(col) != '')
        
        instagram_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                             if 'instagram' in col and pd.notna(row.get(col)) and row.get(col) != '')
        
        youtube_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                           if 'youtube' in col and pd.notna(row.get(col)) and row.get(col) != '')
        
        tiktok_count = sum(1 for _, row in domain_rows.iterrows() for col in df.columns 
                          if 'tiktok' in col and pd.notna(row.get(col)) and row.get(col) != '')
        
        social_count = facebook_count + linkedin_count + twitter_count + instagram_count + youtube_count + tiktok_count
        
        contact_data.append({
            'domain': domain,
            'email_count': email_count,
            'phone_count': phone_count,
            'uncertain_phone_count': uncertain_phone_count,
            'facebook_count': facebook_count,
            'linkedin_count': linkedin_count,
            'twitter_count': twitter_count,
            'instagram_count': instagram_count,
            'youtube_count': youtube_count,
            'tiktok_count': tiktok_count,
            'total_social_count': social_count,
            'total_contacts': email_count + phone_count + uncertain_phone_count + social_count
        })
    
    return pd.DataFrame(contact_data)

def organize_by_social_media(df):
    """Organize data by social media platforms."""
    social_media_data = defaultdict(list)
    
    # Track URLs to avoid duplicates
    platform_seen_urls = {
        'facebook': set(),
        'linkedin': set(),
        'twitter': set(),
        'instagram': set(),
        'youtube': set(),
        'tiktok': set()
    }
    
    social_platforms = {
        'facebook': [col for col in df.columns if 'facebook' in col],
        'linkedin': [col for col in df.columns if 'linkedin' in col],
        'twitter': [col for col in df.columns if 'twitter' in col or 'x.com' in col],
        'instagram': [col for col in df.columns if 'instagram' in col],
        'youtube': [col for col in df.columns if 'youtube' in col],
        'tiktok': [col for col in df.columns if 'tiktok' in col]
    }
    
    for domain in df['domain'].unique():
        domain_rows = df[df['domain'] == domain]
        
        for platform, columns in social_platforms.items():
            # Collect all non-empty social media links for this domain and platform
            for _, row in domain_rows.iterrows():
                for col in columns:
                    if col in row and pd.notna(row[col]) and row[col] != '':
                        # Split long URLs if needed
                        urls = split_long_urls(row[col])
                        
                        for url_chunk in urls:
                            # Normalize URL for comparison
                            normalized_url = str(url_chunk).strip().lower()
                            
                            # Skip if we've seen this URL before
                            if normalized_url in platform_seen_urls[platform]:
                                continue
                                
                            platform_seen_urls[platform].add(normalized_url)
                            
                            social_media_data[platform].append({
                                'domain': domain,
                                'url': url_chunk
                            })
    
    # Convert defaultdict to dict of dataframes
    return {platform: pd.DataFrame(data) for platform, data in social_media_data.items()}

def organize_by_contact_type(df):
    """Organize data by contact type (email, phone)."""
    seen_emails = set()
    seen_phones = set()
    seen_uncertain_phones = set()
    
    emails = []
    phones = []
    uncertain_phones = []
    
    for _, row in df.iterrows():
        domain = row['domain']
        
        # Extract emails
        for col in df.columns:
            if col.startswith('emails_') and pd.notna(row[col]) and row[col] != '':
                # Normalize email for comparison
                normalized_email = str(row[col]).strip().lower()
                
                # Skip if we've seen this email before
                email_key = f"{domain}:{normalized_email}"
                if email_key in seen_emails:
                    continue
                    
                seen_emails.add(email_key)
                
                emails.append({
                    'domain': domain,
                    'email': row[col]
                })
        
        # Extract phones
        for col in df.columns:
            if col.startswith('phones_') and pd.notna(row[col]) and row[col] != '':
                # Normalize phone for comparison
                normalized_phone = str(row[col]).strip().replace('-', '').replace('(', '').replace(')', '').replace(' ', '').replace('.', '')
                
                # Skip if we've seen this phone before
                phone_key = f"{domain}:{normalized_phone}"
                if phone_key in seen_phones:
                    continue
                    
                seen_phones.add(phone_key)
                
                phones.append({
                    'domain': domain,
                    'phone': row[col]
                })
            elif col.startswith('phonesuncertain_') and pd.notna(row[col]) and row[col] != '':
                # Normalize phone for comparison
                normalized_phone = str(row[col]).strip().replace('-', '').replace('(', '').replace(')', '').replace(' ', '').replace('.', '')
                
                # Skip if we've seen this phone before
                phone_key = f"{domain}:{normalized_phone}"
                if phone_key in seen_uncertain_phones:
                    continue
                    
                seen_uncertain_phones.add(phone_key)
                
                uncertain_phones.append({
                    'domain': domain,
                    'phone': row[col]
                })
    
    return {
        'emails': pd.DataFrame(emails),
        'phones': pd.DataFrame(phones),
        'uncertain_phones': pd.DataFrame(uncertain_phones)
    }

def organize_by_domain(df):
    """Organize data by domain with detailed contact information."""
    domains_data = []
    
    for domain in df['domain'].unique():
        domain_rows = df[df['domain'] == domain]
        
        # Get all emails for this domain (with deduplication)
        emails = set()
        for _, row in domain_rows.iterrows():
            for col in [c for c in df.columns if c.startswith('emails_')]:
                if pd.notna(row.get(col)) and row.get(col) != '':
                    emails.add(str(row[col]).strip())
        
        # Get all phones for this domain (with deduplication)
        phones = set()
        for _, row in domain_rows.iterrows():
            for col in [c for c in df.columns if c.startswith('phones_')]:
                if pd.notna(row.get(col)) and row.get(col) != '':
                    phones.add(str(row[col]).strip())
        
        # Get all social media for this domain (with deduplication)
        social_media = {}
        for platform in ['facebook', 'linkedin', 'twitter', 'instagram', 'youtube', 'tiktok']:
            platform_cols = [col for col in df.columns if platform in col.lower()]
            platform_urls = set()
            
            for _, row in domain_rows.iterrows():
                for col in platform_cols:
                    if pd.notna(row.get(col)) and row.get(col) != '':
                        # Split long URLs if needed
                        urls = split_long_urls(row[col])
                        platform_urls.update(urls)
            
            if platform_urls:
                # Join URLs that fit within Excel's limit
                social_media[platform] = split_long_urls(', '.join(platform_urls))
        
        # Create a base record
        base_record = {
            'domain': domain,
            'emails': ', '.join(sorted(emails)),
            'phones': ', '.join(sorted(phones))
        }
        
        # Add social media data, potentially creating multiple records if needed
        max_social_chunks = 1
        for platform in ['facebook', 'linkedin', 'twitter', 'instagram', 'youtube', 'tiktok']:
            if platform in social_media:
                max_social_chunks = max(max_social_chunks, len(social_media[platform]))
        
        # Create records with split social media data
        for i in range(max_social_chunks):
            record = base_record.copy()
            for platform in ['facebook', 'linkedin', 'twitter', 'instagram', 'youtube', 'tiktok']:
                if platform in social_media and i < len(social_media[platform]):
                    record[platform] = social_media[platform][i]
                else:
                    record[platform] = ''
            domains_data.append(record)
    
    return pd.DataFrame(domains_data)

def format_excel_sheet(worksheet, dataframe, header_format, data_format=None):
    """Format an Excel worksheet with appropriate styling."""
    # Format headers
    for col_num, value in enumerate(dataframe.columns):
        worksheet.write(0, col_num, str(value), header_format)
    
    # Auto-filter for headers
    last_col = len(dataframe.columns) - 1
    if last_col >= 0:
        worksheet.autofilter(0, 0, 0, last_col)
    
    # Set freeze panes to keep headers visible
    worksheet.freeze_panes(1, 0)
    
    # Auto-fit columns
    for i, col in enumerate(dataframe.columns):
        # Set a minimum width
        width = max(15, len(str(col)) + 2)
        # Increase width for certain column types
        if 'email' in str(col).lower() or 'url' in str(col).lower() or any(s in str(col).lower() for s in ['facebook', 'linkedin', 'twitter', 'instagram']):
            width = max(width, 40)
        worksheet.set_column(i, i, width)

def create_domain_summary_sheet(df, writer):
    """Create a summary sheet with all domains and their contact info counts."""
    summary_df = get_contact_counts(df)
    
    # Sort by total contacts (descending)
    summary_df = summary_df.sort_values('total_contacts', ascending=False)
    
    # Write to Excel
    summary_df.to_excel(writer, sheet_name='Domain_Summary', index=False)
    
    # Format the worksheet
    worksheet = writer.sheets['Domain_Summary']
    header_format = writer.book.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'bg_color': '#D7E4BC',
        'border': 1
    })
    format_excel_sheet(worksheet, summary_df, header_format)
    
    return summary_df

def process_excel_to_excel(input_file, output_file=None, formats=None):
    """Process the Excel file and create multiple Excel sheets with different organizations."""
    if output_file is None:
        # Create output filename based on input filename
        input_filename = os.path.basename(input_file)
        output_filename = f"{os.path.splitext(input_filename)[0]}_contact_info_organized.xlsx"
        output_file = os.path.join('output', output_filename)
    
    if formats is None:
        formats = ['all']
    
    # Read the Excel file
    df = pd.read_excel(input_file)
    
    # Clean column names
    df = clean_column_names(df)
    
    # Create output directory if it doesn't exist
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Create Excel writer
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'bg_color': '#D7E4BC',
            'border': 1
        })
        
        # Order of sheets:
        # 1. Domain consolidated (commented out)
        # 2. Contact emails
        # 3. Contact phones
        # 4. Social media (LinkedIn, Instagram, others)
        # 5. Domain summary
        # 6. Original data
        
        # Organize by domain - consolidated view (commented out)
        # if 'all' in formats or 'domain' in formats:
        #     domain_df = organize_by_domain(df)
        #     domain_df.to_excel(writer, sheet_name='Domain_Consolidated', index=False)
        #     ws = writer.sheets['Domain_Consolidated']
        #     format_excel_sheet(ws, domain_df, header_format)
        
        # Organize by contact type
        if 'all' in formats or 'contact' in formats:
            contact_dfs = organize_by_contact_type(df)
            
            # First emails
            if 'emails' in contact_dfs:
                sheet_name = 'Contact_Emails'
                contact_dfs['emails'].to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                format_excel_sheet(ws, contact_dfs['emails'], header_format)
            
            # Then phones
            if 'phones' in contact_dfs:
                sheet_name = 'Contact_Phones'
                contact_dfs['phones'].to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                format_excel_sheet(ws, contact_dfs['phones'], header_format)
            
            # Then uncertain phones
            if 'uncertain_phones' in contact_dfs:
                sheet_name = 'Contact_Uncertain_phones'
                contact_dfs['uncertain_phones'].to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                format_excel_sheet(ws, contact_dfs['uncertain_phones'], header_format)
        
        # Organize by social media
        if 'all' in formats or 'social' in formats:
            social_media_dfs = organize_by_social_media(df)
            
            # Define preferred order for social media
            social_media_order = ['linkedin', 'instagram', 'facebook', 'twitter', 'youtube', 'tiktok']
            
            # Process social media in the specified order
            for platform in social_media_order:
                if platform in social_media_dfs:
                    sheet_name = f'Social_{platform.capitalize()}'
                    social_media_dfs[platform].to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    format_excel_sheet(ws, social_media_dfs[platform], header_format)
            
            # Add any remaining platforms not in the preferred order
            for platform, platform_df in social_media_dfs.items():
                if platform not in social_media_order:
                    sheet_name = f'Social_{platform.capitalize()}'
                    platform_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    format_excel_sheet(ws, platform_df, header_format)
        
        # Domain summary sheet (now towards the end)
        summary_df = create_domain_summary_sheet(df, writer)
        
        # Include original data last if requested
        if 'all' in formats or 'original' in formats:
            df.to_excel(writer, sheet_name='Original_Data', index=False)
            ws = writer.sheets['Original_Data']
            format_excel_sheet(ws, df, header_format)
    
    print(f"Excel file created successfully: {output_file}")
    return output_file

def process_all_excel_files():
    """Process all Excel files in the excels folder."""
    # Create excels directory if it doesn't exist
    os.makedirs('excels', exist_ok=True)
    
    # Create output directory if it doesn't exist
    os.makedirs('output', exist_ok=True)
    
    # Get all Excel files from excels folder
    excel_files = [f for f in os.listdir('excels') if f.endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print("No Excel files found in the 'excels' folder.")
        print("Please place your Excel files in the 'excels' folder and try again.")
        return
    
    for excel_file in excel_files:
        input_path = os.path.join('excels', excel_file)
        print(f"\nProcessing: {excel_file}")
        try:
            process_excel_to_excel(input_path)
        except Exception as e:
            print(f"Error processing {excel_file}: {str(e)}")

def main():
    """Main function to run the script."""
    parser = argparse.ArgumentParser(description='Process Excel files to organized contact information sheets')
    parser.add_argument('--formats', type=str, nargs='+', 
                        choices=['all', 'original', 'domain', 'social', 'contact'],
                        default=['all'],
                        help='Formats to include (default: all)')
    
    args = parser.parse_args()
    
    process_all_excel_files()
    print("\nProcessing complete. Check the 'output' folder for results.")

if __name__ == "__main__":
    main() 