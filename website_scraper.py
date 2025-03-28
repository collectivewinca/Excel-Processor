import pandas as pd
import requests
from bs4 import BeautifulSoup
import google.generativeai as genai
import time
from urllib.parse import urlparse
import os
from typing import Dict, Optional

# Configure Gemini API
GOOGLE_API_KEY = "AIzaSyBc5SYK2pvvdNq4mobecMaRzobZHdeOr2A"
if not GOOGLE_API_KEY:
    raise ValueError("Please set GOOGLE_API_KEY environment variable")

genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel('gemini-2.0-flash-lite')

def clean_url(url: str) -> str:
    """Clean and normalize URL."""
    if not url.startswith(('http://', 'https://')):
        url = 'https://' + url
    return url

def extract_domain(url: str) -> str:
    """Extract domain from URL."""
    return urlparse(clean_url(url)).netloc

def scrape_website(url: str) -> Optional[str]:
    """Scrape website content."""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(clean_url(url), headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Remove script and style elements
        for script in soup(["script", "style"]):
            script.decompose()
            
        # Get text content
        text = soup.get_text()
        # Clean up text
        lines = (line.strip() for line in text.splitlines())
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        text = ' '.join(chunk for chunk in chunks if chunk)
        
        return text[:5000]  # Limit text length for API
    except Exception as e:
        print(f"Error scraping {url}: {str(e)}")
        return None

def extract_company_info(text: str) -> Dict[str, str]:
    """Extract company name and bio using separate Gemini API calls."""
    
    # Prompt for company name
    name_prompt = f"""
    Extract only the company name from the following text.
    Return just the name without any explanation.
    If you can't find the company name, return "Not found".
    
    Text: {text}
    """
    
    # Prompt for company bio
    bio_prompt = f"""
    Extract a concise company bio/description from the following text (max 400 words).
    Focus on what the company does, their mission, and key offerings.
    Return just the bio without any explanation.
    If you can't find enough information, return "Not found".
    
    Text: {text}
    """
    
    try:
        # Make separate API calls
        name_response = model.generate_content(name_prompt)
        time.sleep(0.5)  # Add small delay between calls
        bio_response = model.generate_content(bio_prompt)
        
        # Extract text from responses
        company_name = name_response.text.strip()
        company_bio = bio_response.text.strip()
        
        return {
            "company_name": company_name,
            "bio": company_bio
        }
    except Exception as e:
        print(f"Error with Gemini API: {str(e)}")
        return {"company_name": "Error", "bio": "Error"}

def process_excel_file(input_file: str, output_file: str):
    """Process Excel file and add company information."""
    # Dictionary to store all processed domains across sheets
    processed_domains = {}  # Store results to avoid reprocessing
    
    # Create Excel writer
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # Get all sheet names from input file
    excel_file = pd.ExcelFile(input_file)
    sheet_names = excel_file.sheet_names
    print(f"Processing {len(sheet_names)} sheets:", sheet_names)
    
    # Process each sheet
    for sheet_name in sheet_names:
        print(f"\nProcessing sheet: {sheet_name}")
        
        # Read the sheet
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Ensure 'domain' column exists in this sheet
        if 'domain' not in df.columns:
            print(f"Warning: 'domain' column not found in sheet {sheet_name}, skipping...")
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            continue
        
        # Add new columns if they don't exist
        if 'company_name' not in df.columns:
            df['company_name'] = ''
        if 'bio' not in df.columns:
            df['bio'] = ''
        
        # Process each unique domain in this sheet
        for idx, row in df.iterrows():
            domain = str(row['domain']).strip()  # Convert to string and clean
            
            # Skip empty domains
            if pd.isna(domain) or not domain:
                continue
            
            # If domain was already processed in any sheet, use cached results
            if domain in processed_domains:
                print(f"Using cached results for domain: {domain}")
                domain_mask = df['domain'].str.strip() == domain
                df.loc[domain_mask, 'company_name'] = processed_domains[domain]['company_name']
                df.loc[domain_mask, 'bio'] = processed_domains[domain]['bio']
                continue
            
            print(f"Processing new domain: {domain}")
            
            # Scrape website
            website_text = scrape_website(domain)
            if website_text:
                # Extract company info
                company_info = extract_company_info(website_text)
                
                # Store results in cache
                processed_domains[domain] = company_info
                
                # Update all rows with this domain
                domain_mask = df['domain'].str.strip() == domain
                df.loc[domain_mask, 'company_name'] = company_info['company_name']
                df.loc[domain_mask, 'bio'] = company_info['bio']
                
                print(f"Updated {domain} with:")
                print(f"Company Name: {company_info['company_name']}")
                print(f"Bio: {company_info['bio'][:100]}...")  # Print first 100 chars of bio
                
                time.sleep(1)  # Rate limiting
            else:
                print(f"Failed to scrape {domain}")
                processed_domains[domain] = {"company_name": "Error - Could not scrape", "bio": "Error - Could not scrape"}
        
        # Save the processed sheet
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Completed processing sheet: {sheet_name}")
    
    # Save and close the Excel file
    writer.close()
    print(f"\nProcessing complete. Results saved to {output_file}")
    
    # Print summary
    print(f"\nProcessed {len(processed_domains)} unique domains across all sheets")
    print(f"Failed domains: {sum(1 for info in processed_domains.values() if 'Error' in info['company_name'])}")

if __name__ == "__main__":
    input_file = "excels/CorporateNYC_contact_info_organized.xlsx"
    output_file = "output2/CorporateNYC_contact_info_organized.xlsx"
    process_excel_file(input_file, output_file) 