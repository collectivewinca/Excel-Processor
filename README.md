# Website Scraper with Gemini AI

This project scrapes websites and uses Google's Gemini AI to extract company information from the scraped content.

## Setup Instructions

1. Create a virtual environment:
```bash
# Windows
python -m venv venv

# Linux/Mac
python3 -m venv venv
```

2. Activate the virtual environment:
```bash
# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Create a `.env` file in the project root and add your Google API key:
```
GOOGLE_API_KEY=your_api_key_here
```

## Project Structure

### Files Description

1. `website_scraper.py`
   - Main script that processes Excel files containing website domains
   - Input: Excel file with a 'domain' column containing website URLs
   - Output: New Excel file with added 'company_name' and 'bio' columns
   - Features:
     - Processes multiple sheets in the Excel file
     - Caches results to avoid reprocessing same domains
     - Handles errors gracefully
     - Provides progress updates during processing

2. `excel_processor.py`
   - Processes and organizes contact information from Excel files
   - Input: Raw Excel file with contact information
   - Output: Organized Excel file with multiple sheets:
     - Domain summary with contact counts
     - Social media links organized by platform
     - Contact information (emails, phones) organized by type
   - Features:
     - Deduplicates contact information
     - Handles long URLs and splits them if needed
     - Organizes data by domain, social media, and contact type
     - Maintains Excel formatting and readability

3. `requirements.txt`
   - Lists all Python package dependencies with their versions
   - Used for setting up the project environment

4. `.env`
   - Stores environment variables (API keys)
   - Not tracked in git for security
   - Must be created manually with your Google API key

5. `.gitignore`
   - Specifies which files Git should ignore
   - Prevents sensitive data and unnecessary files from being committed

## Usage

1. Place your input Excel file in the project directory
2. Update the input and output file paths in `website_scraper.py`
3. Run the script:
```bash
python website_scraper.py
```

## Notes

- The script processes each unique domain only once, even if it appears in multiple sheets
- Results are cached to avoid redundant API calls
- Progress updates are printed to the console during processing
- Failed domains are marked with "Error" in the output file 