import os
import sys
import PyPDF2
import pandas as pd
import re
from typing import List, Dict

def clean_cyrillic(text: str) -> str:
    """Remove Cyrillic characters from text."""
    return re.sub(r'[\u0400-\u04FF\u0500-\u052F]', '', text)

def extract_invoice_number(text: str) -> str:
    """Extract invoice number from text following specific pattern."""
    pattern = r'Invoice.*?№\s*(\d+\s*/\s*\d+)'
    match = re.search(pattern, text)
    if match:
        return f"№ {clean_cyrillic(match.group(1)).strip()}"
    return ""

def extract_invoice_date(text: str) -> str:
    """Extract date in format dd.MM.yyyy."""
    pattern = r'Date of invoice.*?(\d{2}\.\d{2}\.\d{4})'
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return ""

def extract_supplier(text: str) -> str:
    """Extract supplier name until the first comma."""
    pattern = r'Supplier\s+([^,]+)'  # Using \s+ to ensure at least one whitespace
    match = re.search(pattern, text)
    if match:
        supplier_text = match.group(1).strip()
        return clean_cyrillic(supplier_text)
    return ""

def extract_price(text: str) -> str:
    """Extract price number."""
    pattern = r'Price of the services rendered:\s*(\d+\.?\d*)'
    match = re.search(pattern, text)
    if match:
        return match.group(1).strip()
    return ""

def extract_currency(text: str) -> str:
    """Extract currency word after 'Currency:'."""
    pattern = r'Currency:\s*(\w+)'
    match = re.search(pattern, text)
    if match:
        return clean_cyrillic(match.group(1)).strip()
    return ""

def process_match(keyword: str, matched_text: str) -> str:
    """Process matched text according to keyword-specific rules."""
    if keyword.startswith("Invoice"):
        return extract_invoice_number(matched_text)
    elif keyword.startswith("Date of invoice:"):
        return extract_invoice_date(matched_text)
    elif keyword == "Supplier":
        return extract_supplier(matched_text)
    elif keyword == "Price of the services rendered:":
        return extract_price(matched_text)
    elif keyword == "Currency:":
        return extract_currency(matched_text)
    return clean_cyrillic(matched_text)

def search_pdf(pdf_path: str, keywords: List[str]) -> Dict:
    """Search a PDF file for given keywords and return all matches in a single row."""
    result = {
        'file_path': pdf_path,
        'date_of_invoice': '',
        'supplier': '',
        'invoice': '',
        'currency': '',
        'price_of_the_services_rendered': ''
    }
    supplier_found = False
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            full_text = ""
            
            for page in pdf_reader.pages:
                full_text += page.extract_text()
            
            for keyword in keywords:
                if keyword == "Supplier" and supplier_found:
                    continue
                    
                index = 0
                text_lower = full_text.lower()
                keyword_lower = keyword.lower()
                
                while True:
                    index = text_lower.find(keyword_lower, index)
                    if index == -1:
                        break
                    
                    context = get_context(full_text[index:], keyword)
                    if context:  # Only add non-empty matches
                        # Map keyword to specific column
                        if keyword.startswith("Invoice"):
                            result['invoice'] = context
                        elif keyword.startswith("Date of invoice:"):
                            result['date_of_invoice'] = context
                        elif keyword == "Supplier":
                            result['supplier'] = context
                            supplier_found = True
                            break
                        elif keyword == "Price of the services rendered:":
                            result['price_of_the_services_rendered'] = context
                        elif keyword == "Currency:":
                            result['currency'] = context
                    
                    index += len(keyword)
        
        return result
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
        return result

def get_context(text: str, keyword: str) -> str:
    """Extract and process text after the keyword match."""
    keyword_pos = text.lower().find(keyword.lower())
    if keyword_pos == -1:
        return ""
    
    full_context = text[keyword_pos:keyword_pos + 200].replace('\n', ' ').strip()
    processed = process_match(keyword, full_context)
    return processed if processed else ""  # Return empty string if no match


def find_pdf_files(directory: str) -> List[str]:
    """Recursively find all PDF files in the given directory."""
    pdf_files = []
    
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    return pdf_files

def main():
    if len(sys.argv) < 3:
        print("\nPDF Parser Tool")
        print("==============")
        print("\nUsage:")
        print("pdf_parser.exe <directory> <output.csv>")
        print("\nExample:")
        print('pdf_parser.exe "C:\\Documents\\PDFs" "output.csv"')
        print("\nPress Enter to exit...")
        input()
        sys.exit(1)
    
    directory = sys.argv[1]
    output_file = sys.argv[2]
    
    # Predefined keywords
    keywords = [
        "Invoice",
        "Date of invoice:",
        "Supplier",
        "Price of the services rendered:",
        "Currency:"
    ]
    
    pdf_files = find_pdf_files(directory)
    print(f"\nFound {len(pdf_files)} PDF files to search.")
    
    all_results = []
    for pdf_file in pdf_files:
        print(f"Searching {pdf_file}...")
        result = search_pdf(pdf_file, keywords)
        all_results.append(result)
    
    if all_results:
        df = pd.DataFrame(all_results)
        # Reorder columns to match desired output
        columns = ['file_path', 'date_of_invoice', 'supplier', 'invoice', 'currency', 'price_of_the_services_rendered']
        df = df[columns]
        df.to_csv(output_file, index=False)
        print(f"\nSearch complete! Found {len(all_results)} matches.")
        print(f"Results saved to {output_file}")
    else:
        print("\nNo matches found.")

if __name__ == "__main__":
    main()