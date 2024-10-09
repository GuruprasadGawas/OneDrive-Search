import os
import pandas as pd
import re

onedrive_path = "C:\\Users\\...\\OneDrive"  

def search_linkedin_in_excel(file_path, linkedin_urls):
    matches = []  
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None)

        for sheet_name, df in all_sheets.items():
            for col in df.columns:
                for linkedin_url in linkedin_urls:
                    matches_in_col = df[col].astype(str).str.contains(re.escape(linkedin_url), case=False, na=False)
                    
                    if matches_in_col.any():
                        matches.append((linkedin_url, file_path, sheet_name, col))
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    
    return matches  

def load_linkedin_urls(linkedin_excel_path):
    try:
        df = pd.read_excel(linkedin_excel_path)
        linkedin_urls = df['LinkedIn URL'].dropna().tolist()  
        return linkedin_urls
    except Exception as e:
        print(f"Error reading LinkedIn URLs from {linkedin_excel_path}: {e}")
        return []

def search_linkedin_in_onedrive(linkedin_excel_path, output_file):
    linkedin_urls = load_linkedin_urls(linkedin_excel_path)
    
    if not linkedin_urls:
        print("No LinkedIn URLs found in the input file.")
        return

    with open(output_file, 'w') as f:
        f.write("LinkedIn URL,File Path,Sheet Name,Column Name\n")  

        for root, dirs, files in os.walk(onedrive_path):
            for file in files:
                if file.endswith(".xlsx"):  
                    file_path = os.path.join(root, file)
                    matches = search_linkedin_in_excel(file_path, linkedin_urls)
                    
                    for linkedin_url, file_path, sheet_name, col_name in matches:
                        f.write(f"{linkedin_url},{file_path},{sheet_name},{col_name}\n")
                        print(f"Match found: {linkedin_url} in {file_path} (Sheet: {sheet_name}, Column: {col_name})")

if __name__ == "__main__":
    LINKEDIN_URL_EXCEL = "C:\\...\\Check_LinkedIn_URL.xlsx"  
    OUTPUT_FILE = "linkedin_matches.csv"
    
    search_linkedin_in_onedrive(LINKEDIN_URL_EXCEL, OUTPUT_FILE)

    print(f"Search completed. Results saved in {OUTPUT_FILE}.")
