import os
import pandas as pd
import re

# Specify the local path to your OneDrive folder
onedrive_path = "C:\\Users\\...\\OneDrive"

# Function to load LinkedIn URLs from an Excel file
def load_linkedin_urls(linkedin_excel_path):
    try:
        df = pd.read_excel(linkedin_excel_path)
        linkedin_urls = df['LinkedIn URL'].dropna().tolist()  # Drop missing values
        return linkedin_urls
    except Exception as e:
        print(f"Error reading LinkedIn URLs from {linkedin_excel_path}: {e}")
        return []

# Function to search for LinkedIn URLs and delete matched rows in Excel files
def search_and_delete_linkedin_rows(linkedin_urls, output_file):
    # Walk through the OneDrive folder and its subfolders
    for root, dirs, files in os.walk(onedrive_path):
        for file in files:
            if file.endswith(".xlsx"):  # Only consider Excel files
                file_path = os.path.join(root, file)
                try:
                    # Load all sheets from the Excel file
                    all_sheets = pd.read_excel(file_path, sheet_name=None)
                    changes_made = False
                    
                    # Prepare an Excel writer to save the updated sheets
                    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        # Loop through each sheet
                        for sheet_name, df in all_sheets.items():
                            original_shape = df.shape
                            
                            # Loop through each column in the sheet and check for LinkedIn URLs
                            for col in df.columns:
                                for linkedin_url in linkedin_urls:
                                    # Find matches for the LinkedIn URL in the column
                                    matches_in_col = df[col].astype(str).str.contains(re.escape(linkedin_url), case=False, na=False)
                                    
                                    if matches_in_col.any():
                                        print(f"Match found and deleting row with {linkedin_url} in file: {file_path}, sheet: {sheet_name}, column: {col}")
                                        # Drop rows with matching LinkedIn URL
                                        df = df[~matches_in_col]
                                        changes_made = True
                                        
                            # If any changes were made, write the updated DataFrame back to the sheet
                            if changes_made and df.shape != original_shape:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    if changes_made:
                        print(f"Changes saved for file: {file_path}")
                
                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")

# Example usage
if __name__ == "__main__":
    # Path to the LinkedIn URLs Excel file
    linkedin_urls_file = "C:\\...\\Delete.xlsx"
    
    # Load LinkedIn URLs from the Excel file
    linkedin_urls = load_linkedin_urls(linkedin_urls_file)
    
    # Call the function to search and delete matched rows
    search_and_delete_linkedin_rows(linkedin_urls, "output_log.txt")

    print("Process completed.")

