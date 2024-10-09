import os
import pandas as pd

onedrive_path = "C:\\Users\\" 

def search_name_in_excel(file_path, first_name, last_name):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None)

        for sheet_name, df in all_sheets.items():
            first_name_found = df.apply(lambda col: col.astype(str).str.contains(first_name, case=False, na=False)).any().any()
            last_name_found = df.apply(lambda col: col.astype(str).str.contains(last_name, case=False, na=False)).any().any()

            if first_name_found and last_name_found:
                return True
            elif last_name_found:
                return True

    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return False  

def search_names_in_onedrive(names_list, output_file):
    with open(output_file, 'w') as f:
        f.write("First Name,Last Name,File Path\n")  
        for root, dirs, files in os.walk(onedrive_path):
            for file in files:
                if file.endswith(".xlsx"):  
                    file_path = os.path.join(root, file)
                    for first_name, last_name in names_list:
                        if search_name_in_excel(file_path, first_name, last_name):
                            f.write(f"{first_name},{last_name},{file_path}\n")
                            print(f"Match found for {first_name} {last_name} in {file_path}")

def load_names_from_excel(name_file_path):
    try:
        names_df = pd.read_excel(name_file_path)

        if 'First Name' in names_df.columns and 'Last Name' in names_df.columns:
            names_list = list(zip(names_df['First Name'], names_df['Last Name']))
            return names_list
        else:
            print("The input Excel file must have 'First Name' and 'Last Name' columns.")
            return []
    except Exception as e:
        print(f"Error reading name file: {e}")
        return []
    
if __name__ == "__main__":
    NAME_FILE_PATH = "C:\\...\\Check_First_Last_Name.xlsx"
    names_to_search = load_names_from_excel(NAME_FILE_PATH)
    
    if names_to_search:
        OUTPUT_FILE = "matched_files.csv"
        search_names_in_onedrive(names_to_search, OUTPUT_FILE)
        print(f"Search completed. Results saved in {OUTPUT_FILE}.")
    else:
        print("No names to search.")
