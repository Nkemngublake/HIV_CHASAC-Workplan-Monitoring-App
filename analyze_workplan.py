import pandas as pd
import os
import sys

# Redirect stdout to a file
sys.stdout = open("analysis_results.txt", "w", encoding="utf-8")

file_path = r"WorkPlan/ACMS-HIV CHASAC WorkPlan-FY26-COP25 Updated 18.11.25.xlsx"

if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
    file_path = r"c:\Users\hp\Documents\CHASAC Workplan Agent\WorkPlan\ACMS-HIV CHASAC WorkPlan-FY26-COP25 Updated 18.11.25.xlsx"
    if not os.path.exists(file_path):
        print(f"File still not found at {file_path}")
        exit(1)

    print(f"Loading {file_path}...")
    try:
        from tqdm import tqdm
        
        xl = pd.ExcelFile(file_path)
        print(f"Sheet names: {xl.sheet_names}")
    
        for sheet in tqdm(xl.sheet_names, desc="Analyzing Sheets", unit="sheet"):
            print(f"\n{'='*30}\nSheet: {sheet}\n{'='*30}")
            try:
                df = xl.parse(sheet)
                print(f"Columns: {list(df.columns)}")
                
                # Search for SI Manager
                mask = df.astype(str).apply(lambda x: x.str.contains('SI Manager', case=False, na=False)).any(axis=1)
                si_rows = df[mask]
                
                if not si_rows.empty:
                    print(f"\n*** Found {len(si_rows)} rows mentioning 'SI Manager' in {sheet} ***")
                    # Print all columns for these rows to understand the context
                    print(si_rows.to_string())
                else:
                    print(f"No 'SI Manager' found in {sheet}")
                    
                # Check for Region info
                if 'Region' in df.columns:
                    print(f"\nUnique Regions in {sheet}: {df['Region'].unique()}")
                    
            except Exception as e:
                print(f"Error parsing sheet {sheet}: {e}")

except Exception as e:
    print(f"Error reading excel: {e}")
