import pandas as pd
import os
import sys

# Redirect stdout to a file
sys.stdout = open("analysis_results_v2.txt", "w", encoding="utf-8")

file_path = r"WorkPlan/ACMS-HIV CHASAC WorkPlan-FY26-COP25 Updated 18.11.25.xlsx"

print(f"Loading {file_path}...")
try:
    xl = pd.ExcelFile(file_path)
    sheet = '4. ACMS WorkPlan detail v1'
    print(f"\n{'='*30}\nSheet: {sheet}\n{'='*30}")
    
    # Read with header=1
    df = xl.parse(sheet, header=1)
    print(f"Columns: {list(df.columns)}")
    
    # Search for SI Manager in all columns
    mask = df.astype(str).apply(lambda x: x.str.contains('SI Manager', case=False, na=False)).any(axis=1)
    si_rows = df[mask]
    
    if not si_rows.empty:
        print(f"\n*** Found {len(si_rows)} rows mentioning 'SI Manager' ***")
        # Print the first few rows to see where it appears
        print(si_rows.head(5).to_string())
        
        # Identify which column contains "SI Manager"
        for col in df.columns:
            if df[col].astype(str).str.contains('SI Manager', case=False, na=False).any():
                print(f"\n'SI Manager' found in column: {col}")
    else:
        print("No 'SI Manager' found in this sheet with header=1.")
        
        # Try searching for just "SI" or "Strategic Information"
        mask_si = df.astype(str).apply(lambda x: x.str.contains('Strategic Information', case=False, na=False)).any(axis=1)
        if mask_si.any():
             print(f"\nFound {mask_si.sum()} rows mentioning 'Strategic Information'")
             print(df[mask_si].head(3).to_string())

except Exception as e:
    print(f"Error reading excel: {e}")
