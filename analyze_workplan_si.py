import pandas as pd
import os
import sys

# Redirect stdout to a file
sys.stdout = open("analysis_results_si.txt", "w", encoding="utf-8")

file_path = r"WorkPlan/ACMS-HIV CHASAC WorkPlan-FY26-COP25 Updated 18.11.25.xlsx"

print(f"Loading {file_path}...")
try:
    xl = pd.ExcelFile(file_path)
    sheet = '4. ACMS WorkPlan detail v1'
    print(f"\n{'='*30}\nSheet: {sheet}\n{'='*30}")
    
    df = xl.parse(sheet, header=1)
    
    # Check Program Areas
    if 'Program Area' in df.columns:
        print(f"\nUnique Program Areas: {df['Program Area'].unique()}")
        
        # Filter for SI related Program Areas
        si_program_areas = df[df['Program Area'].astype(str).str.contains('SI|Strategic Information|M&E|Data', case=False, na=False)]
        if not si_program_areas.empty:
            print(f"\n*** Found {len(si_program_areas)} rows with SI-related Program Area ***")
            print(si_program_areas[['Program Area', 'ACMS Sub-Activities']].head(10).to_string())
            
    # Filter for Objective 7 (row index might vary, but we can search in Activities column)
    obj7_mask = df['Activities'].astype(str).str.contains('Objective 7', case=False, na=False)
    if obj7_mask.any():
        print(f"\n*** Found Objective 7 ***")
        # We want to see the activities UNDER this objective. 
        # Usually objectives are headers and activities follow. 
        # But here let's just see where it is.
        print(df[obj7_mask][['Activities']].to_string())
        
    # Search for "Data" or "SI" in ACMS Sub-Activities
    keyword_mask = df['ACMS Sub-Activities'].astype(str).str.contains('Data|SI |Strategic Information|M&E|DQA', case=False, na=False)
    keyword_rows = df[keyword_mask]
    if not keyword_rows.empty:
        print(f"\n*** Found {len(keyword_rows)} rows with SI keywords in Sub-Activities ***")
        print(keyword_rows[['ACMS Sub-Activities', 'Program Area']].head(10).to_string())

except Exception as e:
    print(f"Error reading excel: {e}")
