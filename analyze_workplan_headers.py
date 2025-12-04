import pandas as pd
import os
import sys

# Redirect stdout to a file
sys.stdout = open("analysis_results_headers.txt", "w", encoding="utf-8")

file_path = r"WorkPlan/ACMS-HIV CHASAC WorkPlan-FY26-COP25 Updated 18.11.25.xlsx"

print(f"Loading {file_path}...")
try:
    xl = pd.ExcelFile(file_path)
    sheet = '4. ACMS WorkPlan detail v1'
    print(f"\n{'='*30}\nSheet: {sheet}\n{'='*30}")
    
    # Read first 20 rows without header to find the real header
    df = xl.parse(sheet, header=None, nrows=20)
    print(df.to_string())

except Exception as e:
    print(f"Error reading excel: {e}")
