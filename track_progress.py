import pandas as pd
import os
import sys
from openpyxl import load_workbook

TRACKER_FILE = "Full_Workplan_Tracker.xlsx"
SHEET_NAME = "All_Tasks"

def load_data():
    if not os.path.exists(TRACKER_FILE):
        print(f"Error: Tracker file not found at {TRACKER_FILE}")
        return None
    return pd.read_excel(TRACKER_FILE, sheet_name=SHEET_NAME)

def save_data(df):
    try:
        # We use openpyxl to save to preserve formatting if possible, 
        # but pandas to_excel overwrites. 
        # For simplicity in this CLI, we'll overwrite and re-apply basic formatting logic 
        # or just use pandas for data updates.
        # To avoid losing formatting, we can use openpyxl directly for updates, 
        # but pandas is easier for querying.
        # Let's just overwrite for now, the extraction script handles initial formatting.
        df.to_excel(TRACKER_FILE, index=False, sheet_name=SHEET_NAME)
        print("Changes saved successfully.")
    except PermissionError:
        print("Error: Could not save file. Please close Excel if it is open.")

def show_summary(df):
    total = len(df)
    status_counts = df['Status'].value_counts()
    
    print("\n" + "="*40)
    print("       WORKPLAN PROGRESS SUMMARY       ")
    print("="*40)
    print(f"Total Tasks: {total}")
    for status, count in status_counts.items():
        print(f"{status}: {count} ({count/total*100:.1f}%)")
    print("="*40 + "\n")

def list_tasks(df, filter_status=None):
    if filter_status:
        tasks = df[df['Status'].astype(str).str.lower() == filter_status.lower()]
    else:
        tasks = df
        
    if tasks.empty:
        print("No tasks found.")
        return

    print(f"\nListing {len(tasks)} tasks:")
    print("-" * 100)
    # Adjust columns for display
    display_cols = ['ID', 'Status', 'Activities', 'Oct -Dec 2025', 'Jan - Mar 2026']
    
    # Truncate long text
    display_df = tasks[display_cols].copy()
    display_df['Activities'] = display_df['Activities'].astype(str).apply(lambda x: x[:60] + '...' if len(x) > 60 else x)
    
    print(display_df.to_string(index=False))
    print("-" * 100)

def update_task(df):
    try:
        task_id = int(input("Enter Task ID to update: "))
        if task_id not in df['ID'].values:
            print("Invalid Task ID.")
            return df
        
        print("\nCurrent Status:", df.loc[df['ID'] == task_id, 'Status'].values[0])
        print("1. Pending")
        print("2. In Progress")
        print("3. Completed")
        print("4. Delayed")
        
        choice = input("Select new status (1-4): ")
        status_map = {'1': 'Pending', '2': 'In Progress', '3': 'Completed', '4': 'Delayed'}
        
        if choice in status_map:
            df.loc[df['ID'] == task_id, 'Status'] = status_map[choice]
            
            # Optional: Update Progress %
            if status_map[choice] == 'Completed':
                df.loc[df['ID'] == task_id, 'Progress (%)'] = 100
            elif status_map[choice] == 'Pending':
                df.loc[df['ID'] == task_id, 'Progress (%)'] = 0
            
            comment = input("Add a comment (optional): ")
            if comment:
                current_comment = str(df.loc[df['ID'] == task_id, 'Comments'].values[0])
                if current_comment == 'nan': current_comment = ""
                if current_comment:
                    new_comment = current_comment + " | " + comment
                else:
                    new_comment = comment
                df.loc[df['ID'] == task_id, 'Comments'] = new_comment
                
            save_data(df)
        else:
            print("Invalid choice.")
            
    except ValueError:
        print("Invalid input.")
    
    return df

def main():
    print("Welcome to the SI Manager Workplan Tracker")
    df = load_data()
    if df is None:
        return

    while True:
        print("\nMenu:")
        print("1. Show Summary")
        print("2. List All Tasks")
        print("3. List Pending Tasks")
        print("4. Update Task Status")
        print("5. Exit")
        
        choice = input("Enter choice: ")
        
        if choice == '1':
            show_summary(df)
        elif choice == '2':
            list_tasks(df)
        elif choice == '3':
            list_tasks(df, filter_status='Pending')
        elif choice == '4':
            df = update_task(df)
        elif choice == '5':
            print("Goodbye!")
            break
        else:
            print("Invalid choice.")

if __name__ == "__main__":
    main()
