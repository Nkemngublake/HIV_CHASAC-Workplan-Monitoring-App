import pandas as pd
import os
import shutil
from datetime import datetime
import streamlit as st
from . import config

class DataManager:
    @staticmethod
    def load_data():
        """Loads data from the Excel file and performs migration if needed."""
        if not os.path.exists(config.TRACKER_FILE):
            return None
        
        try:
            df = pd.read_excel(config.TRACKER_FILE, sheet_name=config.SHEET_NAME)
            # Ensure Comments is string to avoid Streamlit editing errors
            if 'Comments' in df.columns:
                df['Comments'] = df['Comments'].fillna("").astype(str)
            
            # Data Cleaning: Remove empty rows and columns
            df.dropna(how='all', inplace=True) # Drop rows where all elements are NaN
            df.dropna(axis=1, how='all', inplace=True) # Drop columns where all elements are NaN
            
            # Remove rows where 'Program Area' is NaN or empty (critical for stats)
            if 'Program Area' in df.columns:
                df = df[df['Program Area'].notna() & (df['Program Area'].astype(str).str.strip() != '')]
            
            # Check for migration to Regional structure
            if 'Region' not in df.columns:
                df = DataManager._migrate_to_regions(df)
                # Save immediately to persist migration
                DataManager.save_data(df)
                
            return df
        except Exception as e:
            st.error(f"Error loading data: {e}")
            return None

    @staticmethod
    def _migrate_to_regions(df):
        """Splits each task into 3 regions and divides budget."""
        with st.spinner("Migrating data to Regional structure (North, Adamawa, Extreme North)... Please wait."):
            regions = ["North", "Adamawa", "Extreme North"]
            new_rows = []
            
            # Budget columns to split
            q1_col = "Oct -Dec 2025"
            q2_col = "Jan - Mar 2026"
            
            for _, row in df.iterrows():
                # Convert budget to numeric for division
                try:
                    q1_val = pd.to_numeric(row.get(q1_col, 0), errors='coerce')
                    if pd.isna(q1_val): q1_val = 0
                except: q1_val = 0
                
                try:
                    q2_val = pd.to_numeric(row.get(q2_col, 0), errors='coerce')
                    if pd.isna(q2_val): q2_val = 0
                except: q2_val = 0
                
                for region in regions:
                    new_row = row.copy()
                    new_row['Region'] = region
                    new_row[q1_col] = q1_val / 3
                    new_row[q2_col] = q2_val / 3
                    new_row['Last Modified By'] = ''
                    new_row['Last Modified Date'] = ''
                    new_rows.append(new_row)
            
            new_df = pd.DataFrame(new_rows)
            
            # Re-assign IDs to be unique
            new_df['ID'] = range(1, len(new_df) + 1)
            
            # Ensure column order - put Region near the beginning
            cols = list(new_df.columns)
            if 'Region' in cols:
                cols.insert(1, cols.pop(cols.index('Region')))
                new_df = new_df[cols]
                
            return new_df

    @staticmethod
    def create_backup():
        """Creates a timestamped backup of the tracker file."""
        if not os.path.exists(config.TRACKER_FILE):
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"tracker_backup_{timestamp}.xlsx"
        backup_path = os.path.join(config.BACKUP_DIR, backup_filename)
        
        try:
            shutil.copy2(config.TRACKER_FILE, backup_path)
            # Optional: Clean up old backups (keep last 10)
            DataManager._cleanup_old_backups()
        except Exception as e:
            st.warning(f"Failed to create backup: {e}")

    @staticmethod
    def _cleanup_old_backups(keep=10):
        """Keeps only the last `keep` backups."""
        try:
            files = [os.path.join(config.BACKUP_DIR, f) for f in os.listdir(config.BACKUP_DIR) if f.endswith('.xlsx')]
            files.sort(key=os.path.getmtime, reverse=True)
            
            for f in files[keep:]:
                os.remove(f)
        except Exception:
            pass # Fail silently on cleanup

    @staticmethod
    def save_data(df):
        """Saves the dataframe to the Excel file after creating a backup."""
        try:
            # Create backup first
            DataManager.create_backup()
            
            # Save new data
            df.to_excel(config.TRACKER_FILE, index=False, sheet_name=config.SHEET_NAME)
            st.success("Changes saved successfully!")
            st.cache_data.clear() # Clear cache to reload new data
            return True
        except Exception as e:
            st.error(f"Error saving data: {e}")
            return False
