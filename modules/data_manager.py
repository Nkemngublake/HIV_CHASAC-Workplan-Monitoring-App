import pandas as pd
import os
import shutil
from datetime import datetime
import streamlit as st
from . import config

class DataManager:
    @staticmethod
    def load_data():
        """Loads data from the Excel file."""
        if not os.path.exists(config.TRACKER_FILE):
            return None
        
        try:
            df = pd.read_excel(config.TRACKER_FILE, sheet_name=config.SHEET_NAME)
            # Ensure Comments is string to avoid Streamlit editing errors
            if 'Comments' in df.columns:
                df['Comments'] = df['Comments'].fillna("").astype(str)
            return df
        except Exception as e:
            st.error(f"Error loading data: {e}")
            return None

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
