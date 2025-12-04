import os

# File Paths
TRACKER_FILE = "Full_Workplan_Tracker.xlsx"
SHEET_NAME = "All_Tasks"
BACKUP_DIR = "backups"

# Ensure backup directory exists
if not os.path.exists(BACKUP_DIR):
    os.makedirs(BACKUP_DIR)
