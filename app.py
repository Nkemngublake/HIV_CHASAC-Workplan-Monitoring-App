import streamlit as st
from modules.data_manager import DataManager
from modules.ui import setup_page, render_metrics, render_filters, render_data_editor, render_financial_summary
from modules import config

def main():
    # 1. Setup Page
    setup_page()
    st.title("📊 CHASAC Workplan Tracker")
    st.markdown("---")

    # 2. Load Data
    df = DataManager.load_data()
    if df is None:
        st.error(f"Tracker file not found: {config.TRACKER_FILE}. Please run extraction script first.")
        return

    # 3. Sidebar Filters
    filtered_df = render_filters(df)

    # 4. Metrics
    # 4. Metrics
    render_metrics(df, filtered_df)
    
    # 5. Financial Summary
    render_financial_summary(filtered_df)
    
    st.markdown("---")

    # 6. Data Editor
    edited_df = render_data_editor(filtered_df)

    # 6. Save Logic
    if st.button("Save Changes", type="primary"):
        # Map changes back to original dataframe using ID
        updated_rows = edited_df.set_index('ID')
        
        # Iterate and update
        # Note: This could be optimized for very large datasets, but works for <1000 rows
        changes_made = False
        for idx, row in updated_rows.iterrows():
            mask = df['ID'] == idx
            if mask.any():
                # Check for changes before assigning to avoid unnecessary updates if we wanted to be strict
                # But for now, just update
                df.loc[mask, 'Status'] = row['Status']
                df.loc[mask, 'Progress (%)'] = row['Progress (%)']
                df.loc[mask, 'Comments'] = row['Comments']
                changes_made = True
        
        if changes_made:
            DataManager.save_data(df)

if __name__ == "__main__":
    main()
