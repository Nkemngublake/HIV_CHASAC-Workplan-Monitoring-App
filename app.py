import streamlit as st
from modules.data_manager import DataManager
from modules.ui import setup_page, render_metrics, render_filters, render_data_editor, render_financial_summary, render_login
from modules import config
from datetime import datetime

def main():
    # 1. Setup Page
    setup_page()
    
    # 2. Login Check
    if not render_login():
        return

    st.title("📊 CHASAC Workplan Tracker")
    st.markdown("---")

    # 3. Load Data
    df = DataManager.load_data()
    if df is None:
        st.error(f"Tracker file not found: {config.TRACKER_FILE}. Please run extraction script first.")
        return

    # 4. Sidebar Filters
    filtered_df, selected_region = render_filters(df)

    # 5. Metrics
    render_metrics(df, filtered_df)
    
    # 6. Financial Summary
    render_financial_summary(filtered_df)
    
    st.markdown("---")

    # 7. Data Editor
    edited_df = render_data_editor(filtered_df)

    # 8. Save Logic
    if st.button("Save Changes", type="primary"):
        # Map changes back to original dataframe using ID
        updated_rows = edited_df.set_index('ID')
        
        # Iterate and update
        changes_made = False
        for idx, row in updated_rows.iterrows():
            mask = df['ID'] == idx
            if mask.any():
                current_row = df.loc[mask].iloc[0]
                has_changed = (
                    current_row['Status'] != row['Status'] or
                    current_row['Progress (%)'] != row['Progress (%)'] or
                    current_row['Comments'] != row['Comments']
                )
                
                if has_changed:
                    df.loc[mask, 'Status'] = row['Status']
                    df.loc[mask, 'Progress (%)'] = row['Progress (%)']
                    df.loc[mask, 'Comments'] = row['Comments']
                    
                    # Update Tracking Info
                    df.loc[mask, 'Last Modified By'] = st.session_state.user_email
                    df.loc[mask, 'Last Modified Date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    changes_made = True
        
        if changes_made:
            DataManager.save_data(df)
        else:
            st.info("No changes detected.")

if __name__ == "__main__":
    main()
