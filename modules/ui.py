import streamlit as st
import pandas as pd

def setup_page():
    """Configures the Streamlit page and adds custom CSS."""
    st.set_page_config(
        page_title="CHASAC Workplan Tracker",
        page_icon="📊",
        layout="wide"
    )
    
    # Custom CSS for a more premium look
    st.markdown("""
        <style>
        .stApp {
            background-color: #f8f9fa;
        }
        .metric-card {
            background-color: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        /* Improve dataframe styling */
        .stDataFrame {
            border: 1px solid #e0e0e0;
            border-radius: 8px;
        }
        /* Header styling */
        h1 {
            color: #1f2937;
            font-family: 'Inter', sans-serif;
        }
        </style>
        """, unsafe_allow_html=True)

def render_metrics(original_df, filtered_df=None):
    """Displays key metrics, optionally comparing filtered vs overall."""
    
    # If no filtered df provided or it's the same as original, just show original stats
    if filtered_df is None or len(filtered_df) == len(original_df):
        st.subheader("📊 Overall Status")
        _display_metrics_row(original_df)
        
        # Show breakdown by Program Area
        st.markdown("#### Status by Program Area")
        breakdown = original_df.groupby(['Program Area', 'Status']).size().unstack(fill_value=0)
        breakdown['Total'] = breakdown.sum(axis=1)
        breakdown = breakdown.sort_values('Total', ascending=False)
        st.dataframe(breakdown, use_container_width=True)
        
    else:
        # Show Filtered Stats
        st.subheader("📊 Current Selection Status")
        _display_metrics_row(filtered_df)
        
        st.markdown("---")
        
        # Show Overall Stats for context
        with st.expander("View Overall Status"):
            _display_metrics_row(original_df)

def _display_metrics_row(df):
    """Helper to display a row of metrics for a dataframe."""
    col1, col2, col3, col4 = st.columns(4)
    
    total_tasks = len(df)
    completed_tasks = len(df[df['Status'] == 'Completed'])
    in_progress = len(df[df['Status'] == 'In Progress'])
    delayed = len(df[df['Status'] == 'Delayed'])
    
    with col1:
        st.metric("Total Tasks", total_tasks)
    with col2:
        st.metric("Completed", completed_tasks, delta=f"{completed_tasks/total_tasks*100:.1f}%" if total_tasks > 0 else "0%")
    with col3:
        st.metric("In Progress", in_progress)
    with col4:
        st.metric("Delayed", delayed, delta_color="inverse")

def render_login():
    """Renders a login screen."""
    if 'user_email' not in st.session_state:
        st.session_state.user_email = None

    if not st.session_state.user_email:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("## 🔒 Login Required")
            st.info("Please enter your email address to access the tracker. This is used to track changes.")
            email = st.text_input("Email Address")
            if st.button("Login", type="primary"):
                if email and "@" in email:
                    st.session_state.user_email = email
                    st.rerun()
                else:
                    st.error("Please enter a valid email address.")
        return False
    
    # Show logged in user in sidebar
    st.sidebar.markdown(f"👤 **{st.session_state.user_email}**")
    if st.sidebar.button("Logout"):
        st.session_state.user_email = None
        st.rerun()
        
    return True

def render_filters(df):
    """Renders sidebar filters and returns the filtered dataframe and selected region."""
    st.sidebar.header("Filters")
    
    # Region Filter (New)
    # Ensure Region column exists (it should after migration)
    if 'Region' in df.columns:
        region_options = ["All"] + sorted(df['Region'].unique().tolist())
        selected_region = st.sidebar.selectbox("Region", region_options)
    else:
        selected_region = "All"

    # Status Filter
    status_options = ["All"] + sorted(df['Status'].unique().tolist())
    selected_status = st.sidebar.selectbox("Status", status_options)
    
    # Program Area Filter
    # Strip whitespace to avoid duplicates and remove 'nan'
    df['Program Area'] = df['Program Area'].astype(str).str.strip()
    unique_areas = sorted([a for a in df['Program Area'].unique().tolist() if a.lower() != 'nan' and a != ''])
    area_options = ["All"] + unique_areas
    selected_area = st.sidebar.selectbox("Program Area", area_options)
    
    # Search
    search_query = st.sidebar.text_input("Search Activities", "")
    
    # Apply Filters
    filtered_df = df.copy()
    
    if selected_region != "All":
        filtered_df = filtered_df[filtered_df['Region'] == selected_region]
    
    if selected_status != "All":
        filtered_df = filtered_df[filtered_df['Status'] == selected_status]
        
    if selected_area != "All":
        filtered_df = filtered_df[filtered_df['Program Area'] == selected_area]
        
    if search_query:
        filtered_df = filtered_df[filtered_df['Activities'].astype(str).str.contains(search_query, case=False, na=False) | 
                                  filtered_df['ACMS Sub-Activities'].astype(str).str.contains(search_query, case=False, na=False)]
                                  
    return filtered_df, selected_region

def render_financial_summary(df):
    """Displays financial summary statistics."""
    st.subheader("💰 Financial Summary")
    
    # Ensure numeric columns
    q1_col = "Oct -Dec 2025"
    q2_col = "Jan - Mar 2026"
    
    # Convert to numeric, coercing errors to NaN then 0
    df[q1_col] = pd.to_numeric(df[q1_col], errors='coerce').fillna(0)
    df[q2_col] = pd.to_numeric(df[q2_col], errors='coerce').fillna(0)
    
    total_q1 = df[q1_col].sum()
    total_q2 = df[q2_col].sum()
    total_budget = total_q1 + total_q2
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Budget", f"${total_budget:,.2f}")
    with col2:
        st.metric("Q1 Budget", f"${total_q1:,.2f}")
    with col3:
        st.metric("Q2 Budget", f"${total_q2:,.2f}")
        
    # Per Program Area Summary
    st.markdown("#### Budget by Program Area")
    area_summary = df.groupby('Program Area')[[q1_col, q2_col]].sum()
    area_summary['Total'] = area_summary[q1_col] + area_summary[q2_col]
    area_summary = area_summary.sort_values('Total', ascending=False)
    
    # Format for display
    display_summary = area_summary.copy()
    display_summary[q1_col] = display_summary[q1_col].map('${:,.2f}'.format)
    display_summary[q2_col] = display_summary[q2_col].map('${:,.2f}'.format)
    display_summary['Total'] = display_summary['Total'].map('${:,.2f}'.format)
    
    st.dataframe(display_summary, use_container_width=True)


def render_data_editor(df):
    """Renders the editable dataframe."""
    st.subheader(f"Tasks ({len(df)})")
    
    column_config = {
        "ID": st.column_config.NumberColumn("ID", disabled=True, width="small"),
        "Status": st.column_config.SelectboxColumn(
            "Status",
            options=["Pending", "In Progress", "Completed", "Delayed"],
            required=True,
            width="medium"
        ),
        "Progress (%)": st.column_config.ProgressColumn(
            "Progress",
            min_value=0,
            max_value=100,
            format="%d%%"
        ),
        "Activities": st.column_config.TextColumn("Activity", width="large", disabled=True),
        "ACMS Sub-Activities": st.column_config.TextColumn("Sub-Activity", width="large", disabled=True),
        "Comments": st.column_config.TextColumn("Comments", width="large"),
        "Oct -Dec 2025": st.column_config.NumberColumn("Q1 Budget", width="small", format="$%.2f"),
        "Jan - Mar 2026": st.column_config.NumberColumn("Q2 Budget", width="small", format="$%.2f"),
    }
    
    display_cols = ['ID', 'Status', 'Progress (%)', 'Comments', 'Program Area', 'Activities', 'ACMS Sub-Activities', 'Oct -Dec 2025', 'Jan - Mar 2026']
    
    # Ensure columns exist before selecting
    available_cols = [c for c in display_cols if c in df.columns]
    
    edited_df = st.data_editor(
        df[available_cols],
        column_config=column_config,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="data_editor"
    )
    
    return edited_df
