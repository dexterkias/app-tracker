import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from sqlalchemy import create_engine
import datetime

# --- Page Configuration ---
st.set_page_config(page_title="Master Project Tracker", page_icon="📊", layout="wide")
st.title("📊 Master Infrastructure & Activity Tracker")

# --- Database Connection Setup ---
db_url = st.secrets.get("DATABASE_URL", "sqlite:///project_data.db")
engine = create_engine(db_url)

# --- Admin: Initialize Multi-Tab Database ---
with st.sidebar.expander("🛠️ Admin: Initialize Database"):
    st.write("Upload your multi-tab Excel file (`RTMC_Trackers.xlsx`).")
    uploaded_file = st.file_uploader("Upload Excel Tracker", type=["xlsx"])
    
    # Let the admin decide where the headers are (default 10 for the legend)
    header_row = st.number_input("How many legend rows to skip?", min_value=0, max_value=50, value=10, help="0 if headers are on Row 1. 10 if headers are on Row 11.")
    
    if uploaded_file and st.button("🚀 Initialize All Tabs"):
        try:
            xls = pd.ExcelFile(uploaded_file)
            
            # Save the list of tabs to a master index table
            pd.DataFrame({'tab_name': xls.sheet_names}).to_sql('app_tabs_index', con=engine, if_exists='replace', index=False)
            
            # Loop through all 13 tabs and save them individually
            progress_bar = st.progress(0)
            for i, sheet in enumerate(xls.sheet_names):
                df_init = pd.read_excel(xls, sheet_name=sheet, skiprows=header_row)
                
                # Automatically add tracking columns if they don't exist
                if "Project Status" not in df_init.columns:
                    df_init["Project Status"] = "Not Started"
                if "Weekly Update" not in df_init.columns:
                    df_init["Weekly Update"] = ""
                
                # Sanitize table name for SQL
                safe_table_name = "tab_" + "".join([c if c.isalnum() else "_" for c in sheet]).lower()
                df_init.to_sql(safe_table_name, con=engine, if_exists="replace", index=False)
                
                # Update progress
                progress_bar.progress((i + 1) / len(xls.sheet_names))
                
            st.success("✅ All tabs successfully initialized! Please refresh the page.")
        except Exception as e:
            st.error(f"❌ Error initializing database:\n\n{e}")

# --- Load Available Tabs ---
try:
    tabs_df = pd.read_sql("SELECT tab_name FROM app_tabs_index", con=engine)
    available_tabs = tabs_df['tab_name'].tolist()
except Exception:
    available_tabs = []

if not available_tabs:
    st.info("👈 Please use the Admin panel in the sidebar to upload `RTMC_Trackers.xlsx` and initialize the database.")
else:
    # --- Sidebar Tab Navigation ---
    st.sidebar.subheader("Navigation")
    selected_tab = st.sidebar.selectbox("📂 Select Tracker Sheet", available_tabs)
    
    # Load the selected tab's data
    safe_table_name = "tab_" + "".join([c if c.isalnum() else "_" for c in selected_tab]).lower()
    df = pd.read_sql(f"SELECT * FROM {safe_table_name}", con=engine)
    
    # --- Dynamic Feature Detection ---
    has_category = 'Implementation Category' in df.columns
    
    # Identify dynamic columns
    status_col = next((c for c in df.columns if str(c).strip().lower() in ['status', 'project status', 'action status', 'task status']), None)
    deadline_col = next((c for c in df.columns if str(c).strip().lower() in ['deadline', 'due date', 'target date']), None)
    
    # Calculate Overdue Automatically
    if deadline_col:
        df[deadline_col] = pd.to_datetime(df[deadline_col], errors='coerce')
        today = pd.Timestamp.now().normalize()
        
        def calc_overdue(row):
            # If completed, it's not overdue
            if status_col and pd.notna(row.get(status_col)):
                v = str(row[status_col]).lower()
                if 'complete' in v or 'close' in v or 'done' in v:
                    return "Completed"
            
            # Check date
            if pd.isna(row.get(deadline_col)):
                return ""
            days = (today - row[deadline_col]).days
            if days > 0:
                return f"🚨 {days} Days Overdue"
            return "On Track"
            
        df["Overdue Status"] = df.apply(calc_overdue, axis=1)

    tab1, tab2, tab3 = st.tabs(["✍️ Collaborate & Update", "📊 Dashboard", "📄 Export Sheet"])
    
    # -- TAB 1: COLLABORATE --
    with tab1:
        st.header(f"Editing: {selected_tab}")
        st.write("Changes made here are saved to the live database for all users.")
        
        # Configure Dropdowns
        col_config = {}
        if status_col:
            existing_opts = [x for x in df[status_col].dropna().unique()]
            standard_opts = ["Not Started", "Open", "In Progress", "Delayed", "Completed", "Closed"]
            options = list(set(existing_opts + standard_opts))
            col_config[status_col] = st.column_config.SelectboxColumn(status_col, options=options)
            
        if 'Weekly Update' in df.columns:
            col_config["Weekly Update"] = st.column_config.TextColumn("Weekly Update")
            
        if deadline_col and "Overdue Status" in df.columns:
            col_config["Overdue Status"] = st.column_config.TextColumn("Overdue Status", disabled=True)

        # Apply Colors
        def style_dataframe(styler):
            if status_col:
                def color_status(val):
                    v = str(val).lower()
                    if 'complete' in v or 'close' in v or 'done' in v:
                        return 'background-color: #d4edda; color: #155724; font-weight: bold;'
                    elif 'progress' in v or 'ongoing' in v or 'open' in v:
                        return 'background-color: #cce5ff; color: #004085; font-weight: bold;'
                    elif 'delay' in v or 'overdue' in v:
                        return 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
                    elif 'not start' in v or 'new' in v:
                        return 'background-color: #e2e3e5; color: #383d41;'
                    return ''
                styler = styler.map(color_status, subset=[status_col])
                
            if deadline_col and "Overdue Status" in df.columns:
                def color_overdue(val):
                    if '🚨' in str(val):
                        return 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
                    elif val == 'On Track' or val == 'Completed':
                        return 'color: #155724;'
                    return ''
                styler = styler.map(color_overdue, subset=['Overdue Status'])
            return styler

        # Display Data Editor with styling
        styled_df = df.style.pipe(style_dataframe)
        edited_df = st.data_editor(styled_df, use_container_width=True, column_config=col_config, hide_index=True)
        
        if st.button("💾 Save Updates to Database"):
            try:
                # Strip out the styling and dynamic calculation before saving back to SQL
                # The user's edits are in edited_df
                save_df = edited_df.copy()
                if deadline_col and "Overdue Status" in save_df.columns:
                    save_df.drop(columns=["Overdue Status"], inplace=True)
                
                save_df.to_sql(safe_table_name, con=engine, if_exists="replace", index=False)
                st.success(f"Updates to '{selected_tab}' successfully saved to the live database!")
                st.rerun() # Refresh to recalculate overdue metrics
            except Exception as e:
                st.error(f"Failed to save: {e}")

    # -- TAB 2: DASHBOARD --
    with tab2:
        if status_col:
            st.subheader(f"Progress Overview: {selected_tab}")
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Items", len(df))
            
            # Dynamic completion matching
            completed_count = len(df[df[status_col].astype(str).str.lower().str.contains('complete|close|done', na=False)])
            in_progress_count = len(df[df[status_col].astype(str).str.lower().str.contains('progress|ongoing|open', na=False)])
            
            c2.metric("Completed/Closed", completed_count)
            c3.metric("In Progress/Open", in_progress_count)
            
            st.markdown("---")
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                fig1 = px.pie(df, names=status_col, title="Status Distribution", hole=0.3)
                st.plotly_chart(fig1, use_container_width=True)
                
            with col_chart2:
                if has_category:
                    status_by_cat = df.groupby(['Implementation Category', status_col]).size().reset_index(name='Count')
                    fig2 = px.bar(status_by_cat, x='Implementation Category', y='Count', color=status_col, title="Status by Category")
                    st.plotly_chart(fig2, use_container_width=True)
                elif 'Assigned To' in df.columns or 'Owner' in df.columns:
                    # Fallback chart for Action trackers
                    owner_col = 'Assigned To' if 'Assigned To' in df.columns else 'Owner'
                    status_by_owner = df.groupby([owner_col, status_col]).size().reset_index(name='Count')
                    fig2 = px.bar(status_by_owner, x=owner_col, y='Count', color=status_col, title=f"Status by {owner_col}")
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.info("Additional charts are hidden because this tab doesn't have an 'Implementation Category' or 'Owner' column.")
        else:
            st.info("This tab does not have a recognizable Status column, so no dashboard is available.")

    # -- TAB 3: EXPORT --
    with tab3:
        st.subheader(f"Export Data for '{selected_tab}'")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False, sheet_name=selected_tab[:30])
            
        st.download_button(
            label="📥 Download current tab as Excel",
            data=output.getvalue(),
            file_name=f"{selected_tab}_tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
