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
    
    # --- Dynamic UI based on columns ---
    has_category = 'Implementation Category' in df.columns
    has_status = 'Project Status' in df.columns
    
    tab1, tab2, tab3 = st.tabs(["✍️ Collaborate & Update", "📊 Dashboard", "📄 Export Sheet"])
    
    # -- TAB 1: COLLABORATE --
    with tab1:
        st.header(f"Editing: {selected_tab}")
        st.write("Changes made here are saved to the live database for all users.")
        
        # Determine column configurations dynamically
        col_config = {}
        if has_status:
            col_config["Project Status"] = st.column_config.SelectboxColumn("Project Status", options=["Not Started", "In Progress", "Completed"])
        if 'Weekly Update' in df.columns:
            col_config["Weekly Update"] = st.column_config.TextColumn("Weekly Update")
            
        edited_df = st.data_editor(df, use_container_width=True, column_config=col_config, hide_index=True)
        
        if st.button("💾 Save Updates to Database"):
            try:
                edited_df.to_sql(safe_table_name, con=engine, if_exists="replace", index=False)
                st.success(f"Updates to '{selected_tab}' successfully saved to the live database!")
            except Exception as e:
                st.error(f"Failed to save: {e}")

    # -- TAB 2: DASHBOARD --
    with tab2:
        if has_status:
            st.subheader(f"Progress Overview: {selected_tab}")
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Items", len(df))
            c2.metric("Completed", len(df[df["Project Status"] == "Completed"]))
            c3.metric("In Progress", len(df[df["Project Status"] == "In Progress"]))
            
            st.markdown("---")
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                fig1 = px.pie(df, names='Project Status', title="Status Distribution", hole=0.3)
                st.plotly_chart(fig1, use_container_width=True)
                
            with col_chart2:
                # If it's the main infrastructure tab, show the category breakdown
                if has_category:
                    status_by_cat = df.groupby(['Implementation Category', 'Project Status']).size().reset_index(name='Count')
                    fig2 = px.bar(status_by_cat, x='Implementation Category', y='Count', color='Project Status', title="Status by Category")
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.info("Chart 'Status by Category' is hidden because this tab does not contain an 'Implementation Category' column.")
        else:
            st.info("This tab does not have a 'Project Status' column, so no dashboard is available.")

    # -- TAB 3: EXPORT --
    with tab3:
        st.subheader(f"Export Data for '{selected_tab}'")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False, sheet_name=selected_tab[:30]) # Sheet names max 31 chars
            
        st.download_button(
            label="📥 Download current tab as Excel",
            data=output.getvalue(),
            file_name=f"{selected_tab}_tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
