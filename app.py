import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from sqlalchemy import create_engine, text, inspect
import datetime

# --- Page Configuration ---
st.set_page_config(page_title="Master Project Tracker", page_icon="📊", layout="wide")
st.title("📊 Master Infrastructure & Activity Tracker")

# --- Database Connection Setup ---
db_url = st.secrets.get("DATABASE_URL", "sqlite:///project_data.db")
engine = create_engine(db_url)

# --- Admin: Initialize Multi-Tab Database ---
with st.sidebar.expander("🛠️ Admin: Database Controls"):
    st.write("Upload your multi-tab Excel file (`RTMC_Trackers.xlsx`).")
    uploaded_file = st.file_uploader("Upload Excel Tracker", type=["xlsx"])
    
    # Let the admin decide where the headers are (default 10 for the legend)
    header_row = st.number_input("How many legend rows to skip?", min_value=0, max_value=50, value=10, help="0 if headers are on Row 1. 10 if headers are on Row 11.")
    
    if uploaded_file and st.button("🚀 Initialize / Refresh All Tabs"):
        try:
            xls = pd.ExcelFile(uploaded_file)
            
            # Save the list of tabs to a master index table
            pd.DataFrame({'tab_name': xls.sheet_names}).to_sql('app_tabs_index', con=engine, if_exists='replace', index=False)
            
            # Loop through all tabs and save them individually
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
                
                # Intelligent Refresh: Merge existing user updates if the data has changed
                try:
                    existing_df = pd.read_sql(f"SELECT * FROM {safe_table_name}", con=engine)
                    pk = df_init.columns[0]
                    if pk in existing_df.columns:
                        status_map = existing_df.set_index(pk)['Project Status'].to_dict()
                        update_map = existing_df.set_index(pk)['Weekly Update'].to_dict()
                        df_init['Project Status'] = df_init[pk].map(status_map).fillna(df_init['Project Status'])
                        df_init['Weekly Update'] = df_init[pk].map(update_map).fillna(df_init['Weekly Update'])
                except Exception:
                    pass
                
                df_init.to_sql(safe_table_name, con=engine, if_exists="replace", index=False)
                
                # Update progress
                progress_bar.progress((i + 1) / len(xls.sheet_names))
                
            st.success("✅ All tabs successfully initialized & refreshed! Please refresh the page.")
        except Exception as e:
            st.error(f"❌ Error initializing database:\n\n{e}")

    st.markdown("---")
    st.write("**Danger Zone**")
    if st.button("🗑️ Clear Entire Database"):
        try:
            inspector = inspect(engine)
            with engine.begin() as conn:
                for table_name in inspector.get_table_names():
                    conn.execute(text(f"DROP TABLE {table_name}"))
            st.success("✅ Database cleared!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Error clearing database: {e}")

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
    
    st.sidebar.markdown("---")
    st.sidebar.write("**Tab Management**")
    if st.sidebar.button(f"🗑️ Delete '{selected_tab}' Data"):
        try:
            with engine.begin() as conn:
                conn.execute(text(f"DROP TABLE {safe_table_name}"))
                conn.execute(text(f"DELETE FROM app_tabs_index WHERE tab_name = :t"), {"t": selected_tab})
            st.sidebar.success(f"Deleted {selected_tab}")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error deleting tab: {e}")

    # --- Dynamic Feature Detection ---
    has_category = 'Implementation Category' in df.columns
    
    # Identify dynamic columns
    status_col = next((c for c in df.columns if str(c).strip().lower() in ['status', 'project status', 'action status', 'task status']), None)
    deadline_col = next((c for c in df.columns if str(c).strip().lower() in ['deadline', 'due date', 'target date']), None)
    
    # Process date columns
    date_cols = [c for c in df.columns if 'date' in str(c).lower() or 'deadline' in str(c).lower()]
    for dc in date_cols:
        df[dc] = pd.to_datetime(df[dc], errors='coerce').dt.date

    # Calculate Overdue Automatically
    if deadline_col:
        today = pd.Timestamp.now().normalize().date()
        
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

    tab1, tab2, tab3, tab4 = st.tabs(["✍️ Collaborate & Update", "📋 Kanban Board", "📊 Dashboard", "📄 Export Sheet"])
    
    # -- TAB 1: COLLABORATE --
    with tab1:
        st.header(f"Editing: {selected_tab}")
        st.write("Changes made here are saved to the live database for all users.")
        
        # Configure Dropdowns and Dates
        col_config = {}
        for dc in date_cols:
            col_config[dc] = st.column_config.DateColumn(dc, format="YYYY-MM-DD")
            
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

    # -- TAB 2: KANBAN --
    with tab2:
        st.subheader(f"Kanban Board: {selected_tab}")
        if status_col:
            statuses = df[status_col].fillna("No Status").unique()
            if len(statuses) > 0:
                cols = st.columns(len(statuses))
                for i, status in enumerate(statuses):
                    with cols[i]:
                        st.markdown(f"### {status}")
                        status_df = df[df[status_col].fillna("No Status") == status]
                        for _, row in status_df.iterrows():
                            title_col = df.columns[0]
                            with st.container(border=True):
                                st.write(f"**{row[title_col]}**")
                                if deadline_col and pd.notna(row.get(deadline_col)):
                                    st.caption(f"📅 {row[deadline_col]}")
                                owner_col = next((c for c in df.columns if 'Assigned' in str(c) or 'Owner' in str(c)), None)
                                if owner_col and pd.notna(row.get(owner_col)):
                                    st.caption(f"👤 {row[owner_col]}")
            else:
                st.info("No tasks to display.")
        else:
            st.info("No status column found for Kanban board.")

    # -- TAB 3: DASHBOARD --
    with tab3:
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
                fig1 = px.pie(df, names=status_col, title="Status Distribution", hole=0.4,
                              color_discrete_sequence=px.colors.qualitative.Pastel)
                fig1.update_traces(textposition='inside', textinfo='percent+label')
                fig1.update_layout(showlegend=False, margin=dict(t=40, b=0, l=0, r=0))
                st.plotly_chart(fig1, use_container_width=True)
                
            with col_chart2:
                if has_category:
                    status_by_cat = df.groupby(['Implementation Category', status_col]).size().reset_index(name='Count')
                    fig2 = px.bar(status_by_cat, x='Implementation Category', y='Count', color=status_col, 
                                  title="Status by Category", text_auto=True,
                                  color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig2.update_layout(barmode='stack', margin=dict(t=40, b=0, l=0, r=0))
                    st.plotly_chart(fig2, use_container_width=True)
                elif 'Assigned To' in df.columns or 'Owner' in df.columns:
                    # Fallback chart for Action trackers
                    owner_col = 'Assigned To' if 'Assigned To' in df.columns else 'Owner'
                    status_by_owner = df.groupby([owner_col, status_col]).size().reset_index(name='Count')
                    fig2 = px.bar(status_by_owner, x=owner_col, y='Count', color=status_col, 
                                  title=f"Status by {owner_col}", text_auto=True,
                                  color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig2.update_layout(barmode='stack', margin=dict(t=40, b=0, l=0, r=0))
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.info("Additional charts are hidden because this tab doesn't have an 'Implementation Category' or 'Owner' column.")
        else:
            st.info("This tab does not have a recognizable Status column, so no dashboard is available.")

    # -- TAB 4: EXPORT --
    with tab4:
        st.subheader(f"Export Data for '{selected_tab}'")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            export_df = edited_df.copy()
            if "Overdue Status" in export_df.columns:
                export_df = export_df.drop(columns=["Overdue Status"])
            export_df.to_excel(writer, index=False, sheet_name=selected_tab[:30])
            
        st.download_button(
            label="📥 Download current tab as Excel",
            data=output.getvalue(),
            file_name=f"{selected_tab}_tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
