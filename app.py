import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from fpdf import FPDF
import datetime
from sqlalchemy import create_engine

# --- Page Configuration ---
st.set_page_config(page_title="Project Tracker", page_icon="📊", layout="wide")
st.title("📊 Collaborative ICT Infrastructure Tracker")

# --- Database Connection Setup ---
# This looks for a cloud database URL in Streamlit Secrets, otherwise defaults to a local SQLite file.
db_url = st.secrets.get("DATABASE_URL", "sqlite:///project_data.db")
engine = create_engine(db_url)


# --- Helper Functions ---
def load_data():
    try:
        # Load live data from the database
        return pd.read_sql("SELECT * FROM project_items", con=engine)
    except Exception:
        return pd.DataFrame()  # Return empty if table doesn't exist yet


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Project Status')
    return output.getvalue()


def to_pdf(df):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Project Progress Report", ln=True, align='C')
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 10, f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "Status Summary:", ln=True)
    pdf.set_font("Arial", size=10)
    status_counts = df['Project Status'].value_counts()
    for status, count in status_counts.items():
        pdf.cell(0, 8, f"- {status}: {count} items", ln=True)

    pdf.ln(10)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "Actionable / In Progress Items:", ln=True)
    pdf.set_font("Arial", size=9)
    actionable = df[df['Project Status'] != 'Completed']
    for idx, row in actionable.iterrows():
        item = str(row['Asset Category & Model'])[:50]
        status = row['Project Status']
        pdf.cell(0, 8, f"> {item} | Status: {status}", ln=True)

    return pdf.output(dest='S').encode('latin-1')


# --- Admin: Initialize Database ---
with st.sidebar.expander("🛠️ Admin: Initialize Database"):
    st.write("Upload the original CSV to populate the database for the first time.")
    uploaded_file = st.file_uploader("Upload CSV", type="csv")
    
    if uploaded_file and st.button("Initialize DB"):
        try:
            # 1. Try UTF-8 with signature (safely handles hidden Byte-Order-Marks)
            df_init = pd.read_csv(uploaded_file, skiprows=10, encoding='utf-8-sig')
        except UnicodeDecodeError:
            # 2. Bulletproof Fallback: read with latin1 and force-replace any broken characters
            uploaded_file.seek(0)  # Reset the file pointer
            df_init = pd.read_csv(
                uploaded_file, 
                skiprows=10, 
                encoding='latin1', 
                encoding_errors='replace'
            )
            
        df_init["Project Status"] = "Not Started"
        df_init["Weekly Update"] = ""
        
        # Save to centralized database
        df_init.to_sql("project_items", con=engine, if_exists="replace", index=False)
        st.success("Database Initialized! Refresh the page.")
        st.rerun()

# --- Main Application UI ---
df = load_data()

if not df.empty:
    st.sidebar.subheader("Filters")
    categories = df['Implementation Category'].dropna().unique().tolist()
    selected_cat = st.sidebar.multiselect("Filter by Category", categories, default=categories)
    filtered_df = df[df['Implementation Category'].isin(selected_cat)]

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "📋 Kanban", "✍️ Collaborate & Update", "📄 Export"])

    # -- TAB 1: DASHBOARD --
    with tab1:
        st.header("Live Project Dashboard")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Line Items", len(filtered_df))
        col2.metric("Completed Items", len(filtered_df[filtered_df["Project Status"] == "Completed"]))
        col3.metric("Items In Progress", len(filtered_df[filtered_df["Project Status"] == "In Progress"]))

        st.markdown("---")
        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.pie(filtered_df, names='Project Status', title="Overall Status", hole=0.3)
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            status_by_cat = filtered_df.groupby(['Implementation Category', 'Project Status']).size().reset_index(
                name='Count')
            fig2 = px.bar(status_by_cat, x='Implementation Category', y='Count', color='Project Status',
                          title="Status by Category")
            st.plotly_chart(fig2, use_container_width=True)

    # -- TAB 2: KANBAN --
    with tab2:
        st.header("Activity Kanban")
        k_col1, k_col2, k_col3 = st.columns(3)
        statuses = ["Not Started", "In Progress", "Completed"]
        columns = [k_col1, k_col2, k_col3]

        for col, status in zip(columns, statuses):
            with col:
                st.subheader(f"{status} ({len(filtered_df[filtered_df['Project Status'] == status])})")
                items = filtered_df[filtered_df['Project Status'] == status]
                for idx, row in items.iterrows():
                    st.info(
                        f"**{row['Asset Category & Model']}**\n\n*Qty:* {row['Qty ']}\n\n*Note:* {row['Weekly Update']}")

    # -- TAB 3: COLLABORATIVE DATA ENTRY --
    with tab3:
        st.header("Team Updates")
        st.write("Edit the table below. Clicking 'Save to Database' updates the tracker for all team members globally.")

        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            column_config={
                "Project Status": st.column_config.SelectboxColumn("Project Status",
                                                                   options=["Not Started", "In Progress", "Completed"]),
                "Weekly Update": st.column_config.TextColumn("Weekly Update")
            },
            hide_index=True
        )

        if st.button("💾 Save Updates to Database"):
            # Update the global database so all members see changes
            edited_df.to_sql("project_items", con=engine, if_exists="replace", index=False)
            st.success("Updates successfully saved to the live database!")
            st.rerun()

    # -- TAB 4: EXPORT --
    with tab4:
        st.header("Download Reports")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📥 Download Excel", data=to_excel(df), file_name="Live_Project_Tracker.xlsx")
        with col2:
            st.download_button("📥 Download PDF", data=to_pdf(df), file_name="Live_Project_Summary.pdf")
else:
    st.warning("⚠️ Database is empty. Please use the Admin menu in the sidebar to upload the initial CSV file.")
