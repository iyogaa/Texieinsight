import pandas as pd
import streamlit as st
from thefuzz import fuzz, process
import re
import io
from all_trans_mvr import all_trans_mvr_app
from mvr_gpt import mvr_gpt_app
from qc_radar import qc_radar_app
from insight_dashboard import insight_dashboard_app
from processor import PDFTextSearcher
import os
import tempfile

# Set page configuration
st.set_page_config(layout="wide")

# Apply custom styling
st.markdown("""
<style>
/* Main page background pure black */
.stApp {
    background-color: #000000;
}
/* Sidebar background medium black */
[data-testid="stSidebar"] {
    background-color: #111111 !important;
}
/* Sidebar text bright white but now in normal font */
[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
    font-weight: normal;
}
/* Left-aligned Heading with smaller font and no border */
.custom-heading {
    font-size: 2rem;
    color: white;
    text-align: left;
    font-weight: bold;
    margin-bottom: 1.5rem;
    margin-left: 2rem;
    background: none;
    border: none;
    padding: 0;
}
/* Remove extra empty box inside file uploader */
[data-testid="stFileUploader"] > div {
    background-color: transparent !important;
    padding: 0 !important;
    margin: 0 !important;
    border: none !important;
    min-height: 0 !important;
    min-width: 0 !important;
}
/* Label and input text white */
label, .stFileUploader, .stNumberInput label, .stSelectbox label {
    color: white !important;
}
/* White text for all content */
body, .stMarkdown, .stText, .stDataFrame, .stMetric {
    color: white !important;
}
/* Custom button styling */
.stButton>button {
    background-color: #000000;
    color: white;
    border-radius: 5px;
    padding: 0.5rem 1rem;
    font-weight: bold;
}
/* Status indicator */
.status-indicator {
    display: inline-block;
    width: 10px;
    height: 10px;
    border-radius: 50%;
    margin-right: 5px;
}
.status-operational {
    background-color: #4CAF50;
}
</style>
""", unsafe_allow_html=True)

# --- Authentication System ---
credentials = {
    "yogaraj": {"password": "afreen", "role": "ADMIN"},
    "Maha": {"password": "Maha@129", "role": "QA"},
    "Gokul": {"password": "reddead", "role": "QA"},
    "user": {"password": "ssapopb", "role": "MAKER"},
    "bharti_sawan": {"password": "sawan@agoy", "role": "QA"},
}

# --- Authentication Function ---
def authenticate(username, password):
    if username in credentials and password == credentials[username]["password"]:
        return credentials[username]["role"]
    return None

# --- Initialize Session State ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["username"] = None
    st.session_state["role"] = None

# --- Show Login if Not Authenticated ---
def show_login():
    with st.sidebar:
        st.title("🔐 Login")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.button("Login"):
            role = authenticate(username, password)
            if role:
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.session_state["role"] = role
                st.rerun()
            else:
                st.error("Invalid username or password")

if not st.session_state["authenticated"]:
    show_login()
    st.stop()

# --- Role-based Menu Generator ---
def get_menu_options(role):
    base = ["QC Radar", "All Trans MVR", "Supplement", "Riscom MVR", "MVR GPT"]
    if role == "ADMIN":
        return base + ["Insight Dashboard"]
    elif role == "QA":
        return base
    elif role == "MAKER":
        return ["MVR GPT"]
    return []

# --- Sidebar Layout (Everything Inside) ---
with st.sidebar:
    st.markdown(f"### 👋 Welcome, **{st.session_state['username']}**")
    st.markdown(f"**Role:** {st.session_state['role']}")
    st.markdown("---")

    menu_options = get_menu_options(st.session_state["role"])
    if menu_options:
        menu = st.radio("📋 Menu", menu_options, label_visibility="collapsed")
    else:
        st.warning("No menu options available.")
        menu = None

    st.markdown("---")
    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()
    st.caption("Built with Yogaraj ")

# --- Main Application Logic ---
if menu == "All Trans MVR":
    all_trans_mvr_app()
# Welcome screen for "App" menu
# HDVI MVR tool
elif menu == "QC Radar":
    qc_radar_app()
elif menu == "MVR GPT":
    mvr_gpt_app()
elif menu == "Alltrans Test":
    import streamlit as st
    import pandas as pd
    import io
    from datetime import datetime
    from alltrans_test import alltrans

    def main():
        st.set_page_config(page_title="Alltrans Test", page_icon="🚛", layout="wide")

        obj = alltrans()

        
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("MVR Data File")
            mvr_file = st.file_uploader("Upload MVR Report", type=["xlsx", "xls"], key="mvr_upload")
        
        with col2:
            st.subheader("Client Driver List")
            client_file = st.file_uploader("Upload Client Driver List", type=["xlsx", "xls", "csv"], key="client_upload")

        if not mvr_file or not client_file:
            st.info("Please upload both MVR and Client files to proceed")
            return

        
        # MVR: Fixed 1 row skip (first row not required)
        mvr_skip = 1
        # Client: User input for rows to skip
        client_skip = st.number_input("Rows to skip in CLIENT file", min_value=0, value=0, key="client_skip")
        st.info(f"📋 Client file will skip {client_skip} rows")

        if st.button("Process Matching", type="primary", key="process_button"):
            try:
                # Load MVR file with fixed 1 row skip, use "MVR" sheet
                mvr_data = pd.read_excel(mvr_file, skiprows=mvr_skip, sheet_name="MVR")
                
                # Load Client file with user-defined rows skip
                if client_file.name.endswith('.csv'):
                    client_data = pd.read_csv(client_file, skiprows=client_skip)
                else:
                    client_data = pd.read_excel(client_file, skiprows=client_skip)

                st.success(f"MVR Data loaded: {len(mvr_data)} records (skipped {mvr_skip} row)")
                st.success(f"Client Data loaded: {len(client_data)} drivers (skipped {client_skip} rows)")

                # Show previews
                col1, col2 = st.columns(2)
                with col1:
                    with st.expander("MVR Data Preview"):
                        st.write("Columns found:", list(mvr_data.columns))
                        st.dataframe(mvr_data.head(5))
                with col2:
                    with st.expander("Client Data Preview"):
                        st.write("Columns found:", list(client_data.columns))
                        st.dataframe(client_data.head(5))

                # Auto-detect columns in client file
                st.subheader("Detected Columns in Client File")
                client_name_col = obj.get_valid_column(client_data, "driver names", ['name', 'driver_name', 'full_name','Name'])
                hire_date_col = obj.get_valid_column(client_data, "hire dates", ['hire_date', 'date_of_hire', 'doh'], False)
                dob_col = obj.get_valid_column(client_data, "date of birth", ['dob', 'date_of_birth', 'birth_date'], False)
                license_col = obj.get_valid_column(client_data, "license state", ['license_state', 'lic_state', 'state'], False)

                st.write("**Detected Columns:**")
                st.write(f"- Driver Name: `{client_name_col}`")
                st.write(f"- Hire Date: `{hire_date_col}`" if hire_date_col else "- Hire Date: Not found")
                st.write(f"- Date of Birth: `{dob_col}`" if dob_col else "- Date of Birth: Not found")
                st.write(f"- License State: `{license_col}`" if license_col else "- License State: Not found")

                # Process the data
                with st.spinner("🔄 Matching drivers and creating All Trans sheet..."):
                    # Extract drivers from MVR
                    mvr_drivers = obj.extract_drivers_from_mvr(mvr_data)
                    
                    st.info(f"📊 Found {len(mvr_drivers)} unique drivers in MVR data")
                    
                    # CORRECTED METHOD CALL: Using match_drivers (not match_drivers_with_hire_date)
                    all_trans_df = obj.match_drivers(
                        client_data,      # client data
                        mvr_drivers,      # MVR drivers  
                        hire_date_col,    # hire date column
                        dob_col,          # DOB column
                        license_col       # license state column
                    )

                    st.success("Processing completed successfully!")
                    
                    # Display results
                    st.header("Results Summary")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    license_matches = len(all_trans_df[all_trans_df['Comment'].str.contains('License')])
                    name_dob_matches = len(all_trans_df[all_trans_df['Comment'].str.contains('Name+DOB')])
                    missing_mvr = len(all_trans_df[all_trans_df['Comment'] == 'MISSING MVR'])
                    extra_mvr = len(all_trans_df[all_trans_df['Comment'] == 'Extra MVR record (no client match)'])
                    total_records = len(all_trans_df)
                    
                    col1.metric("Total Records", total_records)
                    col2.metric("License Matches", license_matches)
                    col3.metric("Name+DOB Matches", name_dob_matches)
                    col4.metric("Missing MVR", missing_mvr)
                    
                    # Show results table
                    st.subheader("All Trans Sheet Preview")
                    st.dataframe(all_trans_df)
                    
                    # Download functionality
                    st.header("💾 Download Results")
                    
                    output_bytes = io.BytesIO()
                    with pd.ExcelWriter(output_bytes, engine='openpyxl') as writer:
                        all_trans_df.to_excel(writer, sheet_name='All_Trans', index=False)
                        mvr_data.to_excel(writer, sheet_name='Raw_MVR_Data', index=False)
                        client_data.to_excel(writer, sheet_name='Raw_Client_Data', index=False)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                    st.download_button(
                        label="📥 Download All Trans Report",
                        data=output_bytes.getvalue(),
                        file_name=f"All_Trans_Report_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
            except Exception as e:
                st.error(f"❌ Error occurred during processing: {str(e)}")
                import traceback
                st.error("Full error details:")
                st.code(traceback.format_exc())

    if __name__ == "__main__":
        main()
elif menu == "Supplement":
    
    # Initialize session state
    if 'text_searcher' not in st.session_state:
        st.session_state.text_searcher = PDFTextSearcher()
        st.session_state.file_processed = False
        st.session_state.search_ready = False

    st.title("📄 PDF Text Searcher")
    st.markdown("Upload a PDF document and ask questions about its content.")

    # File upload section - in main area for better visibility
    with st.container(border=True):
        st.subheader("1. Upload PDF")
        uploaded_file = st.file_uploader(
            "Choose a PDF file", 
            type="pdf",
            label_visibility="visible",
            key="pdf_uploader"
        )

        if uploaded_file and not st.session_state.file_processed:
            with st.spinner("Processing PDF..."):
                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name
                
                # Process the PDF
                processing_result = st.session_state.text_searcher.process_pdf(tmp_path)
                
                # Clean up
                try:
                    os.unlink(tmp_path)
                except:
                    pass

                if st.session_state.text_searcher.processed:
                    st.session_state.file_processed = True
                    st.session_state.search_ready = True
                    st.success("PDF processed successfully! You can now search the document.")
                else:
                    st.error("Failed to process PDF")

    # Search section - only appears after successful upload
    if st.session_state.search_ready:
        with st.container(border=True):
            st.subheader("2. Search Document")
            question = st.text_input(
                "Enter your question about the document:",
                placeholder="e.g. What is the main conclusion?",
                key="question_input"
            )

            if question and st.session_state.file_processed:
                with st.spinner("Searching document..."):
                    answer, pages, confidence, search_time, keywords = st.session_state.text_searcher.semantic_search(question)
                    
                    if pages:
                        # Display results in columns
                        col1, col2 = st.columns([1, 3])
                        
                        with col1:
                            st.metric("Confidence", f"{confidence:.0%}" if confidence >= 0.1 else "Low")
                            st.metric("Found on Page", pages[0])
                        
                        with col2:
                            st.markdown("**Answer:**")
                            st.info(answer)
                            
                            # Show context
                            with st.expander("View in context"):
                                context = st.session_state.text_searcher.get_context(pages[0], answer)
                                st.markdown(context)
                        
                        # Visualize the page
                        st.subheader("Document Preview")
                        highlight_phrases = [answer[:100]]  # Use answer as first phrase
                        if keywords:
                            highlight_phrases.extend(keywords[:3])  # Add up to 3 keywords
                        
                        fig, error = st.session_state.text_searcher.visualize_page(pages[0], highlight_phrases)
                        if fig:
                            st.pyplot(fig)
                        if error:
                            st.warning(error)
                    else:
                        st.warning("No results found for your query.")
