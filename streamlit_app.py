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
    base = ["MVR All Trans", "Supplement", "MVR GPT","MVR All Trans(test)"]
    if role == "ADMIN":
        return base 
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
if menu == "MVR All Trans":
    all_trans_mvr_app()
# Welcome screen for "App" menu
# HDVI MVR tool
#elif menu == "QC Radar":
    #qc_radar_app()
elif menu == "MVR GPT":
    mvr_gpt_app()
elif menu == "MVR All Trans(test)":
    from Alltran import Alltrans
        # app.py
    import streamlit as st

    st.set_page_config(page_title="Final Report — app", layout="wide")
    st.title("Alltrans Test")

    #st.write("Upload MAIN (MVR) and LOOKUP files. Template.xlsx must be in the same folder.")

    main_file = st.file_uploader("Upload MVR File", type=["xlsx"])
    lookup_file = st.file_uploader("Upload Client Excel", type=["xlsx","CSV"])

    if main_file and lookup_file:
        try:
            main_bytes = main_file.read()
            lookup_bytes = lookup_file.read()

            # If lookup workbook has multiple sheets, let user choose
            from openpyxl import load_workbook
            lookup_wb = load_workbook(io.BytesIO(lookup_bytes), read_only=True, data_only=True)
            sheets = lookup_wb.sheetnames
            chosen_sheet = None
            if len(sheets) > 1:
                chosen_sheet = st.selectbox("Lookup workbook has multiple sheets. Choose one:", options=sheets)
            else:
                chosen_sheet = sheets[0]
                st.write(f"From Client Excel: {chosen_sheet}")

            gen = Alltrans(template_path="Template.xlsx",
                        alltrans_sheet="All Trans",
                        alltrans_header_row=4,#Fixed constand row for Column Headers
                        mvr_sheet_name="MVR")
            out = gen.run(main_bytes, lookup_bytes, chosen_lookup_sheet=chosen_sheet, preview_rows=8)

            original_name = getattr(main_file, "name", None)
            base = original_name.rsplit(".", 1)[0] if original_name else "Final_Report"
            out_name = f"{base}.xlsx"
            st.success("Final report generated")
            st.download_button("Download Final Report", data=out, file_name=out_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error: {e}")
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
