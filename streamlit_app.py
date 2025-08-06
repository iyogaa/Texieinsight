import pandas as pd
import streamlit as st
from thefuzz import fuzz, process
import re
import io
from datetime import datetime
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

# --- EXACT CORE LOGIC FROM YOUR PROVIDED CODE ---
def normalize_name(name):
    """Enhanced name normalization with title removal and initials handling"""
    if pd.isna(name) or not name:
        return []
    name = str(name).lower()
    # Remove common prefixes/suffixes
    name = re.sub(r'\b(mr|mrs|ms|dr|jr|sr|iii|ii|iv)\b', '', name)
    # Remove non-alpha chars except spaces
    name = re.sub(r'[^a-z\s]', '', name)
    # Normalize spaces
    name = re.sub(r'\s+', ' ', name).strip()
    parts = name.split()
    if not parts:
        return []

    formats = []
    # Full name normal
    formats.append(' '.join(parts))
    # First last and last first formats
    if len(parts) > 1:
        formats.append(f"{parts[0]} {parts[-1]}")
        formats.append(f"{parts[-1]} {parts[0]}")
        formats.append(f"{parts[0]}{parts[-1]}")
        formats.append(f"{parts[-1]}{parts[0]}")

    # Initial-based formats if middle names exist
    if len(parts) > 2:
        first = parts[0]
        last = parts[-1]
        initials = ''.join([p[0] for p in parts[1:-1]])
        formats.append(f"{first} {initials} {last}")
        formats.append(f"{first} {initials}{last}")
        formats.append(f"{first}{initials} {last}")
        formats.append(f"{first}{initials}{last}")

    # Remove duplicates
    return list(set(formats))

def names_match(name1, name2):
    """Stricter matching with multiple fuzzy strategies"""
    if pd.isna(name1) or pd.isna(name2) or not name1 or not name2:
        return False
    formats1 = normalize_name(name1)
    formats2 = normalize_name(name2)
    for f1 in formats1:
        for f2 in formats2:
            if f1 == f2:
                return True
            if fuzz.token_set_ratio(f1, f2) >= 95:
                return True
            if fuzz.partial_ratio(f1, f2) >= 96:
                return True
            if fuzz.token_sort_ratio(f1, f2) >= 98:
                return True
    return False

def get_valid_column(df, purpose, default_names, required=True):
    """Find column with fuzzy matching, using defaults if possible"""
    # First try exact matches to default names
    for col in default_names:
        if col in df.columns:
            return col
    
    # Then try fuzzy matching
    for col_name in default_names:
        match, score = process.extractOne(col_name, df.columns, scorer=fuzz.ratio)
        if score > 80:
            return match
    
    # If not found and required, return first column
    if required and len(df.columns) > 0:
        return df.columns[0]
    
    return None
# --- END OF EXACT CORE LOGIC ---

# --- Main Application Logic ---
if menu == "All Trans MVR":
    all_trans_mvr_app(get_valid_column, names_match)
# Welcome screen for "App" menu
# HDVI MVR tool
elif menu == "QC Radar":
    qc_radar_app()
# Truckings IFTA tool
elif menu == "Truckings IFTA":
    st.markdown('<div class="custom-heading">Truckings IFTA Tool</div>', unsafe_allow_html=True)
    st.write("Truckings IFTA tool will be available soon.")

# Riscom MVR tool
elif menu == "Riscom MVR":
    st.markdown('<div class="custom-heading">Riscom Tool</div>', unsafe_allow_html=True)
    input_text = st.text_area("You can paste the Fullnames here:", height=150, placeholder="""
    Example:
    Kungfu, Panda
    Chotta, Bheem
    Walter, White
    Yoga, Raj
    Kishoor, Aravindh
    Gokul, Sarvesh
    Jackie, Chan
    """)

    def parse_name(full_name):
        original = full_name.strip()
        if ',' in original:
            last, first = [p.strip() for p in original.split(',', 1)]
        else:
            tokens = original.split()
            if len(tokens) >= 2:
                first = " ".join(tokens[:-1])
                last = tokens[-1]
            else:
                first, last = original, ""
        return {
            "Full Name": original,
            "First Name": first,
            "Last Name": last
        }

    # When user inputs text
    if input_text:
        names = [line.strip() for line in input_text.strip().split('\n') if line.strip()]
        parsed_data = [parse_name(name) for name in names]
        df = pd.DataFrame(parsed_data)

        # Format as tab-separated values
        output_text = df.to_csv(index=False, sep='\t')

        # Display copy-to-clipboard button
        st.subheader("📎 Copy Output")
        st.code(output_text, language='text')

        # Copy button with stateful feedback
        if st.button("📋 Copy to Clipboard"):
            st.toast("✅ Copied successfully!", icon="✅")
            st.session_state.clipboard_text = output_text
    
    input_dates = st.text_area("Enter D.O.B :", height=250, placeholder="""
    9-21-2002
    28-09-1989
    """)

    def parse_and_format_date(date_str):
        # Normalize separators
        clean_date = re.sub(r'[.\s]+', '-', date_str.strip())
        parts = re.split(r'[-/]', clean_date)

        try:
            # Try MM-DD-YYYY
            if int(parts[0]) > 12:
                # Assume DD-MM-YYYY
                day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
            else:
                # Assume MM-DD-YYYY
                month, day, year = int(parts[0]), int(parts[1]), int(parts[2])
            formatted = datetime(year, month, day)
            return formatted.strftime("%m/%d/%Y"), calculate_age(formatted)
        except:
            return "Invalid Date", ""

    def calculate_age(birthdate):
        today = datetime.today()
        return today.year - birthdate.year - ((today.month, today.day) < (birthdate.month, birthdate.day))

    # Process
    if input_dates:
        rows = []
        for line in input_dates.strip().split('\n'):
            original = line.strip()
            if not original:
                continue
            formatted_date, age = parse_and_format_date(original)
            flag = ""
            if isinstance(age, int):
                if age < 21:
                    flag = "<21 (Underage)"
                elif age > 69:
                    flag = ">70 (Overage)"
            rows.append({
                "Formatted DOB": formatted_date,
                "Age": age,
                "Flag": flag
            })

        df = pd.DataFrame(rows)

        st.subheader("Results")
        output_text = df.to_csv(index=False, sep='\t')
        st.code(output_text, language='text')

# MVR GPT tool (accessible to all roles)
elif menu == "MVR GPT":
    mvr_gpt_app()
elif menu == "Insight Dashboard":
    insight_dashboard_app()
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
