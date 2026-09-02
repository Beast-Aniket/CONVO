import streamlit as st
import importlib
import os
import sys

# Setup Page Configuration (Must be first Streamlit command)
st.set_page_config(
    page_title="Convocation Data Management Portal",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium styling CSS
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap');

:root {
    --app-bg: #0b111a;
    --app-panel: #111827;
    --app-panel-soft: #172033;
    --app-border: #2f3b52;
    --app-text: #f8fafc;
    --app-muted: #cbd5e1;
    --app-accent: #38bdf8;
    --app-accent-strong: #fbbf24;
}

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

h1, h2, h3, h4, h5, h6 {
    font-family: 'Outfit', sans-serif;
    color: var(--app-text);
}

.stApp {
    background: var(--app-bg);
    color: var(--app-text) !important;
}

.stApp p,
.stApp li,
.stApp label,
.stApp span,
.stApp div {
    color: inherit;
}

/* Sidebar Styling */
[data-testid="stSidebar"] {
    background: var(--app-panel) !important;
    border-right: 1px solid var(--app-border);
    padding-top: 10px;
    color: var(--app-text) !important;
}

[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] strong {
    color: var(--app-text) !important;
    opacity: 1 !important;
}

[data-testid="stSidebar"] small,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
    color: var(--app-muted) !important;
    opacity: 1 !important;
}

.sidebar-brand {
    text-align: center;
    margin-top: 5px;
    margin-bottom: 20px;
}

.sidebar-brand-icon {
    font-size: 2.5rem;
}

.sidebar-brand-title {
    font-family: 'Outfit', sans-serif;
    font-weight: 700;
    color: var(--app-text) !important;
    margin: 5px 0 2px 0;
    font-size: 1.25rem;
}

.sidebar-brand-subtitle {
    color: var(--app-muted) !important;
    font-size: 0.8rem;
    margin: 0;
}

.sidebar-section-title {
    color: var(--app-muted) !important;
    font-size: 0.9rem;
    margin-bottom: 10px;
}

/* System Status Cards */
.status-card {
    background: var(--app-panel-soft);
    border: 1px solid var(--app-border);
    border-radius: 8px;
    padding: 12px;
    margin-bottom: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.18);
    font-size: 0.85rem;
    color: var(--app-text) !important;
    transition: all 0.25s ease;
}

.status-card * {
    color: var(--app-text) !important;
}
.status-card:hover {
    border-color: #475569;
    box-shadow: 0 4px 12px rgba(0,0,0,0.24);
    transform: translateY(-1px);
}

.status-indicator {
    display: inline-block;
    width: 8px;
    height: 8px;
    border-radius: 50%;
    margin-right: 8px;
}
.status-indicator.success { background-color: #10b981; }
.status-indicator.warning { background-color: #f59e0b; }
.status-indicator.error { background-color: #ef4444; }

/* Custom Headings & Page Header */
.main-header {
    background: linear-gradient(135deg, #1e3a8a, #111827);
    padding: 25px;
    border-radius: 12px;
    margin-bottom: 30px;
    border: 1px solid #253a7a;
    box-shadow: 0 4px 15px rgba(0,0,0,0.25);
}
.main-header h1 {
    color: var(--app-text);
    margin: 0;
    font-size: 2.2rem;
    font-weight: 700;
}
.main-header p {
    color: #bfdbfe;
    margin: 5px 0 0 0;
    font-size: 1rem;
}

[data-testid="stFileUploader"] section,
[data-testid="stDataFrame"],
[data-testid="stExpander"] details,
[data-baseweb="select"] > div,
.stTextInput input,
.stNumberInput input,
.stTextArea textarea {
    background-color: var(--app-panel);
    color: var(--app-text) !important;
    border-color: var(--app-border);
}

[data-testid="stFileUploader"] button {
    color: #0f172a !important;
}

.stAlert {
    color: var(--app-text);
}
</style>
""", unsafe_allow_html=True)

# Add project root directory to sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# --- Sidebar Content ---
st.sidebar.markdown(
    """
    <div class="sidebar-brand">
        <span class="sidebar-brand-icon">🎓</span>
        <h2 class="sidebar-brand-title">CONVOCATION PORTAL</h2>
        <p class="sidebar-brand-subtitle">Data Mapping & Translation Suite</p>
    </div>
    """,
    unsafe_allow_html=True
)

st.sidebar.markdown("<h3 class='sidebar-section-title'>📂 APPLICATION MODULES</h3>", unsafe_allow_html=True)

# Define modules mapping
modules = {
    "📊 REGULAR DATA": "regular",
    "🔄 RLE_RPVDATA": "rle_rpv",
    "🎓 BA EXAM": "ba",
    "🔬 BSC EXAM": "bsc",
    "🎓 NEP EXAM": "nep",
    "📖 TRANSLATE ANY DATA": "translate",
    "🔍 FETCH NAMES": "fetch"
}

choice = st.sidebar.radio(
    "Select Module",
    list(modules.keys()),
    label_visibility="collapsed"
)

# --- Dynamic System Status Sidebar Section ---
st.sidebar.markdown("<h3 class='sidebar-section-title' style='margin-top: 25px;'>🛠 SYSTEM STATUS</h3>", unsafe_allow_html=True)

# 1. dic.py Status
from core.name_lookup import universal_dictionary_path
dic_path = universal_dictionary_path()
if os.path.exists(dic_path):
    dic_size_mb = os.path.getsize(dic_path) / (1024 * 1024)
    dic_status = f'<div class="status-card"><span class="status-indicator success"></span><strong>dic.py:</strong> Available ({dic_size_mb:.1f} MB)</div>'
else:
    dic_status = '<div class="status-card"><span class="status-indicator warning"></span><strong>dic.py:</strong> Missing</div>'

# 2. program_master.xlsx Status
from core.program_master import get_program_master_path
program_path = get_program_master_path("program_master.xlsx")
if os.path.exists(program_path):
    program_status = '<div class="status-card"><span class="status-indicator success"></span><strong>program_master:</strong> Available</div>'
else:
    program_status = '<div class="status-card"><span class="status-indicator error"></span><strong>program_master:</strong> Not found</div>'

# 3. college_master.py Status
college_path = os.path.join(current_dir, "core", "college_master.py")
if os.path.exists(college_path):
    college_status = '<div class="status-card"><span class="status-indicator success"></span><strong>college_master:</strong> Available</div>'
else:
    college_status = '<div class="status-card"><span class="status-indicator warning"></span><strong>college_master:</strong> Missing</div>'

st.sidebar.markdown(dic_status + program_status + college_status, unsafe_allow_html=True)

# --- Main App Layout Section ---
st.markdown(
    f"""
    <div class="main-header">
        <h1>🎓 Convocation Data Management Portal</h1>
        <p>Active Module: <strong>{choice[2:]}</strong></p>
    </div>
    """,
    unsafe_allow_html=True
)

selected_module = modules[choice]

if selected_module == "regular":
    regular_data_app = importlib.import_module("modules.regular_data_app")
    regular_data_app.run_regular_data_app()
elif selected_module == "rle_rpv":
    rle_rpv = importlib.import_module("modules.rle_rpv_app")
    rle_rpv.run_rle_rpv_app()
elif selected_module == "ba":
    ba_app = importlib.import_module("modules.ba_exam_app")
    ba_app.run_ba_exam_app()
elif selected_module == "bsc":
    bsc_app = importlib.import_module("modules.bsc_exam_app")
    bsc_app.run_bsc_exam_app()
elif selected_module == "nep":
    nep_app = importlib.import_module("modules.nep_exam_app")
    nep_app.run_nep_exam_app()
elif selected_module == "translate":
    translate_app = importlib.import_module("modules.translate_app")
    translate_app.run_translate_app()
elif selected_module == "fetch":
    fetch_app = importlib.import_module("modules.fetch_app")
    fetch_app.run_fetch_app()
