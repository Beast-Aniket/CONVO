import streamlit as st
import os
import sys

# Setup Page Configuration (Must be first Streamlit command)
st.set_page_config(
    page_title="Convocation Data Management Portal",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Premium Styling CSS
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

h1, h2, h3, h4, h5, h6 {
    font-family: 'Outfit', sans-serif;
}

/* Sidebar Styling */
[data-testid="stSidebar"] {
    background-color: #f8fafc;
    border-right: 1px solid #e2e8f0;
    padding-top: 10px;
}

/* System Status Cards */
.status-card {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 12px;
    margin-bottom: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    font-size: 0.85rem;
    color: #334155;
    transition: all 0.25s ease;
}
.status-card:hover {
    box-shadow: 0 4px 6px rgba(0,0,0,0.04);
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
    background: linear-gradient(135deg, #1e3a8a, #0f172a);
    padding: 25px;
    border-radius: 12px;
    margin-bottom: 30px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.05);
}
.main-header h1 {
    color: white;
    margin: 0;
    font-size: 2.2rem;
    font-weight: 700;
}
.main-header p {
    color: #93c5fd;
    margin: 5px 0 0 0;
    font-size: 1rem;
}
</style>
""", unsafe_allow_html=True)

# Add subdirectories to sys.path to allow clean imports
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(current_dir, 'BAPRG'))
sys.path.append(os.path.join(current_dir, 'BSCPRG'))

# Import Module Apps
import regular_data_app
import rle_rpv
import TRANSLATE as translate_app
import FETCH as fetch_app
import BAPRG.APP as baprg_app
import BSCPRG.APP as bscprg_app

# --- Sidebar Content ---
st.sidebar.markdown(
    """
    <div style="text-align: center; margin-top: 5px; margin-bottom: 20px;">
        <span style="font-size: 2.5rem;">🎓</span>
        <h2 style="font-family: 'Outfit', sans-serif; font-weight: 700; color: #1e293b; margin: 5px 0 2px 0; font-size: 1.25rem;">CONVOCATION PORTAL</h2>
        <p style="color: #64748b; font-size: 0.8rem; margin: 0;">Data Mapping & Translation Suite</p>
    </div>
    """,
    unsafe_allow_html=True
)

st.sidebar.markdown("<h3 style='font-size: 0.9rem; color: #475569; margin-bottom: 10px;'>📂 APPLICATION MODULES</h3>", unsafe_allow_html=True)

# Define modules mapping
modules = {
    "📊 REGULAR DATA": "regular",
    "🔄 RLE_RPVDATA": "rle_rpv",
    "🎓 BA EXAM": "ba",
    "🔬 BSC EXAM": "bsc",
    "📖 TRANSLATE ANY DATA": "translate",
    "🔍 FETCH NAMES": "fetch"
}

choice = st.sidebar.radio(
    "Select Module",
    list(modules.keys()),
    label_visibility="collapsed"
)

# --- Dynamic System Status Sidebar Section ---
st.sidebar.markdown("<h3 style='font-size: 0.9rem; color: #475569; margin-top: 25px; margin-bottom: 10px;'>🛠 SYSTEM STATUS</h3>", unsafe_allow_html=True)

# 1. dic.py Status
try:
    from dic import name_translation_dict
    dic_status = f'<div class="status-card"><span class="status-indicator success"></span><strong>dic.py:</strong> Loaded ({len(name_translation_dict):,} words)</div>'
except Exception:
    dic_status = '<div class="status-card"><span class="status-indicator warning"></span><strong>dic.py:</strong> Missing or failed</div>'

# 2. program_master.xlsx Status
try:
    from program_master import load_program_master
    p_master, err = load_program_master()
    if p_master:
        program_status = f'<div class="status-card"><span class="status-indicator success"></span><strong>program_master:</strong> Loaded ({len(p_master)} programs)</div>'
    else:
        program_status = f'<div class="status-card"><span class="status-indicator error"></span><strong>program_master:</strong> Not found</div>'
except Exception:
    program_status = '<div class="status-card"><span class="status-indicator error"></span><strong>program_master:</strong> Error loading</div>'

# 3. college_master.py Status
try:
    from college_master import college_master
    college_status = f'<div class="status-card"><span class="status-indicator success"></span><strong>college_master:</strong> Loaded ({len(college_master)} colleges)</div>'
except Exception:
    college_status = '<div class="status-card"><span class="status-indicator warning"></span><strong>college_master:</strong> Failed to load</div>'

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
    regular_data_app.run_regular_data_app()
elif selected_module == "rle_rpv":
    rle_rpv.run_rle_rpv_app()
elif selected_module == "ba":
    baprg_app.run_ba_exam_app()
elif selected_module == "bsc":
    bscprg_app.run_bsc_exam_app()
elif selected_module == "translate":
    translate_app.run_translate_app()
elif selected_module == "fetch":
    fetch_app.run_fetch_app()