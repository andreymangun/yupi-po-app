import streamlit as st
import base64
import os

def get_base64_image(image_path):
    try:
        if os.path.exists(image_path):
            with open(image_path, "rb") as f:
                return base64.b64encode(f.read()).decode()
    except: pass
    return None

def inject_css():
    bg_base64 = get_base64_image("background_login_form.png")
    is_logged_in = st.session_state.get("logged_in", False)
    
    # CSS DASAR UNTUK SEMUA HALAMAN
    base_css = """
    <style>
    /* 1. Sembunyikan Header Streamlit */
    [data-testid="stHeader"] { visibility: hidden; height: 0px; }
    .stDeployButton { display: none; }
    
    /* 2. RESPONSIVE ROOT: Mengatur ukuran dasar teks agar tidak terlalu besar */
    html { font-size: 14px; } /* Standar pengecilan dari 16px default */
    
    @media (max-width: 1400px) { html { font-size: 13px; } }
    @media (max-width: 1024px) { html { font-size: 12px; } }
    
    /* 3. Pengaturan Kontainer Utama */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        max-width: 95% !important; /* Agar tidak terlalu mepet ke pinggir tapi tetap luas */
    }

    /* 4. Login Logo Typography */
    .login-logo-text { 
        color: white; 
        font-size: 4.5rem; 
        font-weight: 900; 
        text-align: center; 
        text-shadow: 0px 8px 20px rgba(0,0,0,0.5); 
        margin-bottom: 0px;
    }
    
    /* 5. Responsive Form Login */
    div[data-testid="stForm"] {
        background: rgba(255,255,255,0.1) !important;
        backdrop-filter: blur(20px);
        border-radius: 15px;
        border: 1px solid rgba(255,255,255,0.2);
        padding: 2.5rem !important;
        width: 100% !important;
        max-width: 400px !important; /* Maksimal lebar form agar tidak melar */
        margin: 0 auto;
    }
    </style>
    """

    if not is_logged_in:
        bg_style = f"""
        <style>
        .stApp {{
            background-image: url("data:image/png;base64,{bg_base64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        </style>
        """ if bg_base64 else ""
        st.markdown(base_css + bg_style, unsafe_allow_html=True)
    else:
        # Tampilan setelah Login (Dashboard & Operation)
        st.markdown(base_css + """
            <style>
            .stApp { background-color: #F8FAFC !important; }
            
            /* Typography Seimbang */
            h1 { font-size: 2.2rem !important; font-weight: 800 !important; color: #1E293B !important; }
            h2 { font-size: 1.8rem !important; font-weight: 700 !important; color: #1E293B !important; }
            h3 { font-size: 1.4rem !important; font-weight: 600 !important; color: #1E293B !important; }
            
            /* Responsive Grid & Columns */
            [data-testid="column"] {
                width: 100% !important;
                flex: 1 1 auto !important;
            }
            
            /* Card Style untuk elemen Dashboard */
            .stMetric, .stDataFrame, .stExpander {
                background: white !important;
                border: 1px solid #E2E8F0 !important;
                border-radius: 10px !important;
                box-shadow: 0 1px 3px rgba(0,0,0,0.1) !important;
            }
            
            /* Input & Select Scaling */
            .stTextInput input, .stSelectbox div {
                height: 2.8rem !important;
                font-size: 1rem !important;
            }
            </style>
        """, unsafe_allow_html=True)