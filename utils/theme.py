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
    
    # CSS GLOBAL
    base_css = """
    <style>
    /* 1. Sembunyikan Header */
    [data-testid="stHeader"] { visibility: hidden; height: 0px; }
    .stDeployButton { display: none; }
    
    /* 2. Pengaturan Font Dasar agar tidak raksasa */
    html { font-size: 14px; } 

    /* 3. Kontainer Utama */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1.5rem !important;
        max-width: 95% !important;
    }

    /* 4. Tampilan Login */
    .login-logo-text { 
        color: white; 
        font-size: 4rem; 
        font-weight: 900; 
        text-align: center; 
        text-shadow: 0px 8px 20px rgba(0,0,0,0.5); 
    }
    
    div[data-testid="stForm"] {
        background: rgba(255,255,255,0.1) !important;
        backdrop-filter: blur(20px);
        border-radius: 15px;
        border: 1px solid rgba(255,255,255,0.2);
        padding: 2rem !important;
        width: 100% !important;
        max-width: 400px !important;
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
        # --- PERBAIKAN LOGIKA RESPONSIVE DISINI ---
        st.markdown(base_css + """
            <style>
            .stApp { background-color: #F8FAFC !important; }
            
            /* Judul yang proporsional */
            h1 { font-size: 2rem !important; color: #1E293B !important; }
            h2 { font-size: 1.6rem !important; color: #1E293B !important; }
            h3 { font-size: 1.3rem !important; color: #1E293B !important; }
            
            /* HANYA gunakan width 100% (tumpuk vertikal) jika layar di bawah 768px (HP).
               Jika di komputer (layar lebar), biarkan Streamlit mengatur kolom secara horizontal.
            */
            @media (max-width: 768px) {
                [data-testid="column"] {
                    width: 100% !important;
                    flex: 1 1 auto !important;
                    min-width: 100% !important;
                }
            }
            
            /* Berikan sedikit nafas antar kartu */
            .stMetric, .stDataFrame, div[data-testid="stExpander"] {
                background: white !important;
                border: 1px solid #E2E8F0 !important;
                border-radius: 8px !important;
                margin-bottom: 1rem !important;
            }
            
            /* Input yang pas ukurannya */
            .stTextInput input, .stSelectbox div {
                font-size: 0.95rem !important;
            }
            </style>
        """, unsafe_allow_html=True)