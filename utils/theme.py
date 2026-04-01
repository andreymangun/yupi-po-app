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
    
    # CSS GLOBAL: Sembunyikan elemen bawaan Streamlit
    global_css = """
    [data-testid="stHeader"] { visibility: hidden !important; height: 0px !important; margin: 0 !important; padding: 0 !important; }
    .stDeployButton { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    [data-testid="stSidebarNav"] { display: none !important; }
    """

    if not is_logged_in:
        bg_style = f"""
        @keyframes panBackground {{ 0% {{background-position: 0% 50%;}} 50% {{background-position: 100% 50%;}} 100% {{background-position: 0% 50%;}} }}
        .stApp {{ background-image: url("data:image/png;base64,{bg_base64}"); background-size: 150% 150%; background-position: center; background-attachment: fixed; animation: panBackground 40s linear infinite; }}
        """ if bg_base64 else ""

        st.markdown(f"""
            <style>
            {global_css} {bg_style}
            div[data-testid="stForm"] {{ background: rgba(255,255,255,0.1) !important; backdrop-filter: blur(25px); border-radius: 20px; border: 1px solid rgba(255,255,255,0.3); padding: 40px !important; width: 420px; margin: 0 auto; box-shadow: 0 15px 35px rgba(0,0,0,0.3); }}
            
            div[data-testid="stTextInput"] label {{ color: white !important; font-weight: bold !important; text-shadow: 1px 1px 2px rgba(0,0,0,0.5); }}
            
            div[data-testid="stTextInput"] input {{ 
                color: #0F172A !important; 
                -webkit-text-fill-color: #0F172A !important; 
                background-color: rgba(255,255,255,0.9) !important; 
                border: 1px solid rgba(255,255,255,0.4) !important; 
            }}
            .login-logo-text {{ color: white; font-size: 4rem; font-weight: 900; text-align: center; text-shadow: 0px 8px 20px rgba(0,0,0,0.5); margin-bottom: -10px; }}
            </style>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
            <style>
            {global_css}
            
            /* --- FIX: Menggunakan CSS Variables Streamlit agar Adaptif Dark Mode --- */
            .stApp {{ background-color: var(--background-color) !important; }}
            
            /* --- PERBAIKAN SPASI KOSONG --- */
            .block-container {{
                padding-top: 2rem !important; 
                padding-bottom: 2rem !important;
            }}

            /* Menyesuaikan kotak kaca agar isinya lebih padat di atas dan warna mengikuti tema */
            .main-glass-frame, div[data-testid="stForm"], div[data-testid="stExpander"], .stDataFrame, .stTabs [data-baseweb="tab-panel"] {{
                background-color: var(--secondary-background-color) !important; 
                border-radius: 12px !important; 
                padding: 10px 20px 20px 20px !important; 
                border: 1px solid rgba(128, 128, 128, 0.2) !important; 
                box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05) !important;
                margin-top: 0 !important; 
            }}
            
            /* Menghilangkan margin bawaan Streamlit pada elemen Markdown pertama di dalam frame */
            .main-glass-frame > div:first-child > div > div > div > h2,
            .main-glass-frame > div:first-child > div > div > div > h3,
            .main-glass-frame h2:first-of-type,
            .main-glass-frame h3:first-of-type {{
                margin-top: 0 !important;
                padding-top: 0 !important;
            }}

            /* Memastikan warna teks mengikuti tema secara otomatis */
            p, span, h1, h2, h3, h4, h5, h6, li, label, .stMarkdown {{ 
                color: var(--text-color) !important; 
            }}
            
            /* Memastikan Input fields mengikuti warna tema */
            input, textarea, select, 
            div[data-testid="stTextInput"] input, 
            div[data-testid="stTextArea"] textarea, 
            div[data-testid="stDateInput"] input, 
            div[data-testid="stNumberInput"] input, 
            div[data-baseweb="select"] div {{
                color: var(--text-color) !important; 
                -webkit-text-fill-color: var(--text-color) !important;
                background-color: var(--background-color) !important;
                border-color: rgba(128, 128, 128, 0.3) !important;
            }}
            
            /* Memastikan tabel mengikuti warna tema */
            .stDataFrame [data-testid="stTable"] th, 
            .stDataFrame [data-testid="stTable"] td, 
            .stDataFrame [data-baseweb="table-header"] {{
                color: var(--text-color) !important;
                background-color: transparent !important;
            }}
            </style>
        """, unsafe_allow_html=True)