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
    
    # CSS GLOBAL (Sembunyikan elemen bawaan Streamlit yang mengganggu)
    global_css = """
    <style>
    [data-testid="stHeader"] { visibility: hidden; height: 0px; }
    .stDeployButton { display: none; }
    #MainMenu { visibility: hidden; }
    
    /* Ukuran font dasar agar tidak 'raksasa' */
    html { font-size: 14px; } 

    /* Lebar konten utama */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        max-width: 95% !important;
    }
    </style>
    """

    if not is_logged_in:
        # Tampilan Login
        bg_style = f"""
        <style>
        .stApp {{
            background-image: url("data:image/png;base64,{bg_base64}");
            background-size: cover;
            background-position: center;
        }}
        .login-logo-text {{ 
            color: white; font-size: 4rem; font-weight: 900; text-align: center; 
            text-shadow: 0px 8px 20px rgba(0,0,0,0.5); 
        }}
        div[data-testid="stForm"] {{
            background: rgba(255,255,255,0.1) !important;
            backdrop-filter: blur(20px);
            border-radius: 15px;
            padding: 2rem !important;
            max-width: 400px !important;
            margin: 0 auto;
        }}
        </style>
        """ if bg_base64 else ""
        st.markdown(global_css + bg_style, unsafe_allow_html=True)
    else:
        # Tampilan Dashboard & Operation
        st.markdown(global_css + """
            <style>
            .stApp { background-color: #F8FAFC !important; }
            
            /* Typography yang Proporsional */
            h1 { font-size: 2rem !important; }
            h2 { font-size: 1.6rem !important; }
            h3 { font-size: 1.2rem !important; }

            /* --- FIX KOLOM BERTUMPUK --- */
            /* Di Desktop (>768px), paksa kolom agar TIDAK bertumpuk */
            @media (min-width: 769px) {
                div[data-testid="column"] {
                    flex: 1 1 0% !important;
                    min-width: 0 !important;
                }
            }

            /* Di HP (<768px), baru biarkan bertumpuk */
            @media (max-width: 768px) {
                div[data-testid="column"] {
                    min-width: 100% !important;
                }
            }

            /* Styling Card / Frame agar lebih padat */
            .stMetric, .stDataFrame, div[data-testid="stExpander"] {
                background: white !important;
                border: 1px solid #E2E8F0 !important;
                border-radius: 8px !important;
            }
            </style>
        """, unsafe_allow_html=True)