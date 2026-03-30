import streamlit as st
import pandas as pd
import requests
import os

def get_ai_response(prompt, user_name="User", page_context="Dashboard"):
    api_key = ""
    
    # 1. BACA KUNCI GROQ
    try:
        key_file = os.path.join(os.getcwd(), "groq_key.txt")
        if os.path.exists(key_file):
            with open(key_file, "r") as f:
                api_key = f.read().strip()
    except Exception as e:
        return f"❌ Error membaca file kunci: {str(e)}"

    if not api_key or not str(api_key).startswith("gsk_"):
        return "❌ **Kunci Groq tidak valid/tidak ditemukan.**\n\nPastikan Anda telah membuat file `groq_key.txt` di folder `serveone-erp` dan mengisi kunci yang berawalan `gsk_` dari console.groq.com."

    # 2. SIAPKAN KONTEKS
    system_prompt = f"Anda OpenClaw, Asisten AI handal untuk ERP ServeOne. Anda sedang berbicara dengan {user_name} di layar {page_context}. Jawablah pertanyaan operasional dengan ringkas dan profesional."
    
    if "operation_df" in st.session_state and isinstance(st.session_state["operation_df"], pd.DataFrame) and not st.session_state["operation_df"].empty:
        df = st.session_state["operation_df"]
        system_prompt += f"\n\n[Konteks Data Operasional]:\n{df.head(5).to_string()}"

    # 3. PANGGIL GROQ API LANGSUNG (Tanpa library berat)
    try:
        url = "https://api.groq.com/openai/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": "llama3-8b-8192", # Model gratis dan super cepat
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.5,
            "max_tokens": 1024
        }
        
        response = requests.post(url, headers=headers, json=payload, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            return data["choices"][0]["message"]["content"]
        else:
            # Jika Groq menolak, tangkap pesan aslinya
            try:
                error_msg = response.json().get("error", {}).get("message", "Unknown error")
            except:
                error_msg = response.text
                
            return f"⚠️ **Groq API Error {response.status_code}:**\n`{error_msg}`"
            
    except Exception as e:
        return f"❌ **Gagal menghubungi Server Groq:** {str(e)}"