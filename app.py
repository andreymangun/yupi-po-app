import streamlit as st
import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import requests
from io import StringIO
import os
from datetime import date

# ==========================================
# 1. PDF ENGINE
# ==========================================
class ServeonePO(FPDF):
    def header(self):
        if os.path.exists("logo.png"):
            self.image("logo.png", 10, 2, 45)
        self.set_font('Helvetica', 'B', 16)
        self.set_y(10)
        self.cell(0, 10, 'P/O Paper', align='C', new_x=XPos.LMARGIN, new_y=YPos.TOP)
        if os.path.exists("stamp.png"):
            self.image("stamp.png", 160, 8, 40)
        self.ln(20)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()} of {{nb}}', align='C')

def clean_val(val):
    if pd.isna(val) or str(val).lower() == 'nan': return ""
    return str(val).strip()

def safe_float(val):
    try:
        if isinstance(val, str): val = val.replace(',', '').replace('"', '').strip()
        return float(val) if val else 0.0
    except: return 0.0

# ==========================================
# 2. GENERATOR PO (V12 - CUSTOM FILENAME LOGIC)
# ==========================================
def generate_po_pdf(po_data, po_number):
    pdf = ServeonePO()
    pdf.alias_nb_pages()
    pdf.add_page()
    info = po_data.iloc[0]
    
    # Ambil Nama Vendor untuk penamaan file nanti
    vendor_name = clean_val(info.get('Vendor Name'))

    # PO INFO
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(30, 7, "PO YUPI :")
    pdf.set_font('Helvetica', '', 10)
    pdf.cell(0, 7, clean_val(po_number), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(30, 7, "PO Date :")
    pdf.set_font('Helvetica', '', 10)
    pdf.cell(0, 7, date.today().strftime("%d/%m/%Y"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

    # CLIENT & VENDOR
    y_anchor = pdf.get_y()
    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(95, 5, "CLIENT:")
    pdf.set_xy(105, y_anchor)
    pdf.cell(95, 5, "VENDOR:")
    
    pdf.ln(6)
    pdf.set_font('Helvetica', '', 8)
    y_content = pdf.get_y()
    client_info = "PT. SERVEONE MRO INDONESIA\nC-2-2 LANTAI 2, MM2100\nKab. Bekasi, Prov. Jawa Barat"
    vendor_info = f"{vendor_name}\n{clean_val(info.get('Vendor Address'))}"
    
    pdf.set_xy(10, y_content)
    pdf.multi_cell(92, 4, client_info)
    pdf.set_xy(105, y_content)
    pdf.multi_cell(95, 4, vendor_info)
    pdf.set_y(max(pdf.get_y(), y_content + 15) + 5)

    # SITE LOGIC & LABEL
    site_raw = ""
    for col in info.index:
        if 'SITE' in str(col).upper():
            site_raw = clean_val(info[col]).upper()
            break
    
    if 'IDN' in site_raw:
        site_label = "IDN"
        full_addr = "Jl. Pancasila IV No.9, Cicadas, Kec. Gn. Putri, Kabupaten Bogor, Jawa Barat 16964, Gunung Putri, Bogor, Jawa Barat"
    else:
        site_label = "KRG"
        full_addr = "JI.Grompol-Jambangan, Kaliwuluh Lor, Kebakkramat, Karanganyar, Jawa Tengah"

    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(0, 6, f"DELIVERY ADDRESS: PT. YUPI INDO JELLY GUM Tbk ({site_label})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', '', 8)
    pdf.multi_cell(0, 5, full_addr)
    pdf.ln(5)

    # TABLE HEADER
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font('Helvetica', 'B', 7)
    cols = [
        ("Deliv Req", 20), ("SERVEONE PO", 30), ("Item Code", 20), 
        ("Item Name & Spec", 50), ("Qty", 10), ("Unit", 10), ("Price", 25), ("Amount", 25)
    ]
    for txt, w in cols:
        pdf.cell(w, 8, txt, border=1, fill=True, align='C')
    pdf.ln()

    # TABLE BODY
    pdf.set_font('Helvetica', '', 7)
    total_dpp = 0
    remark_list = []

    for _, row in po_data.iterrows():
        name_spec = f"{clean_val(row.get('Item name'))}\n{clean_val(row.get('Spec'))}"
        qty = safe_float(row.get("Ord. Q'ty"))
        price = safe_float(row.get('PURCHASE PRICE'))
        amount = qty * price
        total_dpp += amount
        if clean_val(row.get('REMARK YUPI')): remark_list.append(clean_val(row.get('REMARK YUPI')))

        lines = pdf.multi_cell(50, 4, name_spec, split_only=True)
        row_h = max(10, len(lines) * 4 + 2)

        if pdf.get_y() + row_h > 240: pdf.add_page()

        x, y = pdf.get_x(), pdf.get_y()
        pdf.cell(20, row_h, clean_val(row.get('Req. Dlv Date')), border=1, align='C')
        pdf.cell(30, row_h, clean_val(row.get('PO SEMENTARA')), border=1, align='C')
        pdf.cell(20, row_h, clean_val(row.get('Item Yupi')), border=1, align='C')
        pdf.cell(50, row_h, "", border=1) 
        pdf.cell(10, row_h, f"{qty:,.0f}", border=1, align='C')
        pdf.cell(10, row_h, clean_val(row.get('Unit')), border=1, align='C')
        pdf.cell(25, row_h, f"{price:,.0f}", border=1, align='R')
        pdf.cell(25, row_h, f"{amount:,.0f}", border=1, align='R')

        pdf.set_xy(x + 70, y + 1) 
        pdf.multi_cell(50, 4, name_spec, border=0, align='L')
        pdf.set_xy(10, y + row_h)

    # FOOTER
    ppn = total_dpp * 0.11
    grand_total = total_dpp + ppn
    pdf.set_font('Helvetica', 'B', 8)
    for label, val in [("TOTAL AMOUNT", total_dpp), ("PPN (11%)", ppn), ("GRAND TOTAL", grand_total)]:
        pdf.set_x(150)
        pdf.cell(25, 5, label, border=1, align='R')
        pdf.cell(25, 5, f"{val:,.0f}", border=1, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # REMARKS
    pdf.ln(5)
    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(0, 6, "REMARKS:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', '', 8)
    pdf.multi_cell(0, 5, ", ".join(set(remark_list)) if remark_list else "-")

    # KEMBALIKAN PDF BYTES, SITE_LABEL, DAN VENDOR_NAME
    return bytes(pdf.output()), site_label, vendor_name

# ==========================================
# 3. INTERFACE
# ==========================================
st.set_page_config(page_title="Serveone System", layout="wide")
st.title("📑 Serveone PO Generator for Yupi Project")

link = st.sidebar.text_input("Enter your Google Sheets CSV link here")

if link:
    try:
        res = requests.get(link)
        df_raw = pd.read_csv(StringIO(res.text), low_memory=False)
        mask = df_raw.apply(lambda row: row.astype(str).str.contains('PO YUPI').any(), axis=1)
        idx = df_raw[mask].index[0]
        df = df_raw.iloc[idx+1:].copy()
        df.columns = [str(c).strip() for c in df_raw.iloc[idx]]
        df = df.dropna(subset=['PO YUPI'])

        sel_po = st.selectbox("Pilih Nomor PO:", df['PO YUPI'].unique())
        
        if st.button("Generate PDF"):
            # Memanggil fungsi yang mengembalikan 3 data
            po_bytes, site_label, vendor_name = generate_po_pdf(df[df['PO YUPI'] == sel_po], sel_po)
            
            # LOGIKA PENAMAAN FILE: [SITE] PO SERVEONE - PO_NUMBER - VENDOR_NAME.pdf
            file_name = f"[{site_label}] PO SERVEONE - {sel_po} - {vendor_name}.pdf"
            
            st.download_button(
                label=f"Download {file_name}",
                data=po_bytes,
                file_name=file_name,
                mime="application/pdf"
            )
            st.success(f"File siap diunduh dengan nama: {file_name}")
            
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Please enter the CSV link in the upper-left section before starting.")
