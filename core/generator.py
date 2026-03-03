from fpdf import FPDF

class YupiPO(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 14)
        self.cell(0, 10, 'PURCHASE ORDER', 0, 1, 'L')
        self.ln(5)

def clean_pdf_text(text):
    """Membersihkan karakter non-ASCII (seperti simbol diameter/toleransi) agar PDF tidak crash"""
    if text is None or str(text) == 'nan':
        return ""
    # Mengubah ke string, lalu hilangkan karakter yang tidak didukung Latin-1
    return str(text).encode('latin-1', 'replace').decode('latin-1').replace('?', '')

def generate_po_pdf(po_data, po_number):
    pdf = YupiPO()
    pdf.add_page()
    
    info = po_data.iloc[0]
    
    # Cari kolom tanggal & site secara dinamis
    date_col = [c for c in po_data.columns if 'PO DATE' in c][0]
    site_col = [c for c in po_data.columns if 'SITE' in c][0]
    
    # Header Info
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(100, 6, f"PO YUPI: {clean_pdf_text(po_number)}", 0, 0)
    pdf.cell(0, 6, f"Date: {clean_pdf_text(info[date_col])}", 0, 1)
    pdf.ln(10)

    # Vendor & Delivery
    pdf.set_font('Helvetica', 'B', 9)
    pdf.cell(95, 5, "DELIVERY ADDRESS:", 0, 0)
    pdf.cell(95, 5, "VENDOR:", 0, 1)
    
    pdf.set_font('Helvetica', '', 9)
    y_start_info = pdf.get_y()
    addr = f"PT. YUPI INDO JELLY GUM Tbk\nSite: {info[site_col]}\nReq. Delivery: {info['Req. Dlv Date']}"
    pdf.multi_cell(95, 4, clean_pdf_text(addr), 0, 'L')
    
    pdf.set_xy(105, y_start_info)
    pdf.multi_cell(95, 4, clean_pdf_text(info['Vendor Name']), 0, 'L')
    pdf.ln(15)

    # Tabel Header
    pdf.set_fill_color(230, 230, 230)
    pdf.set_font('Helvetica', 'B', 8)
    pdf.cell(20, 8, "Req Date", 1, 0, 'C', True)
    pdf.cell(30, 8, "SVO ID", 1, 0, 'C', True)
    pdf.cell(75, 8, "Item Name & Spec", 1, 0, 'C', True)
    pdf.cell(15, 8, "Qty", 1, 0, 'C', True)
    pdf.cell(25, 8, "Price", 1, 0, 'C', True)
    pdf.cell(25, 8, "Total", 1, 1, 'C', True)

    # Tabel Isi
    pdf.set_font('Helvetica', '', 7)
    for _, row in po_data.iterrows():
        desc = f"{row['Item name']}\nSpec: {row['Spec']}"
        x, y = pdf.get_x(), pdf.get_y()
        
        # Kolom kiri
        pdf.cell(20, 12, clean_pdf_text(row['Req. Dlv Date']), 1)
        pdf.cell(30, 12, clean_pdf_text(row['PO SEMENTARA']), 1)
        
        # Multi_cell deskripsi
        pdf.multi_cell(75, 4, clean_pdf_text(desc), 1)
        y_end_desc = pdf.get_y()
        
        # Pindah ke kolom kanan
        pdf.set_xy(x + 125, y)
        pdf.cell(15, 12, f"{row['Ord. Q\'ty']:,.0f}", 1, 0, 'C')
        pdf.cell(25, 12, f"{row['Unit Price']:,.0f}", 1, 0, 'R')
        pdf.cell(25, 12, f"{row['AMOUNT']:,.0f}", 1, 1, 'R')
        
        # Reset posisi Y agar tidak tumpang tindih
        if pdf.get_y() < y_end_desc:
            pdf.set_y(y_end_desc)

    # Mengembalikan objek bytes (Penting untuk Streamlit)
    return bytes(pdf.output())