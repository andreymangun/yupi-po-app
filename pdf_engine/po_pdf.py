from __future__ import annotations
import os
from datetime import date
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from utils.formatters import clean_val, format_currency, format_qty, safe_float

class ServeonePO(FPDF):
    def header(self):
        # Logo Kiri
        if os.path.exists("logo.png"):
            self.image("logo.png", 10, 2, 45)
            
        # P/O Paper text
        self.set_font("Helvetica", "B", 16)
        self.set_y(10)
        self.cell(0, 10, "P/O Paper", align="C", new_x=XPos.LMARGIN, new_y=YPos.TOP)
        
        # Stempel Serveone Kanan (Sejajar Logo)
        if os.path.exists("stamp_serveone.png"):
            self.image("stamp_serveone.png", 150, 5, 40)
            
        self.ln(20)

    # --- TAMBAHKAN FUNGSI FOOTER INI ---
    def footer(self):
        # Posisi 1.5 cm dari bawah halaman
        self.set_y(-15)
        # Font italic, ukuran 8
        self.set_font("Helvetica", "I", 8)
        # Warna teks abu-abu gelap
        self.set_text_color(128, 128, 128)
        # Mencetak "Page X of Y" (Teks {nb} akan otomatis diganti dengan total halaman oleh alias_nb_pages)
        self.cell(0, 10, f"Page {self.page_no()} of {{nb}}", align="C")
        # Kembalikan warna ke hitam untuk berjaga-jaga
        self.set_text_color(0, 0, 0)

def generate_po_pdf(po_data, po_number: str):
    pdf = ServeonePO()
    # Wajib memanggil ini agar total halaman ({nb}) bisa dihitung
    pdf.alias_nb_pages() 
    pdf.add_page()
    
    # ... (SISA KODE generate_po_pdf SAMA SEPERTI SEBELUMNYA) ...
    info = po_data.iloc[0]
    vendor_name = clean_val(info.get("Vendor Name", "Unknown Vendor"))

    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(30, 7, "PO YUPI :")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 7, clean_val(po_number), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(30, 7, "PO Date :")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 7, date.today().strftime("%d/%m/%Y"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(5)

    y_anchor = pdf.get_y()
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(95, 5, "CLIENT:")
    pdf.set_xy(105, y_anchor)
    pdf.cell(95, 5, "VENDOR:")
    pdf.ln(6)

    pdf.set_font("Helvetica", "", 8)
    y_content = pdf.get_y()
    client_info = "PT. SERVEONE MRO INDONESIA\nJALAN KENARI RAYA BLOK G NO. 19,\nKAWASAN DELTA SILICON V, LIPPO CIKARANG,\nRT. 000 RW. 000, CICAU, CIKARANG PUSAT,\nKAB. BEKASI, JAWA BARAT"
    vendor_info = f"{vendor_name}\n{clean_val(info.get('Vendor Address'))}"
    pdf.set_xy(10, y_content)
    pdf.multi_cell(92, 4, client_info)
    pdf.set_xy(105, y_content)
    pdf.multi_cell(95, 4, vendor_info)
    
    pdf.set_y(max(pdf.get_y(), y_content + 15) + 10)

    site_raw = clean_val(info.get("SITE (IDN/KRG)", info.get("SITE"))).upper()
    site_label = "IDN" if "IDN" in site_raw else "KRG"
    full_addr = "Jl. Pancasila IV, Desa/Kelurahan Cicadas, Kec Gunung Putri, Kab Bogor, Provinsi Jawa Barat, Indonesia" if site_label == "IDN" else "Jl. Grompol Jambangan Km 5, Muringan, Desa Kaliwuluh, Kecamatan Kebak Kramat Kabupaten Karanganyar Provinsi Jawa Tengah, Indonesia"

    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 6, f"DELIVERY ADDRESS: PT. YUPI INDO JELLY GUM Tbk ({site_label})", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 8)
    pdf.multi_cell(0, 5, full_addr)
    pdf.ln(5)

    pdf.set_fill_color(139, 0, 0)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 7)
    cols = [
        ("Deliv Req", 20),
        ("SERVEONE PO", 25),
        ("Item Code", 15),
        ("Item Name & Spec", 55),
        ("Qty", 15),
        ("Unit", 10),
        ("Price", 25),
        ("Amount", 25),
    ]
    for txt, w in cols:
        pdf.cell(w, 8, txt, border=1, fill=True, align="C")
    pdf.ln()

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 7)
    
    total_dpp = 0
    total_ppn = 0
    total_pph = 0
    remark_list = []
    req_date_col = next((c for c in po_data.columns if "REQ" in str(c).upper()), "Req. Dlv Date")

    for _, row in po_data.iterrows():
        name_spec = f"{clean_val(row.get('Item name'))}\n{clean_val(row.get('Spec'))}"
        
        qty = safe_float(row.get("Qty Final", row.get("Ord. Q'ty")))
        price = safe_float(row.get("PURCHASE PRICE"))
        amount = qty * price
        
        item_ppn_rate = 0.11 if row.get("Pajak PPN 11%", False) else 0.0
        item_pph_rate = 0.02 if row.get("Pajak PPh 2%", False) else 0.0 
        
        total_dpp += amount
        total_ppn += (amount * item_ppn_rate)
        total_pph += (amount * item_pph_rate)

        if clean_val(row.get("REMARK YUPI")):
            remark_list.append(clean_val(row.get("REMARK YUPI")))

        lines = pdf.multi_cell(55, 4, name_spec, dry_run=True, output="LINES")
        row_h = max(10, len(lines) * 4 + 2)

        # Ubah batas y bawah sedikit untuk memberi ruang pada footer
        if pdf.get_y() + row_h > 270: 
            pdf.add_page()
        x, y = pdf.get_x(), pdf.get_y()

        pdf.cell(20, row_h, clean_val(row.get(req_date_col)), border=1, align="C")
        pdf.cell(25, row_h, clean_val(row.get("PO SEMENTARA")), border=1, align="C")
        pdf.cell(15, row_h, clean_val(row.get("Item Yupi")), border=1, align="C")
        pdf.cell(55, row_h, "", border=1)
        pdf.cell(15, row_h, format_qty(qty), border=1, align="C")
        pdf.cell(10, row_h, clean_val(row.get("Unit")), border=1, align="C")
        pdf.cell(25, row_h, format_currency(price), border=1, align="R")
        pdf.cell(25, row_h, format_currency(amount), border=1, align="R")

        pdf.set_xy(x + 60, y + 1)
        pdf.multi_cell(55, 4, name_spec, border=0, align="L")
        pdf.set_xy(10, y + row_h)

    curr_raw = clean_val(info.get("CURRENCY")).upper() or "IDR"
    grand_total = total_dpp + total_ppn - total_pph

    pdf.set_font("Helvetica", "B", 8)
    totals = [
        (f"TOTAL AMOUNT ({curr_raw})", total_dpp),
        (f"PPN ({curr_raw})", total_ppn),
        (f"PPh Potongan ({curr_raw})", -total_pph if total_pph > 0 else 0),
        (f"GRAND TOTAL ({curr_raw})", grand_total)
    ]
    
    # Cek jika ruang tidak cukup untuk blok total
    if pdf.get_y() > 250:
         pdf.add_page()

    for label, val in totals:
        pdf.set_x(140)
        pdf.cell(35, 5, label, border=1, align="R")
        pdf.cell(25, 5, format_currency(val), border=1, align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 6, "REMARKS:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "", 8)
    pdf.multi_cell(130, 5, ", ".join(set(remark_list)) if remark_list else "-")

    return bytes(pdf.output()), site_label, vendor_name
