# ServeOne Stage 1 - Paket Tempel Proyek

## Tujuan tahap 1
Tahap ini **belum memigrasikan database ke Supabase**. Fokusnya adalah:
- menstabilkan struktur proyek,
- menyeragamkan UI,
- memusatkan login, sidebar, dan theme,
- menyiapkan pondasi Supabase untuk tahap berikutnya,
- **tanpa merusak logika 4 file asli**.

## Cara pakai

### 1. Copy file berikut ke root proyek Anda
- `config.py`
- `requirements.txt`
- `.env.example`
- folder `utils/`
- folder `services/`

### 2. Ganti `app.py`
Ganti isi `app.py` lama Anda dengan file `app.py` dari paket ini.

### 3. Attendance dan To-Do List
Anda bisa **langsung** ganti file:
- `pages/2_📸_Attendance.py`
- `pages/3_✅_To_Do_List.py`

Kedua file itu sudah disesuaikan agar tetap mempertahankan logika lama, tetapi memakai:
- theme global,
- sidebar global,
- guard login global.

### 4. Operation
Karena `Operation.py` Anda paling besar dan paling sensitif, paket ini **tidak memaksa mengganti seluruh logika lama**.

Gunakan file `pages/1_🚚_Operation.py` dari paket ini sebagai **header baru**, lalu:
- ambil file Operation lama Anda,
- hapus bagian ini dari file lama:
  - proteksi login manual,
  - CSS sidebar lokal,
  - custom sidebar lokal,
  - inisialisasi `dn_counter`, `last_dn_date`, `dn_vendor`, `no_pol`, `dn_remarks`,
  - helper `clean_val`, `safe_float`, `format_currency`, `format_qty`,
- tempelkan sisa logika lama di bawah header baru.

### 5. Logo
Taruh `logo.png` di root proyek agar muncul di sidebar.

### 6. Jalankan
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Kenapa Operation belum saya ganti total?
Karena modul itu sudah berisi:
- PDF engine,
- parsing data CSV,
- ZIP export,
- logic state DN,
- dan banyak formatting.

Tahap 1 sengaja menjaga modul itu tetap aman.
Tahap 2 baru cocok untuk memecah Operation ke:
- `pdf_engine/po_pdf.py`
- `pdf_engine/dn_pdf.py`
- `services/operation_service.py`
- `utils/formatters.py`

## Struktur proyek yang dihasilkan
```bash
serveone_erp/
├── app.py
├── config.py
├── requirements.txt
├── .env.example
├── utils/
├── services/
└── pages/
```
