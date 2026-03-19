# Dokumen Spesifikasi (Asprova Platform Terpadu)

## 1. Lingkungan Pengembangan
- Target lingkungan: Windows 10/11 (contoh: 10.0.19045)
- Bahasa: Python
- Framework web: Flask
- Library yang digunakan
  - viewer: `flask`, `sqlite3`, `csv`, `openpyxl` (ekspor Excel)
  - bridge: `flask`, `oracledb` (koneksi Oracle & pengambilan data)
- Port
  - viewer: `5000`
  - bridge: `5001`

## 2. Struktur Folder
- `common/`
  - `templates/base.html`: Layout bersama (dipakai oleh viewer & bridge)
  - `static/css/style.css`: CSS bersama
- `core/`
  - `csv_loader.py`: Proses CSV bersama (deteksi delimiter, pembuatan DictReader, output DB→CSV)
  - `asprova_parser.py`: Deteksi kolom & parsing 1 baris format CSV Asprova (untuk insert ke DB)
- `apps/`
  - `viewer/`: Tampilan schedule dan import CSV
  - `bridge/`: Koneksi Oracle DB dan download CSV
- `config/`: Pengaturan (jika diperlukan)
- `run.py`: Menjalankan viewer/bridge

## 3. Cara Menjalankan
- Menjalankan keduanya (default)
  - `py run.py`
- Menjalankan viewer saja
  - `py run.py viewer`
- Menjalankan bridge saja
  - `py run.py bridge`
- URL
  - viewer: `http://localhost:5000`
  - bridge: `http://localhost:5001`

## 4. Struktur Layar
UI bersama disediakan oleh `common/templates/base.html`, dan masing-masing aplikasi menambahkan halaman berikut:

- Umum (header/footer)
  - viewer: Navigasi ke Gantt / PSI / Import CSV, serta tombol Clear (jika ada data)
  - bridge: Tombol Connect dan tampilan status koneksi (melalui modal)
- viewer
  - Schedule (tampilan schedule)
  - Gantt Chart (toolbar, tooltip, link next process)
  - PSI Viewer (tabel Supply/Demand/Stock bulanan, export Excel)
  - Import CSV (upload CSV)
- bridge
  - Index (pilihan jenis download CSV dalam bentuk kartu)
  - Confirm/Connect (konfirmasi download dan modal koneksi Oracle)

## 5. Fitur
- viewer
  - Import CSV
    - Dekode dengan dukungan UTF-8 BOM
    - Deteksi delimiter: `,`/`\t`/`;`/`|`
    - Auto mapping kolom dari header
  - Penyimpanan ke SQLite
    - Memasukkan baris schedule ke `schedule.db`
    - Update skema bila kolom yang diperlukan kurang
  - Tampilan schedule
    - Switch periode: weekly / 2week / 3week / monthly
    - Filter berdasarkan mesin
  - Gantt Chart
    - Tooltip detail
    - Link next process (dengan kondisi filter)
  - PSI Viewer
    - Tampilan bulanan Supply/Demand/Stock
    - Export Excel
  - Clear data (hapus semua)
- bridge
  - Koneksi Oracle DB (disimpan pada sesi)
  - Download CSV
    - Integrated Master
    - Item Table
    - Order Table
    - Resource Table
    - Inventory Table
  - Pemilihan lokasi penyimpanan via UI browser

## Rekomendasi Format File
- Dokumentasi: Markdown (`.md`)
  - Contoh: `docs/spec_id.md`
  - Mudah dikelola dan terlihat jelas untuk review/diff

