# 🧾 Invoice Generator

Aplikasi web untuk generate invoice PDF dari data CSV secara otomatis.

## ✨ Fitur

- 📤 Upload file CSV dengan drag & drop
- 📊 Preview dan pilih data yang ingin di-generate
- ⚙️ Konfigurasi info perusahaan dan bank
- 🧾 Generate invoice PDF profesional
- 📦 Download semua invoice sebagai ZIP

## 🚀 Cara Menjalankan

### Lokal
```bash
pip install -r requirements.txt
streamlit run web_app.py
```

### Online
Aplikasi ini dapat diakses di: [Streamlit Cloud](https://share.streamlit.io)

## 📋 Format CSV

File CSV harus memiliki kolom:
| Kolom | Contoh Nama |
|-------|-------------|
| Order ID | `ORDER-ID`, `Order ID` |
| Nama | `Nama Lengkap`, `Nama` |
| Alamat | `Alamat Pengiriman` |
| Ukuran | `Ukuran Kaos (size)` |
| Jumlah | `Jumlah (QTY)` |
| Metode Bayar | `Metode Pembayaran` |
| Status | `STATUS PEMBAYARAN` |

## 🛠️ Tech Stack

- Python 3.8+
- Streamlit
- Pandas
- ReportLab (PDF generation)
