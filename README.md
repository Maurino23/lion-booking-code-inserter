# 🦁 DCR-PAXLIST Automation Tool

Aplikasi otomatisasi untuk menggabungkan data Booking Code dari file PAXLIST ke dalam file DCR.

## Fitur
- Upload file DCR (.xlsx) dan PAXLIST (.xlsx/.csv)
- Validasi otomatis struktur file
- Preview data sebelum processing
- Formatting otomatis (highlight JUMPSEAT)
- Download hasil dengan timestamp

## Cara Menggunakan
1. Upload file DCR yang berisi kolom 'CREW LIST'
2. Upload file PAXLIST yang berisi kolom 'Crew ID' dan 'Booking Code'
3. Sesuaikan pengaturan jika diperlukan
4. Klik 'Proses File'
5. Download hasil yang sudah diproses

## Persyaratan File
- **DCR**: Harus memiliki kolom 'CREW LIST'
- **PAXLIST**: Harus memiliki kolom 'Crew ID' dan 'Booking Code'

## Instalasi Lokal
```bash
pip install -r requirements.txt
streamlit run app.py
