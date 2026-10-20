# Gabung Excel dari ZIP

Aplikasi Streamlit untuk menggabungkan semua file Excel/CSV di sebuah ZIP.

## Jalankan lokal
1. Install Python 3.11 atau 3.12
2. Buka terminal di folder ini
3. `python -m pip install -r requirements.txt`
4. `python -m streamlit run app.py`

## Publikasikan
1. Buat repository GitHub dan upload `app.py`, `requirements.txt`.
2. Buka https://share.streamlit.io dan login via GitHub.
3. Create app / Deploy app; pilih repository, branch `main`, entrypoint `app.py`.
4. Setelah deployment sukses, bagikan URL `*.streamlit.app`.

## Batasan penting
- ZIP <= 40 MB, maksimal 100 file, total uncompressed <= 150 MB.
- Baris kosong sepenuhnya dilewati; duplikat dipertahankan.
- Kolom dipadankan berdasarkan nama persis setelah trimming spasi. Kolom dengan nama berbeda tidak otomatis disamakan.
- Excel cells numerik yang sudah kehilangan digit presisi tidak dapat dipulihkan.
- Validasi baris hanya menjamin kesetaraan baris yang berhasil dibaca, bukan isi sumber secara sempurna.
- Upload diproses di server Streamlit; jangan unggah data rahasia tanpa perlindungan yang memadai.
