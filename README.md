# aipp-convert 🚀

Aplikasi web personal (privat) untuk konversi dokumen secara mandiri tanpa batasan (unlimited). Dibangun dengan pendekatan **Mobile-First Responsive Design** yang modern, bersih, intuitif, dan aman.

---

## ✨ Fitur Utama

1. **Konversi PDF ke Word (.docx)**
   - **PDF Berbasis Teks Asli**: Menggunakan `pdf2docx` untuk menjaga tata letak, tabel, paragraf, dan format dokumen tetap rapi.
   - **PDF Hasil Scan / Foto (OCR)**: Otomatis mendeteksi dokumen pindaian atau gambar berbayang, lalu menjalankan `pytesseract` + `pdf2image` + `python-docx` untuk mengekstrak teks.
   - **Pilihan Bahasa OCR**: Mendukung Bahasa Indonesia (`ind`), English (`eng`), dan gabungan keduanya.
   - **Opsi Paksa OCR**: Tersedia tombol toggle jika Anda ingin memaksa pemrosesan OCR secara manual.

2. **Konversi Word (.docx / .doc) ke PDF**
   - Menggunakan engine `docx2pdf` (Microsoft Word COM di Windows) atau fallback ke LibreOffice headless untuk hasil PDF yang presisi.

3. **Antarmuka Mobile-First & Touch-Friendly**
   - Didesain secara optimal untuk smartphone dengan pengalaman sentuh yang mulus, serta tetap proporsional dan anggun di tablet dan desktop.
   - Tab navigasi geser modern: `[PDF ke Word]` dan `[Word ke PDF]`.
   - Area Drag-and-Drop file uploader dengan inspeksi nama berkas, ukuran, dan validasi format.
   - Multi-step loading progress bar dan status deteksi dokumen secara *real-time*.
   - Kartu hasil konversi dengan tombol unduh langsung.

4. **Pembersihan Otomatis (Auto File Cleanup)**
   - Berkas upload sementara langsung dihapus setelah konversi selesai.
   - Daemon background thread otomatis membersihkan berkas hasil yang berusia lebih dari 1 jam untuk menghemat ruang penyimpanan server.

5. **Privat & Bebas Batasan**
   - 100% diproses di server lokal Anda sendiri. Data tidak dikirim ke pihak ketiga dan tidak ada kuota harian.

---

## 📂 Struktur Folder Proyek

```text
aipp-convert/
├── app.py                  # Server backend Flask & logika konversi/OCR
├── requirements.txt        # Daftar pustaka Python yang dibutuhkan
├── .env.example            # Contoh variabel lingkungan (jalur OCR & Poppler)
├── .gitignore              # Konfigurasi pengabaian Git
├── README.md               # Dokumentasi instalasi dan panduan lengkap
├── templates/
│   └── index.html          # Antarmuka web modern (Tailwind CSS CDN + Lucide Icons)
└── storage/
    ├── uploads/            # Direktori sementara file yang diunggah
    └── converted/          # Direktori berkas hasil konversi
```

---

## 🛠️ Persyaratan Sistem & Dependensi

Aplikasi ini menggunakan beberapa pustaka Python dan dependensi sistem (terutama untuk fitur OCR dan rendering PDF scan).

### 1. Instalasi Pustaka Python

Buka terminal atau PowerShell di dalam folder proyek, lalu buat lingkungan virtual (opsional namun disarankan) dan pasang dependensi:

```bash
# 1. Buat virtual environment (opsional)
python -m venv venv

# Aktifkan di Windows:
venv\Scripts\activate
# Atau aktifkan di Linux/macOS:
source venv/bin/activate

# 2. Pasang requirements
pip install -r requirements.txt
```

---

### 2. Dependensi Sistem Tambahan (Untuk Fitur OCR Scan & Word to PDF)

#### A. Windows

1. **Tesseract OCR (Untuk Membaca PDF Scan):**
   - Unduh installer resmi Windows di: [UB-Mannheim Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
   - Saat proses instalasi, centang **Additional language data (download)** -> pilih **Indonesian** agar mendukung OCR Bahasa Indonesia.
   - Lokasi instalasi standar biasanya di `C:\Program Files\Tesseract-OCR\tesseract.exe`. Aplikasi `aipp-convert` otomatis mendeteksi lokasi ini.

2. **Poppler (Untuk Mengubah Halaman PDF Scan Menjadi Gambar):**
   - Unduh rilis biner Windows dari: [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases/)
   - Ekstrak arsip ZIP, misalnya ke `C:\poppler`.
   - Pastikan folder `bin` berada di `C:\poppler\Library\bin` atau tambahkan jalur tersebut ke Environment Variables `PATH` Windows.
   - Alternatif cepat via Chocolatey:
     ```powershell
     choco install poppler
     ```

3. **Microsoft Word atau LibreOffice (Untuk Word ke PDF):**
   - Jika Anda memiliki Microsoft Word di Windows, modul `docx2pdf` akan langsung berfungsi otomatis.
   - Jika tidak memiliki MS Word, pasang [LibreOffice](https://www.libreoffice.org/download/download/) (gratis & open-source).

---

#### B. Linux (Ubuntu / Debian)

Jalankan perintah berikut:

```bash
sudo apt update
# Pasang Tesseract OCR dan data bahasa Indonesia + Inggris
sudo apt install -y tesseract-ocr tesseract-ocr-ind tesseract-ocr-eng

# Pasang Poppler utils (pdftoppm)
sudo apt install -y poppler-utils

# Pasang LibreOffice untuk konversi Word ke PDF
sudo apt install -y libreoffice
```

---

## ⚙️ Konfigurasi Variabel Lingkungan (.env)

Jika Tesseract atau Poppler dipasang di direktori khusus (custom directory), Anda dapat menyalin file `.env.example` menjadi `.env`:

```bash
copy .env.example .env   # di Windows
cp .env.example .env     # di Linux
```

Sesuaikan isi `.env`:
```ini
PORT=5000
TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe
POPPLER_PATH=C:\poppler\Library\bin
```

---

## 🚀 Menjalankan Aplikasi

Jalankan server Flask dengan perintah:

```bash
python app.py
```

Buka peramban (browser) di komputer atau ponsel Anda (yang berada dalam satu jaringan Wi-Fi/LAN):
- Akses Lokal: **http://127.0.0.1:5000**
- Akses dari HP via Wi-Fi: **http://IP-KOMPUTER-ANDA:5000** (contoh: `http://192.168.1.15:5000`)

---

## 🔒 Privasi dan Keamanan

- Semua dokumen dikonversi langsung pada memori dan disk lokal komputer/server Anda sendiri.
- File input asli segera dihapus setelah selesai dianalisis dan dikonversi.
- File hasil konversi dihapus otomatis oleh daemon background cleaner setiap 1 jam.
