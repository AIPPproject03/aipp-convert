# aipp-convert

Aplikasi web Flask untuk mengonversi dokumen PDF dan Word secara lokal. Mendukung PDF berbasis teks maupun PDF hasil pindai melalui OCR, serta pemrosesan antrean hingga 30 berkas.

> **Catatan:** Konversi berjalan pada komputer/server tempat aplikasi dijalankan. Fitur OCR dan Word ke PDF memerlukan program sistem tambahan. Jalankan aplikasi hanya di jaringan yang Anda percaya; server pengembangan Flask saat ini menggunakan mode debug.

## Fitur

- **PDF ke Word (`.docx`)**: aplikasi mencoba mendeteksi lapisan teks pada maksimal lima halaman pertama. PDF berbasis teks diproses dengan `pdf2docx`; PDF scan diproses dengan OCR. Jika konversi tata letak gagal, aplikasi mencoba OCR sebagai fallback.
- **OCR PDF scan**: menggunakan Tesseract dan Poppler. Bahasa yang tersedia di antarmuka adalah Indonesia (`ind`), Inggris (`eng`), atau gabungan (`ind+eng`). OCR juga dapat dipaksakan untuk PDF berbasis teks.
- **Word ke PDF**: menerima `.docx` dan `.doc`. Di Windows aplikasi mencoba Microsoft Word melalui COM, lalu `docx2pdf` dan LibreOffice sebagai fallback. Di sistem lain, LibreOffice dapat digunakan sebagai engine konversi.
- **Batch**: antrean antarmuka dibatasi hingga 30 berkas dan diproses satu per satu. Batas ukuran request di backend adalah 200 MB; karena tiap berkas dikirim dalam request tersendiri, batas tersebut berlaku per request.
- **Unduhan**: hasil dapat diunduh satu per satu atau digabungkan menjadi ZIP.
- **Pembersihan**: berkas masukan sementara dihapus setelah konversi. Sesi dihapus saat halaman dimuat ulang/ditutup, dan pembersih latar menghapus berkas atau folder sesi yang lebih tua dari 30 menit, setiap 5 menit. Tombol **Bersihkan Memori** menghapus seluruh isi folder upload dan hasil.
- **Nama keluaran**: menggunakan nama dasar berkas asal dengan ekstensi keluaran yang sesuai.

## Struktur proyek

```text
aipp-convert/
├── app.py                 # Server Flask, endpoint, konversi, OCR, dan pembersihan
├── requirements.txt       # Dependensi Python
├── .env.example           # Contoh konfigurasi Tesseract, Poppler, dan port
├── .gitignore
├── README.md
├── templates/
│   └── index.html         # Antarmuka dan logika antrean di browser
└── storage/
    ├── uploads/           # Masukan sementara; dibuat otomatis jika belum ada
    └── converted/         # Hasil sementara; dibuat otomatis jika belum ada
```

Antarmuka memuat Tailwind CSS, Lucide Icons, dan Google Fonts dari CDN, sehingga tampilan lengkap memerlukan koneksi internet. Konversi dokumen tetap berlangsung di server lokal.

## Persyaratan

- Python 3 dan pip.
- Untuk PDF scan/OCR: **Tesseract OCR**, data bahasa yang ingin digunakan, serta **Poppler**.
- Untuk Word ke PDF: Microsoft Word di Windows atau LibreOffice. Word COM hanya tersedia pada Windows dengan Microsoft Word terpasang.

### Memasang dependensi sistem

#### Windows

1. Pasang [Tesseract OCR untuk Windows](https://github.com/UB-Mannheim/tesseract/wiki) dan data bahasa Indonesia (`ind`) dan/atau Inggris (`eng`). Lokasi umum executable adalah `C:\Program Files\Tesseract-OCR\tesseract.exe`; aplikasi mencoba mendeteksi lokasi ini secara otomatis.
2. Pasang [Poppler untuk Windows](https://github.com/oschwartz10612/poppler-windows/releases/) dan pastikan `pdftoppm` tersedia di `PATH`, atau atur `POPPLER_PATH` ke folder `bin` Poppler.
3. Untuk Word ke PDF, pasang Microsoft Word atau [LibreOffice](https://www.libreoffice.org/download/download/).

#### Ubuntu / Debian

```bash
sudo apt update
sudo apt install -y tesseract-ocr tesseract-ocr-ind tesseract-ocr-eng poppler-utils libreoffice
```

Pilih bahasa Tesseract yang diperlukan. `libreoffice` hanya diperlukan jika ingin memakai engine tersebut untuk Word ke PDF.

## Instalasi aplikasi

Dari direktori proyek, buat lingkungan virtual dan pasang pustaka Python:

```bash
python -m venv .venv
```

Aktifkan lingkungan virtual:

```powershell
# Windows PowerShell
.venv\Scripts\Activate.ps1
```

```bash
# Linux / macOS
source .venv/bin/activate
```

Lalu pasang dependensi:

```bash
python -m pip install -r requirements.txt
```

## Konfigurasi

Sebagian besar instalasi Windows umum dideteksi otomatis. Bila perlu, salin `.env.example` ke `.env` lalu isi jalur sesuai komputer Anda:

```ini
PORT=5000
TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe
POPPLER_PATH=C:\poppler\Library\bin
```

`TESSERACT_CMD` menunjuk ke executable `tesseract.exe` (atau executable Tesseract di sistem lain). `POPPLER_PATH` menunjuk ke direktori bin Poppler. Variabel yang tidak diperlukan dapat dikosongkan atau dihilangkan. Jangan masukkan file `.env` ke repositori.

## Menjalankan

```bash
python app.py
```

Buka `http://127.0.0.1:5000`. Port dapat diubah melalui `PORT` di `.env`.

Secara default, Flask mendengarkan pada `0.0.0.0`, sehingga perangkat lain di jaringan lokal mungkin dapat mengaksesnya melalui `http://IP-KOMPUTER:5000`. Mode debug aktif pada entry point saat ini; gunakan hanya untuk pengembangan di jaringan tepercaya, bukan sebagai server produksi atau layanan internet publik.

## Endpoint aplikasi

| Metode | Endpoint | Kegunaan |
| --- | --- | --- |
| `GET` | `/` | Menampilkan antarmuka konversi |
| `GET` | `/api/system-status` | Mengembalikan status kemampuan tool sistem dalam JSON |
| `POST` | `/api/convert/pdf-to-word` | Menerima satu PDF dan opsi `force_ocr`, `ocr_lang`, serta `session_id` sebagai form data |
| `POST` | `/api/convert/word-to-pdf` | Menerima satu berkas Word dan `session_id` sebagai form data |
| `GET` | `/api/download/<session>/<filename>` | Mengunduh satu hasil konversi |
| `GET` | `/api/download-zip/<session_id>` | Mengunduh hasil sesi sebagai ZIP |
| `POST` | `/api/cleanup-session` | Menghapus file sementara milik sesi yang dikirim |
| `POST` | `/api/cleanup-all` | Menghapus seluruh isi folder upload dan hasil |

## Catatan penggunaan dan batasan

- Deteksi PDF memeriksa hingga lima halaman pertama dan memakai ambang rata-rata 40 karakter per halaman untuk memilih jalur teks atau OCR. Hasil deteksi bersifat heuristik.
- OCR menghasilkan teks yang dapat diedit, tetapi tidak merekonstruksi tata letak halaman seperti tabel atau kolom secara presisi. Kualitas bergantung pada resolusi dan kondisi dokumen.
- Pembuatan DOCX melalui OCR menggunakan resolusi 200 DPI.
- Nama hasil dibentuk dari nama dasar file. File berbeda dengan nama dasar yang sama dalam satu sesi bisa menghasilkan nama tujuan yang sama.
- Pembersihan sesi saat browser ditutup bergantung pada sinyal browser; pembersih latar menjadi mekanisme pembersihan tambahan.
- Pastikan Tesseract mengenali bahasa yang dipilih. Jika tidak, konversi OCR akan gagal.

## Privasi

Secara bawaan, file diproses dan disimpan sementara pada komputer tempat server berjalan, bukan dikirim ke layanan konversi eksternal. Aplikasi tidak memiliki login atau pembatasan akses; siapa pun yang dapat mencapai server dapat menggunakan endpoint-nya, termasuk endpoint pembersihan seluruh penyimpanan. Pertimbangkan hal ini sebelum membuka akses di luar komputer pribadi atau jaringan tepercaya.
