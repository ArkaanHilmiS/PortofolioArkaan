# Portfolio Arkaan Hilmi Suharsoyo

Portfolio website untuk Arkaan Hilmi Suharsoyo - Mahasiswa Sistem Informasi, Institut Teknologi Sepuluh Nopember Surabaya.

## 📁 Struktur Proyek

```
PortofolioArkaan/
├── index.html              # Halaman utama HTML
├── css/
│   └── style.css          # Semua styling dan animasi
├── js/
│   └── script.js          # Semua fungsi JavaScript dan interaktivitas
├── assets/
│   ├── images/            # Folder untuk foto, gambar, dan grafis
│   └── documents/         # Folder untuk file PDF, sertifikat, dll
└── README.md              # Dokumentasi proyek ini
```

## 🎯 Fitur Utama

- **Landing Page Interaktif** - Pemilihan role pengunjung (HRD, Mahasiswa, Kolega, dll)
- **Section Lengkap** - Pendidikan, Pengalaman, Proyek, Pelatihan, Skills, Galeri
- **Form Kontak** - Kirim pesan langsung
- **Download PPT** - Generate dan download portfolio dalam format PowerPoint
- **Freelance Services** - Penawaran layanan freelance
- **Responsive Design** - Tampilan optimal di berbagai ukuran layar
- **Smooth Animations** - Scroll reveal dan animasi interaktif

## 🚀 Cara Menggunakan

### Melihat Portfolio
1. Buka file `index.html` di browser favorit Anda
2. Pilih role Anda di landing page
3. Jelajahi berbagai section yang tersedia

### Menambahkan Gambar
Letakkan file gambar (JPG, PNG, dll) ke folder `assets/images/`, kemudian referensikan di HTML:
```html
<img src="assets/images/nama-gambar.jpg" alt="Deskripsi">
```

### Menambahkan Dokumen
Letakkan file PDF, sertifikat, atau dokumen lainnya ke folder `assets/documents/`, kemudian buat link di HTML:
```html
<a href="assets/documents/sertifikat.pdf" download>Download Sertifikat</a>
```

## 🛠️ Teknologi yang Digunakan

- **HTML5** - Struktur halaman
- **CSS3** - Styling dengan CSS Custom Properties (Variables)
- **Vanilla JavaScript** - Interaktivitas tanpa library tambahan
- **PptxGenJS** - Library untuk generate file PowerPoint
- **Google Fonts** - Tipografi (Playfair Display, DM Sans, DM Mono)

## ⚙️ Customisasi

### Mengubah Warna Tema
Edit variabel CSS di file `css/style.css` section `:root`:
```css
:root {
  --accent: #2563eb;        /* Warna aksen utama */
  --accent-light: #3b82f6;  /* Warna aksen terang */
  --gold: #c9a84c;          /* Warna gold */
  /* ... dan lainnya */
}
```

### Menambah/Mengubah Konten
Edit file `index.html` dan cari section yang ingin diubah (contoh: `<section id="education">`)

### Modifikasi JavaScript
Edit file `js/script.js` untuk mengubah perilaku interaktif seperti animasi, form handling, dll

## 📝 Catatan

- Pastikan struktur folder tetap terjaga agar path CSS dan JS tidak error
- File `index.html.backup` adalah backup dari versi sebelumnya (bisa dihapus jika tidak diperlukan)
- Untuk deployment online, upload semua file dan folder ke web hosting

## 📧 Kontak

- **Email**: arkaan@email.com
- **LinkedIn**: [Profil LinkedIn]
- **GitHub**: [Profil GitHub]

---

**Dibuat dengan ❤️ oleh Arkaan Hilmi Suharsoyo**
