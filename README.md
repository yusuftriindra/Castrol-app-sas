# Castrol - GitHub Pages + Google Apps Script

Frontend aplikasi Castrol dijalankan dari GitHub Pages.
Google Apps Script tetap menjadi backend untuk Google Sheets.

## 1. Deploy backend Apps Script

1. Buka project Google Apps Script yang terhubung dengan Google Sheet.
2. Ganti `Code.gs` dengan `Code.gs` dari repository ini.
3. Pastikan fungsi-fungsi lama dari project Anda tetap berada di `Code.gs`.
4. Deploy sebagai **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Salin URL Web App yang berakhiran `/exec`.

Contoh:

`https://script.google.com/macros/s/XXXXXXXXXXXX/exec`

## 2. Atur URL backend di index.html

Buka `index.html`.

Cari:

```javascript
const CASTROL_APPS_SCRIPT_URL = 'PASTE_APPS_SCRIPT_WEB_APP_URL_HERE';
```

Ganti menjadi URL Web App Anda:

```javascript
const CASTROL_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/XXXXXXXXXXXX/exec';
```

Jangan menambahkan `/bridge=1` sendiri. Script akan menambahkannya otomatis.

## 3. Upload ke GitHub

Repository minimal:

```text
Castrol-GitHub-Repository/
├── index.html
├── Code.gs
└── README.md
```

Untuk GitHub Pages, file utama yang digunakan adalah:

`index.html`

`Code.gs` tetap merupakan source backend Apps Script dan tidak dijalankan oleh GitHub Pages.

## 4. Aktifkan GitHub Pages

GitHub:

**Settings → Pages**

Pilih:

- Source: Deploy from a branch
- Branch: `main`
- Folder: `/ (root)`

Setelah aktif, GitHub memberikan alamat seperti:

`https://USERNAME.github.io/REPOSITORY/`

Buka alamat tersebut untuk menjalankan aplikasi.

## Cara kerja

```text
Browser
   ↓
GitHub Pages / index.html
   ↓
hidden iframe
   ↓
Google Apps Script Web App
   ↓
google.script.run
   ↓
Google Sheet
```

`google.script.run` pada `index.html` tetap dipertahankan melalui bridge. Jadi fungsi-fungsi aplikasi yang sudah ada tidak perlu ditulis ulang satu per satu.

## Catatan

- Halaman utama yang dilihat user berasal dari GitHub Pages, sehingga banner Apps Script tidak menjadi bagian dari halaman utama.
- Jangan menghapus fungsi-fungsi backend lama dari `Code.gs`.
- Jika Apps Script di-deploy ulang dengan URL baru, ubah `CASTROL_APPS_SCRIPT_URL` di `index.html`.
- Login, Input Order, Riwayat, Monitoring, Outlet BWS, Outlet Kontrak, Pencapaian, Pricelist, Pesanan Outlet, dan Tambah Customer tetap menggunakan fungsi backend yang sudah ada.
