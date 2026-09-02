# Penjurian Poster Digital PUICE 2026

Web app rasmi untuk penjurian Poster Digital PUICE 2026.

## Komponen

- GitHub Pages: antaramuka responsif untuk hakim.
- Google Apps Script: API, pengesahan PIN, validasi dan audit markah.
- Google Sheets: senarai poster, hakim, penjurian, audit dan tetapan.

## Keselamatan

PIN tidak disimpan dalam repositori atau HTML. Backend menyimpan hash PIN dan rahsia sesi melalui Apps Script Script Properties. Jumlah markah sentiasa dikira semula pada server.

## Struktur

- `index.html` — struktur aplikasi.
- `styles.css` — reka bentuk responsif dan sokongan reduced motion.
- `app.js` — aliran penjurian, DQ, galeri dan rumusan hakim.
- `backend/Code.gs` — sumber backend Google Apps Script.
- `pemenang/` — halaman keputusan rasmi awam selepas penjurian lengkap.

© Unit Sumber Teknologi Pendidikan PPDPU | PUICE 2026
