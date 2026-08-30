# Persediaan Backend

1. Salin `Code.gs` ke projek Apps Script yang dipautkan kepada sistem.
2. Jalankan fungsi `setupSheets()` sekali untuk menyediakan tab, header dan data dummy.
3. Tetapkan Script Properties berikut:
   - `PIN_SALT`
   - `PIN_HASH_H01`
   - `PIN_HASH_H02`
   - `PIN_HASH_H03`
   - `SESSION_SECRET`
4. Deploy sebagai Web app dan gunakan URL `/exec` dalam `app.js`.

Jangan simpan PIN asal atau `SESSION_SECRET` dalam repositori awam.
