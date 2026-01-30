# Panduan Deployment Aplikasi Absensi MI ISLAMADINA ke Cloudflare

Panduan ini akan membantu Anda mendeloy aplikasi absensi ke Cloudflare Pages dengan database D1.

## Prasyarat
- Akun Cloudflare (Gratis).
- Node.js & NPM terinstall di komputer (untuk development/upload).
- Git terinstall (disarankan) atau bisa upload folder langsung.

---

## Langkah 1: Persiapan Database D1

1. **Login ke Dashboard Cloudflare**.
2. Masuk ke menu **Workers & Pages** > **D1**.
3. Klik **Create database**.
4. Beri nama database, misal: `db-absensi-mi`.
5. Klik **Create**.
6. Setelah jadi, masuk ke detail database tersebut.
7. Pilih tab **Console**.
8. **PENTING**: Kita perlu membuat tabel (schema). Salin kode dari file `migrations/0001_initial.sql` (yang ada di folder proyek Anda) dan paste ke Console D1, lalu klik **Execute**.
   - Ini akan membuat tabel `users` dan `attendance` serta mengisi data awal guru.

---

## Langkah 2: Buat Cloudflare Pages

1. Masuk ke menu **Workers & Pages**.
2. Klik **Create application** > **Pages**.
3. Pilih **Upload assets** (jika ingin upload manual folder) ATAU **Connect to Git** (jika source code sudah di GitHub/GitLab).
   - **Rekomendasi**: Gunakan **Connect to Git** agar update lebih mudah.
   
   **Opsi A: Connect to Git (Disarankan)**
   - Pilih repository Git Anda (misal `absensi-mi`).
   - Setup Builds:
     - **Production branch**: `main` (atau master).
     - **Framework preset**: `None`.
     - **Build command**: (kosongkan).
     - **Build output directory**: `public`.
   - Klik **Save and Deploy**.

   **Opsi B: Direct Upload**
   - Beri nama project: `absensi-mi-islamadina`.
   - Upload folder `public` dari komputer Anda.
   - *Catatan: Jika pakai Direct Upload, Fungsi Backend (Functions) mungkin butuh upload terpisah via Wrangler CLI.* **Sangat disarankan memakai Git integration atau Wrangler CLI untuk aplikasi fullstack.**

---

## Langkah 3: Menghubungkan Database (Binding)

Agar aplikasi bisa membaca database D1, kita harus melakukan "Binding".

1. Buka settings project Pages Anda di dashboard Cloudflare.
2. Pergi ke **Settings** > **Functions**.
3. Scroll ke bagian **D1 Database Bindings**.
4. Klik **Add binding**.
5. Isi konfigurasi berikut:
   - **Variable name**: `DB` (Harus PERSIS besar semua).
   - **D1 Database**: Pilih `db-absensi-mi` yang dibuat di Langkah 1.
6. Klik **Save**.
7. **Redeploy**: Anda harus men-deploy ulang agar perubahan binding ini efek.
   - Jika pakai Git: Masuk tab **Deployments** > di commit terakhir klik titik tiga `...` > **Retry deployment**.

---

## Langkah 4: Setup Domain

1. Buka settings project Pages Anda.
2. Pergi ke tab **Custom domains**.
3. Klik **Set up a custom domain**.
4. Masukkan domain: `absen.miislamadina.sch.id`.
   - Pastikan domain utama `miislamadina.sch.id` sudah terdaftar di Cloudflare DNS Anda.
5. Cloudflare akan otomatis mengatur DNS record (CNAME).
6. Klik **Activate domain**. Tunggu propagasi DNS (bisa instan atau beberapa menit).

---

## Langkah 5: Testing Aplikasi

1. Buka `https://absen.miislamadina.sch.id`.
2. **Cek Dropdown**: Apakah nama-nama guru muncul?
   - Jika tidak muncul, cek tab Network inspect (F12). Jika error 500, cek Logs di dashboard Cloudflare Pages. Pastikan Binding D1 Variable name adalah `DB`.
3. **Coba Absen**: Pilih nama -> Klik Hadir.
   - Harus muncul popup sukses.
   - Nama harus muncul di daftar bawah.
4. **Coba Absen Lagi**: Pilih nama yang sama -> Klik Hadir lagi.
   - Harus muncul error "Sudah absen hari ini".
5. **Cek Laporan**: Klik link Laporan.
   - Coba filter tanggal.
   - Download Excel.

---

## Troubleshooting Error Umum

- **Error 500 pada API**:
  - Biasanya karena lupa Binding D1. Cek Langkah 3.
  - Variable name salah (harus `DB`).
  - Schema database belum dibuat (Langkah 1).

- **Muncul "Error: Date parameter required" di Console**:
  - API dipanggil tanpa parameter. Pastikan JS mengirim tanggal dengan benar.

- **Jam Absensi Salah**:
  - Server Cloudflare menggunakan UTC. Kode kita sudah mengonversi display ke WIB, tapi data mentah di database tersimpan dengan tanggal UTC/server. Pastikan `date` yang dikirim dari Frontend sudah benar (YYYY-MM-DD sesuai WIB), yang sudah dihandle oleh `getLocalDate()` di script kita.

---

## Cara Update Data Guru

Untuk menambah/menghapus guru:
1. Masuk Dashboard Cloudflare > D1 > `db-absensi-mi` > Console.
2. Jalankan SQL:
   ```sql
   INSERT INTO users (name) VALUES ('Nama Baru');
   -- atau --
   DELETE FROM users WHERE name = 'Nama Lama';
   ```
3. Refresh aplikasi web, data otomatis terupdate.

---

Selamat! Aplikasi Absensi MI ISLAMADINA kini siap digunakan.
