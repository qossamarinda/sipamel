# Panduan Pemasangan Cloudflare Worker SiPaMeL

Script ini membuat link tetap gratis (*persistent URL*) untuk aplikasi SiPaMeL menggunakan Cloudflare Worker.

Link hasil akhir: `https://sipamel.<nama-akun>.workers.dev`

## Langkah Pemasangan:
1. Buka [https://dash.cloudflare.com/](https://dash.cloudflare.com/) dan login (atau daftar gratis jika belum punya akun).
2. Di panel menu kiri, klik **Workers & Pages**.
3. Klik tombol **Create** / **Create Application** lalu pilih tab **Workers** -> klik **Create Worker**.
4. Beri nama worker: `sipamel` (atau nama lain yang Anda inginkan), lalu klik **Deploy**.
5. Setelah berhasil dideploy, klik tombol **Edit code**.
6. Salin seluruh isi berkas `worker.js` dan tempelkan ke editor kode di Cloudflare (menggantikan kode bawaan).
7. Klik tombol **Save and Deploy** di pojok kanan atas.
8. Selesai! Link worker Anda yang berakhiran `.workers.dev` sudah aktif dan siap digunakan secara permanen.
