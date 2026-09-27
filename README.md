# 📊 Aplikasi Absensi Karyawan & WhatsApp Reminder

Aplikasi berbasis web interaktif yang dibangun menggunakan **Streamlit** untuk mempermudah proses pengecekan kehadiran karyawan/peserta magang (MagangHub) secara otomatis berdasarkan pencocokan nama dan pembersihan gelar akademik. Dilengkapi juga dengan menu konversi data Excel serta fitur pengingat otomatis via **WhatsApp Gateway (Fonnte)**.

✨ Fitur Utama
🏠 Beranda (Cek Absen):

Upload file Master Karyawan dan Rekap Absen (.xlsx / .xls).

Fitur pembersih gelar akademik otomatis (S.Kom, M.Kom, S.E, Dr, Ir, dll.) dan pemangkasan tanda baca/spasi berlebih.

Pencocokan nama cerdas berbasis kata (partial words matching).

Filter berdasarkan Unit / Departemen secara dinamis.

Ekspor hasil karyawan yang belum absen langsung ke format Excel (.xlsx).

🔄 Convert Data:

Merapikan format angka NIP (menghapus tanda petik atau spasi tersembunyi).

Membersihkan gelar pada kolom nama secara massal.

Fitur pemilihan kolom fleksibel sebelum file diexport ulang.

⏰ WhatsApp Reminder:

Konfigurasi token API WhatsApp Gateway (Fonnte) dan nomor/grup tujuan secara interaktif.

Penyimpanan konfigurasi secara permanen menggunakan format lokal JSON (reminder_config.json).

Manajemen daftar konfigurasi (tambah, lihat detail, hapus).

Fitur Test Kirim Pesan manual langsung dari dashboard Streamlit.

👨‍💻 Author Info:

Informasi profil pengembang lengkap dengan tautan kontak dan media sosial.
