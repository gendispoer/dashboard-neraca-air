DASHBOARD NERACA AIR
Dashboard volume dan kehilangan air (Non-Revenue Water)
===========================================================

STRUKTUR FOLDER
---------------
dashboard_neraca_air/
|-- index.html          Halaman Welcome (tombol Start membuka pages/tabel.html)
|-- pages/
|   |-- tabel.html      Halaman Tabel (neraca air)
|   |-- grafik.html     Halaman Grafik
|   `-- detail.html     Halaman Detail (impor file .xlsx dan edit angka)
|-- css/
|   `-- style.css       Seluruh gaya tampilan
|-- js/
|   `-- dashboard.js    Seluruh logika: baca data WB, filter, tabel, grafik, edit, impor
`-- README.txt          Petunjuk ini

CARA MENJALANKAN
----------------
1. Buka index.html di browser (klik dua kali sudah cukup), lalu tekan Start.
2. Perlu koneksi internet, karena Chart.js, pembaca Excel (SheetJS), dan font
   Plus Jakarta Sans / Great Vibes dimuat dari CDN.
3. Pindah halaman lewat menu di bagian bawah layar (Tabel / Grafik / Detail).
   Arahkan kursor ke garis tipis di bawah layar, menu akan muncul.

CARA MEMASUKKAN DATA
--------------------
- Buka halaman Detail, tekan "Impor file .xlsx", lalu pilih file neraca air.
- Nama file harus berformat WB_Bulan_Tahun, contoh: WB_Ags_26.xlsx.
  Sheet yang dibaca bernama "WB".
- Pilih Year dan Month di kanan atas untuk melihat satu bulan atau
  mengakumulasi beberapa bulan.
- Pilihan Year/Month dan file yang diimpor dibawa antar halaman selama tab
  browser masih terbuka (disimpan di sessionStorage). Menutup tab menghapusnya.

PENYIMPANAN GOOGLE DRIVE
------------------------
Tombol "Tautkan Drive" dan penyimpanan otomatis hanya aktif bila dashboard
dibuka di dalam lingkungan Claude (window.claude). Di browser biasa, data
dimasukkan lewat impor .xlsx dan perubahan di halaman Detail berlaku selama
sesi tab itu saja.

CATATAN
-------
- Tampilan tidak diubah dari index.html asli; isinya hanya dipecah ke file
  masing-masing (HTML per halaman, CSS, JS).
- Setiap file di pages/ memanggil ../css/style.css dan ../js/dashboard.js.
  Jika folder dipindah, pertahankan susunan relatif di atas.
