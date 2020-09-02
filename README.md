# OdooModuleWebScraper
Scrap Odoo App (Module) Store data to Excel File

## Spesifikasi umum:
1. Dibuat dengan bahasa Java
2. Disusun sebagai Maven Project
3. Menggunakan dependensi Jsoup
4. Scrapping dari apps.odoo.com

## Cara menggunakan Odoo Web Crawler:
1. Clone atau Download repositori ini https://github.com/Nael-Nathanael/OdooModuleWebScraper/archive/master.zip
2. Extract folder tersebut
3. Buka file src/main/java/OdooWebCrawler.java
4. Sesuaikan baris - baris berikut sesuai kebutuhan:
   - baris 19, ubah "Master-0-1" menjadi nama sheet yang akan dihasilkan
   - baris 24, ubah angka 1 menjadi halaman pertama dan 10 menjadi halaman terakhir dari apps.odoo.com yang akan di-crawl
   - baris 26, ubah filter price (Paid/Free) dan series (13.0/12.0/11.0/10.0/...) sesuai kebutuhan
   - baris 217 adalah fungsi untuk membuat setiap judul kolom, dapat disesuaikan dengan kebutuhan
   - baris 98 adalah fungsi untuk mengisi baris, anda dapat menyesuaikan baris 133 dengan NPM, 137 dengan nama anda, dan baris lainnya sesuai kebutuhan
5. Build and Run
   - (Disarankan) jika anda memiliki Java IDE, import folder sebagai Maven Project dan Jalankan 
   - jika tidak, anda dapat menginstall maven terlebih dahulu dan ikuti panduan build and run Maven Project 

## Beberapa referensi panduan build and run Maven Project
- https://www.vogella.com/tutorials/ApacheMaven/article.html

## Disclaimer:
- Aplikasi ini saya dibuat untuk mempelajari Web Scraping sederhana dan mempermudah penyusunan tugas Konfigurasi ERP sehingga belum dibuatkan UI untuk interaksi dengan user secara langsung
- Keberlangsungan pengembangan aplikasi OdooModuleWebScraper tidak pasti
- Anda diperbolehkan untuk menyalin sebagian atau keseluruhan aplikasi ini, untuk tujuan komersil atau tujuan pribadi dengan bebas dan tanpa biaya apapun 
