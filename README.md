# TECHINICAL-TEST-LAMDA-APK-3
VB6 + PostgreSQL
## Identitas
- **Nama**: Raihan Wendra Baswara
- **Universitas**: Telkom University
- **Program Studi**: S1 INFORMATIKA (2024)

## Deskripsi
Aplikasi desktop legacy menggunakan Visual Basic 6 dengan koneksi langsung 
ke database PostgreSQL. Tidak menggunakan API — query dan business logic 
ada di dalam aplikasi VB6.

## Teknologi yang Digunakan
- **Aplikasi**: Visual Basic 6 (VB6)
- **Database**: PostgreSQL
- **Arsitektur**: Client-Server (direct connection)
- **Koneksi DB**: ADODB / ODBC

## Struktur Project
```bash
soal-3/
├── app-vb6/
│   ├── MahasiswaApp.vbp    
│   ├── Form_Main.frm
│   ├── Form_Mahasiswa.frm
│   └── Module_DB.bas       
└── init.sql       
```
## Cara Menjalankan
1. Pastikan PostgreSQL sudah running
2. Sesuaikan connection string di `Module_DB.bas`
3. Buka `MahasiswaApp.vbp` di VB6 IDE
4. Jalankan dengan F5

## Database
```sql
CREATE DATABASE mahasiswa_db;
```
Lalu jalankan `database/init.sql`
