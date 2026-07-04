# TECHINICAL-TEST-LAMDA-APK-3
VB6 + PostgreSQL + Minikube

## Identitas
- **Nama**: Raihan Wendra Baswara
- **Universitas**: Telkom University
- **Program Studi**: S1 INFORMATIKA (2024)

## Deskripsi
Aplikasi desktop legacy menggunakan Visual Basic 6 dengan koneksi langsung ke database PostgreSQL. Tidak menggunakan API — query dan business logic ada di dalam aplikasi VB6.

## Teknologi yang Digunakan
- **Aplikasi**: Visual Basic 6 (VB6)
- **Database**: PostgreSQL 15
- **Arsitektur**: Client-Server (direct connection)
- **Koneksi DB**: ADODB / ODBC (PostgreSQL Unicode Driver)
- **Deployment**: Minikube (Kubernetes)

## Struktur Project
```bash
soal-3/
├── k8s/
│   └── db.yaml
├── frmMahasiswa.frm
├── MahasiswaVB6.vbp
└── MahasiswaVB6.vbw
```

## Cara Menjalankan

### Opsi 1 — Lokal (PostgreSQL langsung)

1. Pastikan PostgreSQL sudah running di `localhost:5432`
2. Buat database `mahasiswa_db` dan jalankan init SQL
3. Install psqlODBC driver (PostgreSQL Unicode)
4. Buka `MahasiswaVB6.vbp` di VB6 IDE
5. Sesuaikan connection string di `frmMahasiswa.frm`:
```vb
conn.ConnectionString = "Driver={PostgreSQL Unicode};" & _
                        "Server=localhost;" & _
                        "Port=5432;" & _
                        "Database=mahasiswa_db;" & _
                        "Uid=postgres;" & _
                        "Pwd=YOUR_PASSWORD;"
```
6. Jalankan dengan F5

---

### Opsi 2 — Deploy Database ke Minikube

#### Prerequisites
- Docker Desktop
- Minikube
- kubectl
- psqlODBC driver terinstall (PostgreSQL Unicode)

#### 1. Start Minikube
```bash
minikube start
```

#### 2. Deploy Database
```bash
cd soal-3
kubectl apply -f k8s/db.yaml
kubectl get pods -w
```

#### 3. Akses Database
```bash
minikube service postgres-soal3-service --url
```
Catat host dan port yang muncul (contoh: `127.0.0.1:63408`). Terminal harus tetap terbuka.

#### 4. Update Connection String di VB6
Ganti connection string di `frmMahasiswa.frm`:
```vb
conn.ConnectionString = "Driver={PostgreSQL Unicode};" & _
                        "Server=127.0.0.1;" & _
                        "Port=63408;" & _
                        "Database=mahasiswa_db;" & _
                        "Uid=postgres;" & _
                        "Pwd=YOUR_PASSWORD;"
```

#### 5. Jalankan Aplikasi
Buka `MahasiswaVB6.vbp` di VB6 IDE, lalu jalankan dengan F5.

---

## Fitur Aplikasi
- CRUD Data Mahasiswa (Create, Read, Update, Delete)
- Search mahasiswa berdasarkan nama, NIM, alamat, dan jurusan
- Auto-create jurusan baru saat input data
- Validasi input form
- Export data ke Excel, CSV, JSON, PDF
- Reset semua data mahasiswa
- Mode Insert dan Edit terpisah

## Aturan Pengembangan yang Diterapkan
- Tidak menggunakan API
- Koneksi langsung ke PostgreSQL via ADODB/ODBC
- Query dan business logic ada di dalam aplikasi VB6
