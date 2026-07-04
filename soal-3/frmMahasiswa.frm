VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMahasiswa 
   Caption         =   "Sistem Manajemen Data Mahasiswa"
   ClientHeight    =   9345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   8160
      Width           =   1215
   End
   Begin VB.ComboBox cboExport 
      Height          =   315
      Left            =   7560
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   7680
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpTglLahir 
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   212074497
      CurrentDate     =   46206
   End
   Begin VB.ComboBox cboJenjang 
      Height          =   315
      Left            =   1200
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "Batal "
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   20
      Top             =   6000
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdMahasiswa 
      Height          =   5415
      Left            =   6120
      TabIndex        =   19
      Top             =   1200
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   375
      Index           =   3
      Left            =   8760
      TabIndex        =   18
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   17
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton btnUpdate 
      Caption         =   "Update"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   16
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Index           =   0
      Left            =   10440
      TabIndex        =   14
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtAlamat 
      Height          =   1095
      Index           =   4
      Left            =   1200
      TabIndex        =   13
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox txtFakultas 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtJurusan 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtUmur 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1290
      Width           =   975
   End
   Begin VB.TextBox txtNim 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtNama 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   930
      Width           =   2175
   End
   Begin VB.Label lblExport 
      Caption         =   "Export data ke :"
      Height          =   375
      Left            =   6240
      TabIndex        =   25
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lblFakultas 
      Caption         =   "Cari :"
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   23
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label judul 
      Caption         =   "SITEM MANAGEMEN DATA MAHASISWA"
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblAlamat 
      Caption         =   "Alamat : "
      Height          =   975
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblJenjang 
      Caption         =   "Jenjang : "
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label lblFakultas 
      Caption         =   "Fakultas : "
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblJurusan 
      Caption         =   "Jurusan : "
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label lblTglLahir 
      Caption         =   "Tanggal Lahir : "
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblUmur 
      Caption         =   "Umur : "
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblNim 
      Caption         =   "NIM : "
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblNama 
      Caption         =   "Nama : "
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmMahasiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' SISTEM MANAJEMEN DATA MAHASISWA
' + EXPORT: Excel, CSV, JSON, PDF
' ============================================================

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim selectedId As Integer
Dim isEditMode As Boolean

' ============================================================
' FORM LOAD
' ============================================================

Private Sub Form_Load()
    ConnectDB
    IsiComboJenjang
    IsiComboExport
    LoadData
    SetModeInsert
End Sub

' ============================================================
' ISI COMBOBOX JENJANG
' ============================================================

Private Sub IsiComboJenjang()
    With cboJenjang
        .Clear
        .AddItem "D1"
        .AddItem "D2"
        .AddItem "D3"
        .AddItem "S1"
        .AddItem "S2"
        .AddItem "S3"
        .ListIndex = 3
    End With
End Sub

' ============================================================
' ISI COMBOBOX EXPORT
' ============================================================

Private Sub IsiComboExport()
    With cboExport
        .Clear
        .AddItem "Excel (.xls)"
        .AddItem "CSV (.csv)"
        .AddItem "JSON (.json)"
        .AddItem "PDF (.pdf)"
        .ListIndex = 0
    End With
End Sub

' ============================================================
' KONEKSI DATABASE
' ============================================================

Private Sub ConnectDB()
    On Error GoTo ErrorHandler
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={PostgreSQL Unicode};" & _
                            "Server=localhost;" & _
                            "Port=5432;" & _
                            "Database=mahasiswa_db;" & _
                            "Uid=postgres;" & _
                            "Pwd=password;"
    conn.Open
    
    If conn.State = 1 Then
        MsgBox "Koneksi database berhasil!", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Gagal koneksi: " & Err.Description, vbCritical
End Sub

' ============================================================
' LOAD DATA KE GRID
' ============================================================

Private Sub LoadData(Optional searchKeyword As String = "")
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT m.id, m.nama, m.nim, m.umur, m.tgl_lahir, m.alamat, " & _
          "j.nama_jurusan, j.fakultas, j.jenjang " & _
          "FROM mahasiswa m " & _
          "LEFT JOIN jurusan j ON m.id_jurusan = j.id_jurusan "
    
    If Trim(searchKeyword) <> "" Then
        sql = sql & "WHERE LOWER(m.nama) LIKE LOWER('%" & EscapeSQL(searchKeyword) & "%') " & _
              "OR LOWER(m.nim) LIKE LOWER('%" & EscapeSQL(searchKeyword) & "%') " & _
              "OR LOWER(m.alamat) LIKE LOWER('%" & EscapeSQL(searchKeyword) & "%') " & _
              "OR LOWER(j.nama_jurusan) LIKE LOWER('%" & EscapeSQL(searchKeyword) & "%') "
    End If
    
    sql = sql & "ORDER BY m.id"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Setup Grid Header
    With grdMahasiswa
        .Clear
        .Rows = 2
        .Cols = 9
        
        .Row = 0
        .Col = 0: .Text = "ID": .ColWidth(0) = 500
        .Col = 1: .Text = "Nama": .ColWidth(1) = 1800
        .Col = 2: .Text = "NIM": .ColWidth(2) = 1300
        .Col = 3: .Text = "Umur": .ColWidth(3) = 600
        .Col = 4: .Text = "Tgl Lahir": .ColWidth(4) = 1100
        .Col = 5: .Text = "Alamat": .ColWidth(5) = 1500
        .Col = 6: .Text = "Jurusan": .ColWidth(6) = 1300
        .Col = 7: .Text = "Fakultas": .ColWidth(7) = 800
        .Col = 8: .Text = "Jenjang": .ColWidth(8) = 700
        
        Dim i As Integer
        For i = 0 To 8
            .Col = i
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
        Next
    End With
    
    ' Isi Data
    Dim dataRow As Integer
    dataRow = 1
    
    Do While Not rs.EOF
        With grdMahasiswa
            .Row = dataRow
            
            .Col = 0: .Text = rs("id"): .CellAlignment = flexAlignCenterCenter
            .Col = 1: .Text = rs("nama")
            .Col = 2: .Text = rs("nim")
            .Col = 3: .Text = rs("umur"): .CellAlignment = flexAlignCenterCenter
            .Col = 4: .Text = Format(rs("tgl_lahir"), "dd/MM/yyyy")
            .Col = 5: .Text = IIf(IsNull(rs("alamat")), "", rs("alamat"))
            .Col = 6: .Text = IIf(IsNull(rs("nama_jurusan")), "-", rs("nama_jurusan"))
            .Col = 7: .Text = IIf(IsNull(rs("fakultas")), "-", rs("fakultas"))
            .Col = 8: .Text = IIf(IsNull(rs("jenjang")), "-", rs("jenjang"))
        End With
        
        dataRow = dataRow + 1
        rs.MoveNext
        
        If Not rs.EOF Then
            grdMahasiswa.Rows = grdMahasiswa.Rows + 1
        End If
    Loop
    
    rs.Close
    
    ' Hapus row kosong terakhir
    If grdMahasiswa.Rows > 1 Then
        grdMahasiswa.Row = grdMahasiswa.Rows - 1
        grdMahasiswa.Col = 0
        If Trim(grdMahasiswa.Text) = "" Then
            grdMahasiswa.Rows = grdMahasiswa.Rows - 1
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error load data: " & Err.Description, vbCritical
End Sub

' ============================================================
' SEARCH
' ============================================================

Private Sub txtSearch_Change(Index As Integer)
    LoadData txtSearch(Index).Text
End Sub

' ============================================================
' MODE INSERT
' ============================================================

Private Sub SetModeInsert()
    isEditMode = False
    selectedId = 0
    
    btnSave(0).Enabled = True
    btnUpdate(1).Enabled = False
    btnDelete(2).Enabled = False
    btnBatal(1).Enabled = False
    
    ClearForm
End Sub

' ============================================================
' MODE EDIT
' ============================================================

Private Sub SetModeEdit()
    isEditMode = True
    
    btnSave(0).Enabled = False
    btnUpdate(1).Enabled = True
    btnDelete(2).Enabled = True
    btnBatal(1).Enabled = True
End Sub

' ============================================================
' KLIK ROW GRID
' ============================================================

Private Sub grdMahasiswa_Click()
    If grdMahasiswa.Row = 0 Then Exit Sub
    If grdMahasiswa.Row >= grdMahasiswa.Rows Then Exit Sub
    
    With grdMahasiswa
        .Row = .Row
        
        .Col = 0: selectedId = Val(.Text)
        .Col = 1: txtNama(0).Text = .Text
        .Col = 2: txtNim(1).Text = .Text
        .Col = 3: txtUmur.Text = .Text
        .Col = 4: dtpTglLahir.Value = .Text
        .Col = 5: txtAlamat(4).Text = .Text
        .Col = 6: txtJurusan(0).Text = .Text
        .Col = 7: txtFakultas(2).Text = .Text
        
        ' ComboBox Jenjang
        .Col = 8
        Dim i As Integer
        For i = 0 To cboJenjang.ListCount - 1
            If cboJenjang.List(i) = .Text Then
                cboJenjang.ListIndex = i
                Exit For
            End If
        Next
    End With
    
    SetModeEdit
End Sub

' ============================================================
' CLEAR FORM
' ============================================================

Private Sub ClearForm()
    selectedId = 0
    txtNama(0).Text = ""
    txtNim(1).Text = ""
    txtUmur.Text = ""
    dtpTglLahir.Value = Date
    txtJurusan(0).Text = ""
    txtFakultas(2).Text = ""
    cboJenjang.ListIndex = 3
    txtAlamat(4).Text = ""
End Sub

' ============================================================
' TOMBOL SAVE
' ============================================================

Private Sub btnSave_Click(Index As Integer)
    If isEditMode Then
        MsgBox "Sedang mode edit! Klik 'Batal' untuk input baru.", vbExclamation
        Exit Sub
    End If
    
    If Not ValidateForm Then Exit Sub
    
    Dim cekRs As New ADODB.Recordset
    cekRs.Open "SELECT COUNT(*) FROM mahasiswa WHERE nim = '" & EscapeSQL(txtNim(1).Text) & "'", conn
    
    If cekRs(0) > 0 Then
        MsgBox "NIM sudah terdaftar!", vbExclamation
        cekRs.Close: Set cekRs = Nothing
        Exit Sub
    End If
    cekRs.Close: Set cekRs = Nothing
    
    Dim jurusanId As Integer
    jurusanId = GetOrCreateJurusan()
    
    Dim sql As String
    sql = "INSERT INTO mahasiswa (nama, nim, umur, tgl_lahir, alamat, id_jurusan) " & _
          "VALUES ('" & EscapeSQL(txtNama(0).Text) & "', '" & EscapeSQL(txtNim(1).Text) & "', " & _
          Val(txtUmur.Text) & ", '" & Format(dtpTglLahir.Value, "yyyy-MM-dd") & "', " & _
          "'" & EscapeSQL(txtAlamat(4).Text) & "', " & jurusanId & ")"
    
    On Error GoTo ErrorHandler
    conn.Execute sql
    
    MsgBox "Data berhasil disimpan!", vbInformation
    LoadData txtSearch(0).Text
    ClearForm
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================================
' TOMBOL UPDATE
' ============================================================

Private Sub btnUpdate_Click(Index As Integer)
    If Not isEditMode Or selectedId = 0 Then
        MsgBox "Pilih data dulu di grid!", vbExclamation
        Exit Sub
    End If
    
    If Not ValidateForm Then Exit Sub
    
    Dim jurusanId As Integer
    jurusanId = GetOrCreateJurusan()
    
    Dim sql As String
    sql = "UPDATE mahasiswa SET " & _
          "nama = '" & EscapeSQL(txtNama(0).Text) & "', " & _
          "nim = '" & EscapeSQL(txtNim(1).Text) & "', " & _
          "umur = " & Val(txtUmur.Text) & ", " & _
          "tgl_lahir = '" & Format(dtpTglLahir.Value, "yyyy-MM-dd") & "', " & _
          "alamat = '" & EscapeSQL(txtAlamat(4).Text) & "', " & _
          "id_jurusan = " & jurusanId & _
          " WHERE id = " & selectedId
    
    On Error GoTo ErrorHandler
    conn.Execute sql
    
    MsgBox "Data berhasil diupdate!", vbInformation
    LoadData txtSearch(0).Text
    SetModeInsert
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================================
' TOMBOL DELETE
' ============================================================

Private Sub btnDelete_Click(Index As Integer)
    If Not isEditMode Or selectedId = 0 Then
        MsgBox "Pilih data dulu di grid!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Yakin hapus " & txtNama(0).Text & " (" & txtNim(1).Text & ")?", _
              vbQuestion + vbYesNo, "Konfirmasi Hapus") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    conn.Execute "DELETE FROM mahasiswa WHERE id = " & selectedId
    
    MsgBox "Data berhasil dihapus!", vbInformation
    LoadData txtSearch(0).Text
    SetModeInsert
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================================
' TOMBOL RESET DB
' ============================================================

Private Sub btnReset_Click(Index As Integer)
    
    If MsgBox("PERINGATAN!" & vbCrLf & _
              "Ini akan MENGHAPUS SEMUA DATA mahasiswa!" & vbCrLf & _
              "Tindakan ini TIDAK BISA dibatalkan!" & vbCrLf & vbCrLf & _
              "Yakin lanjutkan?", vbCritical + vbYesNo, "RESET DATABASE") = vbNo Then
        Exit Sub
    End If
    
    If MsgBox("SERIUS? Semua data akan hilang selamanya!" & vbCrLf & _
              "Klik Yes untuk menghapus SEMUA data.", vbCritical + vbYesNo, "KONFIRMASI TERAKHIR") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    conn.Execute "DELETE FROM mahasiswa"
    
    MsgBox "Semua data berhasil dihapus!" & vbCrLf & "Database sekarang kosong.", vbInformation
    
    LoadData
    SetModeInsert
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================================
' TOMBOL BATAL
' ============================================================

Private Sub btnBatal_Click(Index As Integer)
    If Not isEditMode Then
        MsgBox "Tidak sedang mode edit!", vbExclamation
        Exit Sub
    End If
    
    SetModeInsert
End Sub

' ============================================================
' ============================================================
' EXPORT DATA
' ============================================================
' ============================================================

Private Sub btnExport_Click()
    
    Dim formatPilihan As String
    formatPilihan = cboExport.Text
    
    Select Case formatPilihan
        Case "Excel (.xls)"
            ExportToExcel
        Case "CSV (.csv)"
            ExportToCSV
        Case "JSON (.json)"
            ExportToJSON
        Case "PDF (.pdf)"
            ExportToPDF
        Case Else
            MsgBox "Pilih format export!", vbExclamation
    End Select
    
End Sub

' ============================================================
' EXPORT KE EXCEL
' ============================================================

Private Sub ExportToExcel()
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer, j As Integer
    Dim savePath As String
    
    ' Dialog simpan file
    savePath = GetSavePath("Excel Files (*.xls)|*.xls", "DataMahasiswa.xls")
    If savePath = "" Then Exit Sub
    
    ' Buat Excel object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    xlApp.Visible = False
    
    ' Tulis Header
    For j = 0 To grdMahasiswa.Cols - 1
        grdMahasiswa.Row = 0
        grdMahasiswa.Col = j
        xlSheet.Cells(1, j + 1).Value = grdMahasiswa.Text
        xlSheet.Cells(1, j + 1).Font.Bold = True
        xlSheet.Cells(1, j + 1).Interior.Color = RGB(200, 200, 200)
    Next j
    
    ' Tulis Data
    For i = 1 To grdMahasiswa.Rows - 1
        For j = 0 To grdMahasiswa.Cols - 1
            grdMahasiswa.Row = i
            grdMahasiswa.Col = j
            xlSheet.Cells(i + 1, j + 1).Value = grdMahasiswa.Text
        Next j
    Next i
    
    ' AutoFit kolom
    xlSheet.Columns.AutoFit
    
    ' Simpan
    xlBook.SaveAs savePath
    xlBook.Close
    xlApp.Quit
    
    ' Bersihkan memory
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    MsgBox "Data berhasil diexport ke Excel!" & vbCrLf & savePath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error export Excel: " & Err.Description, vbCritical
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

' ============================================================
' EXPORT KE CSV
' ============================================================

Private Sub ExportToCSV()
    On Error GoTo ErrorHandler
    
    Dim fs As Object
    Dim ts As Object
    Dim i As Integer, j As Integer
    Dim savePath As String
    Dim line As String
    
    savePath = GetSavePath("CSV Files (*.csv)|*.csv", "DataMahasiswa.csv")
    If savePath = "" Then Exit Sub
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(savePath, True)
    
    ' Tulis Header
    line = ""
    For j = 0 To grdMahasiswa.Cols - 1
        grdMahasiswa.Row = 0
        grdMahasiswa.Col = j
        line = line & """" & grdMahasiswa.Text & """"
        If j < grdMahasiswa.Cols - 1 Then line = line & ","
    Next j
    ts.WriteLine line
    
    ' Tulis Data
    For i = 1 To grdMahasiswa.Rows - 1
        line = ""
        For j = 0 To grdMahasiswa.Cols - 1
            grdMahasiswa.Row = i
            grdMahasiswa.Col = j
            line = line & """" & grdMahasiswa.Text & """"
            If j < grdMahasiswa.Cols - 1 Then line = line & ","
        Next j
        ts.WriteLine line
    Next i
    
    ts.Close
    Set ts = Nothing
    Set fs = Nothing
    
    MsgBox "Data berhasil diexport ke CSV!" & vbCrLf & savePath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error export CSV: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close
End Sub

' ============================================================
' EXPORT KE JSON
' ============================================================

Private Sub ExportToJSON()
    On Error GoTo ErrorHandler
    
    Dim fs As Object
    Dim ts As Object
    Dim i As Integer, j As Integer
    Dim savePath As String
    Dim json As String
    Dim headers(8) As String
    
    savePath = GetSavePath("JSON Files (*.json)|*.json", "DataMahasiswa.json")
    If savePath = "" Then Exit Sub
    
    ' Ambil header
    For j = 0 To grdMahasiswa.Cols - 1
        grdMahasiswa.Row = 0
        grdMahasiswa.Col = j
        headers(j) = grdMahasiswa.Text
    Next j
    
    ' Build JSON
    json = "[" & vbCrLf
    
    For i = 1 To grdMahasiswa.Rows - 1
        json = json & "  {" & vbCrLf
        
        For j = 0 To grdMahasiswa.Cols - 1
            grdMahasiswa.Row = i
            grdMahasiswa.Col = j
            
            json = json & "    """ & headers(j) & """: """ & EscapeJSON(grdMahasiswa.Text) & """"
            If j < grdMahasiswa.Cols - 1 Then json = json & ","
            json = json & vbCrLf
        Next j
        
        json = json & "  }"
        If i < grdMahasiswa.Rows - 1 Then json = json & ","
        json = json & vbCrLf
    Next i
    
    json = json & "]"
    
    ' Simpan file
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(savePath, True)
    ts.Write json
    ts.Close
    
    Set ts = Nothing
    Set fs = Nothing
    
    MsgBox "Data berhasil diexport ke JSON!" & vbCrLf & savePath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error export JSON: " & Err.Description, vbCritical
    If Not ts Is Nothing Then ts.Close
End Sub

' ============================================================
' EXPORT KE PDF
' ============================================================

Private Sub ExportToPDF()
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer, j As Integer
    Dim savePath As String
    
    savePath = GetSavePath("PDF Files (*.pdf)|*.pdf", "DataMahasiswa.pdf")
    If savePath = "" Then Exit Sub
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    xlApp.Visible = False
    
    ' Tulis Header
    For j = 0 To grdMahasiswa.Cols - 1
        grdMahasiswa.Row = 0
        grdMahasiswa.Col = j
        xlSheet.Cells(1, j + 1).Value = grdMahasiswa.Text
        xlSheet.Cells(1, j + 1).Font.Bold = True
        xlSheet.Cells(1, j + 1).Interior.Color = RGB(0, 51, 102)
        xlSheet.Cells(1, j + 1).Font.Color = RGB(255, 255, 255)
    Next j
    
    ' Tulis Data
    For i = 1 To grdMahasiswa.Rows - 1
        For j = 0 To grdMahasiswa.Cols - 1
            grdMahasiswa.Row = i
            grdMahasiswa.Col = j
            xlSheet.Cells(i + 1, j + 1).Value = grdMahasiswa.Text
        Next j
    Next i
    
    ' Rapikan tampilan sebelum export
    xlSheet.Columns.AutoFit
    xlSheet.PageSetup.Orientation = 2   ' 2 = xlLandscape
    xlSheet.PageSetup.FitToPagesWide = 1
    xlSheet.PageSetup.FitToPagesTall = False
    
    ' Export langsung ke PDF (Type:=0 -> xlTypePDF)
    xlBook.ExportAsFixedFormat Type:=0, FileName:=savePath, _
        Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    xlBook.Close False
    xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    MsgBox "Data berhasil diexport langsung ke PDF!" & vbCrLf & savePath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error export PDF: " & Err.Description, vbCritical
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

' ============================================================
' FUNGSI BANTUAN: DIALOG SIMPAN FILE
' ============================================================

Private Function GetSavePath(filter As String, defaultName As String) As String
    Dim cd As Object
    Set cd = CreateObject("MSComDlg.CommonDialog")
    
    cd.DialogTitle = "Simpan File Export"
    cd.filter = filter
    cd.FileName = defaultName
    cd.Flags = &H2 Or &H80000
    
    On Error Resume Next
    cd.ShowSave
    
    If Err.Number = 32755 Then
        GetSavePath = ""
    Else
        GetSavePath = cd.FileName
    End If
    
    On Error GoTo 0
    Set cd = Nothing
End Function

' ============================================================
' FUNGSI BANTUAN: ESCAPE JSON
' ============================================================

Private Function EscapeJSON(str As String) As String
    str = Replace(str, "\", "\\")
    str = Replace(str, """", "\""")
    str = Replace(str, vbCr, "\r")
    str = Replace(str, vbLf, "\n")
    str = Replace(str, vbTab, "\t")
    EscapeJSON = str
End Function

' ============================================================
' VALIDASI
' ============================================================

Private Function ValidateForm() As Boolean
    ValidateForm = False
    
    If Trim(txtNama(0).Text) = "" Then
        MsgBox "Nama tidak boleh kosong!", vbExclamation
        txtNama(0).SetFocus: Exit Function
    End If
    
    If Trim(txtNim(1).Text) = "" Then
        MsgBox "NIM tidak boleh kosong!", vbExclamation
        txtNim(1).SetFocus: Exit Function
    End If
    
    If Not IsNumeric(txtUmur.Text) Or Val(txtUmur.Text) <= 0 Then
        MsgBox "Umur harus angka positif!", vbExclamation
        txtUmur.SetFocus: Exit Function
    End If
    
    If Trim(txtJurusan(0).Text) = "" Then
        MsgBox "Jurusan tidak boleh kosong!", vbExclamation
        txtJurusan(0).SetFocus: Exit Function
    End If
    
    If cboJenjang.ListIndex = -1 Then
        MsgBox "Jenjang harus dipilih!", vbExclamation
        cboJenjang.SetFocus: Exit Function
    End If
    
    ValidateForm = True
End Function

' ============================================================
' GET OR CREATE JURUSAN
' ============================================================

Private Function GetOrCreateJurusan() As Integer
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT id_jurusan FROM jurusan " & _
          "WHERE nama_jurusan = '" & EscapeSQL(txtJurusan(0).Text) & "' " & _
          "AND fakultas = '" & EscapeSQL(txtFakultas(2).Text) & "' " & _
          "AND jenjang = '" & EscapeSQL(cboJenjang.Text) & "'"
    
    Dim rsJurusan As New ADODB.Recordset
    rsJurusan.Open sql, conn
    
    If Not rsJurusan.EOF Then
        GetOrCreateJurusan = rsJurusan("id_jurusan")
    Else
        rsJurusan.Close
        conn.Execute "INSERT INTO jurusan (nama_jurusan, fakultas, jenjang) " & _
                     "VALUES ('" & EscapeSQL(txtJurusan(0).Text) & "', " & _
                     "'" & EscapeSQL(txtFakultas(2).Text) & "', " & _
                     "'" & EscapeSQL(cboJenjang.Text) & "')"
        
        rsJurusan.Open "SELECT currval('jurusan_id_jurusan_seq')", conn
        GetOrCreateJurusan = rsJurusan(0)
    End If
    
    rsJurusan.Close
    Set rsJurusan = Nothing
    Exit Function
    
ErrorHandler:
    GetOrCreateJurusan = 0
End Function

' ============================================================
' FUNGSI BANTUAN
' ============================================================

Private Function EscapeSQL(str As String) As String
    EscapeSQL = Replace(str, "'", "''")
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
End Sub

