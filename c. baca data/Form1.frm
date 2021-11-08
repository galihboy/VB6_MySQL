VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tes Koneksi Database - Galih Hermawan (Nov 2021)"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Baca Data"
      Height          =   4335
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   9975
      Begin VB.TextBox txtHasilCari 
         Height          =   3975
         Left            =   3240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   26
         Top             =   240
         Width           =   6615
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "Cari"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtCari 
         Height          =   405
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label lblJmlData 
         Caption         =   "Jumlah data."
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Teks yang dicari : "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblSumberData 
         Caption         =   "Sumber data : "
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Akses Database"
      Height          =   2895
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   7575
      Begin VB.ListBox lstAtribut 
         Height          =   1230
         Left            =   5280
         MultiSelect     =   2  'Extended
         TabIndex        =   18
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmdListAtribut 
         Caption         =   "List Atribut"
         Height          =   375
         Left            =   5280
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.ListBox lstTabel 
         Height          =   1230
         Left            =   2760
         TabIndex        =   15
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdListTabel 
         Caption         =   "List Tabel"
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.ListBox lstDatabase 
         Height          =   1230
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdListDatabase 
         Caption         =   "List Database"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblJumlahAtribut 
         Caption         =   "Jumlah Atribut"
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label lblJmlTabel 
         Caption         =   "Jumlah tabel."
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lblJumlahDatabase 
         Caption         =   "Jumlah database."
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdKoneksi 
      Caption         =   "Cek Koneksi"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameter Koneksi"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "root"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "3306"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "localhost"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "User"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Port"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Host"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Koneksi VB6 & MySQL/MariaDB
' Developed by Galih Hermawan @ Sat, 6 Nov 2021
' Last update: 8/11/2021
' https://galih.eu
' https://galihboy.github.io
' https://masgalih.medium.com

Dim koneksi As MYSQL_CONNECTION
Dim MyRS As MYSQL_RS
Dim strDB As String, strTabel As String

Private Sub cmdCari_Click()
    Dim arrAtribut() As String, strAtributSeleksi As String
    Dim i As Integer, teks As String
    
    arrAtribut = SeleksiAtribut
    strAtributSeleksi = vbNullString

    If UkuranArray(arrAtribut) = 0 Then
        MsgBox "Silakan pilih minimal 1 atribut."
    Else
        For i = LBound(arrAtribut) To UBound(arrAtribut) - 1
            If i = UBound(arrAtribut) - 1 Then
                strAtributSeleksi = strAtributSeleksi + arrAtribut(i)
            Else
                strAtributSeleksi = strAtributSeleksi + arrAtribut(i) + ", "
                
            End If
        Next

        Call CariData(strAtributSeleksi)
    End If
End Sub

Private Sub cmdKoneksi_Click()
    Dim NamaHost As String, NoPort As Integer, NamaUser As String, Password As String
    Dim AdaKoneksi As Boolean
    
    AdaKoneksi = TesKoneksi
    If AdaKoneksi Then
        MsgBox "Koneksi sukses!", vbInformation, "Sukes!"
    End If
    
End Sub

Private Sub cmdListDatabase_Click()
    On Error GoTo salah
    Dim AdaKoneksi As Boolean
    
    lstDatabase.Clear
    AdaKoneksi = TesKoneksi
    
    If AdaKoneksi Then
        Dim theTemp As String, tempArray() As String
        Set MyRS = New MYSQL_RS
        Set MyRS = koneksi.Show(MY_SHOW_DATABASES)

        If Not MyRS.EOF Then
            theTemp = Trim$(MyRS.GetString(, ""))
            If theTemp <> "" Then
                Dim i As Integer
                tempArray = Split(theTemp, vbCrLf)
                For i = LBound(tempArray) To UBound(tempArray) - 1
                    lstDatabase.AddItem tempArray(i)
                Next
            End If
        End If
        
        MyRS.CloseRecordset
        Set MyRS = Nothing
        lblJumlahDatabase = "Terdapat" + Str(lstDatabase.ListCount) + " database."
    End If
    Exit Sub
salah:
    MsgBox "Query gagal." & vbNewLine & vbNewLine & _
           "Nomor: " & koneksi.Error.Number & vbNewLine & _
           "Keterangan: " & TulisanError(koneksi.Error.Number), vbExclamation, "Ada kesalahan!"
End Sub

Private Sub cmdListTabel_Click()
    On Error GoTo salah
    Dim AdaKoneksi As Boolean
    Dim namaDatabase As String
    namaDatabase = vbNullString
    
    If lstDatabase.ListCount > 0 Then
        namaDatabase = lstDatabase.Text
    End If
    AdaKoneksi = TesKoneksi(namaDatabase)
    
    If AdaKoneksi Then
        Dim theTemp As String, tempArray() As String
        Dim namaTabel As String
        
        lstTabel.Clear
        Set MyRS = New MYSQL_RS
        Set MyRS = koneksi.Show(MY_SHOW_TABLES)

        If Not MyRS.EOF Then
            theTemp = Trim$(MyRS.GetString(, ""))
            If theTemp <> "" Then
                Dim i As Integer
                tempArray = Split(theTemp, vbCrLf)
                For i = LBound(tempArray) To UBound(tempArray) - 1
                    lstTabel.AddItem tempArray(i)
                Next
            End If
        End If
        
        MyRS.CloseRecordset
        Set MyRS = Nothing

        lblJmlTabel = "Terdapat" + Str(lstTabel.ListCount) + " tabel."
    End If
    Exit Sub
salah:
    MsgBox "Query gagal." & vbNewLine & vbNewLine & _
           "Nomor: " & koneksi.Error.Number & vbNewLine & _
           "Keterangan: " & TulisanError(koneksi.Error.Number), vbExclamation, "Ada kesalahan!"
End Sub

Private Sub cmdListAtribut_Click()
    On Error GoTo salah
    Dim AdaKoneksi As Boolean
    Dim namaDatabase As String
    namaDatabase = vbNullString
    
    If lstDatabase.ListCount > 0 Then
        namaDatabase = lstDatabase.Text
    End If
    AdaKoneksi = TesKoneksi(namaDatabase)
    
    If AdaKoneksi Then
        Dim namaTabel As String
        
        namaTabel = lstTabel.Text
        lstAtribut.Clear
        Set MyRS = New MYSQL_RS
        
        If namaTabel = vbNullString Then
            MsgBox "Tabel belum dipilih"
        Else
            
            Set MyRS = koneksi.Execute("SELECT * FROM " & namaTabel)
            If Not MyRS.EOF Then
                Dim i As Integer
                For i = 0 To MyRS.FieldCount - 1
                    lstAtribut.AddItem MyRS.Fields(i).Name
                Next i
            End If
            
            MyRS.CloseRecordset
            Set MyRS = Nothing

            lblJumlahAtribut = "Terdapat" + Str(lstAtribut.ListCount) + " atribut."
        End If
    End If
    Exit Sub
salah:
    MsgBox "Query gagal." & vbNewLine & vbNewLine & _
           "Nomor: " & koneksi.Error.Number & vbNewLine & _
           "Keterangan: " & TulisanError(koneksi.Error.Number), vbExclamation, "Ada kesalahan!"
End Sub

Private Function TesKoneksi(Optional namaDatabase As String = "") As Boolean
    On Error GoTo salah
    Dim NamaHost As String, NoPort As Integer, NamaUser As String, Password As String
    
    NamaHost = Trim(txtHost)
    NoPort = Trim(Val(txtPort))
    NamaUser = Trim(txtUser)
    Password = Trim(txtPassword)
    
    Set koneksi = New MYSQL_CONNECTION
    koneksi.OpenConnection NamaHost, NamaUser, Password, namaDatabase, NoPort
    
    If koneksi.State = MY_CONN_OPEN Then
        TesKoneksi = True
    End If
    Exit Function
salah:
    TesKoneksi = False
    MsgBox "Koneksi gagal." & vbNewLine & vbNewLine & _
           "Nomor: " & koneksi.Error.Number & vbNewLine & _
           "Keterangan: " & TulisanError(koneksi.Error.Number), vbExclamation, "Ada kesalahan!"
End Function

Private Sub CariData(strAtribut As String)
    On Error GoTo salah
    Dim AdaKoneksi As Boolean
    Dim namaDatabase As String
    Dim strCari As String
    
    strCari = Trim(txtCari)
    namaDatabase = vbNullString
    
    If lstDatabase.ListCount > 0 Then
        namaDatabase = lstDatabase.Text
    End If
    AdaKoneksi = TesKoneksi(namaDatabase)
    
    If AdaKoneksi Then
        Dim theTemp As String, tempArray() As String
        Dim namaTabel As String, strSQL As String, strSQLCari As String
        Dim arrAtribut() As String, jAtribut As Integer
        
        arrAtribut = Split(strAtribut, ",")

        namaTabel = lstTabel.Text
        txtHasilCari = vbNullString
        strSQLCari = vbNullString
        
        For jAtribut = 0 To lstAtribut.ListCount - 1
            If jAtribut = lstAtribut.ListCount - 1 Then
                strSQLCari = strSQLCari & "`" & lstAtribut.List(jAtribut) & "` LIKE '%" & strCari & "%'"
            Else
                strSQLCari = strSQLCari & "`" & lstAtribut.List(jAtribut) & "` LIKE '%" & strCari & "%' OR "
            End If
        Next
        
        strSQL = "SELECT " & strAtribut & " FROM " & namaTabel & _
                 " WHERE " & strSQLCari
        
        'MsgBox strSQL
        Set MyRS = New MYSQL_RS
        Set MyRS = koneksi.Execute(strSQL)

        Dim i As Integer, data As Integer
        data = 1
        
        Do While Not MyRS.EOF
            txtHasilCari = txtHasilCari & "Data ke-" & data & vbCrLf
        
            For i = 0 To MyRS.FieldCount - 1
                txtHasilCari = txtHasilCari + Space(3) + MyRS.Fields(i).Name & ": " & MyRS.Fields(i).Value
                txtHasilCari = txtHasilCari + vbCrLf
            Next
            
            txtHasilCari = txtHasilCari + vbCrLf
            data = data + 1
            MyRS.MoveNext
        Loop
        
        lblJmlData = "Ditemukan " & MyRS.RecordCount & " data."
        MyRS.CloseRecordset
        Set MyRS = Nothing

    End If
    Exit Sub
salah:
    MsgBox "Query gagal." & vbNewLine & vbNewLine & _
           "Nomor: " & koneksi.Error.Number & vbNewLine & _
           "Keterangan: " & TulisanError(koneksi.Error.Number), vbExclamation, "Ada kesalahan!"
End Sub

' Fungsi untuk menangkap nomor/jenis error
Private Function TulisanError(nmrError As Integer) As String
    Select Case nmrError
        Case 1045 'Username/Password salah
            TulisanError = "Login salah." & vbNewLine & "Silakan periksa nama user dan password Anda."
        Case 1046 'no database selected
            TulisanError = "Database belum dipilih." & vbNewLine & "Pilih dulu salah satu database yang ada."
        Case 2003 'Cannot connect - server mati, port salah, dll
            TulisanError = "Tidak bisa terhubung ke mysql server."
        Case 2005 'Server tidak dikenali
            TulisanError = "Nama host (server) tidak dikenal."
        Case Else
            TulisanError = koneksi.Error.Description
    End Select
    
End Function

Private Function SeleksiAtribut() As String()
    Dim k As Integer, i As Integer, isi As String, arrAtribut() As String
    i = 0
    For k = 0 To lstAtribut.ListCount - 1
        If lstAtribut.Selected(k) Then
            ReDim Preserve arrAtribut(i + 1)
            isi = lstAtribut.List(k)
            arrAtribut(i) = isi
            i = i + 1
        End If
    Next
    SeleksiAtribut = arrAtribut
End Function
Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set koneksi = Nothing
End Sub

Private Function UkuranArray(arr As Variant) As Long
  On Error GoTo handler

  Dim lngLower As Long
  Dim lngUpper As Long

  lngLower = LBound(arr)
  lngUpper = UBound(arr)

  UkuranArray = (lngUpper - lngLower) + 1
  Exit Function

handler:
  UkuranArray = 0 'error occured.  must be zero length
End Function

Private Sub lstDatabase_Click()
    cmdListTabel_Click
    strDB = lstDatabase.Text
    lblSumberData = "Sumber data: DB=" & strDB
End Sub

Private Sub lstTabel_Click()
    cmdListAtribut_Click
    strTabel = lstTabel.Text
    lblSumberData = "Sumber data: DB=" & strDB & ", Tabel=" & strTabel
End Sub

Private Sub txtCari_GotFocus()
    Clipboard.Clear
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 32, 8 ' A-Z, 0-9 and backspace
        'Let these key codes pass through
        Case 97 To 122, 32, 8 'a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
            KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub
