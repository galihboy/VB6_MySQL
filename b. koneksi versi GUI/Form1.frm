VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tes Koneksi Database - Galih Hermawan (Nov 2021)"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
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
' https://galih.eu
' https://galihboy.github.io
' https://masgalih.medium.com

Dim koneksi As MYSQL_CONNECTION
Dim MyRS As MYSQL_RS

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

            lblJumlahAtribut = "Terdapat" + Str(lstAtribut.ListCount) + " tabel."
        End If
    End If
    Exit Sub
salah:
    MsgBox "Query gagal." & vbNewLine & vbNewLine & _
           "Nomor: " & koneksi.Error.Number & vbNewLine & _
           "Keterangan: " & TulisanError(koneksi.Error.Number), vbExclamation, "Ada kesalahan!"
End Sub

Private Function TesKoneksi(Optional namaDatabase As String = "")
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

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set koneksi = Nothing
End Sub
