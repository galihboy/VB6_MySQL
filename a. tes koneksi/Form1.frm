VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Tes Koneksi"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Koneksi VB6 & MySQL/MariaDB
' Developed by Galih Hermawan
' https://galih.eu

Dim koneksi As MYSQL_CONNECTION

Private Sub Command1_Click()
    Set koneksi = New MYSQL_CONNECTION
    Const host = "localhost"
    Const user = "root"
    Const passw = ""
    Const port = 3306
    
    koneksi.OpenConnection host, user, passw, , port
    
    If koneksi.State = MY_CONN_OPEN Then
        Dim status
        Dim isi
        Set status = koneksi.Execute("SELECT version()")
        isi = status.GetString(, "")
        MsgBox "Koneksi sukses! Versi database: " + isi, vbInformation, "Sukses!"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set koneksi = Nothing
End Sub

