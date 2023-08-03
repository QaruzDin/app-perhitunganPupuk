VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_HapusTanaman 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UF_HapusTanaman.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_HapusTanaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_Batalkan_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
   
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rngData As Range
    Dim cell As Range

    Set ws = ThisWorkbook.Worksheets("Database Tanaman")
    Set tbl = ws.ListObjects("tabelTanaman")
    Set rngData = tbl.ListColumns("Nama Tanaman").DataBodyRange ' Ganti "Nama Tanaman" dengan nama kolom yang sesuai

    For Each cell In rngData
        CBDataTanaman.AddItem cell.Value
    Next cell
End Sub

Private Sub Button_Hapus_Click()
    ' Hapus data dari tabel berdasarkan pilihan pengguna di ComboBox
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rngData As Range
    Dim selectedData As String
    Dim deleteRow As Range

    ' Pastikan ada data yang dipilih di ComboBox
    If CBDataTanaman.ListIndex = -1 Then
        MsgBox "Harap pilih data yang ingin dihapus.", vbExclamation
        Exit Sub
    End If

    selectedData = CBDataTanaman.Value

    Set ws = ThisWorkbook.Worksheets("Database Tanaman")
    Set tbl = ws.ListObjects("tabelTanaman")
    Set rngData = tbl.ListColumns("Nama Tanaman").DataBodyRange ' Ganti "Nama Tanaman" dengan nama kolom yang sesuai

    ' Cari baris yang berisi data yang ingin dihapus
    Set deleteRow = rngData.Find(what:=selectedData, LookIn:=xlValues, lookat:=xlWhole)
    
    If Not deleteRow Is Nothing Then
        deleteRow.EntireRow.Delete
        MsgBox "Data berhasil dihapus.", vbOKOnly
    Else
        MsgBox "Data tidak ditemukan.", vbExclamation
    End If

    ' Tutup UserForm setelah data dihapus
    Unload Me
End Sub
