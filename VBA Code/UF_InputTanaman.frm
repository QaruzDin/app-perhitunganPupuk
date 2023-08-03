VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_InputTanaman 
   Caption         =   "Input Data Tanaman"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8505.001
   OleObjectBlob   =   "UF_InputTanaman.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_InputTanaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_Cancel_Click()
    Unload Me
End Sub

Private Sub Button_OK_Click()
    
    addDataTanaman
End Sub

Private Sub CBJPK_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If CBJPK.Value = "Pilih jenis pupuk" Then
        MsgBox "Harap pilih jenis pupuk yang valid.", vbExclamation
        CBJPK.Value = ""
        Cancel = True
    End If
End Sub

Private Sub CBJPN_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If CBJPN.Value = "Pilih jenis pupuk" Then
        MsgBox "Harap pilih jenis pupuk yang valid.", vbExclamation
        CBJPN.Value = ""
        Cancel = True
    End If
End Sub


Private Sub CBJPP_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If CBJPP.Value = "Pilih jenis pupuk" Then
        MsgBox "Harap pilih jenis pupuk yang valid.", vbExclamation
        CBJPP.Value = ""
        Cancel = True
    End If
End Sub


Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As ListColumn
    Dim cell As Range
    
    Set ws = ThisWorkbook.Worksheets("Database Pupuk")
    Set tbl = ws.ListObjects("tabelPupuk")
    Set rng = tbl.ListColumns("Nama Pasar")
    
    For Each cell In rng.DataBodyRange
        CBJPN.AddItem cell.Value
        CBJPP.AddItem cell.Value
        CBJPK.AddItem cell.Value
    Next cell
    
End Sub



