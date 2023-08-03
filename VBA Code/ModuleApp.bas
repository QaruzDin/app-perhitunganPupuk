Attribute VB_Name = "ModuleApp"

Sub addDataTanaman()
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    'validasi data kosong
    If UF_InputTanaman.namaTanaman.Value = "" Or _
       UF_InputTanaman.CBJPN.Value = "Pilih jenis pupuk" Or _
       UF_InputTanaman.jmlN.Value = "" Or _
       UF_InputTanaman.CBJPP.Value = "Pilih jenis pupuk" Or _
       UF_InputTanaman.jmlP.Value = "" Or _
       UF_InputTanaman.CBJPK.Value = "Pilih jenis pupuk" Or _
       UF_InputTanaman.jmlK.Value = "" Or _
       UF_InputTanaman.namaPORGANIK.Value = "" Or _
       UF_InputTanaman.jmlPORGANIK.Value = "" Then
        MsgBox "Harap mengisi seluruh form yang ditandai bintang (*).", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets("Database Tanaman")
    Set tbl = ws.ListObjects("tabelTanaman")
    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = UF_InputTanaman.namaTanaman.Value
    newRow.Range(1, 2).Value = UF_InputTanaman.namaVarietas.Value
    newRow.Range(1, 3).Value = UF_InputTanaman.CBJPN.Value
    newRow.Range(1, 4).Value = UF_InputTanaman.jmlN.Value
    newRow.Range(1, 5).Value = UF_InputTanaman.CBJPP.Value
    newRow.Range(1, 6).Value = UF_InputTanaman.jmlP.Value
    newRow.Range(1, 7).Value = UF_InputTanaman.CBJPK.Value
    newRow.Range(1, 8).Value = UF_InputTanaman.jmlK.Value
    newRow.Range(1, 9).Value = UF_InputTanaman.namaPORGANIK.Value
    newRow.Range(1, 10).Value = UF_InputTanaman.jmlPORGANIK.Value
    
    MsgBox "Data berhasil ditambahkan", vbOKOnly
    
    cleanFormDataTanaman
End Sub


Public Sub ShowUserFormTanaman()
    UF_InputTanaman.Show
End Sub


Public Sub junk()
    With CBJPN
        .AddItem "Urea"
        .AddItem "ZN"
    End With
    With CBJPP
        .AddItem "SP-36"
        .AddItem "Phonska"
    End With
    With CBJPK
        .AddItem "KCl"
        .AddItem "KNO3"
    End With
End Sub

Public Sub cleanFormDataTanaman()
    UF_InputTanaman.namaTanaman.Value = ""
    UF_InputTanaman.namaVarietas.Value = ""
    UF_InputTanaman.CBJPN.Value = "Pilih jenis pupuk"
    UF_InputTanaman.jmlN.Value = ""
    UF_InputTanaman.CBJPP.Value = "Pilih jenis pupuk"
    UF_InputTanaman.jmlP.Value = ""
    UF_InputTanaman.CBJPK.Value = "Pilih jenis pupuk"
    UF_InputTanaman.jmlK.Value = ""
    UF_InputTanaman.namaPORGANIK.Value = ""
    UF_InputTanaman.jmlPORGANIK.Value = ""
End Sub

Public Sub ShowUFDeleteTanaman()
    UF_HapusTanaman.Show
End Sub
