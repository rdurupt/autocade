VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RepSystem 
   Caption         =   "Répertoires Système :"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   OleObjectBlob   =   "RepSystem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RepSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NbTxt As Long

Private Sub CommandButton10_Click()
If txt1 = "" Then
    MsgBox ""
Else

    txt7 = Replace(ScanFichier.chargement("xlt", txt7), txt1, "", 1)
    Unload ScanFichier
End If
End Sub

Private Sub CommandButton11_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt8 = Replace(ScanRep.chargement(txt8), txt1, "", 1)
End If
End Sub

Private Sub CommandButton12_Click()
Unload Me
End Sub

Private Sub CommandButton13_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt9 = Replace(ScanRep.chargement(txt9), txt1, "", 1)
End If
End Sub

Private Sub CommandButton14_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt10 = Replace(ScanRep.chargement(txt10), txt1, "", 1)
End If
End Sub

Private Sub CommandButton15_Click()

If txt1 = "" Then
    MsgBox ""
Else
    txt11 = Replace(ScanRep.chargement(txt11), txt1, "", 1)
End If
End Sub

Private Sub CommandButton3_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt2 = Replace(ScanRep.chargement(txt2), txt1, "", 1)
End If
End Sub

Private Sub CommandButton4_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt3 = Replace(ScanRep.chargement(txt3), txt1, "", 1)
End If

End Sub

Private Sub CommandButton5_Click()

    txt1 = Replace(ScanRep.chargement(txt1), "", 1)

End Sub

Private Sub CommandButton6_Click()
NbTxt = 10
Dim sql As String
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub

sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt1) & "' "
sql = sql & "WHERE T_Path.NameVar='PathServer';"
Con.Exequte sql

sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt2) & "' "
sql = sql & "WHERE T_Path.NameVar='PathBlocs';"
Con.Exequte sql


sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt3) & "' "
sql = sql & "WHERE T_Path.NameVar='PathConnecteursDefault';"
Con.Exequte sql



sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt4) & "' "
sql = sql & "WHERE T_Path.NameVar='PathNUMEROFIL';"
Con.Exequte sql


sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt5) & "' "
sql = sql & "WHERE T_Path.NameVar='PathPlantVierge';"
Con.Exequte sql


sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt6) & "' "
sql = sql & "WHERE T_Path.NameVar='PathArchiveAutocad';"
Con.Exequte sql


sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt7) & "' "
sql = sql & "WHERE T_Path.NameVar='PathModelXls';"
Con.Exequte sql




sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt8) & "' "
sql = sql & "WHERE T_Path.NameVar='PathOutils';"
Con.Exequte sql



sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt9) & "' "
sql = sql & "WHERE T_Path.NameVar='PathComposantsDefault';"
Con.Exequte sql

sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt10) & "' "
sql = sql & "WHERE T_Path.NameVar='PathNotasDefault';"
Con.Exequte sql


sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt11) & "' "
sql = sql & "WHERE T_Path.NameVar='PathTorDefault';"
Con.Exequte sql

Unload Me

End Sub

Private Sub CommandButton7_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt4 = Replace(ScanRep.chargement(txt4), txt1, "", 1)
End If
End Sub

Private Sub CommandButton8_Click()
If txt1 = "" Then
    MsgBox ""
Else

    txt5 = Replace(ScanFichier.chargement("dwg", txt5), txt1, "", 1)
    Unload ScanFichier
End If

End Sub

Private Sub CommandButton9_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt6 = Replace(ScanRep.chargement(txt6), txt1, "", 1)
End If
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Activate()
Dim sql As String
Dim Rs As Recordset
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub
sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathServer';"
Set Rs = Con.OpenRecordSet(sql)


txt1 = "" & Rs!PathVar

sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathBlocs';"
Set Rs = Con.OpenRecordSet(sql)
txt2 = "" & Rs!PathVar

sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathConnecteursDefault';"
Set Rs = Con.OpenRecordSet(sql)

txt3 = "" & Rs!PathVar


sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathNUMEROFIL';"
Set Rs = Con.OpenRecordSet(sql)
txt4 = "" & Rs!PathVar

sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathPlantVierge';"
Set Rs = Con.OpenRecordSet(sql)
txt5 = "" & Rs!PathVar

sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathArchiveAutocad';"
Set Rs = Con.OpenRecordSet(sql)
txt6 = "" & Rs!PathVar

sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathModelXls';"
Set Rs = Con.OpenRecordSet(sql)

txt7 = "" & Rs!PathVar


sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathOutils';"
Set Rs = Con.OpenRecordSet(sql)
txt8 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)

sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathComposantsDefault';"
Set Rs = Con.OpenRecordSet(sql)
txt9 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)

sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathNotasDefault';"
Set Rs = Con.OpenRecordSet(sql)
txt10 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)


sql = "select T_Path.PathVar from T_Path "
sql = sql & "WHERE T_Path.NameVar='PathTorDefault';"
Set Rs = Con.OpenRecordSet(sql)
txt11 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)


End Sub

