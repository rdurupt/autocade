VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RepSystem 
   Caption         =   "Répertoires Système :"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   Icon            =   "RepSystem.dsx":0000
   OleObjectBlob   =   "RepSystem.dsx":030A
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

    txt7 = Replace(ScanFichier.Chargement("xlt", txt7), txt1, "", 1)
    Unload ScanFichier
End If
End Sub

Private Sub CommandButton11_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt8 = Replace(ScanRep.Chargement(txt8), txt1, "", 1)
End If
End Sub

Private Sub CommandButton12_Click()
Unload Me
End Sub

Private Sub CommandButton13_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt9 = Replace(ScanRep.Chargement(txt9), txt1, "", 1)
End If
End Sub

Private Sub CommandButton14_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt10 = Replace(ScanRep.Chargement(txt10), txt1, "", 1)
End If
End Sub

Private Sub CommandButton15_Click()

If txt1 = "" Then
    MsgBox ""
Else
    txt11 = Replace(ScanRep.Chargement(txt11), txt1, "", 1)
End If
End Sub

Private Sub CommandButton16_Click()
If txt1 = "" Then
    MsgBox ""
Else
     txt12 = Replace(ScanFichier.Chargement("dot", txt12), txt1, "", 1)
End If
End Sub

Private Sub CommandButton3_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt2 = Replace(ScanRep.Chargement(txt2), txt1, "", 1)
End If
End Sub

Private Sub CommandButton4_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt3 = Replace(ScanRep.Chargement(txt3), txt1, "", 1)
End If

End Sub

Private Sub CommandButton5_Click()

    txt1 = Replace(ScanRep.Chargement(txt1), "", 1)

End Sub

Private Sub CommandButton6_Click()
NbTxt = 10
Dim Sql As String
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt1) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathServer';"
Con.Exequte Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt2) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathBlocs';"
Con.Exequte Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt3) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathConnecteursDefault';"
Con.Exequte Sql



Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt4) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathNUMEROFIL';"
Con.Exequte Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt5) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathPlantVierge';"
Con.Exequte Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt6) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathArchiveAutocad';"
Con.Exequte Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt7) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathModelXls';"
Con.Exequte Sql




Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt8) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathOutils';"
Con.Exequte Sql



Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt9) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathComposantsDefault';"
Con.Exequte Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt10) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathNotasDefault';"
Con.Exequte Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt11) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathTorDefault';"
Con.Exequte Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt12) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathModelWord';"
Con.Exequte Sql
Unload Me

End Sub

Private Sub CommandButton7_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt4 = Replace(ScanRep.Chargement(txt4), txt1, "", 1)
End If
End Sub

Private Sub CommandButton8_Click()
If txt1 = "" Then
    MsgBox ""
Else

    txt5 = Replace(ScanFichier.Chargement("dwg", txt5), txt1, "", 1)
    Unload ScanFichier
End If

End Sub

Private Sub CommandButton9_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt6 = Replace(ScanRep.Chargement(txt6), txt1, "", 1)
End If
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Activate()
Dim Sql As String
Dim Rs As Recordset
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub
Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathServer';"
Set Rs = Con.OpenRecordSet(Sql)


txt1 = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathBlocs';"
Set Rs = Con.OpenRecordSet(Sql)
txt2 = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathConnecteursDefault';"
Set Rs = Con.OpenRecordSet(Sql)

txt3 = "" & Rs!PathVar


Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathNUMEROFIL';"
Set Rs = Con.OpenRecordSet(Sql)
txt4 = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathPlantVierge';"
Set Rs = Con.OpenRecordSet(Sql)
txt5 = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathArchiveAutocad';"
Set Rs = Con.OpenRecordSet(Sql)
txt6 = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathModelXls';"
Set Rs = Con.OpenRecordSet(Sql)

txt7 = "" & Rs!PathVar


Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathOutils';"
Set Rs = Con.OpenRecordSet(Sql)
txt8 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathComposantsDefault';"
Set Rs = Con.OpenRecordSet(Sql)
txt9 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathNotasDefault';"
Set Rs = Con.OpenRecordSet(Sql)
txt10 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)


Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathTorDefault';"
Set Rs = Con.OpenRecordSet(Sql)
txt11 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)


Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathModelWord';"
Set Rs = Con.OpenRecordSet(Sql)
txt12 = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)

End Sub

