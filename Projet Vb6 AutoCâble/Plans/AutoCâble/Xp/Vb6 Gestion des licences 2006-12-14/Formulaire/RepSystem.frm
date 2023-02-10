VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RepSystem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Répertoires Système :"
   ClientHeight    =   14475
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   11025
   Icon            =   "RepSystem.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "RepSystem.dsx":030A
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
    Unload ScanRep

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
    Unload ScanRep

End If
End Sub

Private Sub CommandButton14_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt10 = Replace(ScanRep.chargement(txt10), txt1, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub CommandButton15_Click()

If txt1 = "" Then
    MsgBox ""
Else
    txt11 = Replace(ScanRep.chargement(txt11), txt1, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub CommandButton16_Click()
If txt1 = "" Then
    MsgBox ""
Else
     txt12 = Replace(ScanFichier.chargement("dot", txt12), txt1, "", 1)
     Unload ScanFichier
End If
End Sub

Private Sub CommandButton3_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt2 = Replace(ScanRep.chargement(txt2), txt1, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub CommandButton4_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt3 = Replace(ScanRep.chargement(txt3), txt1, "", 1)
    Unload ScanRep

End If

End Sub

Private Sub CommandButton41_Click()
If txt1 = "" Then
    MsgBox ""
Else

    MNC = Replace(ScanFichier.chargement("dot", MNC), txt1, "", 1)
    Unload ScanFichier
End If
End Sub

Private Sub CommandButton42_Click()

If txt1 = "" Then
    MsgBox ""
Else
     PathModelWordMarc = Replace(ScanFichier.chargement("dot", PathModelWordMarc), txt1, "", 1)
     Unload ScanFichier
End If
End Sub

Private Sub CommandButton43_Click()

If txt1 = "" Then
    MsgBox ""
Else
    Eboutique = Replace(ScanRep.chargement(Eboutique), txt1, "", 1)
    Unload ScanRep

End If

End Sub

Private Sub CommandButton44_Click()
If txt1 <> "" Then
    Param_E_Boutique.chargement txt1
    Unload Param_E_Boutique
End If
End Sub

Private Sub CommandButton5_Click()

    txt1 = Replace(ScanRep.chargement(txt1), "", 1)
Unload ScanRep
End Sub

Private Sub CommandButton6_Click()
NbTxt = 10
Dim Sql As String
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt1) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathServer';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt2) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathBlocs';"
Con.Execute Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt3) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathConnecteursDefault';"
Con.Execute Sql



Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt4) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathNUMEROFIL';"
Con.Execute Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt5) & "' "
Sql = Sql & "WHERE T_Path.NameVar='ModelAC';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.MNC) & "' "
Sql = Sql & "WHERE T_Path.NameVar='ModelNC';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt6) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathArchiveAutocad';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(PathModelWordMarc) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathModelWordMarc';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(PDF) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PDF';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt7) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathModelXls';"
Con.Execute Sql

SetUpdate "NbPalachEt", NbPalachEt
SetUpdate "NbPlacheEtiM", NbPlacheEtiM

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt8) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathOutils';"
Con.Execute Sql



Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt9) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathComposantsDefault';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt10) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathNotasDefault';"
Con.Execute Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt11) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathTorDefault';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(txt12) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PathModelWord';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.LI) & "' "
Sql = Sql & "WHERE T_Path.NameVar='LI';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.OU) & "' "
Sql = Sql & "WHERE T_Path.NameVar='OU';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.PL) & "' "
Sql = Sql & "WHERE T_Path.NameVar='PL';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.Fab) & "' "
Sql = Sql & "WHERE T_Path.NameVar='FAB';"

Con.Execute Sql
Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.NC) & "' "
Sql = Sql & "WHERE T_Path.NameVar='NC';"
Con.Execute Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.Synt) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Synt';"
Con.Execute Sql


Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.SyntG) & "' "
Sql = Sql & "WHERE T_Path.NameVar='SyntG';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.AC) & "' "
Sql = Sql & "WHERE T_Path.NameVar='DAC';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.LIEC) & "' "
Sql = Sql & "WHERE T_Path.NameVar='LIEC';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Me.Eboutique) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eboutique';"
Con.Execute Sql
Unload Me

End Sub

Private Sub CommandButton7_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt4 = Replace(ScanRep.chargement(txt4), txt1, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub CommandButton8_Click()
If txt1 = "" Then
    MsgBox ""
Else

    txt5 = Replace(ScanFichier.chargement("dot", txt5), txt1, "", 1)
    Unload ScanFichier
End If

End Sub

Private Sub CommandButton9_Click()
If txt1 = "" Then
    MsgBox ""
Else
    txt6 = Replace(ScanRep.chargement(txt6), txt1, "", 1)
    Unload ScanRep

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

NbPalachEt = GetDefault("NbPalachEt", "3")
NbPlacheEtiM = GetDefault("NbPlacheEtiM", "3")
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
Sql = Sql & "WHERE T_Path.NameVar='ModelAC';"
Set Rs = Con.OpenRecordSet(Sql)
txt5 = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='ModelNC';"
Set Rs = Con.OpenRecordSet(Sql)
Me.MNC = "" & Rs!PathVar


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

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='LI';"
Set Rs = Con.OpenRecordSet(Sql)

Me.LI = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PL';"
Set Rs = Con.OpenRecordSet(Sql)
Me.PL = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='OU';"
Set Rs = Con.OpenRecordSet(Sql)
Me.OU = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='FAB';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Fab = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='DNC';"
Set Rs = Con.OpenRecordSet(Sql)
Me.NC = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PathModelWordMarc';"
Set Rs = Con.OpenRecordSet(Sql)
Me.PathModelWordMarc = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Synt';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Synt = "" & Rs!PathVar


Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='SyntG';"
Set Rs = Con.OpenRecordSet(Sql)
Me.SyntG = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='DAC';"
Set Rs = Con.OpenRecordSet(Sql)
Me.AC = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='PDF';"
Set Rs = Con.OpenRecordSet(Sql)
Me.PDF = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='LIEC';"
Set Rs = Con.OpenRecordSet(Sql)
Me.LIEC = "" & Rs!PathVar


Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eboutique';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eboutique = "" & Rs!PathVar
Set Rs = Con.CloseRecordSet(Rs)
End Sub

