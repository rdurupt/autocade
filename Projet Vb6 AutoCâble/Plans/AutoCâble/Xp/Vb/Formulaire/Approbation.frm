VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Approbation 
   Caption         =   "Approbation Plan :"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   Icon            =   "Approbation.dsx":0000
   OleObjectBlob   =   "Approbation.dsx":08CA
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Approbation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdProjet As Long
Public IdPieces As Long
Public IdIndiceProjet As Long
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    txt20 = Format(Date, "dd/mm/yyyy")
Else
    txt20 = ""
End If
End Sub

Private Sub CommandButton1_Click()
Dim Sql As String
Dim Rs As Recordset

CherchPices.Charge Me, " LiAutoCadSave <>  Null and NbErr=0 and  VerifieDate<> Null  and Archiver=False", True
Unload CherchPices



Sql = "SELECT RqCartouche.Projet AS txt1,  "
Sql = Sql & "RqCartouche.Vague AS txt2,  "
Sql = Sql & "RqCartouche.Equipement AS txt3,  "
Sql = Sql & "RqCartouche.Responsable AS txt4,  "
Sql = Sql & "RqCartouche.Ensemble AS txt5,  "
Sql = Sql & "RqCartouche.CleAc AS txt6,  "
Sql = Sql & "[PI] & '_' & Trim('' & [PI_Indice]) AS txt7,  "
Sql = Sql & "[PL] & '_' & Trim('' & [PL_Indice]) AS txt8,  "
Sql = Sql & "[OU] & '_' & Trim('' & [OU_Indice]) AS txt9,  "
Sql = Sql & "[Li] & '_' & Trim('' & [LI_Indice]) AS txt10,  "
Sql = Sql & "RqCartouche.Client AS txt11,  "
Sql = Sql & "RqCartouche.Destinataire AS txt12,  "
Sql = Sql & "RqCartouche.Service AS txt13,  "
Sql = Sql & "RqCartouche.RefPF AS txt14,  "
Sql = Sql & "RqCartouche.RefP AS txt15,  "
Sql = Sql & "RqCartouche.DessineDate AS txt16,  "
Sql = Sql & "RqCartouche.DessineNOM AS txt17,  "
Sql = Sql & "RqCartouche.VerifieDate AS txt18,  "
Sql = Sql & "RqCartouche.VerifieNom AS txt19,  "
Sql = Sql & "RqCartouche.ApprouveDate AS txt20,  "
Sql = Sql & "RqCartouche.ApprouveNom AS txt21 "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Val(Me.Tag) & " ;"
Debug.Print Sql
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
 Me.Controls("txt" & CStr(1)).Caption = "" & Rs.Fields("txt" & CStr(1))
For i = 2 To 3
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
 Me.Controls("txt" & CStr(4)).Caption = "" & Rs.Fields("txt" & CStr(4))
  Me.Controls("txt" & CStr(5)) = "" & Rs.Fields("txt" & CStr(5))
For i = 6 To 12
    Me.Controls("txt" & CStr(i)).Caption = "" & Rs.Fields("txt" & CStr(i))
Next i
For i = 13 To 15
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
For i = 16 To 18 Step 2
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
 Me.Controls("txt" & CStr(20)) = "" & Rs.Fields("txt" & CStr(20))
For i = 17 To 21 Step 2
    Me.Controls("txt" & CStr(i)).Caption = "" & Rs.Fields("txt" & CStr(i))
Next i
End If
If txt20 <> "" Then
    Me.CheckBox1.Value = True
Else
     Me.CheckBox1.Value = False
End If

End Sub

Private Sub CommandButton7_Click()
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As Recordset
boolValideMOD = False
Set FormBarGrah = Me
If MyFormat("DATE", txt16, "Déssiné par") = False Then Exit Sub
If MyFormat("DATE", txt18, "Vérifié par") = False Then Exit Sub
If MyFormat("DATE", txt20, "Approuvé par") = False Then Exit Sub
If Trim("" & Me.Tag) = "" Then
    CommandButton1_Click
    Exit Sub
End If


Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs!Pere > 0 Then Me.Tag = Rs!Pere

Sql = "SELECT T_indiceProjet.IdStatus, T_indiceProjet.IdStatusSave "
Sql = Sql & "FROM T_indiceProjet "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Trim("" & Rs!IdStatusSave) = "" Then
If Rs!IdStatus = "2" And Trim("" & Rs!IdStatusSave) = "" Then
    If FrmIndice.Charge(txt1, txt2, txt3, txt5, txt11, txt6, txt7, txt8, txt9, txt10) = False Then Exit Sub
    boolValideMOD = True
   Sql = "UPDATE T_indiceProjet SET T_indiceProjet.PI = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 1)) & "', "
Sql = Sql & " T_indiceProjet.PI_Indice = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 5)) & "',  "
Sql = Sql & "T_indiceProjet.PL = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 1)) & "',  "
Sql = Sql & "T_indiceProjet.PL_Indice = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 5)) & "',  "
Sql = Sql & "T_indiceProjet.[OU] = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 1)) & "',  "
Sql = Sql & "T_indiceProjet.OU_Indice = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 2)) & "',  "
Sql = Sql & "T_indiceProjet.Li = '" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 1)) & "',  "
Sql = Sql & "T_indiceProjet.LI_Indice ='" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 2)) & "',  "
Sql = Sql & "T_indiceProjet.DessineNOM = '" & MyReplace(FrmIndice.txt10) & "',  "
Sql = Sql & "T_indiceProjet.VerifieNom = '" & MyReplace(FrmIndice.txt11) & "',  "
Sql = Sql & "T_indiceProjet.ApprouveNom = '" & MyReplace(FrmIndice.txt12) & "', "
Sql = Sql & "T_indiceProjet.ReffIndice= '" & MyReplace(FrmIndice.ReffIndice) & "', "
Sql = Sql & "T_indiceProjet.Version=1,T_indiceProjet.Description='" & MyReplace(FrmIndice.DescIndice) & "' "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Con.Exequte Sql


 Sql = "UPDATE T_indiceProjet SET T_indiceProjet.PI = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 1)) & "', "
Sql = Sql & " T_indiceProjet.PI_Indice = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 5)) & "',  "
Sql = Sql & "T_indiceProjet.PL = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 1)) & "',  "
Sql = Sql & "T_indiceProjet.PL_Indice = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 5)) & "',  "
Sql = Sql & "T_indiceProjet.[OU] = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 1)) & "',  "
Sql = Sql & "T_indiceProjet.OU_Indice = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 2)) & "',  "
Sql = Sql & "T_indiceProjet.Li = '" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 1)) & "',  "
Sql = Sql & "T_indiceProjet.LI_Indice ='" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 2)) & "',  "
Sql = Sql & "T_indiceProjet.DessineNOM = '" & MyReplace(FrmIndice.txt10) & "',  "
Sql = Sql & "T_indiceProjet.VerifieNom = '" & MyReplace(FrmIndice.txt11) & "',  "
Sql = Sql & "T_indiceProjet.ApprouveNom = '" & MyReplace(FrmIndice.txt12) & "', "
Sql = Sql & "T_indiceProjet.ReffIndice= '" & MyReplace(FrmIndice.ReffIndice) & "', "
Sql = Sql & "T_indiceProjet.Version=1,T_indiceProjet.Description='" & MyReplace(FrmIndice.DescIndice) & "' "
Sql = Sql & "WHERE T_indiceProjet.pere=" & Me.Tag & ";"
Con.Exequte Sql

    Unload FrmIndice
End If
    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.IdStatusSave = [T_indiceProjet].[IdStatus] "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
    Con.Exequte Sql
     Sql = "UPDATE T_indiceProjet SET T_indiceProjet.IdStatusSave = [T_indiceProjet].[IdStatus] "
    Sql = Sql & "WHERE T_indiceProjet.pere=" & Me.Tag & ";"
    Con.Exequte Sql
End If
Rs.Requery

Sql = "UPDATE RqCartouche SET "
Sql = Sql & "RqCartouche.ApprouveDate = " & MyReplaceDate(txt20) & ", "
If CheckBox1.Value = True Then
    Sql = Sql & "RqCartouche.IdStatus =3 "
Else
    Sql = Sql & "RqCartouche.IdStatus =" & Rs!IdStatusSave & " "
End If
Sql2 = "WHERE RqCartouche.T_indiceProjet.Id=" & Me.Tag & ";"
Sql3 = "WHERE RqCartouche.pere=" & Me.Tag & ";"
Con.Exequte Sql & Sql2
Con.Exequte Sql & Sql3
IdIndiceProjet = Me.Tag
Sql = "SELECT T_indiceProjet.Id_Pieces FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
IdPieces = Rs!Id_Pieces
Sql = "SELECT T_Pieces.IdProjet FROM T_Pieces "
Sql = Sql & "WHERE T_Pieces.Id=" & IdPieces & ";"
Set Rs = Con.OpenRecordSet(Sql)
IdProjet = Rs!IdProjet
bool_Plan_L_cartouches = True: bool_Plan_E_cartouches = True
 bool_Outil_L_cartouches = True: bool_Outil_E_cartouches = True

 ModifierUnCartouche Me.Tag, True
 bool_Plan_L_cartouches = False: bool_Plan_E_cartouches = False
 bool_Outil_L_cartouches = False: bool_Outil_E_cartouches = False
 MajEcartIndice IdPieces
Me.Hide

End Sub

Private Sub CommandButton8_Click()
Me.Hide

End Sub

