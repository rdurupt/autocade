VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Approbation 
   Caption         =   "Approbation Plan :"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   OleObjectBlob   =   "Approbation.frx":0000
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
Dim sql As String
Dim Rs As Recordset

CherchPices.Charge Me, " LiAutoCadSave <>  Null and NbErr=0 and  VerifieDate<> Null  and Archiver=False", True
Unload CherchPices



sql = "SELECT RqCartouche.Projet AS txt1,  "
sql = sql & "RqCartouche.Vague AS txt2,  "
sql = sql & "RqCartouche.Equipement AS txt3,  "
sql = sql & "RqCartouche.Responsable AS txt4,  "
sql = sql & "RqCartouche.Ensemble AS txt5,  "
sql = sql & "RqCartouche.CleAc AS txt6,  "
sql = sql & "[PI] & '_' & Trim('' & [PI_Indice]) AS txt7,  "
sql = sql & "[PL] & '_' & Trim('' & [PL_Indice]) AS txt8,  "
sql = sql & "[OU] & '_' & Trim('' & [OU_Indice]) AS txt9,  "
sql = sql & "[Li] & '_' & Trim('' & [LI_Indice]) AS txt10,  "
sql = sql & "RqCartouche.Client AS txt11,  "
sql = sql & "RqCartouche.Destinataire AS txt12,  "
sql = sql & "RqCartouche.Service AS txt13,  "
sql = sql & "RqCartouche.RefPF AS txt14,  "
sql = sql & "RqCartouche.RefP AS txt15,  "
sql = sql & "RqCartouche.DessineDate AS txt16,  "
sql = sql & "RqCartouche.DessineNOM AS txt17,  "
sql = sql & "RqCartouche.VerifieDate AS txt18,  "
sql = sql & "RqCartouche.VerifieNom AS txt19,  "
sql = sql & "RqCartouche.ApprouveDate AS txt20,  "
sql = sql & "RqCartouche.ApprouveNom AS txt21 "
sql = sql & "FROM RqCartouche "
sql = sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Val(Me.Tag) & " ;"
Debug.Print sql
Set Rs = Con.OpenRecordSet(sql)
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
Dim sql As String
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


sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs!Pere > 0 Then Me.Tag = Rs!Pere

sql = "SELECT T_indiceProjet.IdStatus, T_indiceProjet.IdStatusSave "
sql = sql & "FROM T_indiceProjet "
 sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(sql)
If Trim("" & Rs!IdStatusSave) = "" Then
If Rs!IdStatus = "2" And Trim("" & Rs!IdStatusSave) = "" Then
    If FrmIndice.Charge(txt1, txt2, txt3, txt5, txt11, txt6, txt7, txt8, txt9, txt10) = False Then Exit Sub
    boolValideMOD = True
   sql = "UPDATE T_indiceProjet SET T_indiceProjet.PI = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 1)) & "', "
sql = sql & " T_indiceProjet.PI_Indice = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 5)) & "',  "
sql = sql & "T_indiceProjet.PL = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 1)) & "',  "
sql = sql & "T_indiceProjet.PL_Indice = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 5)) & "',  "
sql = sql & "T_indiceProjet.[OU] = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 1)) & "',  "
sql = sql & "T_indiceProjet.OU_Indice = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 2)) & "',  "
sql = sql & "T_indiceProjet.Li = '" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 1)) & "',  "
sql = sql & "T_indiceProjet.LI_Indice ='" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 2)) & "',  "
sql = sql & "T_indiceProjet.DessineNOM = '" & MyReplace(FrmIndice.txt10) & "',  "
sql = sql & "T_indiceProjet.VerifieNom = '" & MyReplace(FrmIndice.txt11) & "',  "
sql = sql & "T_indiceProjet.ApprouveNom = '" & MyReplace(FrmIndice.txt12) & "', "
sql = sql & "T_indiceProjet.Version=1,T_indiceProjet.Description='" & MyReplace(FrmIndice.DescIndice) & "' "
sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Con.Exequte sql


 sql = "UPDATE T_indiceProjet SET T_indiceProjet.PI = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 1)) & "', "
sql = sql & " T_indiceProjet.PI_Indice = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 5)) & "',  "
sql = sql & "T_indiceProjet.PL = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 1)) & "',  "
sql = sql & "T_indiceProjet.PL_Indice = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 5)) & "',  "
sql = sql & "T_indiceProjet.[OU] = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 1)) & "',  "
sql = sql & "T_indiceProjet.OU_Indice = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 2)) & "',  "
sql = sql & "T_indiceProjet.Li = '" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 1)) & "',  "
sql = sql & "T_indiceProjet.LI_Indice ='" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 2)) & "',  "
sql = sql & "T_indiceProjet.DessineNOM = '" & MyReplace(FrmIndice.txt10) & "',  "
sql = sql & "T_indiceProjet.VerifieNom = '" & MyReplace(FrmIndice.txt11) & "',  "
sql = sql & "T_indiceProjet.ApprouveNom = '" & MyReplace(FrmIndice.txt12) & "', "
sql = sql & "T_indiceProjet.Version=1,T_indiceProjet.Description='" & MyReplace(FrmIndice.DescIndice) & "' "
sql = sql & "WHERE T_indiceProjet.pere=" & Me.Tag & ";"
Con.Exequte sql

    Unload FrmIndice
End If
    sql = "UPDATE T_indiceProjet SET T_indiceProjet.IdStatusSave = [T_indiceProjet].[IdStatus] "
    sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
    Con.Exequte sql
     sql = "UPDATE T_indiceProjet SET T_indiceProjet.IdStatusSave = [T_indiceProjet].[IdStatus] "
    sql = sql & "WHERE T_indiceProjet.pere=" & Me.Tag & ";"
    Con.Exequte sql
End If
Rs.Requery

sql = "UPDATE RqCartouche SET "
sql = sql & "RqCartouche.ApprouveDate = " & MyReplaceDate(txt20) & ", "
If CheckBox1.Value = True Then
    sql = sql & "RqCartouche.IdStatus =3 "
Else
    sql = sql & "RqCartouche.IdStatus =" & Rs!IdStatusSave & " "
End If
Sql2 = "WHERE RqCartouche.T_indiceProjet.Id=" & Me.Tag & ";"
Sql3 = "WHERE RqCartouche.pere=" & Me.Tag & ";"
Con.Exequte sql & Sql2
Con.Exequte sql & Sql3
IdIndiceProjet = Me.Tag
sql = "SELECT T_indiceProjet.Id_Pieces FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(sql)
IdPieces = Rs!Id_Pieces
sql = "SELECT T_Pieces.IdProjet FROM T_Pieces "
sql = sql & "WHERE T_Pieces.Id=" & IdPieces & ";"
Set Rs = Con.OpenRecordSet(sql)
IdProjet = Rs!IdProjet


 ModifierUnCartouche Me.Tag, True
Me.Hide

End Sub

Private Sub CommandButton8_Click()
Me.Hide

End Sub

