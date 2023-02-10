VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Approbation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Approbation Plan :"
   ClientHeight    =   7260
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   12240
   Icon            =   "Approbation.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "Approbation.dsx":08CA
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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

CherchPices.Charge Me, " LiAutoCadSave <>  Null and NbErr=0 and  VerifieDate<> Null  and  IdStatus<>4 And Archiver = False", True
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
For I = 2 To 3
    Me.Controls("txt" & CStr(I)) = "" & Rs.Fields("txt" & CStr(I))
Next I
 Me.Controls("txt" & CStr(4)).Caption = "" & Rs.Fields("txt" & CStr(4))
  Me.Controls("txt" & CStr(5)) = "" & Rs.Fields("txt" & CStr(5))
For I = 6 To 12
    Me.Controls("txt" & CStr(I)).Caption = "" & Rs.Fields("txt" & CStr(I))
Next I
For I = 13 To 15
    Me.Controls("txt" & CStr(I)) = "" & Rs.Fields("txt" & CStr(I))
Next I
For I = 16 To 18 Step 2
    Me.Controls("txt" & CStr(I)) = "" & Rs.Fields("txt" & CStr(I))
Next I
 Me.Controls("txt" & CStr(20)) = "" & Rs.Fields("txt" & CStr(20))
For I = 17 To 21 Step 2
    Me.Controls("txt" & CStr(I)).Caption = "" & Rs.Fields("txt" & CStr(I))
Next I
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

Sql = "SELECT  T_indiceProjet.Id,T_indiceProjet.IdStatus, T_indiceProjet.IdStatusSave, T_indiceProjet.Vague, T_indiceProjet.Equipement, T_indiceProjet.Ensemble, [pi] & '_' & Trim([pi_Indice]) AS Piece, [PL] & '_' & Trim([PL_Indice]) AS Plan, [OU] & '_' & Trim([OU_Indice]) AS Outil, [Li] & '_' & Trim([LI_Indice]) AS Liste, T_indiceProjet.DessineNOM AS Dessin, T_indiceProjet.VerifieNom AS Verif, T_indiceProjet.ApprouveNom AS Approuv "
Sql = Sql & "FROM T_indiceProjet "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
 
Set Rs = Con.OpenRecordSet(Sql)
'Rs!IdStatusSave = Null
'Rs.Update
If Trim("" & Rs!IdStatusSave) = "" Then
If Rs!IdStatus = "2" And Trim("" & Rs!IdStatusSave) = "" And Trim("" & txt20) <> "" Then
    If FrmIndice.Charge(Val(Me.Tag), txt1, txt2, "" & Rs!Equipement, "" & Rs!Ensemble, txt11, txt6, "" & Rs!Piece, "" & Rs!Plan, "" & Rs!Outil, "" & Rs!Liste, "" & Rs!Dessin, "" & Rs!Verif, "" & Rs!Approuv) = False Then
    Unload FrmIndice
    Exit Sub
    End If
'    boolValideMOD = True
'   Sql = "UPDATE T_indiceProjet SET T_indiceProjet.PI = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 1)) & "', "
'Sql = Sql & " T_indiceProjet.PI_Indice = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 5)) & "',  "
'Sql = Sql & "T_indiceProjet.PL = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 1)) & "',  "
'Sql = Sql & "T_indiceProjet.PL_Indice = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 5)) & "',  "
'Sql = Sql & "T_indiceProjet.[OU] = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 1)) & "',  "
'Sql = Sql & "T_indiceProjet.OU_Indice = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 2)) & "',  "
'Sql = Sql & "T_indiceProjet.Li = '" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 1)) & "',  "
'Sql = Sql & "T_indiceProjet.LI_Indice ='" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 2)) & "',  "
'Sql = Sql & "T_indiceProjet.DessineNOM = '" & MyReplace(FrmIndice.txt10) & "',  "
'Sql = Sql & "T_indiceProjet.VerifieNom = '" & MyReplace(FrmIndice.txt11) & "',  "
'Sql = Sql & "T_indiceProjet.ApprouveNom = '" & MyReplace(FrmIndice.txt12) & "', "
'Sql = Sql & "T_indiceProjet.ReffIndice= '" & MyReplace(FrmIndice.ReffIndice) & "', "
'Sql = Sql & "T_indiceProjet.Version=1,T_indiceProjet.Description='" & MyReplace(FrmIndice.DescIndice) & "' "
'Sql = Sql & ", T_indiceProjet.DNC = '" & MyReplace(FrmIndice.lstNc) & "' , T_indiceProjet.LIEC = '" & MyReplace(FrmIndice.lstLi) & "' "
'
'Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'Con.Execute Sql
'
'
' Sql = "UPDATE T_indiceProjet SET T_indiceProjet.PI = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 1)) & "', "
'Sql = Sql & " T_indiceProjet.PI_Indice = '" & MyReplace(FrmIndice.txt5.List(FrmIndice.txt5.ListIndex, 5)) & "',  "
'Sql = Sql & "T_indiceProjet.PL = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 1)) & "',  "
'Sql = Sql & "T_indiceProjet.PL_Indice = '" & MyReplace(FrmIndice.txt6.List(FrmIndice.txt6.ListIndex, 5)) & "',  "
'Sql = Sql & "T_indiceProjet.[OU] = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 1)) & "',  "
'Sql = Sql & "T_indiceProjet.OU_Indice = '" & MyReplace(FrmIndice.txt7.List(FrmIndice.txt7.ListIndex, 2)) & "',  "
'Sql = Sql & "T_indiceProjet.Li = '" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 1)) & "',  "
'Sql = Sql & "T_indiceProjet.LI_Indice ='" & MyReplace(FrmIndice.txt8.List(FrmIndice.txt8.ListIndex, 2)) & "',  "
'Sql = Sql & "T_indiceProjet.DessineNOM = '" & MyReplace(FrmIndice.txt10) & "',  "
'Sql = Sql & "T_indiceProjet.VerifieNom = '" & MyReplace(FrmIndice.txt11) & "',  "
'Sql = Sql & "T_indiceProjet.ApprouveNom = '" & MyReplace(FrmIndice.txt12) & "', "
'Sql = Sql & "T_indiceProjet.ReffIndice= '" & MyReplace(FrmIndice.ReffIndice) & "', "
'Sql = Sql & "T_indiceProjet.Version=1,T_indiceProjet.Description='" & MyReplace(FrmIndice.DescIndice) & "' "
'Sql = Sql & ", T_indiceProjet.DNC = '" & MyReplace(FrmIndice.lstNc) & "' , T_indiceProjet.LIEC = '" & MyReplace(FrmIndice.lstLi) & "' "
'Sql = Sql & "WHERE T_indiceProjet.pere=" & Me.Tag & ";"
'Con.Execute Sql

    Unload FrmIndice
End If

    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.IdStatusSave = [T_indiceProjet].[IdStatus] "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
    Con.Execute Sql
     Sql = "UPDATE T_indiceProjet SET T_indiceProjet.IdStatusSave = [T_indiceProjet].[IdStatus] "
    Sql = Sql & "WHERE T_indiceProjet.pere=" & Me.Tag & ";"
    Con.Execute Sql
End If
Rs.Requery

Sql = "UPDATE RqCartouche SET "
Sql = Sql & "RqCartouche.ApprouveDate = " & MyReplaceDate(txt20) & ", "
If CheckBox1.Value = True Then
    Sql = Sql & "RqCartouche.IdStatus =3 "
     strStatus = "VAL"
     boolValideMOD = True
Else
    boolValideMOD = False
    Sql = Sql & "RqCartouche.IdStatus =" & Rs!IdStatusSave & " "
    If Rs!IdStatusSave = 2 Then strStatus = "MOD"
    If Rs!IdStatusSave = 1 Then strStatus = "CRE"
End If
Sql2 = "WHERE RqCartouche.T_indiceProjet.Id=" & Me.Tag & ";"
Sql3 = "WHERE RqCartouche.pere=" & Me.Tag & ";"
Con.Execute Sql & Sql2
Con.Execute Sql & Sql3
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
If IsCilent = False Then
 ModifierUnCartouche Me.Tag, True
 MajEcartIndice IdPieces
 
 End If
 bool_Plan_L_cartouches = False: bool_Plan_E_cartouches = False
 bool_Outil_L_cartouches = False: bool_Outil_E_cartouches = False
 
 If IsCilent = True Then

'Sql = "INSERT INTO T_Job ( Id_Piece, Id_Fils, Plan_L_Fils, Plan_L_Composants, Plan_L_Noeuds,  "
'Sql = Sql & "Plan_L_Notas, Plan_L_cartouches, Plan_Ouvrir, Outil_L_Fils, Outil_L_Composants,  "
'Sql = Sql & "Outil_L_Noeuds, Outil_L_Notas, Outil_L_cartouches, Outil_Ouvrir,Machine ) "
'Sql = Sql & "values ( " & Id & ", " & IdFils & ", " & MyReplaceBool(Me.Plan_L_Fils) & ", " & MyReplaceBool(Me.Plan_L_Composants) & ",  "
'Sql = Sql & MyReplaceBool(Me.Plan_L_Noeuds) & "," & MyReplaceBool(Me.Plan_L_Notas) & ", " & MyReplaceBool(Me.Plan_L_cartouches) & ","
'Sql = Sql & MyReplaceBool(Me.Plan_Ouvrir) & "," & MyReplaceBool(Me.Outil_L_Fils) & ", " & MyReplaceBool(Me.Outil_L_Composants) & ", "
'Sql = Sql & MyReplaceBool(Me.Outil_L_Noeuds) & ", " & MyReplaceBool(Me.Outil_L_Notas) & "," & MyReplaceBool(Me.Outil_L_cartouches) & ", "
'Sql = Sql & MyReplaceBool(Me.Outil_Ouvrir) & ",'" &  MyReplace(Machine) & "' );"
Sql = "SELECT [PI] & '_' & Trim([PI_Indice]) AS Name  "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then

    Sql = "DELETE T_Job.* FROM T_Job "
    Sql = Sql & "WHERE T_Job.Id_Piece=" & Me.Tag & ";"
    Con.Execute Sql
    
    Sql = "INSERT INTO T_Job ( Id_Piece, Id_Fils, Action,Outil_E_cartouches, Outil_E_Connecteurs, Outil_E_Criteres, "
Sql = Sql & "Outil_E_Etiquettes, Outil_E_Fils, Outil_E_Noeuds, Outil_E_Notas, Outil_E_Options,  Outil_E_Preconisations,  "
Sql = Sql & "Outil_E_Vignettes, Outil_L_cartouches, Outil_L_Composants,  Outil_L_Connecteurs, Outil_L_Criteres, Outil_L_Etiquettes,  "
Sql = Sql & "Outil_L_Fils, Outil_L_Noeuds,  Outil_L_Notas, Outil_L_Options, Outil_L_Preconisations, Outil_L_Vignettes, Outil_Ouvrir,   "
Sql = Sql & "Plan_E_cartouches, Plan_E_Composants, Plan_E_Connecteurs, Plan_E_Criteres, Plan_E_Etiquettes,  Plan_E_Fils, Plan_E_Noeuds,  "
Sql = Sql & "Plan_E_Notas, Plan_E_Options, Plan_E_Preconisations, Plan_E_Vignettes,  Plan_L_cartouches, Plan_L_Composants, Plan_L_Connecteurs,  "
Sql = Sql & "Plan_L_Criteres, Plan_L_Etiquettes,  Plan_L_Fils, Plan_L_Noeuds, Plan_L_Notas, Plan_L_Options, Plan_L_Preconisations,  "
Sql = Sql & "Plan_L_Vignettes,  Plan_Ouvrir,Outil_E_Composants, Machine,Name )VALUES (" & Me.Tag & ", " & IdFils & ",'Modifier Plan', true, false, false,  false,  "
Sql = Sql & "false, false, false,  true, false, false, true,  false, false, false, false,  false, false, false, false, false,   "
Sql = Sql & "false, true, true, false, false,  false, false, false, false, false, true,  false, false, true, false,  false,  "
Sql = Sql & "false, false, false,  false, false, false, false, false, true,false, '" & MyReplace(UserName) & "','" & MyReplace(Me.txt7) & "' );"
 Con.Execute Sql
 MsgBox "Votre demande a été prise en compte vous pouvez suivre l'évolution de votre travail dans la fenêtre (Liste des JOB)"
 End If
End If
 
Me.Hide

End Sub

Private Sub CommandButton8_Click()
Me.Hide

End Sub

