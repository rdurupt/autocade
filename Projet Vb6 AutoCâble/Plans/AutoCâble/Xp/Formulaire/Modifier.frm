VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modifier 
   Caption         =   "Modifier :"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   OleObjectBlob   =   "Modifier.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdIndiceProjet As Long
Dim Id_Pere As Long
Dim Noquite As Boolean

Private Sub CommandButton1_Click()
CherchPices.Charge Me, "(VerifieDate= Null  and Archiver=False) OR (IdStatus=3 and Archiver=False)"
Unload CherchPices
End Sub

Private Sub CommandButton2_Click()
Dim Piece As Long
Dim pathTmpXls As String
Dim sql As String
Dim Rs As Recordset
Dim Fso As New FileSystemObject
Dim UserForm2_boolExcute As Boolean
Dim Planche_Clous_boolAnnuler As Boolean
If Trim("" & Me.txt3.Tag) = "" Then
    MsgBox "Le champ Pièce est obligatoire.", vbCritical, "Auto-Câble"
    CommandButton1_Click
    Exit Sub
End If
    
sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs!Pere > 0 Then Me.txt3.Tag = Rs!Pere
Set FormBarGrah = Me


If strStatus = "VAL" Then
If MsgBox("La Pièce: " & txt5 & " fait l 'objet d'une validation." & vbCrLf & "Voulez vous effectuer une copie en vue d'un changement d'indice", vbYesNo + vbQuestion, "Pièce déjà validée :") = vbNo Then Exit Sub








    sql = "INSERT INTO T_indiceProjet ( Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice,  "
    sql = sql & "Li, LI_Indice, PI, IdStatus,IdStatusSave, PI_Indice,   PlAutoCadSaveAs,  "
    sql = sql & "PlAutoCadSave, Client, Destinataire, Service, DessineDate, DessineNOM,   "
    sql = sql & "VerifieNom,  ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble,  "
    sql = sql & "CleAc, RefP, Masse, OuAutoCadSaveAs, OuAutoCadSave, Cartouche, Version ) "
    sql = sql & "SELECT T_indiceProjet.Id_Pieces, T_indiceProjet.Description, T_indiceProjet.PL,  "
    sql = sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU], T_indiceProjet.OU_Indice,  "
    sql = sql & "T_indiceProjet.Li, T_indiceProjet.LI_Indice, T_indiceProjet.PI,2,null,  "
    sql = sql & "T_indiceProjet.PI_Indice, "
    sql = sql & "T_indiceProjet.PlAutoCadSaveAs, T_indiceProjet.PlAutoCadSave, T_indiceProjet.Client,  "
    sql = sql & "T_indiceProjet.Destinataire, T_indiceProjet.Service, T_indiceProjet.DessineDate,  "
    sql = sql & "T_indiceProjet.DessineNOM, T_indiceProjet.VerifieNom,  "
    sql = sql & " T_indiceProjet.ApprouveNom, T_indiceProjet.Responsable,  "
    sql = sql & "T_indiceProjet.Vague, T_indiceProjet.Equipement, T_indiceProjet.RefPF,  "
    sql = sql & "T_indiceProjet.Ensemble, T_indiceProjet.CleAc, T_indiceProjet.RefP, T_indiceProjet.Masse,  "
    sql = sql & "T_indiceProjet.OuAutoCadSaveAs, T_indiceProjet.OuAutoCadSave,  "
    sql = sql & "T_indiceProjet.Cartouche, " & VersionPices(txt5.Caption) & " "
    sql = sql & "FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & CLng(Me.txt3.Tag) & ";"

    Con.Exequte sql
    
    
 

    sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = True "
    sql = sql & "WHERE T_indiceProjet.Id=" & CLng(Me.txt3.Tag) & ";"


Con.Exequte sql

  sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = True "
    sql = sql & "WHERE T_indiceProjet.pere=" & CLng(Me.txt3.Tag) & ";"
  Con.Exequte sql
  
  sql = "SELECT  [PI] & '_' & [PI_Indice] AS Piece  "
    sql = sql & "FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & CLng(Me.txt3.Tag) & ";"
Set Rs = Con.OpenRecordSet(sql)
  
  sql = "SELECT T_indiceProjet.Id "
sql = sql & "FROM T_indiceProjet "
    sql = sql & "WHERE  [PI] & '_' & [PI_Indice] ='" & Replace(Rs!Piece, ":", "_", 1) & "' "
    sql = sql & "AND T_indiceProjet.Archiver=False "
    sql = sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(sql)

sql = "INSERT INTO T_indiceProjet ( Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice,  "
    sql = sql & "Li, LI_Indice, PI, IdStatus,IdStatusSave, PI_Indice,   PlAutoCadSaveAs,  "
    sql = sql & "PlAutoCadSave, Client, Destinataire, Service, DessineDate, DessineNOM,   "
    sql = sql & "VerifieNom,  ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble,  "
    sql = sql & "CleAc, RefP, Masse, OuAutoCadSaveAs, OuAutoCadSave, Cartouche, Version,Pere ) "
    
    sql = sql & "SELECT T_indiceProjet.Id_Pieces, T_indiceProjet.Description, T_indiceProjet.PL,  "
    sql = sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU], T_indiceProjet.OU_Indice,  "
    sql = sql & "T_indiceProjet.Li, T_indiceProjet.LI_Indice, T_indiceProjet.PI,2,null,  "
    sql = sql & "T_indiceProjet.PI_Indice, "
    sql = sql & "T_indiceProjet.PlAutoCadSaveAs, T_indiceProjet.PlAutoCadSave, T_indiceProjet.Client,  "
    sql = sql & "T_indiceProjet.Destinataire, T_indiceProjet.Service, T_indiceProjet.DessineDate,  "
    sql = sql & "T_indiceProjet.DessineNOM, T_indiceProjet.VerifieNom,  "
    sql = sql & " T_indiceProjet.ApprouveNom, T_indiceProjet.Responsable,  "
    sql = sql & "T_indiceProjet.Vague, T_indiceProjet.Equipement, T_indiceProjet.RefPF,  "
    sql = sql & "T_indiceProjet.Ensemble, T_indiceProjet.CleAc, T_indiceProjet.RefP, T_indiceProjet.Masse,  "
    sql = sql & "T_indiceProjet.OuAutoCadSaveAs, T_indiceProjet.OuAutoCadSave,  "
    sql = sql & "T_indiceProjet.Cartouche, " & VersionPices(txt5.Caption) & "," & Rs!Id & " "
    sql = sql & "FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Pere=" & CLng(Me.txt3.Tag) & ";"


Con.Exequte sql



sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR, [O/N], DESIGNATION, CODE_APP, N°,  "
sql = sql & "POS, [POS-OUT], PRECO1, PRECO2, [100%] ) "
sql = sql & "SELECT " & Rs!Id & " AS Expr1, Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION,  "
sql = sql & "Connecteurs.CODE_APP, Connecteurs.N°, Connecteurs.POS,  "
sql = sql & "Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.[100%] "
sql = sql & "FROM Connecteurs "
  sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Exequte sql



sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION, FIL, SECT, TEINT,  "
sql = sql & "TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP, VOI, POS2, [POS-OUT2],  "
sql = sql & "FA2, APP2, VOI2, PRECO, [OPTION] ) "
sql = sql & "SELECT " & Rs!Id & " AS Expr1, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
sql = sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,  "
sql = sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
sql = sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
sql = sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,  "
sql = sql & "Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2],  "
sql = sql & "Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
sql = sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
sql = sql & "FROM Ligne_Tableau_fils "
sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Exequte sql
Me.txt3.Tag = Rs!Id
End If

If OptionButton1.Value = True Then
    pathTmpXls = Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt6.Caption, ":", "_", 1) & ".XLS"
     If Fso.FileExists(pathTmpXls) = True Then
        Fso.DeleteFile pathTmpXls
    End If
        
        ExporteXls pathTmpXls, CLng(Me.txt3.Tag)
        UserForm2.chargement pathTmpXls, txt9
      UserForm2_boolExcute = UserForm2.boolExcute
       Unload UserForm2
        If UserForm2_boolExcute = False Then
              Exit Sub
        End If
          
        ImporteXls pathTmpXls, CLng(Me.txt3.Tag)
    End If
    
Planche_Clous.Show vbModal



sql = "SELECT T_Path.PathVar FROM T_Path WHERE T_Path.NameVar='PathOutils';"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
    RepPlacheClous = "" & Rs!PathVar
End If
Set Rs = Con.CloseRecordSet(Rs)

RepPlacheClous = RepPlacheClous & "\" & Planche_Clous.PlanchClous
PlanchClous = Planche_Clous.PlanchClous
Planche_Clous_boolAnnuler = Planche_Clous.boolAnnuler
Unload Planche_Clous

If Planche_Clous_boolAnnuler = True Then Exit Sub

subDessinerPlan Me.txt3.Tag
subDessinerOtil Me.txt3.Tag
 Noquite = False
Me.Hide

End Sub

Private Sub CommandButton3_Click()
Noquite = False
Me.Hide
End Sub
Public Sub Charge(MyForm As Object)
Dim sql As String
Dim Rs As Recordset
IdIndiceProjet = MyForm.IdIndiceProjet

sql = "SELECT SelectProjets.* FROM SelectProjets WHERE SelectProjets.Id=" & IdIndiceProjet & " ;"

Set Rs = Con.OpenRecordSet(sql)

Set FormBarGrah = Me
If Rs.EOF = False Then
For i = 0 To 11
    Me.Controls("txt" & CStr(i + 1)) = "" & Rs(i)
     Me.Controls("txt" & CStr(i + 1)).Tag = "" & Rs.Fields(12)

Next i
    
    OptionButton2.Value = True
    OptionButton1.Value = False
    Me.CommandButton1.Enabled = True
 End If
 Set Rs = Con.CloseRecordSet(Rs)
 MyForm.Hide
 Me.Show
End Sub

Private Sub OptionButton1_Click()
OptionButton2.Value = False
End Sub

Private Sub OptionButton2_Click()
OptionButton1.Value = False
End Sub

Private Sub UserForm_Activate()
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
