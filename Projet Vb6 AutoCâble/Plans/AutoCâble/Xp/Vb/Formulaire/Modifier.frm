VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modifier 
   Caption         =   "Modifier :"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   Icon            =   "Modifier.dsx":0000
   OleObjectBlob   =   "Modifier.dsx":030A
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
CherchPicesAnnuler = CherchPices.Annuler
Unload CherchPices
If CherchPicesAnnuler = True Then Exit Sub
IdFils = 0
End Sub


Private Sub CommandButton2_Click()
Dim Piece As Long
Dim pathTmpXls As String
Dim Sql As String
Dim Rs As Recordset
Dim Fso As New FileSystemObject
Dim UserForm2_boolExcute As Boolean
Dim Planche_Clous_boolAnnuler As Boolean
If Trim("" & Me.txt3.Tag) = "" Then
    MsgBox "Le champ Pièce est obligatoire.", vbCritical, "Auto-Câble"
    CommandButton1_Click
    Exit Sub
End If
  If IdFils <> 0 Then Me.txt3.Tag = IdFils
Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
IdFils = 0
If Rs!Pere > 0 Then
IdFils = Me.txt3.Tag
    Me.txt3.Tag = Rs!Pere
End If
Set FormBarGrah = Me


If strStatus = "VAL" Then
If MsgBox("La Pièce: " & txt5 & " fait l 'objet d'une validation." & vbCrLf & "Voulez vous effectuer une copie en vue d'un changement d'indice", vbYesNo + vbQuestion, "Pièce déjà validée :") = vbNo Then Exit Sub
DoEvents
Approbateur = False
Useres.Charger "NULL", "Approbateur", True
Unload Useres

If Approbateur = False Then Exit Sub




    Sql = "INSERT INTO T_indiceProjet ( Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice,  "
    Sql = Sql & "Li, LI_Indice, PI, IdStatus,IdStatusSave, PI_Indice,   PlAutoCadSaveAs,  "
    Sql = Sql & "PlAutoCadSave, Client, Destinataire, Service, DessineDate, DessineNOM,   "
    Sql = Sql & "VerifieNom,  ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble,  "
    Sql = Sql & "CleAc, RefP, Masse, OuAutoCadSaveAs, OuAutoCadSave, Cartouche, Version ) "
    Sql = Sql & "SELECT T_indiceProjet.Id_Pieces, T_indiceProjet.Description, T_indiceProjet.PL,  "
    Sql = Sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU], T_indiceProjet.OU_Indice,  "
    Sql = Sql & "T_indiceProjet.Li, T_indiceProjet.LI_Indice, T_indiceProjet.PI,2,null,  "
    Sql = Sql & "T_indiceProjet.PI_Indice, "
    Sql = Sql & "T_indiceProjet.PlAutoCadSaveAs, T_indiceProjet.PlAutoCadSave, T_indiceProjet.Client,  "
    Sql = Sql & "T_indiceProjet.Destinataire, T_indiceProjet.Service, T_indiceProjet.DessineDate,  "
    Sql = Sql & "T_indiceProjet.DessineNOM, T_indiceProjet.VerifieNom,  "
    Sql = Sql & " T_indiceProjet.ApprouveNom, T_indiceProjet.Responsable,  "
    Sql = Sql & "T_indiceProjet.Vague, T_indiceProjet.Equipement, T_indiceProjet.RefPF,  "
    Sql = Sql & "T_indiceProjet.Ensemble, T_indiceProjet.CleAc, T_indiceProjet.RefP, T_indiceProjet.Masse,  "
    Sql = Sql & "T_indiceProjet.OuAutoCadSaveAs, T_indiceProjet.OuAutoCadSave,  "
    Sql = Sql & "T_indiceProjet.Cartouche, " & VersionPices(txt5.Caption) & " "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & CLng(Me.txt3.Tag) & ";"

    Con.Exequte Sql
    
    
 

    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = True "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & CLng(Me.txt3.Tag) & ";"


Con.Exequte Sql

  Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = True "
    Sql = Sql & "WHERE T_indiceProjet.pere=" & CLng(Me.txt3.Tag) & ";"
  Con.Exequte Sql
  
  Sql = "SELECT  [PI] & '_' & [PI_Indice] AS Piece  "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & CLng(Me.txt3.Tag) & ";"
Set Rs = Con.OpenRecordSet(Sql)
  
  Sql = "SELECT T_indiceProjet.Id "
Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE  [PI] & '_' & [PI_Indice] ='" & Replace(Rs!Piece, ":", "_", 1) & "' "
    Sql = Sql & "AND T_indiceProjet.Archiver=False "
    Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)

Sql = "INSERT INTO T_indiceProjet ( Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice,  "
    Sql = Sql & "Li, LI_Indice, PI, IdStatus,IdStatusSave, PI_Indice,   PlAutoCadSaveAs,  "
    Sql = Sql & "PlAutoCadSave, Client, Destinataire, Service, DessineDate, DessineNOM,   "
    Sql = Sql & "VerifieNom,  ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble,  "
    Sql = Sql & "CleAc, RefP, Masse, OuAutoCadSaveAs, OuAutoCadSave, Cartouche, Version,Pere ) "
    
    Sql = Sql & "SELECT T_indiceProjet.Id_Pieces, T_indiceProjet.Description, T_indiceProjet.PL,  "
    Sql = Sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU], T_indiceProjet.OU_Indice,  "
    Sql = Sql & "T_indiceProjet.Li, T_indiceProjet.LI_Indice, T_indiceProjet.PI,2,null,  "
    Sql = Sql & "T_indiceProjet.PI_Indice, "
    Sql = Sql & "T_indiceProjet.PlAutoCadSaveAs, T_indiceProjet.PlAutoCadSave, T_indiceProjet.Client,  "
    Sql = Sql & "T_indiceProjet.Destinataire, T_indiceProjet.Service, T_indiceProjet.DessineDate,  "
    Sql = Sql & "T_indiceProjet.DessineNOM, T_indiceProjet.VerifieNom,  "
    Sql = Sql & " T_indiceProjet.ApprouveNom, T_indiceProjet.Responsable,  "
    Sql = Sql & "T_indiceProjet.Vague, T_indiceProjet.Equipement, T_indiceProjet.RefPF,  "
    Sql = Sql & "T_indiceProjet.Ensemble, T_indiceProjet.CleAc, T_indiceProjet.RefP, T_indiceProjet.Masse,  "
    Sql = Sql & "T_indiceProjet.OuAutoCadSaveAs, T_indiceProjet.OuAutoCadSave,  "
    Sql = Sql & "T_indiceProjet.Cartouche, " & VersionPices(txt5.Caption) & "," & Rs!Id & " "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Pere=" & CLng(Me.txt3.Tag) & ";"


Con.Exequte Sql



Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, ACTIVER,CONNECTEUR, [O/N], DESIGNATION, CODE_APP, N°,  "
Sql = Sql & "POS, [POS-OUT], PRECO1, PRECO2, [100%] ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, Connecteurs.ACTIVER,Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION,  "
Sql = Sql & "Connecteurs.CODE_APP, Connecteurs.N°, Connecteurs.POS,  "
Sql = Sql & "Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.[100%] "
Sql = Sql & "FROM Connecteurs "
  Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Exequte Sql



Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet,ACTIVER, LIAI, DESIGNATION, FIL, SECT, TEINT,  "
Sql = Sql & "TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP, VOI, POS2, [POS-OUT2],  "
Sql = Sql & "FA2, APP2, VOI2, PRECO, [OPTION] ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, Ligne_Tableau_fils.ACTIVER,Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
Sql = Sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,  "
Sql = Sql & "Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2],  "
Sql = Sql & "Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
Sql = Sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Exequte Sql

Sql = "INSERT INTO Composants ( Id_IndiceProjet, ACTIVER,DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
Sql = Sql & "SELECT  " & Rs!Id & " AS Expr1, Composants.ACTIVER,Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.Path "
Sql = Sql & "FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"
Con.Exequte Sql

Sql = "INSERT INTO Nota ( Id_IndiceProjet, ACTIVER, NOTA, NUMNOTA ) SELECT  " & Rs!Id & " AS Expr1,  "
Sql = Sql & "Nota.ACTIVER,Nota.NOTA, Nota.NUMNOTA FROM Nota  "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Exequte Sql

Sql = "INSERT INTO T_Critères ( Id_IndiceProjet, ACTIVER,CODE_CRITERE, CRITERES ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1,T_Critères.ACTIVER, T_Critères.CODE_CRITERE, T_Critères.CRITERES "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Exequte Sql

Sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet, ACTIVER, NŒUDS, LONGUEUR, DESIGN_HAB, "
Sql = Sql & "CODE_RSA, CODE_PSA, CODE_ENC, DIAMETRE, CLASSE_T, TORON_PRINCIPAL, LONGUEUR_CUMULEE, Fleche_Droite )"
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, T_Noeuds.ACTIVER, T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.DESIGN_HAB, "
Sql = Sql & "T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T, "
Sql = Sql & "T_Noeuds.TORON_PRINCIPAL, T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.Fleche_Droite "
Sql = Sql & "FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Exequte Sql

Me.txt3.Tag = Rs!Id


End If
Me.Enabled = False

If OptionButton1.Value = True Then
    pathTmpXls = Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS"
     If Fso.FileExists(pathTmpXls) = True Then
        Fso.DeleteFile pathTmpXls
    End If
        
        ExporteXls pathTmpXls, CLng(Me.txt3.Tag)
        UserForm2.Chargement pathTmpXls, txt9
      UserForm2_boolExcute = UserForm2.boolExcute
       Unload UserForm2
        If UserForm2_boolExcute = False Then
            Me.Enabled = True

              Exit Sub
        End If
          MsgAutoCad = "Vos données ont bien été enregistrées, toute foie :" & vbCrLf & vbCrLf
        ImporteXls pathTmpXls, CLng(Me.txt3.Tag)
    End If
 If boolAutoCAD = False Then
   


    MsgBox MsgAutoCad & "Vous ne possédez pas de licence AutoCad." & vbCrLf & "Vous ne pouvez pas reporter vos modifications" & vbCrLf & "sur vos différents plans. "
Else
 Planche_Clous.Chargement CLng(Me.txt3.Tag)



    Sql = "SELECT T_Path.PathVar FROM T_Path WHERE T_Path.NameVar='PathOutils';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        RepPlacheClous = "" & Rs!PathVar
    End If
Set Rs = Con.CloseRecordSet(Rs)
    RepPlacheClous = RepPlacheClous & "\" & Planche_Clous.PlanchClous
PlanchClous = Planche_Clous.PlanchClous
Planche_Clous_boolAnnuler = Planche_Clous.boolAnnuler


Unload Planche_Clous
    If Planche_Clous_boolAnnuler = True Then
        Me.Enabled = True
        Exit Sub
    End If
    
    subDessinerPlan Me.txt3.Tag
    subDessinerOtil Me.txt3.Tag


 MsgBox "Fin du traitement" & vbCrLf & NbError & " Erreur(s) détectée(s) !"
 End If
NbError = 0
 Noquite = False
Me.Hide

End Sub

Private Sub CommandButton3_Click()
Noquite = False
Me.Hide
End Sub
Public Sub Charge(MyForm As Object)
Dim Sql As String
Dim Rs As Recordset
IdFils = 0
IdIndiceProjet = MyForm.IdIndiceProjet

Sql = "SELECT SelectProjets.* FROM SelectProjets WHERE SelectProjets.Id=" & IdIndiceProjet & " ;"

Set Rs = Con.OpenRecordSet(Sql)

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
 Me.Show vbModal
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
