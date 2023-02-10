VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modifier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modifier :"
   ClientHeight    =   5415
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9120
   Icon            =   "Modifier.dsx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   OleObjectBlob   =   "Modifier.dsx":030A
End
Attribute VB_Name = "Modifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdIndiceProjet As Long
Dim Id_Pere As Long
Public Noquite As Boolean
Dim Qui As String
Public BooolBloque As Boolean
Dim NewUserForm2 As UserForm2
Dim H As Double
Dim W As Double
Dim L As Double



Private Sub CommandButton1_Click()
If BooolBloque = False Then
CherchPices.Charge Me, "(VerifieDate= Null  and Archiver=false) OR (IdStatus<4  and Archiver=false)"
Else
    CherchPices.Charge Me, "IdStatus<>4"
End If
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
Dim FileOu As String
Dim FileLi As String
Dim FilePL As String
Dim UserForm2_boolExcute As Boolean
Dim MyVersion As Long
Dim Planche_Clous_boolAnnuler As Boolean
Dim PassWordOk As Boolean
 Set TableauPath = funPath
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



If strStatus = "VAL" Then
If BooolBloque = False Then
If MsgBox("La Pièce: " & txt5 & " fait l 'objet d'une validation." & vbCrLf & "Voulez vous effectuer une copie en vue d'un changement d'indice", vbYesNo + vbQuestion, "Pièce déjà validée :") = vbNo Then Exit Sub
Sql = "SELECT T_Users.Id AS Id_Users, T_Droits.Id_Bouton, T_Users.Id, T_Users.Cloturer  "
Sql = Sql & "FROM T_Boutons INNER JOIN (T_Droits LEFT JOIN T_Users ON T_Droits.Id_Useur = T_Users.Id) ON  "
Sql = Sql & "T_Boutons.Id = T_Droits.Id_Bouton "
Sql = Sql & "WHERE T_Users.Id=" & Id_Users & " AND T_Users.Cloturer=False  "
Sql = Sql & "AND T_Boutons.Name='CommandButton1';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
DoEvents
PassWordOk = False
Useres.charger "CommandButton1", ModuleOk:=True
PassWordOk = Useres.DroitsOk
Unload Useres
Else
    PassWordOk = True
End If
If PassWordOk = False Then Exit Sub




    Sql = "INSERT INTO T_indiceProjet (  RefPieceClient, Ref_PF, Ref_Piece_CLI, ReffIndice,  "
    Sql = Sql & "Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice, Li, LI_Indice, PI, IdStatus,  "
    Sql = Sql & "IdStatusSave, PI_Indice, PlAutoCadSave, LiAutoCadSave, OuAutoCadSave,  "
    Sql = Sql & "Client, Destinataire, Service, DessineDate, DessineNOM, VerifieNom,  "
    Sql = Sql & "ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble, CleAc,  "
    Sql = Sql & "RefP, Masse, Cartouche, Version ) "
    Sql = Sql & "SELECT T_indiceProjet.RefPieceClient, T_indiceProjet.Ref_PF, T_indiceProjet.Ref_Piece_CLI,  "
    Sql = Sql & "T_indiceProjet.ReffIndice, T_indiceProjet.Id_Pieces, T_indiceProjet.Description,  "
    Sql = Sql & "T_indiceProjet.PL, T_indiceProjet.PL_Indice, T_indiceProjet.[OU],  "
    Sql = Sql & "T_indiceProjet.OU_Indice, T_indiceProjet.Li, T_indiceProjet.LI_Indice,  "
    Sql = Sql & "T_indiceProjet.PI, 2 AS Expr1, Null AS Expr2, T_indiceProjet.PI_Indice,  "
    Sql = Sql & "T_indiceProjet.PlAutoCadSave, T_indiceProjet.LiAutoCadSave,  "
    Sql = Sql & "T_indiceProjet.OuAutoCadSave, T_indiceProjet.Client, T_indiceProjet.Destinataire,  "
    Sql = Sql & "T_indiceProjet.Service, T_indiceProjet.DessineDate, T_indiceProjet.DessineNOM,  "
    Sql = Sql & "T_indiceProjet.VerifieNom, T_indiceProjet.ApprouveNom, T_indiceProjet.Responsable,  "
    Sql = Sql & "T_indiceProjet.Vague, T_indiceProjet.Equipement, T_indiceProjet.RefPF,  "
    Sql = Sql & "T_indiceProjet.Ensemble, T_indiceProjet.CleAc, T_indiceProjet.RefP,  "
    MyVersion = 2
    Sql = Sql & "T_indiceProjet.Masse, T_indiceProjet.Cartouche," & MyVersion & " "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & CLng(Me.txt3.Tag) & ";"

    Con.Execute Sql
    
    
 

    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = True "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & CLng(Me.txt3.Tag) & ";"


Con.Execute Sql



   
  Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = True "
    Sql = Sql & "WHERE T_indiceProjet.pere=" & CLng(Me.txt3.Tag) & ";"
  Con.Execute Sql
  
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

'Sql = "INSERT INTO T_indiceProjet ( Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice,  "
'    Sql = Sql & "Li, LI_Indice, PI, IdStatus,IdStatusSave, PI_Indice,   PlAutoCadSaveAs,  "
'    Sql = Sql & "PlAutoCadSave, Client, Destinataire, Service, DessineDate, DessineNOM,   "
'    Sql = Sql & "VerifieNom,  ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble,  "
'    Sql = Sql & "CleAc, RefP, Masse, OuAutoCadSaveAs, OuAutoCadSave, Cartouche, Version,Pere ) "
'
'    Sql = Sql & "SELECT T_indiceProjet.Id_Pieces, T_indiceProjet.Description, T_indiceProjet.PL,  "
'    Sql = Sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU], T_indiceProjet.OU_Indice,  "
'    Sql = Sql & "T_indiceProjet.Li, T_indiceProjet.LI_Indice, T_indiceProjet.PI,2,null,  "
'    Sql = Sql & "T_indiceProjet.PI_Indice, "
'    Sql = Sql & "T_indiceProjet.PlAutoCadSaveAs, T_indiceProjet.PlAutoCadSave, T_indiceProjet.Client,  "
'    Sql = Sql & "T_indiceProjet.Destinataire, T_indiceProjet.Service, T_indiceProjet.DessineDate,  "
'    Sql = Sql & "T_indiceProjet.DessineNOM, T_indiceProjet.VerifieNom,  "
'    Sql = Sql & " T_indiceProjet.ApprouveNom, T_indiceProjet.Responsable,  "
'    Sql = Sql & "T_indiceProjet.Vague, T_indiceProjet.Equipement, T_indiceProjet.RefPF,  "
'    Sql = Sql & "T_indiceProjet.Ensemble, T_indiceProjet.CleAc, T_indiceProjet.RefP, T_indiceProjet.Masse,  "
'    Sql = Sql & "T_indiceProjet.OuAutoCadSaveAs, T_indiceProjet.OuAutoCadSave,  "
'    Sql = Sql & "T_indiceProjet.Cartouche, " & VersionPices(txt5.Caption) & "," & rs!Id & " "
'    Sql = Sql & "FROM T_indiceProjet "
'    Sql = Sql & "WHERE T_indiceProjet.Pere=" & CLng(Me.txt3.Tag) & ";"
    
    
     Sql = "INSERT INTO T_indiceProjet (  RefPieceClient, Ref_PF, Ref_Piece_CLI, ReffIndice,  "
    Sql = Sql & "Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice, Li, LI_Indice, PI, IdStatus,  "
    Sql = Sql & "IdStatusSave, PI_Indice, PlAutoCadSave, LiAutoCadSave, OuAutoCadSave,  "
    Sql = Sql & "Client, Destinataire, Service, DessineDate, DessineNOM, VerifieNom,  "
    Sql = Sql & "ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble, CleAc,  "
    Sql = Sql & "RefP, Masse, Cartouche, Version,Pere ) "
    Sql = Sql & "SELECT T_indiceProjet.RefPieceClient, T_indiceProjet.Ref_PF, T_indiceProjet.Ref_Piece_CLI,  "
    Sql = Sql & "T_indiceProjet.ReffIndice, T_indiceProjet.Id_Pieces, T_indiceProjet.Description,  "
    Sql = Sql & "T_indiceProjet.PL, T_indiceProjet.PL_Indice, T_indiceProjet.[OU],  "
    Sql = Sql & "T_indiceProjet.OU_Indice, T_indiceProjet.Li, T_indiceProjet.LI_Indice,  "
    Sql = Sql & "T_indiceProjet.PI, 2 AS Expr1, Null AS Expr2, T_indiceProjet.PI_Indice,  "
    Sql = Sql & "T_indiceProjet.PlAutoCadSave, T_indiceProjet.LiAutoCadSave,  "
    Sql = Sql & "T_indiceProjet.OuAutoCadSave, T_indiceProjet.Client, T_indiceProjet.Destinataire,  "
    Sql = Sql & "T_indiceProjet.Service, T_indiceProjet.DessineDate, T_indiceProjet.DessineNOM,  "
    Sql = Sql & "T_indiceProjet.VerifieNom, T_indiceProjet.ApprouveNom, T_indiceProjet.Responsable,  "
    Sql = Sql & "T_indiceProjet.Vague, T_indiceProjet.Equipement, T_indiceProjet.RefPF,  "
    Sql = Sql & "T_indiceProjet.Ensemble, T_indiceProjet.CleAc, T_indiceProjet.RefP,  "
    Sql = Sql & "T_indiceProjet.Masse, T_indiceProjet.Cartouche," & MyVersion & "," & Rs!Id & " "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Pere=" & CLng(Me.txt3.Tag) & ";"


Con.Execute Sql


'
'Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, ACTIVER,CONNECTEUR, [O/N], DESIGNATION, CODE_APP, N°,  "
'Sql = Sql & "POS, [POS-OUT], PRECO1, PRECO2, [100%] ) "
'Sql = Sql & "SELECT " & rs!Id & " AS Expr1, Connecteurs.ACTIVER,Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION,  "
'Sql = Sql & "Connecteurs.CODE_APP, Connecteurs.N°, Connecteurs.POS,  "
'Sql = Sql & "Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.[100%] "
'Sql = Sql & "FROM Connecteurs "
'  Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

'sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR, [O/N], DESIGNATION, CODE_APP, N°, "
'sql = sql & "POS, [POS-OUT], PRECO1, PRECO2, [100%], [OPTION], Pylone, Colonne, Ligne, ACTIVER, "
'sql = sql & "RefBouchon, RefBouchonFour, ReFCapot, ReFCapotFour, RefVerrou, RefVerrouFour, "
'sql = sql & "RefConnecteurFour, LongueurF_Choix )"
'sql = sql & "SELECT " & Rs!Id & "  AS Expr1, Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, "
'sql = sql & "Connecteurs.CODE_APP, Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT], "
'sql = sql & "Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.[100%], Connecteurs.OPTION, "
'sql = sql & "Connecteurs.Pylone, Connecteurs.Colonne, Connecteurs.Ligne, Connecteurs.ACTIVER, "
'sql = sql & "Connecteurs.RefBouchon, Connecteurs.RefBouchonFour, Connecteurs.ReFCapot, "
'sql = sql & "Connecteurs.ReFCapotFour, Connecteurs.RefVerrou, Connecteurs.RefVerrouFour, "
'sql = sql & "Connecteurs.RefConnecteurFour, Connecteurs.LongueurF_Choix "
'sql = sql & "FROM Connecteurs "
'sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"


Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, ACTIVER, CONNECTEUR, RefConnecteurFour, [O/N],  "
Sql = Sql & "DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2, [OPTION], [100%], Pylone, Colonne,  "
Sql = Sql & "Ligne, RefBouchon, RefBouchonFour, ReFCapot, ReFCapotFour, RefVerrou, RefVerrouFour,  "
Sql = Sql & "LongueurF_Choix ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, Connecteurs.ACTIVER, Connecteurs.CONNECTEUR,  "
Sql = Sql & "Connecteurs.RefConnecteurFour, Connecteurs.[O/N], Connecteurs.DESIGNATION,  "
Sql = Sql & "Connecteurs.CODE_APP, Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT],  "
Sql = Sql & "Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.OPTION, Connecteurs.[100%],  "
Sql = Sql & "Connecteurs.Pylone, Connecteurs.Colonne, Connecteurs.Ligne, Connecteurs.RefBouchon,  "
Sql = Sql & "Connecteurs.RefBouchonFour, Connecteurs.ReFCapot, Connecteurs.ReFCapotFour,  "
Sql = Sql & "Connecteurs.RefVerrou, Connecteurs.RefVerrouFour, Connecteurs.LongueurF_Choix "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"


Con.Execute Sql



Sql = "INSERT INTO Ligne_Tableau_fils (  "
Sql = Sql & "Id_IndiceProjet,  "
Sql = Sql & "LIAI,  "
Sql = Sql & "DESIGNATION,  "
Sql = Sql & "FIL,  "
Sql = Sql & "SECT,  "
Sql = Sql & "TEINT,  "
Sql = Sql & "TEINT2,   "
Sql = Sql & "ISO,  "
Sql = Sql & "[LONG],  "
Sql = Sql & "[LONG CP],  "
Sql = Sql & "COUPE, POS,  "
Sql = Sql & "[POS-OUT],  "
Sql = Sql & "FA,  "
Sql = Sql & "APP,  "
Sql = Sql & "VOI,  "
Sql = Sql & "[Ref Connecteur],   "
Sql = Sql & "[Ref Connecteur_Four],  "
Sql = Sql & "Long_Add,  "
Sql = Sql & "[Ref Clip],  "
Sql = Sql & "[Ref Clip Four],  "
Sql = Sql & "[Ref Joint],  "
Sql = Sql & "[Ref Joint four],   "
Sql = Sql & "POS2,  "
Sql = Sql & "[POS-OUT2],  "
Sql = Sql & "FA2,  "
Sql = Sql & "APP2,  "
Sql = Sql & "VOI2,  "
Sql = Sql & "[Ref Connecteur2],   "
Sql = Sql & "[Ref Connecteur_Four2],  "
Sql = Sql & "Long_Add2,  "
Sql = Sql & "[Ref Clip2],  "
Sql = Sql & "[Ref Clip Four2],  "
Sql = Sql & "[Ref Joint2],   "
Sql = Sql & "[Ref Joint Four2],  "

'sql = sql & "PRECOG, "
'sql = sql & "PRECO2, PRECO,   "
'sql = sql & "[OPTION],  "
'sql = sql & "ACTIVER,  "
'sql = sql & "[Critères spécifiques] )  "
'sql = sql & "SELECT " & Rs!Id & " AS Expr1, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,   "
'sql = sql & "Ligne_Tableau_fils.FIL,Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,   "
'sql = sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,   "
'sql = sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,   "
'sql = sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,   "
'sql = sql & "Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.[Ref Connecteur],   "
'sql = sql & "Ligne_Tableau_fils.[Ref Connecteur_Four], Ligne_Tableau_fils.Long_Add,   "
'sql = sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four],   "
'sql = sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint four],   "
'sql = sql & " Ligne_Tableau_fils.POS2,   "
'sql = sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2,   "
'sql = sql & "Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.[Ref Connecteur2], "
'sql = sql & "Ligne_Tableau_fils.[Ref Connecteur_Four2], Ligne_Tableau_fils.Long_Add2,   "
'sql = sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],   "
'sql = sql & "Ligne_Tableau_fils.[Ref Joint2], Ligne_Tableau_fils.[Ref Joint Four2],   "
'sql = sql & " Ligne_Tableau_fils.PRECOG,   "
'sql = sql & " Ligne_Tableau_fils.PRECO2, Ligne_Tableau_fils.PRECO,   "
'sql = sql & "Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.ACTIVER,   "
'sql = sql & "Ligne_Tableau_fils.[Critères spécifiques]  "
'sql = sql & "FROM Ligne_Tableau_fils "
'sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"


Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, ACTIVER, LIAI, DESIGNATION, FIL, SECT, TEINT,   "
Sql = Sql & "TEINT2, ISO, [LONG], [LONG CP], Long_Add, Long_Add2, COUPE, POS, [POS-OUT], FA, APP, VOI,   "
Sql = Sql & "[Ref Connecteur], [Ref Connecteur_Four], [Ref Clip], [Ref Clip Four], PRECO, [Ref Joint],   "
Sql = Sql & "[Ref Joint four], POS2, [POS-OUT2], FA2, APP2, VOI2, [Ref Connecteur2],   "
Sql = Sql & "[Ref Connecteur_Four2], [Ref Clip2], [Ref Clip Four2], PRECO2, [Ref Joint2],   "
Sql = Sql & "[Ref Joint Four2], PRECOG, [OPTION], [Critères spécifiques] )  "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI,   "
Sql = Sql & "Ligne_Tableau_fils.DESIGNATION,Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT,   "
Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,   "
Sql = Sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.Long_Add,   "
Sql = Sql & "Ligne_Tableau_fils.Long_Add2, Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,   "
Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,   "
Sql = Sql & "Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.[Ref Connecteur],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Connecteur_Four], Ligne_Tableau_fils.[Ref Clip],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip Four], Ligne_Tableau_fils.PRECO,   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint four],   "
Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,   "
Sql = Sql & "Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Connecteur2], Ligne_Tableau_fils.[Ref Connecteur_Four2],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],   "
Sql = Sql & "Ligne_Tableau_fils.PRECO2, Ligne_Tableau_fils.[Ref Joint2],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint Four2], Ligne_Tableau_fils.PRECOG,   "
Sql = Sql & "Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.[Critères spécifiques]  "
Sql = Sql & "FROM Ligne_Tableau_fils  "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"


Con.Execute Sql

'sql = "INSERT INTO NomeclatureConnecteurs ( Id_IndiceProjet, App, Designation, Connecteur,  "
'sql = sql & "Connecteur_Four, Liaison, Voie, Long_Add, Famille, Bouchon, Capot, Capot_Four,  "
'sql = sql & "Verrou, Verrout_Four, Options, Clip, ClipFour, Joint, JointFour ) "
'sql = sql & "SELECT " & Rs!Id & "  AS Expr1, NomeclatureConnecteurs.App, NomeclatureConnecteurs.Designation,  "
'sql = sql & "NomeclatureConnecteurs.Connecteur, NomeclatureConnecteurs.Connecteur_Four,  "
'sql = sql & "NomeclatureConnecteurs.Liaison, NomeclatureConnecteurs.Voie, NomeclatureConnecteurs.Long_Add, "
'sql = sql & "NomeclatureConnecteurs.Famille, NomeclatureConnecteurs.Bouchon, NomeclatureConnecteurs.Capot,  "
'sql = sql & "NomeclatureConnecteurs.Capot_Four, NomeclatureConnecteurs.Verrou,  "
'sql = sql & "NomeclatureConnecteurs.Verrout_Four, NomeclatureConnecteurs.Options,  "
'sql = sql & "NomeclatureConnecteurs.Clip, NomeclatureConnecteurs.ClipFour, NomeclatureConnecteurs.Joint,  "
'sql = sql & "NomeclatureConnecteurs.JointFour "
'sql = sql & "FROM NomeclatureConnecteurs "
'sql = sql & "WHERE NomeclatureConnecteurs.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Sql = "INSERT INTO NomeclatureConnecteurs ( Id_IndiceProjet, App, Designation, Connecteur,  "
Sql = Sql & "Connecteur_Four, Liaison, SECT, TEINT, TEINT2, ISO, Voie, [LONG], COUPE, [LONG CP],  "
Sql = Sql & "Long_Add, Famille, Bouchon, Capot, Capot_Four, Verrou, Verrout_Four, Options, Clip,  "
Sql = Sql & "ClipFour, Joint, JointFour ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, NomeclatureConnecteurs.App, NomeclatureConnecteurs.Designation,  "
Sql = Sql & "NomeclatureConnecteurs.Connecteur, NomeclatureConnecteurs.Connecteur_Four,  "
Sql = Sql & "NomeclatureConnecteurs.Liaison, NomeclatureConnecteurs.SECT, NomeclatureConnecteurs.TEINT,  "
Sql = Sql & "NomeclatureConnecteurs.TEINT2, NomeclatureConnecteurs.ISO, NomeclatureConnecteurs.Voie, "
Sql = Sql & " NomeclatureConnecteurs.LONG, NomeclatureConnecteurs.COUPE,  "
Sql = Sql & "NomeclatureConnecteurs.[LONG CP], NomeclatureConnecteurs.Long_Add,  "
Sql = Sql & "NomeclatureConnecteurs.Famille, NomeclatureConnecteurs.Bouchon,  "
Sql = Sql & "NomeclatureConnecteurs.Capot, NomeclatureConnecteurs.Capot_Four,  "
Sql = Sql & "NomeclatureConnecteurs.Verrou, NomeclatureConnecteurs.Verrout_Four,  "
Sql = Sql & "NomeclatureConnecteurs.Options, NomeclatureConnecteurs.Clip,  "
Sql = Sql & "NomeclatureConnecteurs.ClipFour, NomeclatureConnecteurs.Joint,  "
Sql = Sql & "NomeclatureConnecteurs.JointFour "
Sql = Sql & "FROM NomeclatureConnecteurs "
Sql = Sql & "WHERE NomeclatureConnecteurs.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"


Con.Execute Sql

'sql = "INSERT INTO NomenclaturFinal ( Id_IndiceProjet, Designation, Ref, RefFour, Qts, Info_Lib1, Info1,  "
'sql = sql & "Info_Lib2, Info2, Info_Lib3, Info3, Options ) "
'sql = sql & "SELECT " & Rs!Id & " AS Expr1, NomenclaturFinal.Designation, NomenclaturFinal.Ref,  "
'sql = sql & "NomenclaturFinal.RefFour, NomenclaturFinal.Qts, NomenclaturFinal.Info_Lib1,  "
'sql = sql & "NomenclaturFinal.Info1, NomenclaturFinal.Info_Lib2, NomenclaturFinal.Info2,  "
'sql = sql & "NomenclaturFinal.Info_Lib3, NomenclaturFinal.Info3, NomenclaturFinal.Options "
'sql = sql & "FROM NomenclaturFinal "
'sql = sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"


Sql = "INSERT INTO NomenclaturFinal ( Id_IndiceProjet, Designation, Famille, Fournisseur, Ref,  "
Sql = Sql & "RefFour, Qts, ISO, TEINT, TEINT2, SECT, Qts_Encelade, Qts_E_Boutique, Qts_Appro,  "
Sql = Sql & "Prix_Revient, Prix_Vente, Options ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, NomenclaturFinal.Designation, NomenclaturFinal.Famille,  "
Sql = Sql & "NomenclaturFinal.Fournisseur, NomenclaturFinal.Ref, NomenclaturFinal.RefFour,  "
Sql = Sql & "NomenclaturFinal.Qts, NomenclaturFinal.ISO, NomenclaturFinal.TEINT,  "
Sql = Sql & "NomenclaturFinal.TEINT2, NomenclaturFinal.SECT, NomenclaturFinal.Qts_Encelade,  "
Sql = Sql & "NomenclaturFinal.Qts_E_Boutique, NomenclaturFinal.Qts_Appro,  "
Sql = Sql & "NomenclaturFinal.Prix_Revient, NomenclaturFinal.Prix_Vente, NomenclaturFinal.Options "
Sql = Sql & "FROM NomenclaturFinal "
Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"


Con.Execute Sql


'sql = "INSERT INTO Composants ( Id_IndiceProjet, ACTIVER,DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
'sql = sql & "SELECT  " & Rs!Id & " AS Expr1, Composants.ACTIVER,Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.Path "
'sql = sql & "FROM Composants "
'sql = sql & "WHERE Composants.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Sql = "INSERT INTO Composants ( Id_IndiceProjet, DESIGNCOMP, NUMCOMP, REFCOMP, Code_APP_Lier, Voie,  "
Sql = Sql & "POS, [POS-OUT], Path, ACTIVER, [OPTION] ) "
Sql = Sql & "SELECT " & Rs!Id & "  AS Expr1, Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP,  "
Sql = Sql & "Composants.Code_APP_Lier, Composants.Voie, Composants.POS, Composants.[POS-OUT],  "
Sql = Sql & "Composants.Path, Composants.ACTIVER, Composants.OPTION "
Sql = Sql & "FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"
Con.Execute Sql

'sql = "INSERT INTO Nota ( Id_IndiceProjet, ACTIVER, NOTA, NUMNOTA ) SELECT  " & Rs!Id & " AS Expr1,  "
'sql = sql & "Nota.ACTIVER,Nota.NOTA, Nota.NUMNOTA FROM Nota  "
'sql = sql & "WHERE Nota.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Sql = "INSERT INTO Nota ( Id_IndiceProjet, NOTA, NUMNOTA, ACTIVER, [OPTION] ) "
Sql = Sql & "SELECT " & Rs!Id & "  AS Expr1, Nota.NOTA, Nota.NUMNOTA, Nota.ACTIVER, Nota.OPTION "
Sql = Sql & "FROM Nota "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"
Con.Execute Sql

Sql = "INSERT INTO T_Critères ( Id_IndiceProjet, ACTIVER,CODE_CRITERE, CRITERES ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1,T_Critères.ACTIVER, T_Critères.CODE_CRITERE, T_Critères.CRITERES "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Sql = "INSERT INTO T_Critères ( Id_IndiceProjet, ACTIVER, CODE_CRITERE, CRITERES, DESIGNATION ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, T_Critères.ACTIVER, T_Critères.CODE_CRITERE, T_Critères.CRITERES,  "
Sql = Sql & "T_Critères.DESIGNATION "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Execute Sql
'
'sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet, ACTIVER, NŒUDS, LONGUEUR, DESIGN_HAB, "
'sql = sql & "CODE_RSA, CODE_PSA, CODE_ENC, DIAMETRE, CLASSE_T, TORON_PRINCIPAL, LONGUEUR_CUMULEE, Fleche_Droite )"
'sql = sql & "SELECT " & Rs!Id & " AS Expr1, T_Noeuds.ACTIVER, T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.DESIGN_HAB, "
'sql = sql & "T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T, "
'sql = sql & "T_Noeuds.TORON_PRINCIPAL, T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.Fleche_Droite "
'sql = sql & "FROM T_Noeuds "
'sql = sql & "WHERE T_Noeuds.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet, ACTIVER, NŒUDS, LONGUEUR, DESIGN_HAB, CODE_RSA, CODE_PSA,  "
Sql = Sql & "CODE_ENC, DIAMETRE, CLASSE_T, TORON_PRINCIPAL, LONGUEUR_CUMULEE, Fleche_Droite, [OPTION] ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, T_Noeuds.ACTIVER, T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.DESIGN_HAB,  "
Sql = Sql & "T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE,  "
Sql = Sql & "T_Noeuds.CLASSE_T, T_Noeuds.TORON_PRINCIPAL, T_Noeuds.LONGUEUR_CUMULEE,  "
Sql = Sql & "T_Noeuds.Fleche_Droite, T_Noeuds.OPTION "
Sql = Sql & "FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"

Con.Execute Sql

Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, LIAI, Designation, App, Voie, Ref, Fournisseur,  "
Sql = Sql & "RefFour, App2, Voie2, Options, ISO, Longueur, [Longueur Total], TEINT, TEINT2, SECT, Qts ) "
Sql = Sql & "SELECT " & Rs!Id & "  AS Expr1, Nomenclature2.LIAI, Nomenclature2.Designation, Nomenclature2.App,  "
Sql = Sql & "Nomenclature2.Voie, Nomenclature2.Ref, Nomenclature2.Fournisseur, Nomenclature2.RefFour,  "
Sql = Sql & "Nomenclature2.App2, Nomenclature2.Voie2, Nomenclature2.Options, Nomenclature2.ISO,  "
Sql = Sql & "Nomenclature2.Longueur, Nomenclature2.[Longueur Total], Nomenclature2.TEINT,  "
Sql = Sql & "Nomenclature2.TEINT2, Nomenclature2.SECT, Nomenclature2.Qts "
Sql = Sql & "FROM Nomenclature2 WHERE Nomenclature2.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"
Con.Execute Sql

Sql = "INSERT INTO T_Nomenclature ( Id_IndiceProjet, CONNECTEUR, [Nb Voies], [OPTION], Qté, [Prix U],  "
Sql = Sql & "[Prix Total], CODE_APP, DESIGNATION, Voie, Couleur, [Lib Connecteur], Fournisseur, [Ref Four],  "
Sql = Sql & "[Ref Bouch], [Bouchon Qté], [Bouchon Prix U], [Bouchon Prix Total], [Lib Bouch], "
Sql = Sql & " [Bouch Fourr], [Bouch Réf Four], [Ref Capot], [Ref Verrou], [Ref Joint], [Joint Qté],  "
Sql = Sql & "[Joint Prix U], [Joint Prix Total], [Lib Joint], [Joint Four], [Joint Four Réf],  "
Sql = Sql & "[Nb Alvé], Famille, [Famille Lib], [Alvé Réf], [Alvé Qté], [Alvé Prix U],  "
Sql = Sql & "[Alvé Prix Total], [Alvé Réf Fourr], [Alvéole Mini en mm2], [Alvéole Maxi en mm2] ) "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, T_Nomenclature.CONNECTEUR, T_Nomenclature.[Nb Voies],  "
Sql = Sql & "T_Nomenclature.OPTION,  T_Nomenclature.Qté, T_Nomenclature.[Prix U],  "
Sql = Sql & "T_Nomenclature.[Prix Total], T_Nomenclature.CODE_APP, T_Nomenclature.DESIGNATION,  "
Sql = Sql & "T_Nomenclature.Voie, T_Nomenclature.Couleur, T_Nomenclature.[Lib Connecteur],  "
Sql = Sql & "T_Nomenclature.Fournisseur, T_Nomenclature.[Ref Four], T_Nomenclature.[Ref Bouch],  "
Sql = Sql & "T_Nomenclature.[Bouchon Qté], T_Nomenclature.[Bouchon Prix U],  "
Sql = Sql & "T_Nomenclature.[Bouchon Prix Total], T_Nomenclature.[Lib Bouch],  "
Sql = Sql & "T_Nomenclature.[Bouch Fourr], T_Nomenclature.[Bouch Réf Four],  "
Sql = Sql & "T_Nomenclature.[Ref Capot], T_Nomenclature.[Ref Verrou], T_Nomenclature.[Ref Joint],  "
Sql = Sql & "T_Nomenclature.[Joint Qté], T_Nomenclature.[Joint Prix U],  "
Sql = Sql & "T_Nomenclature.[Joint Prix Total], T_Nomenclature.[Lib Joint],  "
Sql = Sql & "T_Nomenclature.[Joint Four], T_Nomenclature.[Joint Four Réf],  "
Sql = Sql & "T_Nomenclature.[Nb Alvé], T_Nomenclature.Famille, T_Nomenclature.[Famille Lib],  "
Sql = Sql & "T_Nomenclature.[Alvé Réf], T_Nomenclature.[Alvé Qté], T_Nomenclature.[Alvé Prix U],  "
Sql = Sql & "T_Nomenclature.[Alvé Prix Total] , T_Nomenclature.[Alvé Réf Fourr],  "
Sql = Sql & "T_Nomenclature.[Alvéole Mini en mm2], T_Nomenclature.[Alvéole Maxi en mm2] "
Sql = Sql & "FROM T_Nomenclature "
Sql = Sql & "WHERE T_Nomenclature.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"
Con.Execute Sql

Sql = "INSERT INTO T_Dossier_Contrôle ( Id_IndiceProjet, Onglet, ACTIVER, LIAI, DESIGNATION, FIL, SECT,   "
Sql = Sql & "TEINT, TEINT2, ISO, [LONG], [LONG CP], LONG_ADD, LONG_ADD2, COUPE, POS, [POS-OUT], FA, APP,   "
Sql = Sql & "VOI, [REF CONNECTEUR], [REF CONNECTEUR_FOUR], [REF CLIP], [REF CLIP FOUR], PRECO, [REF JOINT],   "
Sql = Sql & "[REF JOINT FOUR], POS2, [POS-OUT2], FA2, APP2, VOI2, [REF CONNECTEUR2],   "
Sql = Sql & "[REF CONNECTEUR_FOUR2], [REF CLIP2], [REF CLIP FOUR2], PRECO2, [REF JOINT2],   "
Sql = Sql & "[REF JOINT FOUR2], PRECOG, [OPTION], [CRITÈRES SPÉCIFIQUES] )  "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.ACTIVER,   "
Sql = Sql & "T_Dossier_Contrôle.LIAI,T_Dossier_Contrôle.DESIGNATION, T_Dossier_Contrôle.FIL,   "
Sql = Sql & "T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT, T_Dossier_Contrôle.TEINT2,   "
Sql = Sql & "T_Dossier_Contrôle.ISO, T_Dossier_Contrôle.LONG, T_Dossier_Contrôle.[LONG CP],   "
Sql = Sql & "T_Dossier_Contrôle.LONG_ADD, T_Dossier_Contrôle.LONG_ADD2, T_Dossier_Contrôle.COUPE,   "
Sql = Sql & "T_Dossier_Contrôle.POS, T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.FA,   "
Sql = Sql & "T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI, T_Dossier_Contrôle.[REF CONNECTEUR],   "
Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[REF CLIP],   "
Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR], T_Dossier_Contrôle.PRECO, T_Dossier_Contrôle.[REF JOINT],   "
Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR], T_Dossier_Contrôle.POS2, T_Dossier_Contrôle.[POS-OUT2],  "
Sql = Sql & " T_Dossier_Contrôle.FA2, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2,  "
Sql = Sql & " T_Dossier_Contrôle.[REF CONNECTEUR2], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2],   "
Sql = Sql & "T_Dossier_Contrôle.[REF CLIP2], T_Dossier_Contrôle.[REF CLIP FOUR2] ,   "
Sql = Sql & "T_Dossier_Contrôle.PRECO2, T_Dossier_Contrôle.[REF JOINT2],   "
Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR2], T_Dossier_Contrôle.PRECOG,   "
Sql = Sql & "T_Dossier_Contrôle.Option, T_Dossier_Contrôle.[CRITÈRES SPÉCIFIQUES]  "
Sql = Sql & "FROM T_Dossier_Contrôle  "
Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"
Con.Execute Sql

Sql = "INSERT INTO T_Dossier_Fabrication ( Id_IndiceProjet, Onglet, ACTIVER, LIAI, DESIGNATION, FIL, SECT,   "
Sql = Sql & "TEINT, TEINT2, ISO, [LONG], [LONG CP], LONG_ADD, LONG_ADD2, COUPE, POS, [POS-OUT],   "
Sql = Sql & "FA, APP, VOI, [REF CONNECTEUR], [REF CONNECTEUR_FOUR], [REF CLIP], [REF CLIP FOUR],   "
Sql = Sql & "PRECO, [REF JOINT], [REF JOINT FOUR], POS2, [POS-OUT2], FA2, APP2, VOI2, [REF CONNECTEUR2],   "
Sql = Sql & "[REF CONNECTEUR_FOUR2], [REF CLIP2], [REF CLIP FOUR2], PRECO2, [REF JOINT2],   "
Sql = Sql & "[REF JOINT FOUR2], PRECOG, [OPTION], [CRITÈRES SPÉCIFIQUES] )  "
Sql = Sql & "SELECT " & Rs!Id & " AS Expr1, T_Dossier_Fabrication.Onglet, T_Dossier_Fabrication.ACTIVER,   "
Sql = Sql & "T_Dossier_Fabrication.LIAI, T_Dossier_Fabrication.DESIGNATION, T_Dossier_Fabrication.FIL,   "
Sql = Sql & "T_Dossier_Fabrication.SECT, T_Dossier_Fabrication.TEINT, T_Dossier_Fabrication.TEINT2,   "
Sql = Sql & "T_Dossier_Fabrication.ISO, T_Dossier_Fabrication.LONG, T_Dossier_Fabrication.[LONG CP],   "
Sql = Sql & "T_Dossier_Fabrication.LONG_ADD, T_Dossier_Fabrication.LONG_ADD2,   "
Sql = Sql & "T_Dossier_Fabrication.COUPE, T_Dossier_Fabrication.POS, T_Dossier_Fabrication.[POS-OUT],   "
Sql = Sql & "T_Dossier_Fabrication.FA, T_Dossier_Fabrication.APP, T_Dossier_Fabrication.VOI,   "
Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR], T_Dossier_Fabrication.[REF CONNECTEUR_FOUR],   "
Sql = Sql & "T_Dossier_Fabrication.[REF CLIP], T_Dossier_Fabrication.[REF CLIP FOUR],   "
Sql = Sql & "T_Dossier_Fabrication.PRECO, T_Dossier_Fabrication.[REF JOINT],   "
Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR], T_Dossier_Fabrication.POS2,   "
Sql = Sql & "T_Dossier_Fabrication.[POS-OUT2], T_Dossier_Fabrication.FA2, T_Dossier_Fabrication.APP2,   "
Sql = Sql & "T_Dossier_Fabrication.VOI2, T_Dossier_Fabrication.[REF CONNECTEUR2],   "
Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR_FOUR2], T_Dossier_Fabrication.[REF CLIP2],   "
Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR2], T_Dossier_Fabrication.PRECO2,   "
Sql = Sql & "T_Dossier_Fabrication.[REF JOINT2], T_Dossier_Fabrication.[REF JOINT FOUR2],   "
Sql = Sql & "T_Dossier_Fabrication.PRECOG, T_Dossier_Fabrication.OPTION,   "
Sql = Sql & "T_Dossier_Fabrication.[CRITÈRES SPÉCIFIQUES]  "
Sql = Sql & "FROM T_Dossier_Fabrication  "
Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & CLng(Me.txt3.Tag) & ";"
Con.Execute Sql


Me.txt3.Tag = Rs!Id
End If
Sql = "SELECT T_indiceProjet.PlAutoCadSave, T_indiceProjet.LiAutoCadSave, T_indiceProjet.OuAutoCadSave  "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
 Set RsCartouche = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
FileOu = Trim("" & Rs!OUAutoCadSave)
FilePL = Trim("" & Rs!PlAutoCadSave)
FileLi = Trim("" & Rs!LiAutoCadSave)
    If Trim("" & Rs!PlAutoCadSave) <> "" Then
        If Fso.FileExists(DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), _
            Rs!PlAutoCadSave & ".dwg")) = True Then
             Fso.CopyFile DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), _
             Rs!PlAutoCadSave & ".dwg"), PathArchive(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), "" & RsCartouche!Client, "" & RsCartouche!CleAc, "" & RsCartouche!Pieces, "Pl", RsCartouche.Fields("Pl"), Me.txt3.Tag, RsCartouche.Fields("PI_Indice"), RsCartouche.Fields("Pl_Indice"), RsCartouche!Version) & ".dwg"
        End If
    End If
    

    If Trim("" & Rs!LiAutoCadSave) <> "" Then
         If Fso.FileExists(DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), _
            Rs!LiAutoCadSave & ".XLS")) = True Then
              Fso.CopyFile DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), _
              Rs!LiAutoCadSave & ".XLS"), PathArchive(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), "" & RsCartouche!Client, "" & RsCartouche!CleAc, "" & RsCartouche!Pieces, "Li", RsCartouche.Fields("Li"), Me.txt3.Tag, RsCartouche.Fields("PI_Indice"), RsCartouche.Fields("Li_Indice"), RsCartouche!Version) & ".XLS"
         
        End If
    End If
    
    If Trim("" & Rs!OUAutoCadSave) <> "" Then
        If Fso.FileExists(DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), _
            Rs!OUAutoCadSave & ".dwg")) = True Then
             Fso.CopyFile DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), _
             Rs!OUAutoCadSave & ".dwg"), PathArchive(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), "" & RsCartouche!Client, "" & RsCartouche!CleAc, "" & RsCartouche!Pieces, "Ou", RsCartouche.Fields("Ou"), Me.txt3.Tag, RsCartouche.Fields("PI_Indice"), RsCartouche.Fields("Ou_Indice"), RsCartouche!Version) & ".dwg"
        End If
   End If
    
End If

Rs.Requery
FileOu = Trim("" & Rs!OUAutoCadSave)
FilePL = Trim("" & Rs!PlAutoCadSave)
FileLi = Trim("" & Rs!LiAutoCadSave)
Sql = "SELECT T_indiceProjet.PlAutoCadSave, T_indiceProjet.LiAutoCadSave, T_indiceProjet.OuAutoCadSave, T_indiceProjet.Id "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Pere=" & Me.txt3.Tag & " AND T_indiceProjet.Archiver=False;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & Rs!Id & ";"
        Set RsCartouche = Con.OpenRecordSet(Sql)
        If FileOu <> "" Then
            Racourci PathArchive(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), "" & RsCartouche!Client, "" & RsCartouche!CleAc, "" & RsCartouche!Pieces, "Ou", RsCartouche.Fields("Ou"), Rs!Id, RsCartouche.Fields("pi_Indice"), RsCartouche.Fields("Ou_Indice"), RsCartouche!Version), _
             DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), FileOu), "dwg"
        End If
        If FilePL <> "" Then
            Racourci PathArchive(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), "" & RsCartouche!Client, "" & RsCartouche!CleAc, "" & RsCartouche!Pieces, "Pl", RsCartouche.Fields("Pl"), Rs!Id, RsCartouche.Fields("pi_Indice"), RsCartouche.Fields("Pl_Indice"), RsCartouche!Version), _
             DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), FilePL), "dwg"
        End If
        If FileLi <> "" Then
           Racourci PathArchive(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), "" & RsCartouche!Client, "" & RsCartouche!CleAc, "" & RsCartouche!Pieces, "Li", RsCartouche.Fields("Li"), Rs!Id, RsCartouche.Fields("pi_Indice"), RsCartouche.Fields("Li_Indice"), RsCartouche!Version), _
             DefinirChemienComplet(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("PathArchiveAutocad")), FileLi), "Xls"
        End If
        
    Rs.MoveNext
Wend
 strStatus = "MOD"
End If
Me.Enabled = False
Set FormBarGrah = Me
If OptionButton1.Value = True Then
'    pathTmpXls = Environ("USERPROFILE") & "\Mes Documents"
'    If Fso.FolderExists(pathTmpXls) = False Then Fso.CreateFolder pathTmpXls
'    pathTmpXls = pathTmpXls & "\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS"
'
'     If Fso.FileExists(pathTmpXls) = True Then
'        Fso.DeleteFile pathTmpXls
'    End If
         Set NewUserForm2 = New UserForm2
         
'        ExporteFrmModifier NewUserForm2, CLng(Me.txt3.Tag), txt9, Me, Edition:=True
       
        
        NewUserForm2.chargement txt6, CLng(Me.txt3.Tag), txt9, Me, True, BooolBloque
     
Else
  If boolAutoCAD = False And IsCilent = False Then
   


    MsgBox MsgAutoCad & "Vous ne possédez pas de licence AutoCad." & vbCrLf & "Vous ne pouvez pas reporter vos modifications" & vbCrLf & "sur vos différents plans. "
Else
 Planche_Clous.chargement CLng(Me.txt3.Tag)



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
    If IsCilent = False Then
        
        subDessinerPlan Me.txt3.Tag
        subDessinerOutil Me.txt3.Tag
        
        
        MsgBox "Fin du traitement" & vbCrLf & NbError & " Erreur(s) détectée(s) !"
    Else
        MsgBox "Votre demande a été prise en compte vous pouvez suivre l'évolution de votre travail dans la fenêtre (Liste des JOB)"
    End If
 End If
NbError = 0
 Noquite = False
 
Unload Me
End If
End Sub
Sub continuer(Optional import As Boolean)
Dim pathTmpXls As String
Dim Sql As String
Dim Rs As Recordset
Me.Visible = True
If import = True Then
        pathTmpXls = UserForm2.Caption
        UserForm2_boolExcute = UserForm2.boolExcute
       Unload UserForm2
        If UserForm2_boolExcute = False Then
            Me.Enabled = True

              Exit Sub
        End If
          MsgAutoCad = "Vos données ont bien été enregistrées, toustefois :" & vbCrLf & vbCrLf
        ImporteXls pathTmpXls, CLng(Me.txt3.Tag), Edition:=True
    End If


End Sub
Public Sub charger(MyDroit As String, BooolBloque As Boolean)
Qui = MyDroit
If BooolBloque = True Then
OptionButton1.Value = True
OptionButton2.Locked = True
End If
DoEvents
Me.Show

End Sub
Public Sub chargement(FRM As Object, Optional BooolBloque As Boolean)
Unload FRM

If BooolBloque = True Then
OptionButton1.Value = True
OptionButton2.Locked = True
End If
Me.Show
End Sub

Public Sub Charge(MyForm As Object, MyDroit As String)
Dim Sql As String
Dim Rs As Recordset
Qui = MyDroit
IdFils = 0
Dim Ofset As Long
'Debug.Print MyForm.Name
'MyForm.Hide
IdIndiceProjet = MyForm.IdIndiceProjet

Sql = "SELECT SelectProjets.* FROM SelectProjets WHERE SelectProjets.Id=" & IdIndiceProjet & " ;"

Set Rs = Con.OpenRecordSet(Sql)

Set FormBarGrah = Me

If Rs.EOF = False Then
On Error Resume Next
For I = 0 To 12
Debug.Print Rs(I).Name
If Rs(I).Name = "CleAc" Then
Ofset = 1
End If
    Me.Controls("txt" & CStr(I + 1)) = "" & Rs(I + Ofset)
     Me.Controls("txt" & CStr(I + 1)).Tag = "" & Rs.Fields(13)

Next I
    
    OptionButton2.Value = True
    OptionButton1.Value = False
    If BooolBloque = True Then OptionButton1.Locked = True
    Me.CommandButton1.Enabled = False
 End If
 Set Rs = Con.CloseRecordSet(Rs)
Unload MyForm
 Me.Show
End Sub

Private Sub OptionButton1_Click()
OptionButton2.Value = False
End Sub

Private Sub OptionButton2_Click()
OptionButton1.Value = False
End Sub

Private Sub UserForm_Activate()
frmAutocâble.EnabledMenu
End Sub

Private Sub UserForm_Initialize()
 H = Me.Height
 W = Me.Width
 L = CommandButton3.Left
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
LibertPice
MajDroitsFrm Id_Users
frmAutocâble.DesEnabledMenu
End Sub


Private Sub CommandButton3_Click()
 Noquite = False
Unload Me
End Sub

'Private Sub UserForm_Resize()
'Dim P_H As Double
'Dim P_W As Double
'P_W = W / Me.Width
'P_H = H / Me.Height
'CommandButton3.Left = L / P_W
'End Sub
Private Sub UserForm_Resize()
Dim Sql As String
Sql = "UPDATE T_indiceProjet SET T_indiceProjet.UserName = '" & Replace(Machine, "'", "''") & "' "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Controls("txt1").Tag & " OR T_indiceProjet.Pere=" & Me.Controls("txt1").Tag & ";"
Con.Execute Sql

Con.Execute Sql
End Sub
