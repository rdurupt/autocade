VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportXls 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Créer un plan import des données :"
   ClientHeight    =   9480
   ClientLeft      =   30
   ClientTop       =   195
   ClientWidth     =   12180
   Icon            =   "ImportXls.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "ImportXls.dsx":08CA
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "DESSINE.PAR;NOM Déssiné par;QRY;TXT;TXT5"
End
Attribute VB_Name = "ImportXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdProjet As Long
Public IdPieces As Long
Public IdIndiceProjet As Long
Dim Noquite As Boolean
Dim boolChrono As Boolean
Dim Extension As String
Dim Qui As String
Dim txtOption As String



Private Sub CommandButton1_Click()
Dim NewUserForm2 As UserForm2
Dim Fso As New FileSystemObject
Dim Sql As String
Dim Rs As Recordset
 Dim Rs2 As Recordset
Set FormBarGrah = Me

 Set TableauPath = funPath
 

If Me.OptionButton1.Value = True Then
    txtOption = "A"
     Me.FichierXLS = Trim("" & Me.FichierXLS)
    If Trim("" & Me.FichierXLS) = "" Then
    MsgBox "Vous devez saisir le chemin ainsi que le nom du fichier AUTOCAD à importer", vbExclamation, "Erreur"
    Me.FichierXLS.SetFocus
    Set Fso = Nothing
    Exit Sub
    End If
    If UCase(Right(Me.FichierXLS, 4)) <> UCase(".dwg") Then Me.FichierXLS = Me.FichierXLS & ".dwg"
If MsgBox("Voulez vous exécuter l'importation de :" & Me.FichierXLS, vbQuestion + vbYesNo, "Import EXCEL") = vbNo Then Exit Sub

DoEvents

    
'Dim Rs As Recordset
'Dim sql As String

    
End If

If Me.OptionButton2.Value = True Then
    txtOption = "E"
    Me.FichierXLS = Trim("" & Me.FichierXLS)
    If Trim("" & Me.FichierXLS) = "" Then
    MsgBox "Vous devez saisir le chemin ainsi que le nom du fichier EXCEL à importer", vbExclamation, "Erreur"
    Me.FichierXLS.SetFocus
    Set Fso = Nothing
    Exit Sub
'Dim Rs As Recordset
'Dim sql As String

Exit Sub
End If
Me.Enabled = False

If UCase(Right(Me.FichierXLS, 4)) <> ".XLS" Then Me.FichierXLS = Me.FichierXLS & ".XLS"
If MsgBox("Voulez vous exécuter l'importation de :" & Me.FichierXLS, vbQuestion + vbYesNo, "Import EXCEL") = vbNo Then Exit Sub

DoEvents

    
End If


If Me.OptionButton3.Value = True Then
    txtOption = "N"
End If
  
If Me.OptionButton7.Value = True Then
    txtOption = "P"
End If
  
Select Case txtOption
         Case "A"
                If MsgBox("Voulez vous garder :" & vbCrLf & Me.FichierXLS & vbCrLf & "comme modèle", vbYesNo) = vbYes Then
                     ScanDessin Me.FichierXLS, IdIndiceProjet, True
                Else
                    ScanDessin Me.FichierXLS, IdIndiceProjet
                End If
'             Exit Sub
         Case "E"
                 'me.hide
                 ImporteXls Me.FichierXLS, IdIndiceProjet
         
         Case "N"
         Dim pathTmpXls As String

 

   
    
        pathTmpXls = Environ("USERPROFILE") & "\Mes Documents\" & Replace(Replace(txt10, ":", ""), ".", "") & ".XLS"
   Me.FichierXLS = pathTmpXls
 
'pathTmpXls = Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt6.Caption, ":", "_", 1) & ".XLS"
     If Fso.FileExists(pathTmpXls) = True Then
        Fso.DeleteFile pathTmpXls
    End If
        
        ExporteXls pathTmpXls, IdIndiceProjet
        Set NewUserForm2 = New UserForm2
        NewUserForm2.chargement txt7, IdIndiceProjet, txt11, Me, Edition:=True
        Noquite = False
'     UserForm2.Chargement pathTmpXls, txt11, IdIndiceProjet, Me, False
     Me.Hide
'     UserForm2.SetFocus
      
        GoTo Fin
      Case "P"
        AffaireExistante.Show vbModal
      AffaireExistanteAnnuler = AffaireExistante.Annuler
      AffaireExistante_txt3_Tag = AffaireExistante.txt3.Tag
      Unload AffaireExistante
        If AffaireExistanteAnnuler = True Then Exit Sub
        Dim txtArchive As String
        If PlanArchive = True Then txtArchive = ""
      Sql = "UPDATE T_indiceProjet, T_indiceProjet AS T_indiceProjet_1  "
        Sql = Sql & "SET T_indiceProjet.Masse = [T_indiceProjet_1].[Masse],  "
        Sql = Sql & "T_indiceProjet.PlAutoCadSaveAs = [T_indiceProjet_1].[PlAutoCadSaveAs],  "
        Sql = Sql & "T_indiceProjet.PlAutoCadSave = [T_indiceProjet_1].[PlAutoCadSave],  "
         Sql = Sql & "T_indiceProjet.OuAutoCadSaveAs = [T_indiceProjet_1].[OuAutoCadSaveAs],  "
        Sql = Sql & "T_indiceProjet.OuAutoCadSave = [T_indiceProjet_1].[OuAutoCadSave] ,"
        Sql = Sql & "T_indiceProjet.NbCartouche = [T_indiceProjet_1].[NbCartouche], T_indiceProjet.Cartouche = [T_indiceProjet_1].[Cartouche] "
        Sql = Sql & "WHERE T_indiceProjet.Id= " & IdIndiceProjet & " "
        Sql = Sql & "AND T_indiceProjet_1.Id=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

        
        

 Con.Execute Sql
    

        
        
  
   Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet,ACTIVER, CONNECTEUR, [O/N],  "
        Sql = Sql & "DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1,  "
        Sql = Sql & "PRECO2, [100%] ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet, Connecteurs.ACTIVER, "
        Sql = Sql & "Connecteurs.CONNECTEUR, Connecteurs.[O/N],  "
        Sql = Sql & "Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
        Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT],  "
        Sql = Sql & "Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.[100%] "
        Sql = Sql & "FROM Connecteurs "
        Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

    Con.Execute Sql
    
    
Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, ACTIVER, LIAI, DESIGNATION, FIL, SECT, TEINT,   "
Sql = Sql & "TEINT2, ISO, [LONG], [LONG CP], Long_Add, Long_Add2, COUPE, POS, [POS-OUT], FA, APP, VOI,   "
Sql = Sql & "[Ref Connecteur], [Ref Connecteur_Four], [Ref Clip], [Ref Clip Four], PRECO, [Ref Joint],   "
Sql = Sql & "[Ref Joint four], POS2, [POS-OUT2], FA2, APP2, VOI2, [Ref Connecteur2],   "
Sql = Sql & "[Ref Connecteur_Four2], [Ref Clip2], [Ref Clip Four2], PRECO2, [Ref Joint2],   "
Sql = Sql & "[Ref Joint Four2], PRECOG, [OPTION], [Critères spécifiques] )  "
Sql = Sql & "SELECT " & IdIndiceProjet & " AS Expr1, Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI,   "
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
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"


Con.Execute Sql
'
    
    Sql = "INSERT INTO Composants ( Id_IndiceProjet, ACTIVER, DESIGNCOMP, NUMCOMP, REFCOMP, Path, Code_APP_Lier, Voie, POS, [POS-OUT] )"
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_Indice, Composants.ACTIVER, Composants.DESIGNCOMP,  "
        Sql = Sql & "Composants.NUMCOMP, Composants.REFCOMP, Composants.Path, Composants.Code_APP_Lier,  "
        Sql = Sql & "Composants.Voie, Composants.POS, Composants.[POS-OUT] "
        Sql = Sql & "FROM Composants "
        Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
Con.Execute Sql

Sql = "INSERT INTO nota ( Id_IndiceProjet,ACTIVER,NOTA, NUMNOTA ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet,Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA "
        Sql = Sql & "FROM Nota "
        Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

Con.Execute Sql
        Sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet, ACTIVER, NŒUDS,  "
        Sql = Sql & "LONGUEUR, DESIGN_HAB, CODE_RSA, CODE_PSA, CODE_ENC,  "
        Sql = Sql & "DIAMETRE, CLASSE_T, TORON_PRINCIPAL, LONGUEUR_CUMULEE,  "
        Sql = Sql & "Fleche_Droite ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet, T_Noeuds.ACTIVER, T_Noeuds.NŒUDS, "
          Sql = Sql & "T_Noeuds.LONGUEUR,  "
        Sql = Sql & "T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC,  "
        Sql = Sql & "T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T, T_Noeuds.TORON_PRINCIPAL, T_Noeuds.LONGUEUR_CUMULEE,  "
        Sql = Sql & "T_Noeuds.Fleche_Droite "
        Sql = Sql & "FROM T_Noeuds "
        Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
Con.Execute Sql

Sql = "INSERT INTO NomeclatureConnecteurs ( Id_IndiceProjet, App, Designation, Connecteur, Connecteur_Four, Liaison,   "
        Sql = Sql & "SECT, TEINT, TEINT2, ISO, Voie, [LONG], COUPE, [LONG CP], Long_Add, Famille, Bouchon,   "
        Sql = Sql & "Capot, Capot_Four, Verrou, Verrout_Four, Options, Clip, ClipFour, Joint, JointFour )  "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_Indice, NomeclatureConnecteurs.App, NomeclatureConnecteurs.Designation,   "
        Sql = Sql & "NomeclatureConnecteurs.Connecteur, NomeclatureConnecteurs.Connecteur_Four, NomeclatureConnecteurs.Liaison,   "
        Sql = Sql & "NomeclatureConnecteurs.SECT, NomeclatureConnecteurs.TEINT, NomeclatureConnecteurs.TEINT2,   "
        Sql = Sql & "NomeclatureConnecteurs.ISO, NomeclatureConnecteurs.Voie, NomeclatureConnecteurs.LONG, NomeclatureConnecteurs.COUPE,   "
        Sql = Sql & "NomeclatureConnecteurs.[LONG CP], NomeclatureConnecteurs.Long_Add, NomeclatureConnecteurs.Famille,   "
        Sql = Sql & "NomeclatureConnecteurs.Bouchon, NomeclatureConnecteurs.Capot, NomeclatureConnecteurs.Capot_Four,   "
        Sql = Sql & "NomeclatureConnecteurs.Verrou, NomeclatureConnecteurs.Verrout_Four, NomeclatureConnecteurs.Options,   "
        Sql = Sql & "NomeclatureConnecteurs.Clip, NomeclatureConnecteurs.ClipFour, NomeclatureConnecteurs.Joint, NomeclatureConnecteurs.JointFour   "
         Sql = Sql & "FROM NomeclatureConnecteurs  "
        Sql = Sql & "WHERE NomeclatureConnecteurs.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
    Con.Execute Sql

    Sql = "INSERT INTO Nomenclature ( Id_IndiceProjet, Designation, App, Ref, RefFour, Options )   "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_Indice, Nomenclature.Designation, Nomenclature.App, Nomenclature.Ref,    "
        Sql = Sql & "Nomenclature.RefFour, Nomenclature.Options   "
        Sql = Sql & "FROM Nomenclature   "
        Sql = Sql & "WHERE Nomenclature.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
Con.Execute Sql


Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, LIAI, Designation, App, Voie, Ref, Fournisseur, RefFour, App2, Voie2, Options, ISO,   "
        Sql = Sql & "Longueur, [Longueur Total], TEINT, TEINT2, SECT, Qts )  "
        Sql = Sql & "SELECT " & IdIndiceProjet & "  AS Id_Indice, Nomenclature2.LIAI, Nomenclature2.Designation, Nomenclature2.App, Nomenclature2.Voie,   "
        Sql = Sql & "Nomenclature2.Ref, Nomenclature2.Fournisseur, Nomenclature2.RefFour, Nomenclature2.App2, Nomenclature2.Voie2,   "
        Sql = Sql & "Nomenclature2.Options, Nomenclature2.ISO, Nomenclature2.Longueur, Nomenclature2.[Longueur Total], Nomenclature2.TEINT,   "
        Sql = Sql & "Nomenclature2.TEINT2, Nomenclature2.SECT, Nomenclature2.Qts  "
        Sql = Sql & "FROM Nomenclature2  "
        Sql = Sql & "WHERE Nomenclature2.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
Con.Execute Sql

Sql = "INSERT INTO NomenclaturFinal ( Id_IndiceProjet, Designation, Famille, Fournisseur, Ref, RefFour, Qts, ISO, TEINT, TEINT2,   "
        Sql = Sql & "SECT, Qts_Encelade, Qts_E_Boutique, Qts_Appro, Prix_Revient, Prix_Vente, Options )  "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_Indice, NomenclaturFinal.Designation, NomenclaturFinal.Famille, NomenclaturFinal.Fournisseur,   "
        Sql = Sql & "NomenclaturFinal.Ref, NomenclaturFinal.RefFour, NomenclaturFinal.Qts, NomenclaturFinal.ISO,   "
        Sql = Sql & "NomenclaturFinal.TEINT, NomenclaturFinal.TEINT2, NomenclaturFinal.SECT, NomenclaturFinal.Qts_Encelade,   "
        Sql = Sql & "NomenclaturFinal.Qts_E_Boutique, NomenclaturFinal.Qts_Appro, NomenclaturFinal.Prix_Revient,   "
        Sql = Sql & "NomenclaturFinal.Prix_Vente, NomenclaturFinal.Options  "
        Sql = Sql & "FROM NomenclaturFinal  "
        Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

Con.Execute Sql

Sql = "INSERT INTO T_Nomenclature ( Id_IndiceProjet, CONNECTEUR, [Nb Voies], [OPTION], Qté, [Prix U], [Prix Total], CODE_APP, DESIGNATION,   "
        Sql = Sql & "Voie, Couleur, [Lib Connecteur], Fournisseur, [Ref Four], [Ref Bouch], [Bouchon Qté], [Bouchon Prix U],   "
        Sql = Sql & "[Bouchon Prix Total], [Lib Bouch], [Bouch Fourr], [Bouch Réf Four], [Ref Capot], [Ref Verrou], [Ref Joint],   "
        Sql = Sql & "[Joint Qté], [Joint Prix U], [Joint Prix Total], [Lib Joint], [Joint Four], [Joint Four Réf], [Nb Alvé], Famille,   "
        Sql = Sql & "[Famille Lib], [Alvé Réf], [Alvé Qté], [Alvé Prix U], [Alvé Prix Total], [Alvé Réf Fourr], [Alvéole Mini en mm2],   "
        Sql = Sql & "[Alvéole Maxi en mm2] )  "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_Indice, T_Nomenclature.CONNECTEUR, T_Nomenclature.[Nb Voies], T_Nomenclature.OPTION, T_Nomenclature.Qté,   "
        Sql = Sql & "T_Nomenclature.[Prix U], T_Nomenclature.[Prix Total], T_Nomenclature.CODE_APP, T_Nomenclature.DESIGNATION,   "
        Sql = Sql & "T_Nomenclature.Voie, T_Nomenclature.Couleur, T_Nomenclature.[Lib Connecteur], T_Nomenclature.Fournisseur,   "
        Sql = Sql & "T_Nomenclature.[Ref Four], T_Nomenclature.[Ref Bouch], T_Nomenclature.[Bouchon Qté], T_Nomenclature.[Bouchon Prix U],   "
        Sql = Sql & "T_Nomenclature.[Bouchon Prix Total], T_Nomenclature.[Lib Bouch], T_Nomenclature.[Bouch Fourr],   "
        Sql = Sql & "T_Nomenclature.[Bouch Réf Four], T_Nomenclature.[Ref Capot], T_Nomenclature.[Ref Verrou],   "
        Sql = Sql & "T_Nomenclature.[Ref Joint], T_Nomenclature.[Joint Qté], T_Nomenclature.[Joint Prix U],  "
        Sql = Sql & " T_Nomenclature.[Joint Prix Total], T_Nomenclature.[Lib Joint], T_Nomenclature.[Joint Four],   "
        Sql = Sql & "T_Nomenclature.[Joint Four Réf], T_Nomenclature.[Nb Alvé], T_Nomenclature.Famille, T_Nomenclature.[Famille Lib],   "
        Sql = Sql & "T_Nomenclature.[Alvé Réf], T_Nomenclature.[Alvé Qté], T_Nomenclature.[Alvé Prix U], T_Nomenclature.[Alvé Prix Total] ,   "
        Sql = Sql & "T_Nomenclature.[Alvé Réf Fourr], T_Nomenclature.[Alvéole Mini en mm2], T_Nomenclature.[Alvéole Maxi en mm2]  "
        Sql = Sql & "FROM T_Nomenclature  "
        Sql = Sql & "WHERE T_Nomenclature.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

Con.Execute Sql


Sql = "INSERT INTO T_Dossier_Contrôle ( Id_IndiceProjet, Onglet, ACTIVER, LIAI, DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, [LONG],  "
        Sql = Sql & "[LONG CP], LONG_ADD, LONG_ADD2, COUPE, POS, [POS-OUT], FA, APP, VOI, [REF CONNECTEUR], [REF CONNECTEUR_FOUR],  "
        Sql = Sql & "[REF CLIP], [REF CLIP FOUR], PRECO, [REF JOINT], [REF JOINT FOUR], POS2, [POS-OUT2], FA2, APP2, VOI2,  "
        Sql = Sql & "[REF CONNECTEUR2], [REF CONNECTEUR_FOUR2], [REF CLIP2], [REF CLIP FOUR2], PRECO2, [REF JOINT2], [REF JOINT FOUR2],  "
        Sql = Sql & "PRECOG, [OPTION], [CRITÈRES SPÉCIFIQUES] ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_Indice, T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.ACTIVER, T_Dossier_Contrôle.LIAI,  "
        Sql = Sql & "T_Dossier_Contrôle.DESIGNATION, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
        Sql = Sql & "T_Dossier_Contrôle.TEINT2, T_Dossier_Contrôle.ISO, T_Dossier_Contrôle.LONG, T_Dossier_Contrôle.[LONG CP],  "
        Sql = Sql & "T_Dossier_Contrôle.LONG_ADD, T_Dossier_Contrôle.LONG_ADD2, T_Dossier_Contrôle.COUPE, T_Dossier_Contrôle.POS,  "
        Sql = Sql & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.FA, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI,  "
        Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[REF CLIP],  "
        Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR], T_Dossier_Contrôle.PRECO, T_Dossier_Contrôle.[REF JOINT],  "
        Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR], T_Dossier_Contrôle.POS2, T_Dossier_Contrôle.[POS-OUT2],  "
        Sql = Sql & "T_Dossier_Contrôle.FA2, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2],  "
        Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2], T_Dossier_Contrôle.[REF CLIP2], T_Dossier_Contrôle.[REF CLIP FOUR2] ,  "
        Sql = Sql & "T_Dossier_Contrôle.PRECO2, T_Dossier_Contrôle.[REF JOINT2], T_Dossier_Contrôle.[REF JOINT FOUR2],  "
        Sql = Sql & "T_Dossier_Contrôle.PRECOG, T_Dossier_Contrôle.Option, T_Dossier_Contrôle.[CRITÈRES SPÉCIFIQUES] "
        Sql = Sql & "FROM T_Dossier_Contrôle "
        Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
    Con.Execute Sql

Sql = "INSERT INTO T_Dossier_Fabrication ( Id_IndiceProjet, Onglet, ACTIVER, LIAI, DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP],  "
        Sql = Sql & "LONG_ADD, LONG_ADD2, COUPE, POS, [POS-OUT], FA, APP, VOI, [REF CONNECTEUR], [REF CONNECTEUR_FOUR], [REF CLIP], [REF CLIP FOUR],  "
        Sql = Sql & "PRECO, [REF JOINT], [REF JOINT FOUR], POS2, [POS-OUT2], FA2, APP2, VOI2, [REF CONNECTEUR2], [REF CONNECTEUR_FOUR2],  "
        Sql = Sql & "[REF CLIP2], [REF CLIP FOUR2], PRECO2, [REF JOINT2], [REF JOINT FOUR2], PRECOG, [OPTION], [CRITÈRES SPÉCIFIQUES] ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_Indice, T_Dossier_Fabrication.Onglet, T_Dossier_Fabrication.ACTIVER, T_Dossier_Fabrication.LIAI,  "
        Sql = Sql & "T_Dossier_Fabrication.DESIGNATION, T_Dossier_Fabrication.FIL, T_Dossier_Fabrication.SECT, T_Dossier_Fabrication.TEINT,  "
        Sql = Sql & "T_Dossier_Fabrication.TEINT2, T_Dossier_Fabrication.ISO, T_Dossier_Fabrication.LONG, T_Dossier_Fabrication.[LONG CP],  "
        Sql = Sql & "T_Dossier_Fabrication.LONG_ADD, T_Dossier_Fabrication.LONG_ADD2, T_Dossier_Fabrication.COUPE, T_Dossier_Fabrication.POS,  "
        Sql = Sql & "T_Dossier_Fabrication.[POS-OUT], T_Dossier_Fabrication.FA, T_Dossier_Fabrication.APP, T_Dossier_Fabrication.VOI,  "
        Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR], T_Dossier_Fabrication.[REF CONNECTEUR_FOUR], T_Dossier_Fabrication.[REF CLIP],  "
        Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR], T_Dossier_Fabrication.PRECO, T_Dossier_Fabrication.[REF JOINT],  "
        Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR], T_Dossier_Fabrication.POS2, T_Dossier_Fabrication.[POS-OUT2],  "
        Sql = Sql & "T_Dossier_Fabrication.FA2, T_Dossier_Fabrication.APP2, T_Dossier_Fabrication.VOI2, T_Dossier_Fabrication.[REF CONNECTEUR2],  "
        Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR_FOUR2], T_Dossier_Fabrication.[REF CLIP2], T_Dossier_Fabrication.[REF CLIP FOUR2],  "
        Sql = Sql & "T_Dossier_Fabrication.PRECO2, T_Dossier_Fabrication.[REF JOINT2], T_Dossier_Fabrication.[REF JOINT FOUR2],  "
        Sql = Sql & "T_Dossier_Fabrication.PRECOG, T_Dossier_Fabrication.OPTION, T_Dossier_Fabrication.[CRITÈRES SPÉCIFIQUES] "
        Sql = Sql & "FROM T_Dossier_Fabrication "
        Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
Con.Execute Sql

    Sql = "INSERT INTO T_Critères ( Id_IndiceProjet, ACTIVER,CODE_CRITERE, CRITERES )  "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet,T_Critères.ACTIVER, T_Critères.CODE_CRITERE, T_Critères.CRITERES "
        Sql = Sql & "FROM T_Critères "
        Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

   Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
        
 Sql = "SELECT T_indiceProjet.* FROM T_indiceProjet WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
        Set Rs2 = Con.OpenRecordSet(Sql)
       
       
      

Dim FilSaveAs As String
Dim FilSource As String
Set Fso = New FileSystemObject

If Trim("" & Rs2!PlAutoCadSave) <> "" Then
    FilSource = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs2!PlAutoCadSave & ".dwg")
    FilSaveAs = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2.Fields("PI"), "Pl", Rs2.Fields("PL"), IdIndiceProjet, Rs2.Fields("PI_Indice"), Rs2.Fields("pl_Indice"), Rs2!Version) & ".dwg"
    If Fso.FileExists(FilSource) = True Then
        If Fso.FileExists(FilSaveAs) = True Then Fso.DeleteFile FilSaveAs, True
        Fso.CopyFile FilSource, FilSaveAs
    End If
End If



If Trim("" & Rs2!OUAutoCadSave) <> "" Then
    FilSource = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs2!OUAutoCadSave & ".dwg")
    FilSaveAs = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2.Fields("PI"), "Ou", Rs2.Fields("OU"), IdIndiceProjet, Rs2.Fields("Li_Indice"), Rs2.Fields("Ou_Indice"), Rs2!Version) & ".dwg"
    
    If Fso.FileExists(FilSource) = True Then
        If Fso.FileExists(FilSaveAs) = True Then Fso.DeleteFile FilSaveAs, True
        Fso.CopyFile FilSource, FilSaveAs
    End If
End If
End Select
PlanArchive = False
 Noquite = False
 Me.Hide
Modifier.Charge Me, "CommandButton1"
'Unload Modifier
'Unload Me

Fin:
End Sub

Private Sub CommandButton12_Click()

End Sub

Private Sub CommandButton2_Click()
Noquite = False
Me.Hide
End Sub




Private Sub CommandButton5_Click()


UserForm1.charger txt2, ";", "Equipement:", " "

End Sub

Private Sub CommandButton6_Click()

UserForm1.charger txt1, vbCrLf, "Ensemble:"


End Sub

'
Sub Maj(MyControl As Object)
Dim Rs As Recordset
Dim Sql As String

MyControl.Clear
Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    MyControl.AddItem Trim("" & Rs!Client)
        If MyControl.ListCount = 1 Then MyControl.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)

End Sub

Private Sub Label44_Click()

End Sub

Private Sub CommandButton3_Click()
Extension = "dwg"
FichierXLS = ScanFichier.chargement(Extension, FichierXLS)
Unload ScanFichier
End Sub

Private Sub OptionButton1_Click()
  EXCEL.Enabled = True
  Label1.Caption = "Chemin & nom du fichier AUTOCAD :"
  Extension = "dwg"
End Sub

Private Sub OptionButton2_Click()
EXCEL.Enabled = True
Label1.Caption = "Chemin & nom du fichier EXCEL :"
 Extension = "xls"
End Sub

Private Sub OptionButton3_Click()
EXCEL.Enabled = False
 Extension = ""
End Sub



Private Sub OptionButton7_Click()
 Extension = ""
End Sub

Private Sub UserForm_Activate()
Noquite = True
OptionButton1_Click
End Sub

Public Sub Charge(MyForm As Object, MyDroit As String)
Qui = MyDroit
    IdProjet = MyForm.IdProjet
 IdPieces = MyForm.IdPieces
 IdIndiceProjet = MyForm.IdIndiceProjet
 NbTxt = MyForm.NbTxt
For I = 1 To NbTxt
Debug.Print MyForm.Controls("txt" & CStr(I))
    Me.Controls("txt" & CStr(I)).Caption = MyForm.Controls("txt" & CStr(I))
Next I

Unload MyForm
Me.Show vbModal
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
Public Sub continuer(Optional aa As Boolean)
Dim Fso As New FileSystemObject
Me.Visible = True
UserForm2_boolExcute = UserForm2.boolExcute
       Unload UserForm2
If UserForm2_boolExcute = True Then
       
          '
                ImporteXls Me.FichierXLS, IdIndiceProjet
                If txtOption = "N" Then
                    If Fso.FileExists(Me.FichierXLS) = True Then
                        Fso.DeleteFile Me.FichierXLS
                    End If
                End If
          Else
              If txtOption = "N" Then
                    If Fso.FileExists(Me.FichierXLS) = True Then
                        Fso.DeleteFile Me.FichierXLS
                    End If
                End If
             Me.Enabled = True
              Me.FichierXLS = ""
             
          
          End If


PlanArchive = False
 Noquite = False
Modifier.Charge Me, Qui
'Unload Modifier
Me.Hide
End Sub
