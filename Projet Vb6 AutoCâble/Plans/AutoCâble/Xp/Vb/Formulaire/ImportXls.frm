VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportXls 
   Caption         =   "Créer un plan import des données :"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12180
   Icon            =   "ImportXls.dsx":0000
   OleObjectBlob   =   "ImportXls.dsx":08CA
   StartUpPosition =   1  'CenterOwner
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



Private Sub CommandButton1_Click()
Dim TxtOption As String
Dim Fso As New FileSystemObject
Dim Sql As String
Dim Rs As Recordset
Set FormBarGrah = Me

 Set TableauPath = funPath
 

If Me.OptionButton1.Value = True Then
    TxtOption = "A"
     Me.FichierXLS = Trim("" & Me.FichierXLS)
    If Trim("" & Me.FichierXLS) = "" Then
    MsgBox "Vous devez saisir le chemin ainsi que le nom du fichier AUTUCAD à importer", vbExclamation, "Erreur"
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
    TxtOption = "E"
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
    TxtOption = "N"
End If
  
If Me.OptionButton7.Value = True Then
    TxtOption = "P"
End If
  
Select Case TxtOption
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
        UserForm2.Chargement pathTmpXls, txt11, False
      UserForm2_boolExcute = UserForm2.boolExcute
       Unload UserForm2
         If UserForm2_boolExcute = True Then
       
          '
                ImporteXls Me.FichierXLS, IdIndiceProjet
                If TxtOption = "N" Then
                    If Fso.FileExists(Me.FichierXLS) = True Then
                        Fso.DeleteFile Me.FichierXLS
                    End If
                End If
          Else
              If TxtOption = "N" Then
                    If Fso.FileExists(Me.FichierXLS) = True Then
                        Fso.DeleteFile Me.FichierXLS
                    End If
                End If
             Me.Enabled = True
              Me.FichierXLS = ""
             Exit Sub
          
          End If
      Case "P"
        AffaireExistante.Show vbModal
      AffaireExistanteAnnuler = AffaireExistante.Annuler
      AffaireExistante_txt3_Tag = AffaireExistante.txt3.Tag
      Unload AffaireExistante
        If AffaireExistanteAnnuler = True Then Exit Sub
        Dim txtArchive As String
        If PlanArchive = True Then txtArchive = "Archive_"
      Sql = "UPDATE T_indiceProjet, " & txtArchive & "T_indiceProjet AS T_indiceProjet_1  "
        Sql = Sql & "SET T_indiceProjet.Masse = [T_indiceProjet_1].[Masse],  "
        Sql = Sql & "T_indiceProjet.PlAutoCadSaveAs = [T_indiceProjet_1].[PlAutoCadSaveAs],  "
        Sql = Sql & "T_indiceProjet.PlAutoCadSave = [T_indiceProjet_1].[PlAutoCadSave],  "
         Sql = Sql & "T_indiceProjet.OuAutoCadSaveAs = [T_indiceProjet_1].[OuAutoCadSaveAs],  "
        Sql = Sql & "T_indiceProjet.OuAutoCadSave = [T_indiceProjet_1].[OuAutoCadSave] ,"
        Sql = Sql & "T_indiceProjet.NbCartouche = [T_indiceProjet_1].[NbCartouche] "
        Sql = Sql & "WHERE T_indiceProjet.Id= " & IdIndiceProjet & " "
        Sql = Sql & "AND T_indiceProjet_1.Id=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

        
        

 Con.Exequte Sql
    

        
        
  
   Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet,ACTIVER, CONNECTEUR, [O/N],  "
        Sql = Sql & "DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1,  "
        Sql = Sql & "PRECO2, [100%] ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet, " & txtArchive & "Connecteurs.ACTIVER, "
        Sql = Sql & "" & txtArchive & "Connecteurs.CONNECTEUR, " & txtArchive & "Connecteurs.[O/N],  "
        Sql = Sql & "" & txtArchive & "Connecteurs.DESIGNATION, " & txtArchive & "Connecteurs.CODE_APP,  "
        Sql = Sql & "" & txtArchive & "Connecteurs.N°, " & txtArchive & "Connecteurs.POS, " & txtArchive & "Connecteurs.[POS-OUT],  "
        Sql = Sql & "" & txtArchive & "Connecteurs.PRECO1, " & txtArchive & "Connecteurs.PRECO2, " & txtArchive & "Connecteurs.[100%] "
        Sql = Sql & "FROM " & txtArchive & "Connecteurs "
        Sql = Sql & "WHERE " & txtArchive & "Connecteurs.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

    Con.Exequte Sql
    
    
   
        Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet,ACTIVER, LIAI, DESIGNATION, "
        Sql = Sql & "FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, "
        Sql = Sql & "[POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2, VOI2, PRECO, [OPTION] )"
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet," & txtArchive & "Ligne_Tableau_fils.ACTIVER, " & txtArchive & "Ligne_Tableau_fils.LIAI, "
        Sql = Sql & "" & txtArchive & "Ligne_Tableau_fils.DESIGNATION, " & txtArchive & "Ligne_Tableau_fils.FIL, "
        Sql = Sql & "" & txtArchive & "Ligne_Tableau_fils.SECT, " & txtArchive & "Ligne_Tableau_fils.TEINT, " & txtArchive & "Ligne_Tableau_fils.TEINT2, "
        Sql = Sql & "" & txtArchive & "Ligne_Tableau_fils.ISO, " & txtArchive & "Ligne_Tableau_fils.LONG, " & txtArchive & "Ligne_Tableau_fils.[LONG CP], "
        Sql = Sql & "" & txtArchive & "Ligne_Tableau_fils.COUPE, " & txtArchive & "Ligne_Tableau_fils.POS, " & txtArchive & "Ligne_Tableau_fils.[POS-OUT], "
        Sql = Sql & "" & txtArchive & "Ligne_Tableau_fils.FA, " & txtArchive & "Ligne_Tableau_fils.APP, " & txtArchive & "Ligne_Tableau_fils.VOI, "
        Sql = Sql & "" & txtArchive & "Ligne_Tableau_fils.POS2, " & txtArchive & "Ligne_Tableau_fils.[POS-OUT2], " & txtArchive & "Ligne_Tableau_fils.FA2, "
        Sql = Sql & "" & txtArchive & "Ligne_Tableau_fils.APP2, " & txtArchive & "Ligne_Tableau_fils.VOI2, " & txtArchive & "Ligne_Tableau_fils.PRECO, "
        Sql = Sql & "" & txtArchive & "Ligne_Tableau_fils.OPTION "
        Sql = Sql & "FROM " & txtArchive & "Ligne_Tableau_fils "
        Sql = Sql & "WHERE " & txtArchive & "Ligne_Tableau_fils.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
 Con.Exequte Sql
    
    
    Sql = "INSERT INTO Composants (  Id_IndiceProjet,ACTIVER, DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet," & txtArchive & "Composants.ACTIVER,  "
        Sql = Sql & "" & txtArchive & "Composants.DESIGNCOMP, " & txtArchive & "Composants.NUMCOMP,  "
        Sql = Sql & "" & txtArchive & "Composants.REFCOMP, " & txtArchive & "Composants.Path "
        Sql = Sql & "FROM " & txtArchive & "Composants "
         Sql = Sql & "WHERE " & txtArchive & "Composants.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
Con.Exequte Sql

Sql = "INSERT INTO nota ( Id_IndiceProjet,ACTIVER,NOTA, NUMNOTA ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet," & txtArchive & "Nota.ACTIVER, " & txtArchive & "Nota.NOTA, " & txtArchive & "Nota.NUMNOTA "
        Sql = Sql & "FROM " & txtArchive & "Nota "
        Sql = Sql & "WHERE " & txtArchive & "Nota.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

Con.Exequte Sql
        Sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet, ACTIVER, NŒUDS,  "
        Sql = Sql & "LONGUEUR, DESIGN_HAB, CODE_RSA, CODE_PSA, CODE_ENC,  "
        Sql = Sql & "DIAMETRE, CLASSE_T, TORON_PRINCIPAL, LONGUEUR_CUMULEE,  "
        Sql = Sql & "Fleche_Droite ) "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet, " & txtArchive & "T_Noeuds.ACTIVER, " & txtArchive & "T_Noeuds.NŒUDS, "
          Sql = Sql & "" & txtArchive & "T_Noeuds.LONGUEUR,  "
        Sql = Sql & "" & txtArchive & "T_Noeuds.DESIGN_HAB, " & txtArchive & "T_Noeuds.CODE_RSA, " & txtArchive & "T_Noeuds.CODE_PSA, " & txtArchive & "T_Noeuds.CODE_ENC,  "
        Sql = Sql & "" & txtArchive & "T_Noeuds.DIAMETRE, " & txtArchive & "T_Noeuds.CLASSE_T, " & txtArchive & "T_Noeuds.TORON_PRINCIPAL, " & txtArchive & "T_Noeuds.LONGUEUR_CUMULEE,  "
        Sql = Sql & "" & txtArchive & "T_Noeuds.Fleche_Droite "
        Sql = Sql & "FROM " & txtArchive & "T_Noeuds "
        Sql = Sql & "WHERE " & txtArchive & "T_Noeuds.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"
Con.Exequte Sql

    Sql = "INSERT INTO T_Critères ( Id_IndiceProjet, ACTIVER,CODE_CRITERE, CRITERES )  "
        Sql = Sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet," & txtArchive & "T_Critères.ACTIVER, " & txtArchive & "T_Critères.CODE_CRITERE, " & txtArchive & "T_Critères.CRITERES "
        Sql = Sql & "FROM " & txtArchive & "T_Critères "
        Sql = Sql & "WHERE " & txtArchive & "T_Critères.Id_IndiceProjet=" & Trim("" & AffaireExistante_txt3_Tag) & ";"

       
End Select
PlanArchive = False
 Noquite = False
Modifier.Charge Me
Unload Modifier
Me.Hide
Fin:
End Sub

Private Sub CommandButton12_Click()

End Sub

Private Sub CommandButton2_Click()
Noquite = False
Me.Hide
End Sub




Private Sub CommandButton5_Click()


UserForm1.Charger txt2, ";", "Equipement:", " "

End Sub

Private Sub CommandButton6_Click()

UserForm1.Charger txt1, vbCrLf, "Ensemble:"


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
FichierXLS = ScanFichier.Chargement(Extension, FichierXLS)
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

Public Sub Charge(MyForm As Object)
    IdProjet = MyForm.IdProjet
 IdPieces = MyForm.IdPieces
 IdIndiceProjet = MyForm.IdIndiceProjet
 NbTxt = MyForm.NbTxt
For i = 1 To NbTxt
Debug.Print MyForm.Controls("txt" & CStr(i))
    Me.Controls("txt" & CStr(i)).Caption = MyForm.Controls("txt" & CStr(i))
Next i
MyForm.Hide
Me.Show vbModal
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
