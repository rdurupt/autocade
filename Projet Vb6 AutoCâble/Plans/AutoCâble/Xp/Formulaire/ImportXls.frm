VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportXls 
   Caption         =   "Créer un plan import des données :"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12180
   OleObjectBlob   =   "ImportXls.frx":0000
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
Dim sql As String
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
 
pathTmpXls = Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt6.Caption, ":", "_", 1) & ".XLS"
     If Fso.FileExists(pathTmpXls) = True Then
        Fso.DeleteFile pathTmpXls
    End If
        
        ExporteXls pathTmpXls, IdIndiceProjet
        UserForm2.chargement pathTmpXls, txt11
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
             
             Exit Sub
          
          End If
      Case "P"
        AffaireExistante.Show
     
        If AffaireExistante.Annuler = True Then Exit Sub
        Dim txtArchive As String
        If PlanArchive = True Then txtArchive = "Archive_"
      sql = "UPDATE T_indiceProjet, " & txtArchive & "T_indiceProjet AS T_indiceProjet_1  "
        sql = sql & "SET T_indiceProjet.Masse = [T_indiceProjet_1].[Masse],  "
        sql = sql & "T_indiceProjet.PlAutoCadSaveAs = [T_indiceProjet_1].[PlAutoCadSaveAs],  "
        sql = sql & "T_indiceProjet.PlAutoCadSave = [T_indiceProjet_1].[PlAutoCadSave],  "
         sql = sql & "T_indiceProjet.OuAutoCadSaveAs = [T_indiceProjet_1].[OuAutoCadSaveAs],  "
        sql = sql & "T_indiceProjet.OuAutoCadSave = [T_indiceProjet_1].[OuAutoCadSave] "
        sql = sql & "WHERE T_indiceProjet.Id= " & IdIndiceProjet & " "
        sql = sql & "AND T_indiceProjet_1.Id=" & Trim("" & AffaireExistante.txt3.Tag) & ";"

        
        

 Con.Exequte sql
    

        
        
  
   sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR, [O/N],  "
        sql = sql & "DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1,  "
        sql = sql & "PRECO2, [100%] ) "
        sql = sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet,  "
        sql = sql & "" & txtArchive & "Connecteurs.CONNECTEUR, " & txtArchive & "Connecteurs.[O/N],  "
        sql = sql & "" & txtArchive & "Connecteurs.DESIGNATION, " & txtArchive & "Connecteurs.CODE_APP,  "
        sql = sql & "" & txtArchive & "Connecteurs.N°, " & txtArchive & "Connecteurs.POS, " & txtArchive & "Connecteurs.[POS-OUT],  "
        sql = sql & "" & txtArchive & "Connecteurs.PRECO1, " & txtArchive & "Connecteurs.PRECO2, " & txtArchive & "Connecteurs.[100%] "
        sql = sql & "FROM " & txtArchive & "Connecteurs "
        sql = sql & "WHERE " & txtArchive & "Connecteurs.Id_IndiceProjet=" & Trim("" & AffaireExistante.txt3.Tag) & ";"

    Con.Exequte sql
    
    
   
        sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION, "
        sql = sql & "FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, "
        sql = sql & "[POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2, VOI2, PRECO, [OPTION] )"
        sql = sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet, " & txtArchive & "Ligne_Tableau_fils.LIAI, "
        sql = sql & "" & txtArchive & "Ligne_Tableau_fils.DESIGNATION, " & txtArchive & "Ligne_Tableau_fils.FIL, "
        sql = sql & "" & txtArchive & "Ligne_Tableau_fils.SECT, " & txtArchive & "Ligne_Tableau_fils.TEINT, " & txtArchive & "Ligne_Tableau_fils.TEINT2, "
        sql = sql & "" & txtArchive & "Ligne_Tableau_fils.ISO, " & txtArchive & "Ligne_Tableau_fils.LONG, " & txtArchive & "Ligne_Tableau_fils.[LONG CP], "
        sql = sql & "" & txtArchive & "Ligne_Tableau_fils.COUPE, " & txtArchive & "Ligne_Tableau_fils.POS, " & txtArchive & "Ligne_Tableau_fils.[POS-OUT], "
        sql = sql & "" & txtArchive & "Ligne_Tableau_fils.FA, " & txtArchive & "Ligne_Tableau_fils.APP, " & txtArchive & "Ligne_Tableau_fils.VOI, "
        sql = sql & "" & txtArchive & "Ligne_Tableau_fils.POS2, " & txtArchive & "Ligne_Tableau_fils.[POS-OUT2], " & txtArchive & "Ligne_Tableau_fils.FA2, "
        sql = sql & "" & txtArchive & "Ligne_Tableau_fils.APP2, " & txtArchive & "Ligne_Tableau_fils.VOI2, " & txtArchive & "Ligne_Tableau_fils.PRECO, "
        sql = sql & "" & txtArchive & "Ligne_Tableau_fils.OPTION "
        sql = sql & "FROM " & txtArchive & "Ligne_Tableau_fils "
        sql = sql & "WHERE " & txtArchive & "Ligne_Tableau_fils.Id_IndiceProjet=" & Trim("" & AffaireExistante.txt3.Tag) & ";"
 Con.Exequte sql
    
    
    sql = "INSERT INTO Composants (  Id_IndiceProjet, DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
        sql = sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet,  "
        sql = sql & "" & txtArchive & "Composants.DESIGNCOMP, " & txtArchive & "Composants.NUMCOMP,  "
        sql = sql & "" & txtArchive & "Composants.REFCOMP, " & txtArchive & "Composants.Path "
        sql = sql & "FROM Archive_Composants "
         sql = sql & "WHERE " & txtArchive & "Composants.Id_IndiceProjet=" & Trim("" & AffaireExistante.txt3.Tag) & ";"
Con.Exequte sql

sql = "INSERT INTO nota ( Id_IndiceProjet, NOTA, NUMNOTA ) "
        sql = sql & "SELECT " & IdIndiceProjet & " AS Id_IndiceProjet, " & txtArchive & "Nota.NOTA, " & txtArchive & "Nota.NUMNOTA "
        sql = sql & "FROM " & txtArchive & "Nota "
        sql = sql & "WHERE " & txtArchive & "Nota.Id_IndiceProjet=" & Trim("" & AffaireExistante.txt3.Tag) & ";"

Con.Exequte sql
        
    Unload AffaireExistante
       
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
Dim sql As String

MyControl.Clear
sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
sql = sql & "FROM T_Clients "
sql = sql & "ORDER BY T_Clients.Client;"

Set Rs = Con.OpenRecordSet(sql)
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
FichierXLS = ScanFichier.chargement(Extension, FichierXLS)
Unload ScanFichier
End Sub

Private Sub Label35_Click()

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
Me.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
