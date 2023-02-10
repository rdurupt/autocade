VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmAutocâble 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Autocâble:"
   ClientHeight    =   12060
   ClientLeft      =   1050
   ClientTop       =   585
   ClientWidth     =   17460
   Icon            =   "frmAutocâble.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmAutocâble.frx":0A7A
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutocâble.frx":A8BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAutocâble.frx":A8EFE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   11490
      Width           =   17460
      _ExtentX        =   30798
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   11685
      Width           =   17460
      _ExtentX        =   30798
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4419
            MinWidth        =   4419
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   28681
            MinWidth        =   28681
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1482
            MinWidth        =   1482
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   2
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu Fichier 
      Caption         =   "Fichier"
      Begin VB.Menu CommandButton1 
         Caption         =   "Nouveau"
      End
      Begin VB.Menu CommandButton7 
         Caption         =   "Modifier"
      End
      Begin VB.Menu CommandButton5 
         Caption         =   "Vérifier"
      End
      Begin VB.Menu CommandButton9 
         Caption         =   "Approuver"
      End
      Begin VB.Menu Supprimer_Archiver 
         Caption         =   "Supprimer/Archiver"
      End
      Begin VB.Menu ImpArch 
         Caption         =   "Importer Archive"
      End
      Begin VB.Menu Visu 
         Caption         =   "Visualiser"
         Begin VB.Menu CommandButton31 
            Caption         =   "Synthèse"
         End
         Begin VB.Menu CommandButton29 
            Caption         =   "Etude"
         End
         Begin VB.Menu Utliser_Par 
            Caption         =   "Utliser Par"
         End
      End
      Begin VB.Menu CommandButton21 
         Caption         =   "Suivi des Jobs"
         Shortcut        =   ^J
      End
      Begin VB.Menu Autre_Utilisateur 
         Caption         =   "Changer d'Utilisateur"
      End
      Begin VB.Menu Quitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Editer"
      Begin VB.Menu CommandButton6 
         Caption         =   "Etude"
         Shortcut        =   ^E
      End
      Begin VB.Menu CommandButton11 
         Caption         =   "Nomenclature"
         Begin VB.Menu Nomenclature0 
            Caption         =   "Prénomenclature par Code Appareil"
         End
         Begin VB.Menu Nomenclature1 
            Caption         =   "Préparer nomenclature par Code Appareil"
         End
         Begin VB.Menu Nomenclature2 
            Caption         =   "Nomenclature par Code Appareil"
         End
         Begin VB.Menu Nomenclature3 
            Caption         =   "Préparer la liste d'approvisinnement finale"
         End
         Begin VB.Menu MajEboutique 
            Caption         =   "Maj Stock eboutique"
         End
      End
      Begin VB.Menu Exporter 
         Caption         =   "Documents"
         Begin VB.Menu CommandButton19 
            Caption         =   "Préparer documents"
         End
         Begin VB.Menu CommandButton30 
            Caption         =   "Etiquette"
         End
         Begin VB.Menu CommandButton22 
            Caption         =   "Etats"
         End
      End
      Begin VB.Menu BD 
         Caption         =   "Base Données"
         Begin VB.Menu Projets 
            Caption         =   "Liste Projets "
         End
         Begin VB.Menu Equipement 
            Caption         =   "Equipement"
         End
         Begin VB.Menu Vagues 
            Caption         =   "Vagues"
         End
         Begin VB.Menu Ensemble 
            Caption         =   "Ensemble"
         End
         Begin VB.Menu Codes_Liaisons 
            Caption         =   "Codes Liaisons"
         End
         Begin VB.Menu A_Client 
            Caption         =   "Client"
         End
         Begin VB.Menu Habillage 
            Caption         =   "Habillage"
         End
      End
   End
   Begin VB.Menu Autre 
      Caption         =   "Outils"
      Begin VB.Menu CommandButton10 
         Caption         =   "Liens"
         Begin VB.Menu Utilitaire 
            Caption         =   "Utilitaire"
            Index           =   0
         End
      End
      Begin VB.Menu Modules 
         Caption         =   "Modules"
         Begin VB.Menu ModuleDetail 
            Caption         =   "ModuleDetail"
            Index           =   0
         End
      End
      Begin VB.Menu A_Utilitaires 
         Caption         =   "Ajouter Modfier Liens"
      End
      Begin VB.Menu A_Module 
         Caption         =   "Ajouter Modfier Module"
      End
      Begin VB.Menu Connecteur 
         Caption         =   "Connecteur"
         Begin VB.Menu CommandButton14 
            Caption         =   "Test Connecteur(s)"
         End
         Begin VB.Menu CommandButton18 
            Caption         =   "Créer Attributs"
         End
      End
   End
   Begin VB.Menu Fenêtre 
      Caption         =   "Fenêtre"
      WindowList      =   -1  'True
      Begin VB.Menu MosVer 
         Caption         =   "Mosaïque Verticale"
      End
      Begin VB.Menu MosHor 
         Caption         =   "Mosaïque Horizontal"
      End
      Begin VB.Menu cascade 
         Caption         =   "En cascade"
      End
   End
   Begin VB.Menu Admin 
      Caption         =   "Administration"
      Begin VB.Menu CommandButton27 
         Caption         =   "Menu Magasin"
         Begin VB.Menu Import_Cables 
            Caption         =   "Import Câbles"
         End
         Begin VB.Menu Export_Cables 
            Caption         =   "Export Câbles"
         End
         Begin VB.Menu Import_Habillage 
            Caption         =   "Import Habillage"
         End
         Begin VB.Menu Export_Habillage 
            Caption         =   "Export Habillage"
         End
      End
      Begin VB.Menu Gestion_droits 
         Caption         =   "Gestion des droits"
         Begin VB.Menu A_Utilisateur 
            Caption         =   "Utilisateur"
         End
         Begin VB.Menu A_Groupe 
            Caption         =   "Groupe"
         End
         Begin VB.Menu Message_Droits 
            Caption         =   "Message Droits"
         End
      End
      Begin VB.Menu A_Répertoires 
         Caption         =   "Répertoires"
      End
      Begin VB.Menu SMTP 
         Caption         =   "Serveur SMTP"
      End
      Begin VB.Menu A_Boutons 
         Caption         =   "Boutons"
      End
      Begin VB.Menu Generateur_Etats 
         Caption         =   "Generateur d 'Etats"
      End
   End
   Begin VB.Menu SUBAddAtrib 
      Caption         =   "AddAtrib"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmAutocâble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MenuOf As Boolean
Dim txt1 As Object



Private Sub AG_Utilisateur_Click()
EditUser.Show vbModal
End Sub





Private Sub A_Boutons_Click()
MenuSys.Show vbModal

End Sub

Private Sub A_Client_Click()
UserForm3.Show vbModal

End Sub

Private Sub A_Groupe_Click()
EditGroupe.Show vbModal
End Sub

Private Sub A_Module_Click()
ModuleListes.Show vbModal
End Sub

Private Sub A_Répertoires_Click()
 RepSystem.Show vbModal
End Sub

Private Sub A_Utilisateur_Click()
EditUser.Show vbModal
End Sub

Private Sub A_Utilitaires_Click()
UtilitairesListes.Show vbModal

End Sub

Private Sub AddAtrib_Click()
AddAtrib
End Sub

Private Sub Autre_Utilisateur_Click()
GestionDesDroit "Application"
End Sub

Private Sub cascade_Click()
Me.Arrange vbCascade
End Sub

Private Sub Codes_Liaisons_Click()
Set FormBarGrah = Me
MousePointer = fmMousePointerHourGlass
UserForm6.chargement
Unload UserForm6
MousePointer = fmMousePointerDefault
End Sub

Private Sub CommandButton1_Click()
'If GestionDesDroit("CommandButton1") = True Then
SubCreer CommandButton1.Caption

End Sub

Private Sub CommandButton10_Click()
'If GestionDesDroit("CommandButton10") = True Then
'Utilitaire
End Sub

Private Sub CommandButton11_Click()

'If GestionDesDroit("CommandButton11") = True Then
'subExporter
End Sub

Private Sub CommandButton14_Click()
If boolAutoCAD = False Then
    MsgBox "Vous ne possédez pas de licence AutoCad." & vbCrLf & "Vous ne pouvez pas effectuer ce test."
Else
Set FormBarGrah = Me
LireRepEval
'If GestionDesDroit("CommandButton14") = True Then LireRepEval
End If

End Sub

Private Sub CommandButton18_Click()
If boolAutoCAD = False Then
    MsgBox "Vous ne possédez pas de licence AutoCad." & vbCrLf & "Vous ne pouvez pas effectuer ce test."
Else
AddAtrib
'If GestionDesDroit("CommandButton18") = True Then AddAtrib
End If
End Sub

Private Sub CommandButton19_Click()

'If GestionDesDroit("CommandButton19") = True Then
subTestWord
LibertPice
End Sub

Private Sub CommandButton21_Click()
'If GestionDesDroit("CommandButton21") = True Then
subJob
End Sub

Private Sub CommandButton22_Click()

'If GestionDesDroit("CommandButton22") = True Then
subGenerateur
End Sub

Private Sub CommandButton26_Click()
'If GestionDesDroit("CommandButton26") = True Then

End Sub

Private Sub CommandButton27_Click()
'If GestionDesDroit("CommandButton27") = True Then
End Sub

Private Sub CommandButton29_Click()
'EnabledMenu
'If GestionDesDroit("CommandButton29") = True Then
    
    subEDITER CommandButton29.Caption, True
    
'End If
End Sub

Private Sub CommandButton30_Click()

subGenerateurEtiquette
End Sub

Private Sub CommandButton31_Click()
'If GestionDesDroit("CommandButton31") = True Then
Synthese

End Sub

Private Sub CommandButton5_Click()
'EnabledMenu
'If GestionDesDroit("CommandButton5") = True Then
subVerifierEtude
LibertPice
End Sub

Private Sub CommandButton6_Click()
EnabledMenu
'If GestionDesDroit("CommandButton6") = True Then
subEDITER CommandButton1.Caption, False
End Sub

Private Sub CommandButton7_Click()
'If GestionDesDroit("CommandButton7") = True Then
subModifierCartouche
LibertPice
End Sub

Private Sub CommandButton8_Click()
'If GestionDesDroit("CommandButton8") = True Then
subUtilisateur
End Sub

Private Sub CommandButton9_Click()
'If GestionDesDroit("CommandButton9") = True Then
subApprobation
LibertPice
End Sub

Private Sub Ensemble_Click()
On Error Resume Next
UserForm1.charger txt1, " ", "Ensemble:", " "
Unload UserForm1
txt1 = ""

End Sub

Private Sub Equipement_Click()
On Error Resume Next
UserForm1.charger txt1, " ", "Equipement:", " "
Unload UserForm1
txt1 = ""
Set txt1 = Nothing
End Sub

Private Sub Export_Cables_Click()
XlsPrix = "CablePrix"
subImport
End Sub

Private Sub Export_Habillage_Click()
XlsPrix = "CablePrix"
subImport
End Sub

Private Sub Generateur_Etats_Click()
FrmEtats.Show vbModal
End Sub

Private Sub Habillage_Click()
FrmHabillage.chargement
End Sub

Private Sub ImpArch_Click()
UserForm5.Charge Me, "IdStatus=4"
Unload UserForm5

End Sub

Private Sub Import_Cables_Click()
XlsPrix = "CablePrix"
subImport
End Sub

Private Sub Import_Habillage_Click()
XlsPrix = "CablePrix"
subImport
End Sub


Private Sub MajEboutique_Click()
MajStockEboutique.Show
LibertPice

End Sub

Private Sub MDIForm_Activate()
Dim NbMenu As Long
Dim MyControl As New Collection
Dim Rs As Recordset
Dim I As Long
Dim Sql As String
'Dim sql As String
'Dim Rs As Recordset

'LoadDb

'UPDATE T_indiceProjet SET T_indiceProjet.UserName = Null
'WHERE (((T_indiceProjet.Id)=94));

Set Rs = Con.OpenRecordSet("SELECT T_Boutons.Bouton, T_Boutons.Name FROM T_Boutons where T_Boutons.ContonTotal=false ;")
'While Rs.EOF = False
'    Me.Controls(MyControl(Rs!Name)).Caption = Trim("" & Rs!Bouton)
'    Rs.MoveNext
'Wend

'Me.CommandButton10.Visible = True
'For I = Me.ModuleDetail.Count - 1 To 1 Step -1
'    If I <> 0 Then
'        Unload Me.ModuleDetail(I)
'    End If
'Next
'Load Me.ModuleDetail(1)
'Me.ModuleDetail(1).Visible = True
'Unload Me.ModuleDetail(0)
NotSaveRacourci = True
Bool_Fichier_Li = False
NoClose = True
'Form1.Visible = True
End Sub

Private Sub MDIForm_Initialize()
On Error Resume Next
'ChDir "C:\Program Files\AutoCAD 2002 Fra\"
NomenclatureOk = True
If IsCilent = False Then
    boolAutoCAD = True
    If MsgBox("Voulez vous ouvrir une licence AUTOCAD.", vbQuestion + vbYesNo) = vbYes Then
        SetAutocad
    Else
        boolAutoCAD = False
    End If
End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 If ArretKill = True Then
 Call keybd_event(145, 1, 0, 0)
End If
End Sub

Private Sub ModuleDetail_Click(Index As Integer)
Dim Fso As FileSystemObject
Dim Sql As String
Dim Rs As Recordset
'If Me.List1.ListIndex = -1 Then
'    MsgBox "devez sélectionner un utilitaire.", vbExclamation
'    Exit Sub
'End If
Sql = "SELECT Module.Utilitaire FROM Module "
Sql = Sql & "WHERE Module.NameBouton='" & Me.ModuleDetail(Index).Caption & "' "
Sql = Sql & "ORDER BY Module.NameBouton;"

Set Rs = Con.OpenRecordSet(Sql)
'Set Fso = New FileSystemObject
If Rs.EOF = False Then
'   If Fso.FileExists("" & Rs!Utilitaire) = True Then
'   MsgBox ""
'   End If
     'Execute explorer.exe
     '"\\autocable\Autocable Access\AutoCable Client\SychroXml.exe"  "[N est pas peur fils papa est la]"
    MyExecute "" & Rs!Utilitaire, "[N ais pas peur fils papa est la]"
     'Execute la calculette
''    WinExec "Calc.exe", 1
     'autre astuce :Le vrai mode plein ecran est là !!
'    Shell "c:\program files\internet explorer\iexplore -k c:"
'    Shell "" & Rs!Utilitaire, vbMaximizedFocus
End If
Set Rs = Con.CloseRecordSet(Rs)

End Sub

Private Sub SUBAddAtrib_Click()
AddAtrib
End Sub

Private Sub Timer1_Timer()
Dim a As Long
Dim Bouton As Long
DoEvents
'************************
'   Code de TFlorian ;-)
'  TFlorian@IFrance.com
'------------------------
Static SaveTouche As Long
Static SaveTouche2 As Long
Dim msg As String
Dim NumBouton As Integer
Me.Timer1.Enabled = False
Me.StatusBar1.Panels(5).Text = Time
Me.StatusBar1.Panels(5).ToolTipText = Date
NumBouton = 0
For a = 0 To 1 'on scanne toute les touche du clavier
 If NumBouton = 0 Then
    NumBouton = 19
 Else
    If NumBouton = 19 Then NumBouton = 145
 End If
DoEvents
Bouton = GetAsyncKeyState(NumBouton)

If Bouton <> 0 Then 'filtre si la touche consideree a ete appuiller
Select Case NumBouton
    Case 19
      
     
    Restart.Show vbModal
   
    SaveTouche = 0
    

    Case 145
    If PossibleArretKill = True Then
   
    If -32767 = Bouton Then
    
 DoEvents
        If ArretKill = True Then
             ArretKill = False
             msg = "Arrêt après suppression annulé."
                Me.StatusBar1.Panels(3).Text = ""
        Else
            ArretKill = True
                msg = "Arrêt après suppression activé."
                Me.StatusBar1.Panels(3).Text = "Arrêt Activé."
        End If
        MsgBox msg
    End If
    End If
End Select
  
SaveTouche = a
'    Select Case a
'        Case 0:
'        Case 1: Text1.Text = "Bouton gauche de la souris" & vbCrLf & Text1.Text ' vbKeyLButton
'        Case 2: Text1.Text = "Bouton droit de la souris" & vbCrLf & Text1.Text ' vbKeyRButton
'        Case 3: 'je sais pas a quoi elle correspond ! -> elle est toujour appuiller
'        'Case 3: Text1.Text = "Touche ANNUL" & vbCrLf & Text1.Text ' vbKeyCancel
'        Case 4: Text1.Text = "Bouton central de la souris" & vbCrLf & Text1.Text ' vbKeyMButton
'        Case 8: Text1.Text = "Touche RET.ARR" & vbCrLf & Text1.Text ' vbKeyBack
'        Case 9: Text1.Text = "Touche TAB" & vbCrLf & Text1.Text ' vbKeyTab
'        Case 12: Text1.Text = "Touche EFFACER" & vbCrLf & Text1.Text ' vbKeyClear
'        Case 13: Text1.Text = "Touche ENTRÉE" & vbCrLf & Text1.Text ' vbKeyReturn
'        Case 16: Text1.Text = "Touche MAJ" & vbCrLf & Text1.Text ' vbKeyShift
'        Case 17: Text1.Text = "Touche CTRL" & vbCrLf & Text1.Text ' vbKeyControl
'        Case 18: Text1.Text = "Touche MENU" & vbCrLf & Text1.Text ' vbKeyMenu
'        Case 19: Text1.Text = "Touche PAUSE" & vbCrLf & Text1.Text ' vbKeyPause
'        Case 20: Text1.Text = "Touche VERR.MAJ" & vbCrLf & Text1.Text ' vbKeyCapital
'        Case 27: Text1.Text = "Touche ÉCHAP." & vbCrLf & Text1.Text ' vbKeyEscape
'        Case 32: Text1.Text = "Touche ESPACE" & vbCrLf & Text1.Text ' vbKeySpace
'        Case 33: Text1.Text = "Touche PG PRÉC." & vbCrLf & Text1.Text ' vbKeyPageUp
'        Case 34: Text1.Text = "Touche PG SUIV." & vbCrLf & Text1.Text ' vbKeyPageDown
'        Case 35: Text1.Text = "Touche FIN" & vbCrLf & Text1.Text ' vbKeyEnd
'        Case 36: Text1.Text = "Touche ORIGINE" & vbCrLf & Text1.Text ' vbKeyHome
'        Case 37: Text1.Text = "Touche FLÈCHE VERS LA GAUCHE " & vbCrLf & Text1.Text ' vbKeyLeft
'        Case 38: Text1.Text = "Touche FLÈCHE VERS LE HAUT " & vbCrLf & Text1.Text ' vbKeyUp
'        Case 39: Text1.Text = "Touche FLÈCHE VERS LA DROITE " & vbCrLf & Text1.Text ' vbKeyRight
'        Case 40: Text1.Text = "Touche FLÈCHE VERS LE BAS " & vbCrLf & Text1.Text ' vbKeyDown
'        Case 41: Text1.Text = "Touche SELECT" & vbCrLf & Text1.Text ' vbKeySelect
'        Case 42: Text1.Text = "Touche IMPR.ÉCRAN" & vbCrLf & Text1.Text ' vbKeyPrint
'        Case 43: Text1.Text = "Touche EXÉCUTE" & vbCrLf & Text1.Text ' vbKeyExecute
'        Case 44: Text1.Text = "Touche INSTANTANÉ" & vbCrLf & Text1.Text ' vbKeySnapshot
'        Case 45: Text1.Text = "Touche INSER" & vbCrLf & Text1.Text ' vbKeyInsert
'        Case 46: Text1.Text = "Touche SUPPR." & vbCrLf & Text1.Text ' vbKeyDelete
'        Case 47: Text1.Text = "Touche AIDE" & vbCrLf & Text1.Text ' vbKeyHelp
'        Case 48: Text1.Text = "Touche 0" & vbCrLf & Text1.Text ' vbKey0
'        Case 49: Text1.Text = "Touche 1" & vbCrLf & Text1.Text ' vbKey1
'        Case 50: Text1.Text = "Touche 2" & vbCrLf & Text1.Text ' vbKey2
'        Case 51: Text1.Text = "Touche 3" & vbCrLf & Text1.Text ' vbKey3
'        Case 52: Text1.Text = "Touche 4" & vbCrLf & Text1.Text ' vbKey4
'        Case 53: Text1.Text = "Touche 5" & vbCrLf & Text1.Text ' vbKey5
'        Case 54: Text1.Text = "Touche 6" & vbCrLf & Text1.Text ' vbKey6
'        Case 55: Text1.Text = "Touche 7" & vbCrLf & Text1.Text ' vbKey7
'        Case 56: Text1.Text = "Touche 8" & vbCrLf & Text1.Text ' vbKey8
'        Case 57: Text1.Text = "Touche 9" & vbCrLf & Text1.Text ' vbKey9
'        Case 65: Text1.Text = "Touche A" & vbCrLf & Text1.Text ' vbKeyA
'        Case 66: Text1.Text = "Touche B" & vbCrLf & Text1.Text ' vbKeyB
'        Case 67: Text1.Text = "Touche C" & vbCrLf & Text1.Text ' vbKeyC
'        Case 68: Text1.Text = "Touche D" & vbCrLf & Text1.Text ' vbKeyD
'        Case 69: Text1.Text = "Touche E" & vbCrLf & Text1.Text ' vbKeyE
'        Case 70: Text1.Text = "Touche F" & vbCrLf & Text1.Text ' vbKeyF
'        Case 71: Text1.Text = "Touche G" & vbCrLf & Text1.Text ' vbKeyG
'        Case 72: Text1.Text = "Touche H" & vbCrLf & Text1.Text ' vbKeyH
'        Case 73: Text1.Text = "Touche I" & vbCrLf & Text1.Text ' vbKeyI
'        Case 74: Text1.Text = "Touche J" & vbCrLf & Text1.Text ' vbKeyJ
'        Case 75: Text1.Text = "Touche K" & vbCrLf & Text1.Text ' vbKeyK
'        Case 76: Text1.Text = "Touche L" & vbCrLf & Text1.Text ' vbKeyL
'        Case 77: Text1.Text = "Touche M" & vbCrLf & Text1.Text ' vbKeyM
'        Case 78: Text1.Text = "Touche N" & vbCrLf & Text1.Text ' vbKeyN
'        Case 79: Text1.Text = "Touche O" & vbCrLf & Text1.Text ' vbKeyO
'        Case 80: Text1.Text = "Touche P" & vbCrLf & Text1.Text ' vbKeyP
'        Case 81: Text1.Text = "Touche Q" & vbCrLf & Text1.Text ' vbKeyQ
'        Case 82: Text1.Text = "Touche R" & vbCrLf & Text1.Text ' vbKeyR
'        Case 83: Text1.Text = "Touche S" & vbCrLf & Text1.Text ' vbKeyS
'        Case 84: Text1.Text = "Touche T" & vbCrLf & Text1.Text ' vbKeyT
'        Case 85: Text1.Text = "Touche U" & vbCrLf & Text1.Text ' vbKeyU
'        Case 86: Text1.Text = "Touche V" & vbCrLf & Text1.Text ' vbKeyV
'        Case 87: Text1.Text = "Touche W" & vbCrLf & Text1.Text ' vbKeyW
'        Case 88: Text1.Text = "Touche X" & vbCrLf & Text1.Text ' vbKeyX
'        Case 89: Text1.Text = "Touche Y" & vbCrLf & Text1.Text ' vbKeyY
'        Case 90: Text1.Text = "Touche Z" & vbCrLf & Text1.Text '  vbKeyZ
'        Case 91: Text1.Text = "Touche Windows gauche" & vbCrLf & Text1.Text '
'        Case 92: Text1.Text = "Touche contextuel" & vbCrLf & Text1.Text '
'        Case 93: Text1.Text = "Touche Windows droite" & vbCrLf & Text1.Text '
'        Case 96: Text1.Text = "Touche 0" & vbCrLf & Text1.Text ' vbKeyNumpad0
'        Case 97: Text1.Text = "Touche 1" & vbCrLf & Text1.Text ' vbKeyNumpad1
'        Case 98: Text1.Text = "Touche 2" & vbCrLf & Text1.Text ' vbKeyNumpad2
'        Case 99: Text1.Text = "Touche 3" & vbCrLf & Text1.Text ' vbKeyNumpad3
'        Case 100: Text1.Text = "Touche 4" & vbCrLf & Text1.Text ' vbKeyNumpad4
'        Case 101: Text1.Text = "Touche 5" & vbCrLf & Text1.Text ' vbKeyNumpad5
'        Case 102: Text1.Text = "Touche 6" & vbCrLf & Text1.Text ' vbKeyNumpad6
'        Case 103: Text1.Text = "Touche 7" & vbCrLf & Text1.Text ' vbKeyNumpad7
'        Case 104: Text1.Text = "Touche 8" & vbCrLf & Text1.Text ' vbKeyNumpad8
'        Case 105: Text1.Text = "Touche 9" & vbCrLf & Text1.Text ' vbKeyNumpad9
'        Case 106: Text1.Text = "Touche SIGNE MULTIPLICATION (*)" & vbCrLf & Text1.Text ' vbKeyMultiply
'        Case 107: Text1.Text = "Touche SIGNE PLUS (+)" & vbCrLf & Text1.Text ' vbKeyAdd
'        Case 108: Text1.Text = "Touche ENTRÉE (pavé numérique)" & vbCrLf & Text1.Text ' vbKeySeparator
'        Case 109: Text1.Text = "Touche SIGNE MOINS (-)" & vbCrLf & Text1.Text ' vbKeySubtract
'        Case 110: Text1.Text = "Touche POINT DÉCIMAL (.)" & vbCrLf & Text1.Text ' vbKeyDecimal
'        Case 111: Text1.Text = "Touche SIGNE DIVISION (/)" & vbCrLf & Text1.Text ' vbKeyDivide
'        Case 112: Text1.Text = "Touche F1" & vbCrLf & Text1.Text ' vbKeyF1
'        Case 113: Text1.Text = "Touche F2" & vbCrLf & Text1.Text ' vbKeyF2
'        Case 114: Text1.Text = "Touche F3" & vbCrLf & Text1.Text ' vbKeyF3
'        Case 115: Text1.Text = "Touche F4" & vbCrLf & Text1.Text ' vbKeyF4
'        Case 116: Text1.Text = "Touche F5" & vbCrLf & Text1.Text ' vbKeyF5
'        Case 117: Text1.Text = "Touche F6" & vbCrLf & Text1.Text ' vbKeyF6
'        Case 118: Text1.Text = "Touche F7" & vbCrLf & Text1.Text ' vbKeyF7
'        Case 119: Text1.Text = "Touche F8" & vbCrLf & Text1.Text ' vbKeyF8
'        Case 120: Text1.Text = "Touche F9" & vbCrLf & Text1.Text ' vbKeyF9
'        Case 121: Text1.Text = "Touche F10" & vbCrLf & Text1.Text ' vbKeyF10
'        Case 122: Text1.Text = "Touche F11" & vbCrLf & Text1.Text ' vbKeyF11
'        Case 123: Text1.Text = "Touche F12" & vbCrLf & Text1.Text ' vbKeyF12
'        Case 124: Text1.Text = "Touche F13" & vbCrLf & Text1.Text ' vbKeyF13
'        Case 125: Text1.Text = "Touche F14" & vbCrLf & Text1.Text ' vbKeyF14
'        Case 126: Text1.Text = "Touche F15" & vbCrLf & Text1.Text ' vbKeyF15
'        Case 127: Text1.Text = "Touche F16" & vbCrLf & Text1.Text ' vbKeyF16
'        Case 144: Text1.Text = "Touche VERR.NUM" & vbCrLf & Text1.Text ' vbKeyNumlock
'        Case 145: Text1.Text = "Touche Arrêt défil" & vbCrLf & Text1.Text '
'        Case 186: Text1.Text = "Touche $ ou £" & vbCrLf & Text1.Text '
'        Case 187: Text1.Text = "Touche + ou =" & vbCrLf & Text1.Text '
'        Case 188: Text1.Text = "Touche , ou ?" & vbCrLf & Text1.Text '
'        Case 190: Text1.Text = "Touche ; ou ." & vbCrLf & Text1.Text '
'        Case 191: Text1.Text = "Touche : ou /" & vbCrLf & Text1.Text '
'        Case 192: Text1.Text = "Touche ù ou %" & vbCrLf & Text1.Text '
'        Case 219: Text1.Text = "Touche ° ou )" & vbCrLf & Text1.Text '
'        Case 220: Text1.Text = "Touche * ou µ" & vbCrLf & Text1.Text '
'        Case 221: Text1.Text = "Touche ^ ou ¨" & vbCrLf & Text1.Text '
'        Case 222: Text1.Text = "Touche ²" & vbCrLf & Text1.Text '
'        Case 223: Text1.Text = "Touche < ou >" & vbCrLf & Text1.Text '
'        Case 226: Text1.Text = "Touche ! ou §" & vbCrLf & Text1.Text '
'        Case Else: Text1.Text = "Touche inconnue : " & a & vbCrLf & Text1.Text 'Touche inconnue
'    End Select
End If
Next

  If boolAutoCAD = True Then
    PossibleArretKill = True
  Else
    PossibleArretKill = False
  End If
Me.Timer1.Enabled = True
End Sub



Private Sub MDIForm_Load()
'Dim aa As Double
'aa = Me.Width

Me.StatusBar1.Panels(1) = "" & UserName
If boolAutoCAD = True Then
    Me.StatusBar1.Panels(4).Picture = Me.ImageList1.ListImages(1).Picture
    Me.StatusBar1.Panels(4).ToolTipText = "Licence AUTOCAD Disponible."
Else
    Me.StatusBar1.Panels(4).Picture = Me.ImageList1.ListImages(2).Picture
     Me.StatusBar1.Panels(4).ToolTipText = "Pas de licence AUTOCAD Disponible."
        
End If
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim hProcess As Long
 
    If Not Visible Then
        Select Case x \ Screen.TwipsPerPixelX
            Case WM_MOUSEMOVE
            Case WM_LBUTTONDOWN
            Case WM_LBUTTONUP
            Case WM_LBUTTONDBLCLK
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
                DeIconify Me
'                Me.StartUpPosition = vbStartUpScreen
            Case WM_RBUTTONDOWN
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
'                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
            Case WM_RBUTTONDBLCLK
        End Select
    End If

End Sub

Private Sub MDIForm_Resize()
Dim aa As Double
On Error Resume Next
aa = Me.Width - Me.StatusBar1.Panels(4).Picture.Width

Me.StatusBar1.Panels(1).Width = aa * 0.1
Me.StatusBar1.Panels(2).Width = aa * 0.7
Me.StatusBar1.Panels(3).Width = aa * 0.1
Me.StatusBar1.Panels(4).Width = Me.StatusBar1.Panels(4).Picture.Width
Me.StatusBar1.Panels(5).Width = aa * 0.1


'Me.StatusBar1.Refresh
If WindowState = vbMinimized Then
        Iconify Me, "Autocâble"
        DoEvents
        Exit Sub
    End If
End Sub

Private Sub MDIForm_Terminate()
CodageX.DcrJenton

Con.Execute "DELETE [Utilise_Par].* FROM [Utilise_Par] WHERE [Utilise_Par].Machine='" & MyReplace(Machine) & "' and [Utilise_Par].User='" & MyReplace(UserName) & "';"


Con.CloseConnection


If boolAutoCAD = True Then
 AutoApp.Quit
 End If
 Set AutoApp = Nothing

Unload Me
On Error Resume Next
End
funCloseConnextion
End Sub

Private Sub Message_Droits_Click()
FrmMesageDroits.Show vbModal
End Sub

Private Sub MosHor_Click()
Me.Arrange vbTileVertical

End Sub

Private Sub MosVer_Click()
Me.Arrange vbTileHorizontal
End Sub




Private Sub Nomenclature0_Click()
ExporterExcel.ChargeNomenclature 0, Me.Nomenclature0.Caption
Unload ExporterExcel
LibertPice
End Sub

Private Sub Nomenclature1_Click()
ExporterExcel.ChargeNomenclature 1, Me.Nomenclature1.Caption
Unload ExporterExcel
LibertPice
End Sub

Private Sub Nomenclature2_Click()
ExporterExcel.ChargeNomenclature 2, Me.Nomenclature0.Caption
Unload ExporterExcel
LibertPice
End Sub

Private Sub Nomenclature3_Click()
ExporterExcel.ChargeNomenclature 3, Me.Nomenclature0.Caption
Unload ExporterExcel
LibertPice
End Sub

Private Sub Projets_Click()
Liste_Projets.Show vbModal
End Sub

Private Sub Quitter_Click()
Unload Me
End Sub
Sub EnabledMenu()
Me.StatusBar1.Enabled = False
'
    Me.Fichier.Enabled = False
    Me.Edit.Enabled = False
    Me.Autre.Enabled = False
    Me.Autre.Enabled = False
'    Me.Fenêtre.Enabled = False
    Me.Admin.Enabled = False
'    Me.Utilisateur.Enabled = False
'    Me.Utilisateur_Avec_Pouvoir.Enabled = False
    
    
End Sub
Public Sub DesEnabledMenu()
Me.StatusBar1.Enabled = True
'Me.Fichier.Enabled = True
'    Me.Edit.Enabled = True
'    Me.Autre.Enabled = True
'    Me.Autre.Enabled = True
'    Me.Fenêtre.Enabled = True
'    Me.Admin.Enabled = True
'    Me.Utilisateur.Enabled = True
'    Me.Utilisateur_Avec_Pouvoir.Enabled = True
'MajDroitsFrm Id_Users
End Sub

Private Sub SMTP_Click()
frmPOP3.Show vbModal
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As ComctlLib.Panel)
On Error Resume Next
If UCase(CherCheInFihier("IsCilent")) = "TRUE" Then IsCilent = True

If UCase(CherCheInFihier("IsServeur")) = "TRUE" Then IsServeur = True

If IsServeur = IsCilent Then IsServeur = False: IsCilent = False

If Panel.Index = 4 Then
    If boolAutoCAD = True Then
            If MsgBox("Voulez vous fermer votre licence AUTOCAD.", vbQuestion + vbYesNo) = vbYes Then
                AutoApp.Quit
                Set AutoApp = Nothing
                boolAutoCAD = False
                Me.StatusBar1.Panels(3).Text = ""
                PossibleArretKill = False
                If ArretKill = True Then
                Me.Timer1.Enabled = False
                Call keybd_event(145, 1, 0, 0)
                ArretKill = False
                Me.Timer1.Enabled = True
            End If
        End If
        
    Else
        If MsgBox("Voulez vous ouvrir une licence AUTOCAD.", vbQuestion + vbYesNo) = vbYes Then
        '        Set AutoApp = New AutoCAD.AcadApplication
        
        NewAutocadAdmin
        End If
        Err.Clear
        Dim a
        
        
        
        
        
       
    End If
End If
DoEvents
 If boolAutoCAD = True Then
        Me.StatusBar1.Panels(4).Picture = Me.ImageList1.ListImages(1).Picture
        Me.StatusBar1.Panels(4).ToolTipText = "Licence AUTOCAD Disponible."
        Else
        Me.StatusBar1.Panels(4).Picture = Me.ImageList1.ListImages(2).Picture
        Me.StatusBar1.Panels(4).ToolTipText = "Pas de licence AUTOCAD Disponible."
        End If
End Sub

Private Sub Supprimer_Archiver_Click()
UserForm4.Charge Me, "IdStatus=3 or (VerifieDate= Null and IdStatus<>4)"
Unload UserForm4

End Sub

Private Sub Utilitaire_Click(Index As Integer)
Dim Fso As FileSystemObject
Dim Sql As String
Dim Rs As Recordset
'If Me.List1.ListIndex = -1 Then
'    MsgBox "devez sélectionner un utilitaire.", vbExclamation
'    Exit Sub
'End If
Sql = "SELECT Utilitaire.Utilitaire FROM Utilitaire "
Sql = Sql & "WHERE Utilitaire.NameBouton='" & Me.Utilitaire(Index).Caption & "' "
Sql = Sql & "ORDER BY Utilitaire.NameBouton;"

Set Rs = Con.OpenRecordSet(Sql)
Set Fso = New FileSystemObject
If Rs.EOF = False Then
'   If Fso.FileExists("" & Rs!Utilitaire) = True Then
'   MsgBox ""
'   End If
     'Execute explorer.exe
    MyExecute "" & Rs!Utilitaire
     'Execute la calculette
''    WinExec "Calc.exe", 1
     'autre astuce :Le vrai mode plein ecran est là !!
'    Shell "c:\program files\internet explorer\iexplore -k c:"
'    Shell "" & Rs!Utilitaire, vbMaximizedFocus
End If
Set Rs = Con.CloseRecordSet(Rs)
End Sub


Private Sub Utliser_Par_Click()
UtliserPar.Show
End Sub

Private Sub Vagues_Click()
On Error Resume Next
UserForm1.charger txt1, " ", "Vagues:", " "
Unload UserForm1
txt1 = ""
End Sub
