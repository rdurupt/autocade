VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0002E550-0000-0000-C000-000000000046}#1.1#0"; "OWC10.DLL"
Begin VB.Form UserForm2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11925
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   17955
   Icon            =   "AutoCable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11925
   ScaleWidth      =   17955
   Begin VB.TextBox txtMacro 
      Height          =   285
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton StopMaco 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "AutoCable.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Reprendre le traitement"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   15960
      TabIndex        =   4
      Top             =   -70
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   525
         Left            =   0
         Picture         =   "AutoCable.frx":0BD4
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Annuler"
      Enabled         =   0   'False
      Height          =   435
      Left            =   13680
      TabIndex        =   3
      Top             =   11040
      Width           =   2130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Valider"
      Enabled         =   0   'False
      Height          =   435
      Left            =   9555
      TabIndex        =   2
      Top             =   11040
      Width           =   2130
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualiser &/ Valider"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5445
      TabIndex        =   1
      Top             =   11040
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "A&ctualiser"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1320
      TabIndex        =   0
      Top             =   11040
      Width           =   2130
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   11640
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10140
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   17850
      _ExtentX        =   31485
      _ExtentY        =   17886
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "Critères"
      TabPicture(0)   =   "AutoCable.frx":10C96
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Crit"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Connecteurs"
      TabPicture(1)   =   "AutoCable.frx":10CB2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Conn"
      Tab(1).Control(1)=   "Picture2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tableau de fils"
      TabPicture(2)   =   "AutoCable.frx":10CCE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Fil"
      Tab(2).Control(1)=   "Picture1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Composants"
      TabPicture(3)   =   "AutoCable.frx":10CEA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Comp"
      Tab(3).Control(1)=   "Combo1"
      Tab(3).Control(2)=   "Picture3"
      Tab(3).Control(3)=   "Picture4"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Notas"
      TabPicture(4)   =   "AutoCable.frx":10D06
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Notas"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Noeuds"
      TabPicture(5)   =   "AutoCable.frx":10D22
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Noeuds"
      Tab(5).Control(1)=   "Longueur"
      Tab(5).Control(2)=   "Hab"
      Tab(5).Control(3)=   "RSA"
      Tab(5).Control(4)=   "PSA"
      Tab(5).Control(5)=   "ENC"
      Tab(5).Control(6)=   "Command5"
      Tab(5).Control(7)=   "Command6"
      Tab(5).Control(8)=   "Command7"
      Tab(5).Control(9)=   "ACTIVER"
      Tab(5).Control(10)=   "DIAMETRE"
      Tab(5).Control(11)=   "CLASSE_T"
      Tab(5).Control(12)=   "TORON_P"
      Tab(5).Control(13)=   "Long_C"
      Tab(5).Control(14)=   "Fleche_Droite"
      Tab(5).Control(15)=   "txtOption"
      Tab(5).Control(16)=   "Command8"
      Tab(5).Control(17)=   "Label1"
      Tab(5).Control(18)=   "Label2"
      Tab(5).Control(19)=   "Label3"
      Tab(5).Control(20)=   "Label4"
      Tab(5).Control(21)=   "Label5"
      Tab(5).Control(22)=   "Label6"
      Tab(5).Control(23)=   "NOUED"
      Tab(5).Control(24)=   "Label7"
      Tab(5).Control(25)=   "Label8"
      Tab(5).Control(26)=   "Label9"
      Tab(5).Control(27)=   "Label10"
      Tab(5).ControlCount=   28
      TabCaption(6)   =   "Nomenclatures"
      TabPicture(6)   =   "AutoCable.frx":10D3E
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Nom"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Dossier de Fabrication"
      TabPicture(7)   =   "AutoCable.frx":10D5A
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Fab"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Dossier de Contrôle"
      TabPicture(8)   =   "AutoCable.frx":10D76
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Cont"
      Tab(8).ControlCount=   1
      Begin OWC.Spreadsheet Notas 
         Height          =   9735
         Left            =   -75000
         TabIndex        =   43
         Top             =   360
         Width           =   17775
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":10D92
         DataType        =   "HTMLDATA"
         AutoFit         =   0   'False
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin OWC.Spreadsheet Crit 
         Height          =   9735
         Left            =   0
         TabIndex        =   39
         Top             =   360
         Width           =   17775
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":11805
         DataType        =   "HTMLDATA"
         AutoFit         =   0   'False
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -69600
         Picture         =   "AutoCable.frx":11F38
         ScaleHeight     =   255
         ScaleWidth      =   270
         TabIndex        =   50
         Top             =   670
         Width           =   270
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -69960
         Picture         =   "AutoCable.frx":12332
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   49
         Top             =   670
         Width           =   225
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -69240
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   650
         Width           =   2295
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -69960
         Picture         =   "AutoCable.frx":123B8
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   27
         Top             =   670
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -69960
         Picture         =   "AutoCable.frx":1243E
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   10
         Top             =   670
         Width           =   225
      End
      Begin OWC10.Spreadsheet Nom 
         Height          =   9735
         Left            =   -74880
         OleObjectBlob   =   "AutoCable.frx":124C4
         TabIndex        =   47
         Top             =   360
         Width           =   17775
      End
      Begin OWC10.Spreadsheet Cont 
         Height          =   9735
         Left            =   -75000
         OleObjectBlob   =   "AutoCable.frx":13431
         TabIndex        =   46
         Top             =   360
         Width           =   17775
      End
      Begin OWC10.Spreadsheet Fab 
         Height          =   9735
         Left            =   -75000
         OleObjectBlob   =   "AutoCable.frx":13B23
         TabIndex        =   45
         Top             =   360
         Width           =   17775
      End
      Begin OWC.Spreadsheet Noeuds 
         Height          =   8655
         Left            =   -75000
         TabIndex        =   44
         Top             =   1440
         Width           =   17775
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":1428C
         DataType        =   "HTMLDATA"
         AutoFit         =   0   'False
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin OWC.Spreadsheet Fil 
         Height          =   9735
         Left            =   -75000
         TabIndex        =   41
         Top             =   360
         Width           =   17775
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":149BF
         DataType        =   "HTMLDATA"
         AutoFit         =   0   'False
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin OWC.Spreadsheet Conn 
         Height          =   9735
         Left            =   -75000
         TabIndex        =   40
         Top             =   360
         Width           =   17775
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":150F2
         DataType        =   "HTMLDATA"
         AutoFit         =   0   'False
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin VB.TextBox Longueur 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -66960
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox Hab 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         ItemData        =   "AutoCable.frx":15825
         Left            =   -70680
         List            =   "AutoCable.frx":15827
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox RSA 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -66960
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox PSA 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -64440
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox ENC 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         ItemData        =   "AutoCable.frx":15829
         Left            =   -61920
         List            =   "AutoCable.frx":1582B
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   -74880
         Picture         =   "AutoCable.frx":1582D
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ajouter"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   -73800
         Picture         =   "AutoCable.frx":160A3
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Supprimer"
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   -74340
         Picture         =   "AutoCable.frx":1689D
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Modifier"
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox ACTIVER 
         Alignment       =   1  'Right Justify
         Caption         =   "ACTIVER"
         Height          =   315
         Left            =   -73080
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox DIAMETRE 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -61920
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox CLASSE_T 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -59520
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox TORON_P 
         Alignment       =   1  'Right Justify
         Caption         =   "TORON/P"
         Height          =   315
         Left            =   -73080
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Long_C 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -64440
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Fleche_Droite 
         Alignment       =   1  'Right Justify
         Caption         =   "Fleche D"
         Height          =   315
         Left            =   -74520
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtOption 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -59520
         TabIndex        =   12
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   -57720
         Picture         =   "AutoCable.frx":1704B
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   315
      End
      Begin OWC.Spreadsheet Comp 
         Height          =   9735
         Left            =   -75000
         TabIndex        =   42
         Top             =   360
         Width           =   17775
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":17E8D
         DataType        =   "HTMLDATA"
         AutoFit         =   0   'False
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOEUDS"
         Height          =   315
         Left            =   -71880
         TabIndex        =   38
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "LONGUEUR"
         Height          =   315
         Left            =   -68040
         TabIndex        =   37
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DESIGN.HAB."
         Height          =   315
         Left            =   -71880
         TabIndex        =   36
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CODE.RSA."
         Height          =   315
         Left            =   -68040
         TabIndex        =   35
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CODE.PSA."
         Height          =   315
         Left            =   -65640
         TabIndex        =   34
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CODE.ENC."
         Height          =   255
         Left            =   -62760
         TabIndex        =   33
         Top             =   360
         Width           =   870
      End
      Begin VB.Label NOUED 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -70680
         TabIndex        =   32
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DIAMETRE"
         Height          =   315
         Left            =   -62880
         TabIndex        =   31
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CLASSE_T"
         Height          =   315
         Left            =   -60360
         TabIndex        =   30
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "LONGUEUR/C"
         Height          =   315
         Left            =   -65640
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "OPTION"
         Height          =   315
         Left            =   -60360
         TabIndex        =   28
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Label ProgressBar1Caption 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   11640
      Width           =   3015
   End
   Begin VB.Menu Outils 
      Caption         =   "Outils"
      Begin VB.Menu AfficherMasquees 
         Caption         =   "Afficher colonnes masquées"
         Shortcut        =   ^F
      End
      Begin VB.Menu Masquer_Colonnes 
         Caption         =   "Masquer Colonnes"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu Macro 
      Caption         =   "Macro"
      Begin VB.Menu NewMacro 
         Caption         =   "Nouvelle Macro"
         Shortcut        =   ^N
      End
      Begin VB.Menu ExecMacro 
         Caption         =   "Executer Macro"
         Shortcut        =   ^O
      End
      Begin VB.Menu SupMacro 
         Caption         =   "Supprimer Macro"
         Shortcut        =   ^P
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
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim New_frmLstComposants
Dim ColecBool As New Collection
Dim PathComposantstous() As String
Dim MyIdIndiceProjet As Long
Dim Myfrm As Object
Dim MacroName As String
Dim NoMacro As Boolean
Dim Nouveau As Boolean
Public boolExcute As Boolean
Dim NotSortie As Boolean
Dim MyClient As String
Dim msg As String
Dim MyErr As Boolean
Dim IfValidationOk As Boolean
Dim NbFinOuiNon As Long
Dim NoMacro1Change As Boolean
Dim NoMacro1Select As Boolean
Dim NoMacro2 As Boolean
Dim NoMacro3 As Boolean
Dim NoMacro4 As Boolean
Dim NoMacro5 As Boolean
Dim NoMacro5Select As Boolean
Dim CollecCrieres As Collection
Dim CollecCrieresCode As Collection
Dim CollecCrieresDesigne As Collection
Dim CollectionPath As Collection
Dim NoMacro6 As Boolean
Dim boolSelctChange As Boolean
Dim boolMajListe As Boolean
Dim MyTableENC() As String
Dim MyTablePSA() As String
Dim MyTableRSA() As String
Dim MyTableHab() As String
Dim NoMaj As Boolean
Dim MyCollectionENC As New Collection
Dim MyCollectionPSA As New Collection
Dim MyCollectionRSA As New Collection
Dim MyCollectionHab As New Collection
Dim MyCollectionLienHab As New Collection
Public NumCollonne As New Collection
Public ChrCollonne As New Collection
Dim bool_Activate As Boolean
Dim boolActu As Boolean
Dim IdProjet As Long
Dim OnGletName As String
Dim CollecApp As Collection
Dim Mygrid As String
Dim RefCli As String
Dim refFour As String
Public CollectionMenu As Collection

Public Sub Charger_Colection(grid, Lib As String)
Dim MyRange
Dim C As Long
Dim I As Long
Dim Txt As String
Dim Adress
Set MyRange = Nothing
Set MyRange = grid.Range("a1").CurrentRegion
For C = 1 To MyRange.Columns.Count

    NumCollonne.Add C, Lib & Trim("" & MyRange(1, C).Value)
    Adress = MyRange(1, C).Address
    Txt = ""
    For I = 1 To Len(Adress)
        If Not IsNumeric(Mid(Adress, I, 1)) Then
            Txt = Txt & Mid(Adress, I, 1)
        End If
    Next
    ChrCollonne.Add Txt, Lib & Trim("" & MyRange(1, C).Value)
    
Next
End Sub









Private Sub AfficherMasquees_Click()
If Me.StopMaco.Visible = True Then

    txtMacro = txtMacro & Space(5) & "SSTab1.Tab =" & SSTab1.Tab & vbCrLf
    If CollectionMenu("Nomenclatures") = Mygrid Then
        txtMacro = txtMacro & Space(5) & Mygrid & ".Sheets(" & Chr(34) & Controls(Mygrid).ActiveSheet.Name & Chr(34) & ").Select" & vbCrLf
        txtMacro = txtMacro & Space(5) & Mygrid & ".ActiveSheet.Range(""A1"").CurrentRegion.ColumnWidth = 100" & vbCrLf
        txtMacro = txtMacro & Space(5) & Mygrid & ".ActiveSheet.Range(""A1"").CurrentRegion.Cells.EntireColumn.AutoFit" & vbCrLf
               
    Else
    If CollectionMenu("Dossier de Fabrication") = Mygrid Then
        txtMacro = txtMacro & Space(5) & "For I= 1 to " & Mygrid & ".Sheets.Count" & vbCrLf
        txtMacro = txtMacro & Space(10) & Mygrid & ".Sheets(I).Select" & vbCrLf
        txtMacro = txtMacro & Space(10) & Mygrid & ".ActiveSheet.Range(""A1"").CurrentRegion.ColumnWidth = 100" & vbCrLf
        txtMacro = txtMacro & Space(10) & Mygrid & ".ActiveSheet.Range(""A1"").CurrentRegion.Cells.EntireColumn.AutoFit" & vbCrLf
        txtMacro = txtMacro & Space(5) & "NEXT" & vbCrLf
       
    Else
        txtMacro = txtMacro & Space(5) & Mygrid & ".Cells.AutoFitColumns" & vbCrLf
    End If
    End If
End If
On Error Resume Next
Controls(Mygrid).ActiveSheet.Cells.AutoFitColumns
'Controls(Mygrid).ActiveSheet .EntireColumn.AutoFit
Controls(Mygrid).ActiveSheet.Range("a1").CurrentRegion.ColumnWidth = 100
Controls(Mygrid).ActiveSheet.Range("a1").CurrentRegion.Cells.EntireColumn.AutoFit
On Error GoTo 0
End Sub

Private Sub Autre_Click()

End Sub

Private Sub cascade_Click()
frmAutocâble.Arrange vbCascade

End Sub

Private Sub Combo1_Click()
If Comp.ActiveCell.Row > 1 Then
    If Combo1.ListIndex > 0 Then
        Comp.Cells(Comp.ActiveCell.Row, NumCollonne("comppath")) = Me.Combo1.List(Me.Combo1.ListIndex)
    End If
End If
Combo1.ListIndex = 0
'Combo1.Visible = False
End Sub

Private Sub Command3_Click()
Dim Sql As String
Dim Rs As Recordset
Me.Command1.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False

Me.Tag = MyIdIndiceProjet
If boolActu = False Then
    MsgBox "Il est impossible de valide l'étude si un test de d'actualisation na pas été effectué."
    Me.Command1.Enabled = True
    Me.Command2.Enabled = True
    Me.Command3.Enabled = True
    Me.Command4.Enabled = True

    Exit Sub
End If
If Trim(msg) <> "" Then
    MsgBox "Il est impossible de valide l'étude si le test de validation présente des erreurs."
    Me.Command1.Enabled = True
    Me.Command2.Enabled = True
    Me.Command3.Enabled = True
    Me.Command4.Enabled = True
    Exit Sub
End If
NoMacro2 = False
NoMacro5 = False
NoMacro3 = False
NoMacro1Change = False

Importefrm Me, True

 If boolAutoCAD = False And IsCilent = False Then
   


    MsgBox MsgAutoCad & "Vous ne possédez pas de licence AutoCad." & vbCrLf & "Vous ne pouvez pas reporter vos modifications" & vbCrLf & "sur vos différents plans. "
Else
Set FormBarGrah = Me
 Planche_Clous.chargement CLng(Me.Tag)



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
        Myfrm.Enabled = True
        NotSortie = False
       GoTo Fin
    End If
    If IsCilent = False Then
        
        subDessinerPlan Me.Tag
        subDessinerOutil Me.Tag
        
        
        MsgBox "Fin du traitement" & vbCrLf & NbError & " Erreur(s) détectée(s) !"
    Else
        MsgBox "Votre demande a été prise en compte vous pouvez suivre l'évolution de votre travail dans la fenêtre (Liste des JOB)"
        Unload Myfrm
    End If
 End If
Fin:
NbError = 0
 NotSortie = False
'' Unload New_frmLstComposants
'Set New_frmLstComposants = Nothing
 Unload Myfrm
Unload Me
End Sub



Private Sub Command4_Click()

MenuShow = True
 boolExcute = False
 NotSortie = False
 boolActu = False
'' Unload New_frmLstComposants
'Set New_frmLstComposants = Nothing
 Myfrm.Enabled = True
 Unload Myfrm
 Unload Me

End Sub

Private Sub Command1_Click()
'Me.Refresh
DoEvents
Dim MyRange
Dim Sql As String
Me.Command1.Enabled = False
Me.Command2.Enabled = False
Me.Command3.Enabled = False
Me.Command4.Enabled = False


Set CollecApp = Nothing
Set CollecApp = New Collection
Set CollecCrieres = Nothing
Set CollecCrieresCode = Nothing
Set CollecCrieresDesigne = Nothing

Set CollecCrieres = New Collection
Set CollecCrieresCode = New Collection
Set CollecCrieresDesigne = New Collection

Me.Crit.Cells(1, 1).Select
Sql = "DELETE Ajout_LIAISON_CONNECTEURS.* FROM Ajout_LIAISON_CONNECTEURS;"
Con.Execute Sql
Sql = "DELETE Ajout_LIAISON.* FROM Ajout_LIAISON;"
Con.Execute Sql

msg = ""
Me.Refresh
DoEvents
IfValidationOk = True
SSTab1.Tab = 0
RazFiltreEditExcel Me.Crit
Set MyRange = Nothing
Set MyRange = Me.Crit.Range("a1").CurrentRegion

Me.Crit.Cells(1, 1).Select
For I = 2 To MyRange.Rows.Count
'
'Me.Crit.Refresh
DoEvents
IfValidationOk = True
Me.Crit.Cells(I, 1).Select

If msg <> "" Then
    IfValidationOk = False
'    Me.Crit..AutoFilter = True
    GoTo FinTraitement
End If
    Me.Crit.Cells(I, 1).Value = "'" & Me.Crit.Cells(I, 1).Value
    ConverOuiNon MyRange, I
If msg <> "" Then
    IfValidationOk = False
    GoTo FinTraitement
End If

'Me.Crit.Refresh
DoEvents
Next I
'Me.Crit.Refresh
SSTab1.Tab = 1
RazFiltreEditExcel Me.Conn
Set MyRange = Nothing
Set MyRange = Me.Conn.Range("a1").CurrentRegion
OnGletName = "Connecteur"


DoEvents
Me.Conn.Cells(1, 1).Select
For I = 2 To MyRange.Rows.Count

DoEvents
IfValidationOk = True
Me.Conn.Cells(I, 2).Select
Me.Conn.Cells(I, 1).Select
ConverOuiNon MyRange, I
If msg <> "" Then
    IfValidationOk = False
    GoTo FinTraitement
End If
    Me.Conn.Cells(I, 1).Value = "'" & Me.Conn.Cells(I, 1).Value
If msg <> "" Then
    IfValidationOk = False
    GoTo FinTraitement
End If
'Me.Conn.Refresh
DoEvents
Next I
'Me.Conn.Refresh
SSTab1.Tab = 2
RazFiltreEditExcel Me.Fil
Set MyRange = Nothing
Set MyRange = Me.Fil.Range("a1").CurrentRegion


'Me.Refresh
DoEvents
Me.Fil.Cells(1, 1).Select
For I = 2 To MyRange.Rows.Count
'Me.Refresh
DoEvents
Me.Fil.Cells(I, NumCollonne("filsApp")).Select
ConverOuiNon MyRange, I
IfValidationOk = True
    Me.Fil.Cells(I, NumCollonne("filsApp")).Value = UCase("'" & Me.Fil.Cells(I, NumCollonne("filsApp")).Value)
    
If msg <> "" Then
'Me.Fil..AutoFilter = True
    IfValidationOk = False
    GoTo FinTraitement
End If
Me.Fil.Cells(I, NumCollonne("filsApp2")).Select
    Me.Fil.Cells(I, NumCollonne("filsApp2")).Value = UCase("'" & Me.Fil.Cells(I, NumCollonne("filsApp2")).Value)
    If msg <> "" Then
'    Me.Fil..AutoFilterMode = True
ConverOuiNon MyRange, I
        IfValidationOk = False
        GoTo FinTraitement
    End If
'Me.Refresh
DoEvents
Next I
'Me.Crit..AutoFilter = False
SSTab1.Tab = 3
RazFiltreEditExcel Me.Comp
Set MyRange = Nothing
Set MyRange = Me.Comp.Range("a1").CurrentRegion

Me.Comp.Cells(1, 2).Select
For I = 2 To MyRange.Rows.Count

DoEvents
Me.Comp.Cells(I, 2).Select

Me.Comp.Cells(I, 3) = 0
ConverOuiNon MyRange, I
IfValidationOk = True
 If msg <> "" Then GoTo FinTraitement
'Me.Fil.Refresh
DoEvents
Next I
'Me.Fil.Refresh
SSTab1.Tab = 4
RazFiltreEditExcel Me.Notas
Set MyRange = Nothing
Set MyRange = Me.Notas.Range("a1").CurrentRegion
SSTab1.Tab = 4
Me.Notas.Cells(1, 1).Select
For I = 2 To MyRange.Rows.Count
Me.Notas.Cells(I, 3).Select
ConverOuiNon MyRange, I
Me.Notas.Cells(I, 3) = I - 1
IfValidationOk = True

' Me.Notas.RefreshDoEvents
Next I
' Me.Notas.Refresh
RazFiltreEditExcel Me.Noeuds
 
Set MyRange = Nothing
Set MyRange = Me.Noeuds.Range("a1").CurrentRegion
SSTab1.Tab = 5
Me.Noeuds.Cells(1, 1).Select
For I = 2 To MyRange.Rows.Count
'Me.Refresh
DoEvents
IfValidationOk = True
'If I = 80 Then MsgBox ""
ConverOuiNon MyRange, I
Me.Noeuds.Cells(I, 1).Select
'
'Me.Noeuds.Refresh
DoEvents
If msg <> "" Then
'Me.Noeuds..AutoFilter = True
    IfValidationOk = False
    GoTo FinTraitement
End If
'If Noeuds.Cells(i, 4) = "x" Then
'    MsgBox ""
'End If
    Command7_Click
If msg <> "" Then
    IfValidationOk = False
    GoTo FinTraitement
End If
'Me.Noeuds.Refresh
DoEvents
Next I

'Me.Noeuds.Refresh
'Me.Noeuds..AutoFilter = True
If MyErr = True Then
    LoadLiasons.charger MyClient
    Unload LoadLiasons
End If
MyErr = False
    IfValidationOk = False
    boolActu = True
FinTraitement:
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True

End Sub
Private Sub Command2_Click()
Command1_Click
If msg <> "" Then Exit Sub
Command3_Click

End Sub

Private Sub Command5_Click()
Set MyRange = Nothing
Set MyRange = Me.Noeuds.ActiveSheet.Range("a1").CurrentRegion

'If Me.Tag = "" Then
    If Trim("" & Me.Hab) <> "" Then
        boolSelctChange = True
        MyRange(MyRange.Rows.Count + 1, 1).Select
        MyRange(MyRange.Rows.Count + 1, 2) = Me.Fleche_Droite.Value
        MyRange(MyRange.Rows.Count + 1, 3) = Me.TORON_P.Value
        MyRange(MyRange.Rows.Count + 1, 1) = Me.Activer.Value
        MyRange(MyRange.Rows.Count + 1, 5) = Val(Replace("" & Me.Longueur, ",", "."))
        MyRange(MyRange.Rows.Count + 1, 6) = Val(Replace("" & Me.Long_C, ",", "."))
        MyRange(MyRange.Rows.Count + 1, 7) = "'" & Me.Hab
        MyRange(MyRange.Rows.Count + 1, 8) = "'" & Me.RSA
        MyRange(MyRange.Rows.Count + 1, 9) = "'" & Me.PSA
        MyRange(MyRange.Rows.Count + 1, 10) = "'" & Me.ENC
         
        MyRange(MyRange.Rows.Count + 1, 11) = Val(Replace("" & Me.DIAMETRE, ",", "."))
        MyRange(MyRange.Rows.Count + 1, 12) = "'" & Me.CLASSE_T
        MyRange(MyRange.Rows.Count + 1, 13) = "'" & Me.txtOption
        boolSelctChange = False
    End If
'Else
'     If Trim("" & Me.Hab) <> "" Then
'     boolSelctChange = True
'         Me.Noeuds.ActiveSheet.Cells(Me.Tag, 1).InsertRows
'        MyRange(Me.Tag, 1).Select
'        MyRange(Me.Tag, 1) = Me.Fleche_Droite.Value
'         MyRange(Me.Tag, 2) = Me.TORON_P.Value
'         MyRange(Me.Tag, 3) = Me.ACTIVER.Value
'        MyRange(Me.Tag, 5) = Val(Replace("" & Me.Longueur, ",", "."))
'         MyRange(Me.Tag, 6) = Val(Replace("" & Me.Long_C, ",", "."))
'        MyRange(Me.Tag, 7) = "" & Me.Hab
'        MyRange(Me.Tag, 8) = "" & Me.RSA
'        MyRange(Me.Tag, 9) = "" & Me.PSA
'        MyRange(Me.Tag, 10) = "" & Me.ENC
'         MyRange(Me.Tag, 11) = Val(Replace("" & Me.DIAMETRE, ",", "."))
'        MyRange(Me.Tag, 12) = "" & Me.CLASSE_T
'        boolSelctChange = False
'    End If
'
'End If
Me.Fleche_Droite.Value = 0
Me.Activer.Value = 0
TORON_P.Value = 0
Me.Long_C = ""
 Me.DIAMETRE = ""
Me.CLASSE_T = ""
Me.NOUED = ""
  Longueur = ""
Me.Hab.ListIndex = 0
Me.Tag = ""
Me.txtOption = ""
End Sub

Private Sub Command6_Click()
If Me.Tag <> "" Then
    boolSelctChange = True
    Me.Noeuds.ActiveSheet.Rows(Val(Me.Tag)).DeleteRows
    Me.Tag = ""
    Me.Hab.ListIndex = 0
    Longueur = ""
    Me.Tag = ""
    Me.NOUED = ""
    boolSelctChange = False
End If
End Sub

Private Sub Command7_Click()
Set MyRange = Nothing
Set MyRange = Me.Noeuds.ActiveSheet.Range("a1").CurrentRegion

If Me.Tag = "" Then
    Command5_Click
Else
'    If Trim("" & Me.Hab) <> "" Then
        boolSelctChange = True
        MyRange(Val(Me.Tag), 1).Select
        
         MyRange(Val(Me.Tag), 2) = Fleche_Droite.Value
          MyRange(Val(Me.Tag), 3) = TORON_P.Value
         MyRange(Val(Me.Tag), 1) = Me.Activer.Value
        MyRange(Val(Me.Tag), 5) = Val(Replace("" & Me.Longueur, ",", "."))
         MyRange(Val(Me.Tag), 6) = Val(Replace("" & Me.Long_C, ",", "."))
        MyRange(Val(Me.Tag), 7) = "'" & Me.Hab
        MyRange(Val(Me.Tag), 8) = "'" & Me.RSA
        MyRange(Val(Me.Tag), 9) = "'" & Me.PSA
        MyRange(Val(Me.Tag), 10) = "'" & Me.ENC
        MyRange(Val(Me.Tag), 11) = Val(Replace("" & Me.DIAMETRE, ",", "."))
        MyRange(Val(Me.Tag), 12) = "'" & Me.CLASSE_T
        If InStr("" & Me.txtOption, "ALL") <> 0 Then
            If Len("" & Me.txtOption) > Len("ALL;") Then
                MyRange(Val(Me.Tag), 13) = "'" & Replace(Me.txtOption, "ALL", "")
                MyRange(Val(Me.Tag), 13) = Replace(MyRange(Val(Me.Tag), 13), ";;", ";")
                If Left("" & MyRange(Val(Me.Tag), 13), 1) = ";" Then MyRange(Val(Me.Tag), 13) = Right(MyRange(Val(Me.Tag), 13), Len(MyRange(Val(Me.Tag), 13)) - 1)
                
            Else
                MyRange(Val(Me.Tag), 13) = "'" & Me.txtOption
            End If
        Else
         MyRange(Val(Me.Tag), 13) = "'" & Me.txtOption
        End If
        boolSelctChange = False
'    End If
End If
Me.Activer.Value = 0
TORON_P.Value = 0
Me.Long_C = ""
 Me.DIAMETRE = ""
Me.CLASSE_T = ""
Me.NOUED = ""
  Longueur = ""
Me.Hab.ListIndex = 0
Me.Fleche_Droite.Value = 0
Me.Tag = ""
Me.txtOption = ""
End Sub



Private Sub Command8_Click()
Me.txtOption = FrmSelectCriteres.chargement(Crit, Me.txtOption)
Unload FrmSelectCriteres
End Sub

'Private Sub Comp_DblClick(ByVal EventInfo As OWC.SpreadsheetEventInfo)
'Me.Combo1.Top = Comp.Cells(Comp.ActiveCell.Row, NumCollonne("comppath")).Top
'Me.Combo1.Left = Comp.Cells(Comp.ActiveCell.Row, NumCollonne("comppath")).Left
'Me.Combo1.Visible = True
''  Set New_frmLstComposants = New frmLstComposants
''    New_frmLstComposants.chargement PathComposantstous, Comp.Range("a1")
''    Set New_frmLstComposants = Nothing
'End Sub

Private Sub DIAMETRE_LostFocus()
If MyFormat("dbl", DIAMETRE, "DIAMETRE") = False Then Exit Sub
End Sub

Private Sub ENC_Click()
If boolMajListe = False Then
    boolMajListe = True
'     Me.ENC.ListIndex = MyCollectionENC(MyTableENC(MyCollectionENC("N" & Me.ENC.Text), 1))
     Me.PSA.ListIndex = MyCollectionPSA(MyTableENC(MyCollectionENC("N" & Trim(Me.ENC.Text)), 2))
    Me.RSA.ListIndex = MyCollectionRSA(MyTableENC(MyCollectionENC("N" & Trim(Me.ENC.Text)), 3))
    Me.Hab.ListIndex = MyCollectionHab(MyTableENC(MyCollectionENC("N" & Trim(Me.ENC.Text)), 4))
    boolMajListe = False
End If


End Sub

Private Sub ENC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    Me.ENC.ListIndex = MyCollectionENC(MyTableENC(MyCollectionENC("N" & Me.ENC.Text), 1))
On Error GoTo 0
End If
End Sub

Private Sub ExecMacro_Click()

MacroName = ""
Me.txtMacro = ""
If Me.StopMaco.Visible = False Then
frmMacco.charger Me.Name, "EXE"
MacroName = frmMacco.NameMacro
Me.txtMacro = frmMacco.SubMacro

Unload frmMacco
'MacroName = InputBox("Entrez le nom de la macro")
If Trim("" & MacroName) = "" Then Exit Sub
   ' Crée des variables.
   Dim sc, M
   Set sc = CreateObject("ScriptControl")
   sc.language = "VBScript"
   sc.AddObject "Me", Me, True
   ' Ajoute un module.
   Set M = sc.Modules.Add("Module1")
   ' Ajoute du code au module.
   M.AddCode Me.txtMacro
   ' Exécute le script.
   M.Run MacroName
End If
MacroName = ""
 Me.txtMacro = ""
End Sub

Private Sub Form_Initialize()
boolActu = False

End Sub

Private Sub Hab_Click()
If boolMajListe = False Then
    boolMajListe = True
    
     Me.ENC.ListIndex = MyCollectionENC(MyTableHab(MyCollectionHab("N" & Trim(Me.Hab.Text)), 1))
     Me.PSA.ListIndex = MyCollectionPSA(MyTableHab(MyCollectionHab("N" & Trim(Me.Hab.Text)), 2))
    Me.RSA.ListIndex = MyCollectionRSA(MyTableHab(MyCollectionHab("N" & Trim(Me.Hab.Text)), 3))
'    Me.Hab.ListIndex = MyCollectionHab(MyTableHab(MyCollectionHab("N" & Me.Hab.Text), 4))
    boolMajListe = False
End If

End Sub



Private Sub Hab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
 Me.Hab.ListIndex = MyCollectionHab(MyTableHab(MyCollectionHab("N" & Trim(Me.Hab.Text)), 4))
  On Error GoTo 0
 End If
End Sub

Private Sub Long_C_Change()
If MyFormat("dbl", Long_C, "Longueur cumuler") = False Then Exit Sub

End Sub

Private Sub Longueur_LostFocus()
If MyFormat("dbl", Longueur, "Longueur") = False Then Exit Sub
End Sub

Private Sub Masquer_Colonnes_Click()
If Me.StopMaco.Visible = True Then

    txtMacro = txtMacro & Space(5) & "SSTab1.Tab =" & SSTab1.Tab & vbCrLf
    If CollectionMenu("Nomenclatures") = Mygrid Then
        txtMacro = txtMacro & Space(5) & Mygrid & ".Sheets(" & Chr(34) & Controls(Mygrid).ActiveSheet.Name & Chr(34) & ").Select" & vbCrLf
        txtMacro = txtMacro & Space(5) & Mygrid & ".Sheets(" & Chr(34) & Controls(Mygrid).ActiveSheet.Name & Chr(34) & ").Range(" & Chr(34) & Replace(Controls(Mygrid).Selection.Address, "$", "") & Chr(34) & ").Select" & vbCrLf
        txtMacro = txtMacro & Space(5) & Mygrid & ".Selection.ColumnWidth = 0" & vbCrLf
        
    Else
    If CollectionMenu("Dossier de Fabrication") = Mygrid Then
        txtMacro = txtMacro & Space(5) & "For I= 1 to " & Mygrid & ".Sheets.Count" & vbCrLf
        txtMacro = txtMacro & Space(10) & Mygrid & ".Sheets(I).Select" & vbCrLf
        txtMacro = txtMacro & Space(10) & Mygrid & ".Sheets(I).Range(" & Chr(34) & Replace(Controls(Mygrid).Selection.Address, "$", "") & Chr(34) & ").Select" & vbCrLf
        txtMacro = txtMacro & Space(10) & Mygrid & ".Selection.ColumnWidth = 0" & vbCrLf
        txtMacro = txtMacro & Space(5) & "NEXT" & vbCrLf
       
    Else
        txtMacro = txtMacro & Space(5) & Mygrid & ".Range(" & Chr(34) & Controls(Mygrid).Selection.Address & Chr(34) & ").ColumnWidth = 0" & vbCrLf
    End If
    End If
End If

Controls(Mygrid).Selection.ColumnWidth = 0
End Sub

Private Sub NewMacro_Click()
Dim Sql As String
Dim Rs As Recordset
Dim Trouve As Boolean
MacroName = ""
Reprise:
    MacroName = InputBox("Entrez le nom de la macro", "Nouvelle Macro", MacroName)
    If Trim("" & MacroName) <> "" Then
    MacroName = Replace(MacroName, " ", "_")
    Sql = "SELECT T_Macro.Formulaire, T_Macro.Macro "
    Sql = Sql & "FROM T_Macro "
    Sql = Sql & "WHERE T_Macro.Formulaire='" & Me.Name & "'  "
    Sql = Sql & "AND T_Macro.Macro='" & Replace(MacroName, "'", "''") & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Trouve = True
    Else
        Trouve = False
    End If
    Set Rs = Con.CloseRecordSet(Rs)
    If Trouve = True Then
        If MsgBox("La Macro : existe déjà voulez-vous le replacer", vbQuestion + _
            vbYesNo, "Nouvelle Macro") = vbNo Then GoTo Reprise
    End If
    If Trouve = False Then
    Sql = "INSERT INTO T_Macro ( Formulaire, Macro ) "
    Sql = Sql & "VALUES ( '" & Me.Name & "' , '" & Replace(Replace(MacroName, "'", "''"), Chr(34), Chr(34) & Chr(34)) & "' );"
Con.Execute Sql
    End If

        Me.txtMacro.Text = "Sub " & MacroName & "()" & vbCrLf
        Me.txtMacro.Text = Me.txtMacro.Text & "Dim  I" & vbCrLf
        StopMaco.Visible = True
    End If


End Sub

Private Sub Picture1_Click()
frmEditClip.charger Me.Fil, NumCollonne
End Sub

Private Sub Picture2_Click()
frmEditBouchon.charger Me.Conn, NumCollonne
End Sub

Private Sub Picture3_Click()
Combo1.Visible = True
End Sub

Private Sub PSA_Click()
If boolMajListe = False Then
    boolMajListe = True
     Me.ENC.ListIndex = MyCollectionENC(MyTablePSA(MyCollectionPSA("N" & Trim(Me.PSA.Text)), 1))
'     Me.PSA.ListIndex = MyCollectionPSA(MyTablePSA(MyCollectionPSA("N" & Me.PSA.Text), 2))
    Me.RSA.ListIndex = MyCollectionRSA(MyTablePSA(MyCollectionPSA("N" & Trim(Me.PSA.Text)), 3))
    Me.Hab.ListIndex = MyCollectionHab(MyTablePSA(MyCollectionPSA("N" & Trim(Me.PSA.Text)), 4))
    boolMajListe = False
End If
End Sub

Private Sub PSA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
 Me.PSA.ListIndex = MyCollectionPSA(MyTablePSA(MyCollectionPSA("N" & Trim(Me.PSA.Text)), 2))
    On Error GoTo 0
 End If
End Sub

Private Sub RSA_Click()
If boolMajListe = False Then
    boolMajListe = True
   
     Me.ENC.ListIndex = MyCollectionENC(MyTableRSA(MyCollectionRSA("N" & Trim(Me.RSA.Text)), 1))
     Me.PSA.ListIndex = MyCollectionPSA(MyTableRSA(MyCollectionRSA("N" & Trim(Me.RSA.Text)), 2))
'    Me.RSA.ListIndex = MyCollectionRSA(MyTableRSA(MyCollectionRSA("N" & Me.RSA.Text), 3))
    Me.Hab.ListIndex = MyCollectionHab(MyTableRSA(MyCollectionRSA("N" & Trim(Me.RSA.Text)), 4))
    boolMajListe = False
End If


End Sub

Private Sub RSA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
 Me.RSA.ListIndex = MyCollectionRSA(MyTableRSA(MyCollectionRSA("N" & Trim(Me.RSA.Text)), 3))
  On Error GoTo 0
 End If
End Sub

Private Sub Conn_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Static SaveRow As Long

Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Dim Row As Long
Dim Col As Long
Row = Me.Conn.ActiveCell.Row
Col = Me.Conn.ActiveCell.Column

If (NoMacro1Change = True Or NoMacro1Select = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro1Change = True

ValideConneceteur Me.Conn, Me.Crit.Range("a1").CurrentRegion, Row, Col
Dim aaa
On Error GoTo 0
'Set aaa = Me.Crit.Cells.Find("vba", Me.Crit.Cells(1, 1), ssValues, ssPart)

 NoMacro1Change = False
    Col3 = 0
    SaveRow = Row

Fin:
DoEvents
End Sub
Function ValideConneceteur(sheet, RangeSersh As Object, Row, Col)
Dim SplitCon
Dim Sql As String
Dim Rs As Recordset
   ConverOuiNon sheet.Range("a1").CurrentRegion, Row
   If sheet.Cells(Row, NumCollonne("conActiver")) = True Then
   sheet.Cells(Row, NumCollonne("conOPTION")) = UCase("" & sheet.Cells(Row, NumCollonne("conOPTION")))
   sheet.Cells(Row, NumCollonne("conCONNECTEUR")) = UCase("'" & sheet.Cells(Row, NumCollonne("conCONNECTEUR")))
    sheet.Cells(Row, NumCollonne("conCODE_APP")) = UCase("'" & sheet.Cells(Row, NumCollonne("conCODE_APP")))
If Trim("" & sheet.Cells(Row, NumCollonne("conOPTION"))) <> "" Then
If UCase(sheet.Cells(Row, NumCollonne("conOPTION"))) = "TOUS" Then
    sheet.Cells(Row, NumCollonne("conOPTION")) = "ALL"
End If
    If UCase(sheet.Cells(Row, NumCollonne("conOPTION"))) <> "ALL" Then
        
        aa = Split(UCase(Trim("" & sheet.Cells(Row, NumCollonne("conOPTION")))) & ";", ";")
        For Iaa = 0 To UBound(aa) - 1
        I = ChercheXls(UCase(Trim("" & aa(Iaa))), RangeSersh, True)

        If I = 0 Then


            msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbExclamation
             sheet.Cells(Row, NumCollonne("conOPTION")) = Replace(sheet.Cells(Row, NumCollonne("conOPTION")) & ";", aa(Iaa) & ";", "")
             If Right(sheet.Cells(Row, NumCollonne("conOPTION")), 1) = ";" Then sheet.Cells(Row, NumCollonne("conOPTION")) = Left(sheet.Cells(Row, NumCollonne("conOPTION")), Len(sheet.Cells(Row, NumCollonne("conOPTION"))))
                  sheet.Cells(Row, NumCollonne("conOPTION")).Select
                 Conn.SetFocus
        End If
        Next
     Set MyRange = Nothing
     End If

End If





        If Trim("" & sheet.Cells(Row, NumCollonne("conCONNECTEUR"))) <> "" Then
            sheet.Cells(Row, NumCollonne("conN°")) = Row - 1
        End If
End If


    If Row > 1 Then
        If Trim("" & sheet.Cells(Row, NumCollonne("ConCONNECTEUR"))) <> "" Then
            sheet.Cells(Row, NumCollonne("ConN°")) = Row - 1
            If NumCollonne("ConCONNECTEUR") = Col Then
                SplitCon = Trim("" & sheet.Cells(Row, NumCollonne("ConCONNECTEUR")))
                SplitCon = Split(SplitCon & "§", "§")
                Sql = "SELECT con_contacts." & refFour & " "
                Sql = Sql & "FROM con_contacts IN '"
                Sql = Sql & TableauPath("Eb_CONNECTEURS")
                Sql = Sql & "'"
                Sql = Sql & "WHERE con_contacts." & RefCli & "='" & Trim("" & SplitCon(0)) & "';"
                Set Rs = Con.OpenRecordSet(Sql)
                If Rs.EOF = False Then
                     If Trim("" & "" & Rs(0)) <> "" Then
                        sheet.Cells(Row, NumCollonne("ConRefConnecteurFour")) = "'" & Replace(Trim("" & sheet.Cells(Row, NumCollonne("ConCONNECTEUR"))), Trim("" & SplitCon(0)), "" & Rs(0))
                    End If
                Else
                    Sql = "SELECT con_contacts." & refFour & " "
                Sql = Sql & "FROM con_contacts IN '"
                Sql = Sql & TableauPath("Eb_CONNECTEURS")
                Sql = Sql & "'"
                Sql = Sql & "WHERE con_contacts." & refFour & "='" & Trim("" & SplitCon(0)) & "';"
                Set Rs = Con.OpenRecordSet(Sql)
                    If Rs.EOF = False Then
                        If Trim("" & "" & Rs(0)) <> "" Then
                            sheet.Cells(Row, NumCollonne("ConRefConnecteurFour")) = "'" & Replace(Trim("" & sheet.Cells(Row, NumCollonne("ConCONNECTEUR"))), Trim("" & SplitCon(0)), "" & Rs(0))
                        End If
                    End If
                End If
                Set Rs = Con.CloseRecordSet(Rs)
            End If
        End If
        
        If Trim("" & sheet.Cells(Row, NumCollonne("ConCode_app"))) <> "" Then
            Sql = "SELECT LIAISON_CONNECTEURS.LIB FROM LIAISON_CONNECTEURS "
            Sql = Sql & "WHERE LIAISON_CONNECTEURS.CLIENT='" & MyReplace(MyClient) & "' "
            Sql = Sql & "AND LIAISON_CONNECTEURS.LIAISON='" & MyReplace(DecodeCode_APP(sheet.Cells(Row, NumCollonne("ConCode_app")))) & "';"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = False Then
                sheet.Cells(Row, NumCollonne("ConCode_app") - 1) = Trim("'" & Rs!Lib)
            Else
                If IfValidationOk = False Then
'                    If MsgBox("Le code App : " & DecodeCode_APP(sheet.Cells(Row, NumCollonne("ConCode_app"))) & " n'existe pas" & vbCrLf & "Voulez-vous le créer", vbQuestion + vbYesNo, "Liaison Connecteur :") = vbYes Then
'                        LibCode_APP = InputBox("Entrez la désignation du code APP : " & DecodeCode_APP(sheet.Cells(Row, NumCollonne("ConCode_app"))), "Ajout d'un code App")
''                        If Trim(LibCode_APP) <> "" Then
''                            sheet.Cells(Row, NumCollonne("ConCode_app") - 1) = LibCode_APP
''                            sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
''                            sql = sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(DecodeCode_APP(sheet.Cells(Row, NumCollonne("ConCode_app"))))) & "', '" & UCase(MyReplace(sheet.Cells(Row, NumCollonne("ConCode_app") - 1))) & "' );"
''                            Con.Execute sql
''                        End If
'                    End If
                Else
                   Sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON "
                   Sql = Sql & "FROM Ajout_LIAISON_CONNECTEURS "
                   Sql = Sql & "WHERE Ajout_LIAISON_CONNECTEURS.LIAISON='" & UCase(MyReplace(DecodeCode_APP(sheet.Cells(Row, NumCollonne("ConCode_app"))))) & "' "
                   Sql = Sql & "AND Ajout_LIAISON_CONNECTEURS.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(Sql)
                    If Rs.EOF = True Then
                        Sql = "INSERT INTO Ajout_LIAISON_CONNECTEURS ( LIAISON, LIB,Job ) "
                        Sql = Sql & "values ( '" & UCase(MyReplace(DecodeCode_APP(sheet.Cells(Row, NumCollonne("ConCode_app"))))) & "', '" & MyReplace(sheet.Cells(Row, NumCollonne("ConCode_app") - 1)) & "'," & NmJob & ");"
                        Con.Execute Sql
                        MyErr = True
                    End If
                End If
            End If
            Set Rs = Con.CloseRecordSet(Rs)

        End If
    End If






End Function


Private Sub Conn_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Static NoMacro As Boolean
Static SaveRow As Long
Dim Sql As String
Dim Rs As Recordset
Row = Me.Conn.ActiveCell.Row
Col = Me.Conn.ActiveCell.Column

If (NoMacro = False) And (Row > 1) Then
NoMacro = True
NoMacro1Change = True
ValideConneceteur Me.Conn, Me.Crit.Range("a1").CurrentRegion, Row, Col

    NoMacro1Change = False
    NoMacro = False
    Col3 = 0
End If
End Sub

Private Sub Nom_SheetChange(ByVal Sh As OWC10.Worksheet, ByVal Target As OWC10.Range)
If Me.StopMaco.Visible = True Then
    txtMacro = txtMacro & Space(5) & "SSTab1.Tab =" & SSTab1.Tab & vbCrLf
    txtMacro = txtMacro & Space(5) & Mygrid & ".Sheets(" & Chr(34) & Controls(Mygrid).ActiveSheet.Name & ").Select" & Chr(34) & vbCrLf
End If
'
End Sub



Private Sub Fil_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)

ValideFils Me.Fil
End Sub

Private Sub Fil_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Static SaveRow As Long
Dim Row As Long
Dim Sql As String
Dim Rs As Recordset
Row = Me.Fil.ActiveCell.Row
If (NoMacro2 = True) Or (Row = 1) Then GoTo Fin
If SaveRow = 0 Then SaveRow = Me.Fil.ActiveCell.Row

 If Trim("" & Me.Fil.Cells(SaveRow, NumCollonne("filsLIAI"))) <> "" Then
                Sql = "SELECT LIAISON.LIB FROM LIAISON "
                Sql = Sql & "WHERE LIAISON.CLIENT='" & MyReplace(MyClient) & "' "
                Sql = Sql & "AND LIAISON.LIAISON='" & MyReplace(Me.Fil.Cells(SaveRow, NumCollonne("filsLIAI"))) & "';"
                Set Rs = Con.OpenRecordSet(Sql)
                If Rs.EOF = False Then
                Me.Fil.Cells(SaveRow, NumCollonne("filsDESIGNATION")) = Trim("'" & Rs!Lib)
                Else
                    If IfValidationOk = False Then
                        If SaveRow <> Row And SaveRow <> 1 Then
'                            If MsgBox("La liaison : " & Me.Fil.Cells(SaveRow,NumCollonne("filsLIAI")) & " n'existe pas" & vbCrLf & "Voulez-vous la créer", vbYesNo + vbQuestion, "AutoCâble: Tableau de fils") = vbYes Then
'                                LibCode_APP = InputBox("Entrez la désignation de la liaison : " & Me.Fil.Cells(SaveRow, 1), "Ajout de liaison")
''                                If Trim(LibCode_APP) <> "" Then
''                                    Me.Fil.Cells(SaveRow,NumCollonne("filsDESIGNATION")) = LibCode_APP
''                                    sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
''                                    sql = sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(Me.Fil.Cells(SaveRow,NumCollonne("filsLIAI")))) & "', '" & UCase(MyReplace("" & LibCode_APP)) & "' );"
''                                    Con.Execute sql
''                                End If
'                            End If
                        End If
                        Else
                      Set Rs = Con.CloseRecordSet(Rs)
                   Sql = "SELECT Ajout_LIAISON.LIAISON "
                   Sql = Sql & "FROM Ajout_LIAISON "
                   Sql = Sql & "WHERE Ajout_LIAISON.LIAISON='" & UCase(MyReplace(Me.Fil.Cells(SaveRow, NumCollonne("filsLIAI")))) & "' "
                   Sql = Sql & "AND Ajout_LIAISON.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(Sql)
                    If Rs.EOF = True Then
                        Sql = "INSERT INTO Ajout_LIAISON ( LIAISON, LIB,Job ) "
                        Sql = Sql & "values ( '" & UCase(MyReplace(Me.Fil.Cells(SaveRow, NumCollonne("filsLIAI")))) & "', '" & MyReplace(Me.Fil.Cells(SaveRow, NumCollonne("filsDESIGNATION"))) & "'," & NmJob & ");"
                        Con.Execute Sql
                        MyErr = True
                    End If
                    End If
                End If
                Set Rs = Con.CloseRecordSet(Rs)
            If Trim("" & Me.Fil.Cells(SaveRow, NumCollonne("filsLIAI"))) <> "" And SaveRow <> Row Then
                        If UCase(Trim("" & Me.Fil.Cells(SaveRow, NumCollonne("filsACTIVER")))) <> 0 Then
                        If Len(Trim("" & Me.Fil.Cells(SaveRow, NumCollonne("filsApp")))) = 0 And Me.Fil.Cells(Row, NumCollonne("filsACTIVER")) = True Then

                            msg = "Le code APP ne peut être Nul"
                            MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                             Me.Fil.Cells(SaveRow, NumCollonne("filsApp")).Select
                              Me.Fil.SetFocus
                               Row = SaveRow
                              GoTo Fin
                          End If
                         If Len(Trim("" & Me.Fil.Cells(SaveRow, NumCollonne("filsApp2")))) = 0 And Me.Fil.Cells(Row, NumCollonne("filsACTIVER")) = True Then

                            msg = "Le code APP ne peut être Nul"
                            MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                             Me.Fil.Cells(SaveRow, NumCollonne("filsApp2")).Select
                              Me.Fil.SetFocus
                               Row = SaveRow
                          End If
                        End If
                    End If
            End If
Fin:

SaveRow = Row
End Sub

Private Sub Comp_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveRow As Long
Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Dim BoolOui As Boolean
Dim txtCompPath As String
If SaveRow = 0 Then SaveRow = 1


Row = Me.Comp.ActiveCell.Row
Col = Me.Comp.ActiveCell.Column
If (NoMacro3 = True) Or (Row = 1) Then GoTo Fin

If SaveRow = 0 Then SaveRow = Row
If Row = 1 Then GoTo Fin
boolActu = False
NoMacro3 = True
Set MyRange = Nothing
 Set MyRange = Me.Comp.Range("a1").CurrentRegion
   ConverOuiNon MyRange, Row
   If Trim("" & Me.Comp.Cells(Row, 2)) <> "" Then Me.Comp.Cells(Row, 3) = Row - 1
   If Trim("" & Me.Comp.Cells(Row, 5)) <> "" Then
    Me.Comp.Cells(Row, 5) = UCase(Me.Comp.Cells(Row, 5))
'    Set Myrange = Me.Crit.ActiveSheet.Range("a1").CurrentRegion
'        Set Myrange = Me.Crit.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Myrange.Rows.Count))
 aa = Split(Me.Comp.Cells(Row, 5) & ";", ";")
                For Iaa = 0 To UBound(aa) - 1

        I = ChercheXls(UCase(Trim("" & aa(Iaa))), Me.Crit.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Me.Crit.Range("a1").CurrentRegion.Rows.Count)))

        If I = 0 Then
                msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbQuestion, "AutoCâble: Composants"
                     Me.Comp.Cells(Row, 5) = Replace(Me.Comp.Cells(Row, 5) & ";", "" & Trim(aa(Iaa)) & ";", "")
                Comp.SetFocus
        End If
        Next
        If Right(Me.Comp.Cells(Row, 5), 1) = ";" Then Me.Comp.Cells(Row, 5) = Left(Me.Comp.Cells(Row, 5), Len(Me.Comp.Cells(Row, 5)) - 1)
   End If
   If Me.Comp.Cells(Row, NumCollonne("Compactiver")) <> 0 Then
        If Trim("" & Me.Comp.Cells(Row, NumCollonne("CompREFCOMP"))) <> "" Then
            If Trim("" & Me.Comp.Cells(Row, NumCollonne("CompPath"))) <> "" Then
            On Error Resume Next
            txtCompPath = ""
                txtCompPath = CollectionPath(Trim("" & Me.Comp.Cells(Row, NumCollonne("CompPath"))))
                If Err Then
                    Err.Clear
                        msg = "?"
                    MsgBox "Le répertoire que vous avez saisi n'existe pas.", vbQuestion, "AutoCâble: Composants"
                     Me.Comp.Cells(Row, NumCollonne("CompPath")) = ""
                     On Error GoTo 0
                     GoTo Sortie
                End If
                On Error GoTo 0
            Else
                    msg = "?"
                   MsgBox "Vous devez saisir ou sélectionner le nom d'un répertoire.", vbQuestion, "AutoCâble: Composants"
                   GoTo Sortie
                
            End If
        End If
   End If
'   If Col > NumCollonne("CompPOS-OUT") Then
'    For I = NumCollonne("CompPOS-OUT") + 1 To NbFinOuiNon
'        If Me.Comp.Cells(Row, I) = 1 And Me.Comp.Cells(Row, Col) = 1 Then
'            If I <> Col Then
'                MsgBox "Vous ne pouvez pas sélectionner plusieurs répertoires.", vbQuestion, "AutoCâble: Composants"
'                Me.Comp.Cells(Row, Col) = 0
'               Exit For
'            End If
'        End If
'    Next I
'End If
If Col = NumCollonne("CompCODE_APP_LIER") Then

    If ChercheXls(Comp.ActiveCell.Value, Conn.Range("a1").CurrentRegion) = 0 Then
        MsgBox "Le code App : " & Comp.ActiveCell.Value & " n'existe pas dans la liste des connecteur.", vbExclamation
        Comp.ActiveCell.Clear
    End If
End If
Sortie:
   NoMacro3 = False
Fin:
End Sub

Private Sub Comp_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveCol As Long
Static SaveRow As Long
Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Dim BoolOui As Boolean
If SaveRow = 0 Then SaveRow = 1
If SaveCol = 0 Then SaveCol = 1
If NoMacro3 = True Then GoTo Fin

Row = Me.Comp.ActiveCell.Row
Col = Me.Comp.ActiveCell.Column

If Row = 1 Then GoTo Fin
NoMacro3 = True
Set MyRange = Nothing
 Set MyRange = Me.Conn.Range("a1").CurrentRegion
   ConverOuiNon MyRange, Row
   If Trim("" & Me.Comp.Cells(Row, 2)) <> "" Then Me.Comp.Cells(Row, 3) = Row - 1

If Col > NumCollonne("CompPOS-OUT") Then
    For I = NumCollonne("CompPOS-OUT") + 1 To NbFinOuiNon
        If Me.Comp.Cells(Row, I) = 1 And Me.Comp.Cells(Row, Col) = 1 Then
            If I <> Col Then
                MsgBox "Vous ne pouvez pas sélectionner plusieurs répertoires.", vbQuestion, "AutoCâble: Composants"
                Me.Comp.Cells(Row, Col) = 0
               Exit For
            End If
        End If
    Next I
End If
BoolOui = False
If (SaveRow <> 1) And (SaveRow <> Row) And (Trim("" & Me.Comp.Cells(SaveRow, 1)) <> "") And Col > NumCollonne("CompPOS-OUT") Then
 For I = NumCollonne("CompPOS-OUT") + 1 To NbFinOuiNon
    If Val(Me.Comp.Cells(SaveRow, I)) = 1 Then
        BoolOui = True

        Exit For
    End If

    Next I
  If BoolOui = False And Me.Comp.Cells(SaveRow, 1) = 1 Then
    MsgBox "Vous devez sélectionner un répertoire.", vbQuestion, "AutoCâble: Composants"
    Me.Comp.Cells(SaveRow, NumCollonne("CompCODE_APP_LIER") + 1).Select
     Me.Comp.SetFocus
     Row = SaveRow
    msg = "?"
  End If

End If
SaveRow = Row
SaveCol = Col
NoMacro3 = False
Fin:
End Sub

Private Sub Notas_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Row = Me.Notas.ActiveCell.Row
Col = Me.Notas.ActiveCell.Column
If Row = 1 Then GoTo Fin
If (NoMacro4 = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro4 = True
Set MyRange = Nothing
 Set MyRange = Me.Notas.Range("a1").CurrentRegion
   ConverOuiNon MyRange, Row
If Col > 2 Then
   If Trim("" & Me.Notas.Cells(Row, 2)) <> "" Then Me.Notas.Cells(Row, 3) = Row - 1

    If Trim("" & Me.Notas.Cells(Row, 4)) <> "" Then
    Me.Notas.Cells(Row, 4) = UCase(Me.Notas.Cells(Row, 4))
 aa = Split(Me.Notas.Cells(Row, 4) & ";", ";")
                For Iaa = 0 To UBound(aa) - 1

        I = ChercheXls(UCase(Trim("" & aa(Iaa))), Me.Crit.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Me.Crit.Range("a1").CurrentRegion.Rows.Count)))

        If I = 0 Then
                msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbQuestion, "AutoCâble: Notas"
                     Me.Notas.Cells(Row, 4) = Replace(Me.Notas.Cells(Row, 4) & ";", "" & Trim(aa(Iaa)) & ";", "")
                Notas.SetFocus
        End If
        Next
        If Right(Me.Notas.Cells(Row, 4), 1) = ";" Then Me.Notas.Cells(Row, 4) = Left(Me.Notas.Cells(Row, 4), Len(Me.Notas.Cells(Row, 4)) - 1)
   End If
End If
NoMacro4 = False

Fin:
End Sub
Private Sub MosHor_Click()
frmAutocâble.Arrange vbTileVertical

End Sub

Private Sub MosVer_Click()
frmAutocâble.Arrange vbTileHorizontal
End Sub
Public Sub chargement(MeCapTion As String, IdIndiceProjet As Long, Client As String, FRM As Object, Optional Edition As Boolean, Optional Modifiable As Boolean)
Dim Rs As Recordset
Dim Sql As String
Dim txtMyCollectionLienHab As String
Dim RsIdProjet As Recordset
Dim PathModelXls As String
Dim NbEregistrement As Long
Dim RsOnglet As Recordset
Dim MyErr As String
Dim PathComposantsDefault As String
Dim NbColonne As Long
Dim Fso As New FileSystemObject
On Error Resume Next
If Modifiable = True Then
    Me.Command1.Visible = False
    Me.Command2.Visible = False
    Me.Command3.Visible = False
End If
'Sql = "SELECT  T_Clients.ChampCli FROM T_Clients "
'Sql = Sql & "Where T_Clients.Client = '" & MyReplace(Client) & "' "
'Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
'    RefCli = "" & Rs(0)
'Else
   RefCli = GetDefault("DefaultFour" & Client, "txt3")
'End If
refFour = GetDefault("DefaultFour" & Client, "txt3")
Set Rs = Con.CloseRecordSet(Rs)
Set CollectionPath = Nothing
Set CollectionPath = New Collection
Sql = ""
PathComposantsDefault = DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath.Item("PathComposantsDefault"))
    Dim fs, f, f1, s, sf

'  MyExcel.Visible = True
    Set f = Fso.GetFolder(PathComposantsDefault) '\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS\")
    Set sf = f.SubFolders
 Me.Combo1.AddItem ""
 Set sf = f.SubFolders
    For Each f1 In sf
      
        
        Me.Combo1.AddItem f1.Name
        CollectionPath.Add f1.Name, f1.Name

    Next
    Set Fso = Nothing
  
Me.Visible = True
NoMacro1Change = True
NoMacro2 = True
NoMacro5 = True
NoMacro3 = True
Set NumCollonne = Nothing
Set NumCollonne = New Collection
Set Myfrm = FRM
MyIdIndiceProjet = IdIndiceProjet
Me.Tag = IdIndiceProjet
'MyPathXlsMoins1 = BackUp(Xls & ".XLS", True, MyPathXlsMoins1)

 
    Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathCatalogue FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathCatalogue) = "" Then
         DbCatalogue = ""
   Else
             DbCatalogue = RsConnecteur!PathCatalogue
             DbCatalogue = DefinirChemienComplet(TableauPath.Item("PathServer"), DbCatalogue)
'         If Left(DbCatalogue, 2) <> "\\" And Left(DbCatalogue, 1) = "\" Then DbCatalogue = TableauPath.Item("PathServer") & DbCatalogue
'            If Right(DbCatalogue, 2) = "\\" Then DbCatalogue = Mid(DbCatalogue, 1, Len(DbCatalogue) - 1)
    
    End If
Else
    DbCatalogue = ""
End If
Set CollectionMenu = Nothing
Set CollectionMenu = New Collection

 CollectionMenu.Add "Crit", "Critères"
 CollectionMenu.Add "Conn", "Connecteurs"
 CollectionMenu.Add "Fil", "Tableau de fils"
 CollectionMenu.Add "Comp", "Composants"
 CollectionMenu.Add "Notas", "Notas"
 CollectionMenu.Add "Noeuds", "Noeuds"
' CollectionMenu.Add "Spreadsheet7", "Nomenclature Connecteur"
' CollectionMenu.Add "Spreadsheet8", "Nomenclature Fils"
' CollectionMenu.Add "Spreadsheet9", "Nomenclature Habillage"
 CollectionMenu.Add "Nom", "Nomenclatures"
 CollectionMenu.Add "Fab", "Dossier de Fabrication"
 CollectionMenu.Add "Cont", "Dossier de Contrôle"


'
'PathModelXls = TableauPath.Item("PathModelXls")
'PathModelXls = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelXls)
'         If Left(PathModelXls, 2) <> "\\" And Left(PathModelXls, 1) = "\" Then PathModelXls = TableauPath.Item("PathServer") & PathModelXls
'          If Right(PathModelXls, 2) = "\\" Then PathModelXls = Mid(PathModelXls, 1, Len(PathModelXls) - 1)

'***********************************************************************************************************************
''*                                       Ouvre le Modèle Excel.                                                        *
'    If MyPathXlsMoins1 <> "" Then
'    If NomenclatureOk = True Then
'        Set MyWorkbook = OpenModelXlt(MyPathXlsMoins1)
'    Else
'        Set MyWorkbook = OpenModelXlt(PathModelXls)
'    End If
'    Else
'        Set MyWorkbook = OpenModelXlt(PathModelXls)
'    End If
'    MyWorkbook.Application.Visible = True
'    RetournIdApp "EXCEL.EXE", True
'    MyWorkbook.Application.Visible = True
'***********************************************************************************************************************
'*                                      Exporte la liste des T_Noeuds.                                              *


    Sql = "SELECT T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL, T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR,  "
    Sql = Sql & "T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA,  "
    Sql = Sql & "T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T, T_Noeuds.OPTION,T_Noeuds.Commentaires, T_Noeuds.ID "
    Sql = Sql & "FROM T_Noeuds "
    Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Noeuds.NŒUDS;"
    Set Rs = Con.OpenRecordSet(Sql)

'    ExporteXlsNoeuds Rs, IdIndiceProjet
'Noeuds.Range("A1").Find
For I = 0 To Rs.Fields.Count - 1


    If Rs.Fields(I).Type = adBoolean Then
        ColecBool.Add True, Rs.Fields(I).Name
    Else
        ColecBool.Add False, Rs.Fields(I).Name
    End If
Next
    Copy_Rs_Spreadsheet Me, Noeuds, Rs, "Noeu", Me, "Exportation des nœuds"
   
    
''    Noeuds.Columns(NumCollonne("NoeuActiver")).NumberFormat = "Yes/No"
'Noeuds.Columns(NumCollonne("NoeuFleche_Droite")).NumberFormat = "Yes/No"
'Noeuds.Columns(NumCollonne("NoeuTORON_PRINCIPAL")).NumberFormat = "Yes/No"

    
'***********************************************************************************************************************
'*                                      Exporte la liste des Critères.                                              *

    Sql = "SELECT T_Critères.ACTIVER, T_Critères.CODE_CRITERE, T_Critères.CRITERES,T_Critères.DESIGNATION,T_Critères.Commentaires,T_Critères.ID FROM T_Critères "
    Sql = Sql & "WHERE T_Critères.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE;"
    Set Rs = Con.OpenRecordSet(Sql)
     Charger_Colection Conn.ActiveSheet, "Con"


For I = 0 To Rs.Fields.Count - 1


    If Rs.Fields(I).Type = adBoolean Then
        ColecBool.Add True, Rs.Fields(I).Name
    Else
        ColecBool.Add False, Rs.Fields(I).Name
    End If
Next
    Copy_Rs_Spreadsheet Me, Crit, Rs, "Crit", Me, "Exportation des Critères"
'    Crit.Columns(NumCollonne("CritActiver")).NumberFormat = "Yes/No"
'    ExporteXlsCriteres Rs, IdIndiceProjet
    ExporteFrmCriteresFils IdIndiceProjet, Crit
'***********************************************************************************************************************
'*                                      Exporte la liste des connecteurs.                                              *

    Sql = "SELECT Connecteurs.ACTIVER, Connecteurs.CONNECTEUR, Connecteurs.RefConnecteurFour,  "
    Sql = Sql & "Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP, Connecteurs.N°,  "
    Sql = Sql & "Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2,  "
    Sql = Sql & "Connecteurs.OPTION, Connecteurs.[100%], Connecteurs.Pylone, Connecteurs.Colonne,  "
    Sql = Sql & "Connecteurs.Ligne, Connecteurs.RefBouchon, Connecteurs.RefBouchonFour, Connecteurs.ReFCapot,  "
    Sql = Sql & "Connecteurs.ReFCapotFour, Connecteurs.RefVerrou, Connecteurs.RefVerrouFour, Connecteurs.LongueurF_Choix,Connecteurs.Commentaires,Connecteurs.ID  "
    Sql = Sql & "FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Connecteurs.N°;"
    Set Rs = Con.OpenRecordSet(Sql)
    
 For I = 0 To Rs.Fields.Count - 1


    If Rs.Fields(I).Type = adBoolean Then
        ColecBool.Add True, Rs.Fields(I).Name
    Else
        ColecBool.Add False, Rs.Fields(I).Name
    End If
Next
Copy_Rs_Spreadsheet Me, Conn, Rs, "Con", Me, "Exportation des Connecteurs"
Conn.ActiveSheet.Range("a1").Select
'Conn.Columns(NumCollonne("ConActiver")).NumberFormat = "Yes/No"
'Conn.Columns(NumCollonne("ConO/N")).NumberFormat = "Yes/No"
'    ExporteXlsConnecteur Rs, IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte la liste des Fils.                                                     *
'
'    Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
'   Sql = Sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
'   Sql = Sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
'   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
'   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
'   Sql = Sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
'   Sql = Sql & "FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val('' & Ligne_Tableau_fils.FIL);"
'    Set Rs = Con.OpenRecordSet(Sql)
    
'
'     sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
'   sql = sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
'   sql = sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
'   sql = sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI,Ligne_Tableau_fils. Ligne_Tableau_fils.POS2,  "
'   sql = sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
'   sql = sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION,Ligne_Tableau_fils.[Critères spécifiques] "
'   sql = sql & "FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val('' & Ligne_Tableau_fils.FIL);"
   
   Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
   Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,  "
   Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
   Sql = Sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.Long_Add,Ligne_Tableau_fils.Long_Add2,Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,Ligne_Tableau_fils.VOI,Ligne_Tableau_fils.[Ref Connecteur], Ligne_Tableau_fils.[Ref Connecteur_Four],  "
   Sql = Sql & "   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four],Ligne_Tableau_fils.PRECO,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint Four],  "
   Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
   Sql = Sql & "Ligne_Tableau_fils.APP2,Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.[Ref Connecteur2], Ligne_Tableau_fils.[Ref Connecteur_Four2],   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],Ligne_Tableau_fils.PRECO2,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint2], Ligne_Tableau_fils.[Ref Joint Four2],  "
   Sql = Sql & "Ligne_Tableau_fils.PRECOG, Ligne_Tableau_fils.OPTION,  "
   Sql = Sql & "Ligne_Tableau_fils.[Critères spécifiques],Ligne_Tableau_fils.Commentaires,Ligne_Tableau_fils.ID  "
   Sql = Sql & "FROM Ligne_Tableau_fils "
   Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & IdIndiceProjet & " "
   Sql = Sql & "ORDER BY Val('' & Ligne_Tableau_fils.FIL);"
    
    Set Rs = Con.OpenRecordSet(Sql)
'    ExporteXlsFils Rs, IdIndiceProjet
For I = 0 To Rs.Fields.Count - 1


    If Rs.Fields(I).Type = adBoolean Then
        ColecBool.Add True, Rs.Fields(I).Name
    Else
        ColecBool.Add False, Rs.Fields(I).Name
    End If
Next
    Copy_Rs_Spreadsheet Me, Fil, Rs, "Fils", Me, "Exportation des Fils"
'    Fil.Columns(NumCollonne("FilsActiver")).NumberFormat = "Yes/No"
    
    
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Composants.                                              *
       
    Sql = "SELECT Composants.ACTIVER, Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.OPTION, Composants.Code_APP_Lier, Composants.Voie,Composants.POS, Composants.[POS-OUT],Composants.Commentaires ,Composants.Path,Composants.ID "
    Sql = Sql & "FROM Composants "
    Sql = Sql & "WHERE Composants.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Composants.NUMCOMP;"
    Set Rs = Con.OpenRecordSet(Sql)
    
 
     Copy_Rs_Spreadsheet Me, Comp, Rs, "Comp", Me, "Exportation des Composants"
'    Comp.Columns(NumCollonne("CompACTIVER")).NumberFormat = "Yes/No"
'    ExporteXlsComposants Rs, IdIndiceProjet
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Notas.                                                    *
    
    Sql = "SELECT Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA ,Nota.OPTION,Nota.Commentaires,Nota.ID FROM Nota "
    Sql = Sql & "WHERE Nota.Id_IndiceProjet= " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Nota.NUMNOTA ;"

    Set Rs = Con.OpenRecordSet(Sql)
'    ExporteXlsNotas Rs, IdIndiceProjet

    Copy_Rs_Spreadsheet Me, Notas, Rs, "Not", Me, "Exportation des Notas"
'    Notas.Columns(NumCollonne("NotActiver")).NumberFormat = "Yes/No"
'***********************************************************************************************************************
'*                                      Si NomenclatureOk= faux alors génère la nomenclature.                          *

If NomenclatureOk = False Then
Sql = "SELECT Rq_Cable_Prix.* FROM Rq_Cable_Prix "
    Sql = Sql & "WHERE Rq_Cable_Prix.Id_IndiceProjet=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
    ExporteXlsPrixFils Rs, IdIndiceProjet

Sql = "SELECT Rq_Habillages_Prix.* FROM Rq_Habillages_Prix "
    Sql = Sql & "Where Rq_Habillages_Prix.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Rq_Habillages_Prix.DESIGN_HAB;"
Set Rs = Con.OpenRecordSet(Sql)
    
    ExporteXlsHabillages Rs, IdIndiceProjet
'    NomenclatureOk = Nomenclature3(IdIndiceProjet, PathPl, Save)

End If
If Edition = True Then
    Sql = "SELECT T_Nomenclature.CONNECTEUR,T_Nomenclature.[Nb Voies], T_Nomenclature.OPTION, "
    Sql = Sql & "T_Nomenclature.Qté, T_Nomenclature.[Prix U], T_Nomenclature.[Prix Total],  "
    Sql = Sql & "T_Nomenclature.CODE_APP, T_Nomenclature.DESIGNATION, T_Nomenclature.Couleur,  "
    Sql = Sql & "T_Nomenclature.[Lib Connecteur], T_Nomenclature.Fournisseur,  "
    Sql = Sql & "T_Nomenclature.[Ref Four], T_Nomenclature.[Ref Bouch],  "
    Sql = Sql & "T_Nomenclature.[Bouchon Qté], T_Nomenclature.[Bouchon Prix U],  "
    Sql = Sql & "T_Nomenclature.[Bouchon Prix Total], T_Nomenclature.[Lib Bouch],  "
    Sql = Sql & "T_Nomenclature.[Bouch Fourr], T_Nomenclature.[Bouch Réf Four],  "
    Sql = Sql & "T_Nomenclature.[Ref Capot], T_Nomenclature.[Ref Verrou],  "
    Sql = Sql & "T_Nomenclature.[Ref Joint], T_Nomenclature.[Joint Qté],  "
    Sql = Sql & "T_Nomenclature.[Joint Prix U], T_Nomenclature.[Joint Prix Total],  "
    Sql = Sql & "T_Nomenclature.[Lib Joint], T_Nomenclature.[Joint Four],  "
    Sql = Sql & "T_Nomenclature.[Joint Four Réf], T_Nomenclature.[Nb Alvé],  "
    Sql = Sql & "T_Nomenclature.Voie, T_Nomenclature.Famille, T_Nomenclature.[Famille Lib],  "
    Sql = Sql & "T_Nomenclature.[Alvé Réf], T_Nomenclature.[Alvé Qté],  "
    Sql = Sql & "T_Nomenclature.[Alvé Prix U],  "
    Sql = Sql & "T_Nomenclature.[Alvé Prix Total], T_Nomenclature.[Alvé Réf Fourr],  "
    Sql = Sql & "T_Nomenclature.[Alvéole Mini en mm2], T_Nomenclature.[Alvéole Maxi en mm2] "
    Sql = Sql & "FROM T_Nomenclature "
    Sql = Sql & "WHERE T_Nomenclature.Id_IndiceProjet=" & IdIndiceProjet & ";"
    Set Rs = Con.OpenRecordSet(Sql)
'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Connecteur :"
'    Rs.Requery
'    MyWorkbook.Application.Visible = True
'Me.Nom.Sheets(1).Range("a1").Select
Copy_Rs_Spreadsheet Me, Nom.Sheets(1), Rs, "NomCon", Me, "Nomenclature Connecteur"
'me.Charger_Colection me.Nom.Sheets(1), "NomCon"
'    'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Connecteur", IdIndiceProjet
'
'    'ReplaceNull MyWorkbook.Worksheets("Nomenclature Connecteur"), Chr(10), "©"
    
    Sql = "SELECT T_Prix_Fils.TEINT, T_Prix_Fils.OPTION, T_Prix_Fils.ISO, T_Prix_Fils.SECT, T_Prix_Fils.Longeur,  "
    Sql = Sql & "T_Prix_Fils.[Prix U], T_Prix_Fils.[Prix Total] "
    Sql = Sql & "FROM T_Prix_Fils "
    Sql = Sql & "WHERE T_Prix_Fils.Id_IndiceProjet=" & IdIndiceProjet & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    Copy_Rs_Spreadsheet Me, Nom.Sheets(2), Rs, "NomFil", Me, "Nomenclature Fils"
'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Fils :"
'    Rs.Requery
    
    'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Fils", IdIndiceProjet
    'ReplaceNull MyWorkbook.Worksheets("Nomenclature Fils"), Chr(10), "©"
    Sql = "SELECT T_Appro_Habillage.DESIGN_HAB, T_Appro_Habillage.OPTION, T_Appro_Habillage.Qté, T_Appro_Habillage.[Prix U],  "
    Sql = Sql & "T_Appro_Habillage.[Prix Total], T_Appro_Habillage.CODE_ENC "
    Sql = Sql & "FROM T_Appro_Habillage "
    Sql = Sql & "WHERE T_Appro_Habillage.Id_IndiceProjet=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
Copy_Rs_Spreadsheet Me, Nom.Sheets(3), Rs, "NimHab", Me, "Nomenclature Habillage"
'    Copy_Rs_Spreadsheet me, me.Nom(3), Rs, "NimHab"
'
'
'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Habillage :"
'    Rs.Requery
    
    'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Habillage", IdIndiceProjet
'ReplaceNull MyWorkbook.Worksheets("Nomenclature Habillage"), Chr(10), "©"


Sql = "SELECT Nomenclature2.LIAI, Nomenclature2.Designation, Nomenclature2.App, Nomenclature2.Voie, Nomenclature2.Ref,  "
    Sql = Sql & "Nomenclature2.RefFour, Nomenclature2.App2, Nomenclature2.Voie2, Nomenclature2.Options, Nomenclature2.ISO,  "
    Sql = Sql & "Nomenclature2.Longueur, Nomenclature2.[Longueur Total], Nomenclature2.TEINT, Nomenclature2.TEINT2,  "
    Sql = Sql & "Nomenclature2.SECT, Nomenclature2.Qts "
    Sql = Sql & "FROM Nomenclature2 "
    Sql = Sql & "WHERE Nomenclature2.Id_IndiceProjet=" & IdIndiceProjet & ";"
     Set Rs = Con.OpenRecordSet(Sql)

'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature :"
'    Rs.Requery
Copy_Rs_Spreadsheet Me, Nom.Sheets(4), Rs, "Nom", Me, "Nomenclature"
        'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature", IdIndiceProjet
'      MyWorkbook.Application.Visible = False
    Sql = "SELECT NomenclaturFinal.Designation, NomenclaturFinal.Famille, NomenclaturFinal.Ref,NomenclaturFinal.RefFour, NomenclaturFinal.Fournisseur, NomenclaturFinal.Qts,  "
    Sql = Sql & " NomenclaturFinal.ISO, NomenclaturFinal.TEINT, NomenclaturFinal.TEINT2, NomenclaturFinal.SECT,  "
    Sql = Sql & "NomenclaturFinal.Qts_Encelade, NomenclaturFinal.Qts_E_Boutique, NomenclaturFinal.Qts_Appro, NomenclaturFinal.Prix_Revient,  "
    Sql = Sql & "NomenclaturFinal.Prix_Vente, NomenclaturFinal.Options  "
    Sql = Sql & "FROM NomenclaturFinal "
    Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & IdIndiceProjet & ";"
     Set Rs = Con.OpenRecordSet(Sql)
     Copy_Rs_Spreadsheet Me, Nom.Sheets(5), Rs, "NomF", Me, "Nomenclature Finale"
'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Finale :"
'    Rs.Requery
    
'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Finale", IdIndiceProjet


Sql = "SELECT  T_Dossier_Fabrication.Onglet, T_Dossier_Fabrication.ACTIVER, T_Dossier_Fabrication.LIAI,  "
    Sql = Sql & "T_Dossier_Fabrication.DESIGNATION, T_Dossier_Fabrication.FIL, T_Dossier_Fabrication.SECT, T_Dossier_Fabrication.TEINT,  "
    Sql = Sql & "T_Dossier_Fabrication.TEINT2, T_Dossier_Fabrication.ISO, T_Dossier_Fabrication.LONG, T_Dossier_Fabrication.[LONG CP],  "
    Sql = Sql & "T_Dossier_Fabrication.LONG_ADD, T_Dossier_Fabrication.LONG_ADD2, T_Dossier_Fabrication.COUPE, T_Dossier_Fabrication.POS,  "
    Sql = Sql & "T_Dossier_Fabrication.[POS-OUT], T_Dossier_Fabrication.FA, T_Dossier_Fabrication.APP, T_Dossier_Fabrication.VOI,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR], T_Dossier_Fabrication.[REF CONNECTEUR_FOUR], T_Dossier_Fabrication.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR], T_Dossier_Fabrication.PRECO, T_Dossier_Fabrication.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR], T_Dossier_Fabrication.POS2, T_Dossier_Fabrication.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Fabrication.FA2, T_Dossier_Fabrication.APP2, T_Dossier_Fabrication.VOI2, T_Dossier_Fabrication.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR_FOUR2], T_Dossier_Fabrication.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR2], T_Dossier_Fabrication.PRECO2, T_Dossier_Fabrication.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR2], T_Dossier_Fabrication.PRECOG, T_Dossier_Fabrication.Option,  "
    Sql = Sql & "T_Dossier_Fabrication.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Dossier_Fabrication.Id;"
    
    
    Sql = "SELECT T_Dossier_Fabrication.Onglet  "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & " "
    Sql = Sql & "GROUP BY Onglet; "

Set Rs = Con.OpenRecordSet(Sql)

    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Max = NbEregistrement
    Me.ProgressBar1Caption.Caption = " Exporter liste Dossier de Fabrication :"
    Rs.Requery
    Dim Ionglet As Long
    Ionglet = 0
    Do While Rs.EOF = False
    Ionglet = Ionglet + 1
   IncremanteBarGrah Me
   Sql = "SELECT  T_Dossier_Fabrication.Onglet, T_Dossier_Fabrication.ACTIVER, T_Dossier_Fabrication.LIAI,  "
    Sql = Sql & "T_Dossier_Fabrication.DESIGNATION, T_Dossier_Fabrication.FIL, T_Dossier_Fabrication.SECT, T_Dossier_Fabrication.TEINT,  "
    Sql = Sql & "T_Dossier_Fabrication.TEINT2, T_Dossier_Fabrication.ISO, T_Dossier_Fabrication.LONG, T_Dossier_Fabrication.[LONG CP],  "
    Sql = Sql & "T_Dossier_Fabrication.LONG_ADD, T_Dossier_Fabrication.LONG_ADD2, T_Dossier_Fabrication.COUPE, T_Dossier_Fabrication.POS,  "
    Sql = Sql & "T_Dossier_Fabrication.[POS-OUT], T_Dossier_Fabrication.FA, T_Dossier_Fabrication.APP, T_Dossier_Fabrication.VOI,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR], T_Dossier_Fabrication.[REF CONNECTEUR_FOUR], T_Dossier_Fabrication.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR], T_Dossier_Fabrication.PRECO, T_Dossier_Fabrication.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR], T_Dossier_Fabrication.POS2, T_Dossier_Fabrication.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Fabrication.FA2, T_Dossier_Fabrication.APP2, T_Dossier_Fabrication.VOI2, T_Dossier_Fabrication.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR_FOUR2], T_Dossier_Fabrication.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR2], T_Dossier_Fabrication.PRECO2, T_Dossier_Fabrication.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR2], T_Dossier_Fabrication.PRECOG, T_Dossier_Fabrication.Option,  "
    Sql = Sql & "T_Dossier_Fabrication.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & "  and T_Dossier_Fabrication.Onglet ='" & Rs!Onglet & "' "
    Sql = Sql & "ORDER BY T_Dossier_Fabrication.Id;"
  
   
   Set RsOnglet = Con.OpenRecordSet(Sql)
   
    IncremanteBarGrah Me
    If Ionglet > 1 Then
        Me.Fab.Sheets.Add After:=Me.Fab.Sheets.Count
    End If
    Debug.Print Me.Fab.Sheets(Ionglet).Name
     Me.Fab.Sheets(Ionglet).Name = Replace("" & Rs!Onglet, "-", "_")
    Debug.Print Me.Fab.Sheets(Ionglet).Name
      Copy_Rs_Spreadsheet Me, Fab.Sheets(Ionglet), RsOnglet, "Fab_" & Ionglet, Me, Me.ProgressBar1Caption.Caption
      
     
        Rs.MoveNext
    Loop
    
    Sql = "SELECT  T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.ACTIVER, T_Dossier_Contrôle.LIAI,  "
    Sql = Sql & "T_Dossier_Contrôle.DESIGNATION, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
    Sql = Sql & "T_Dossier_Contrôle.TEINT2, T_Dossier_Contrôle.ISO, T_Dossier_Contrôle.LONG, T_Dossier_Contrôle.[LONG CP],  "
    Sql = Sql & "T_Dossier_Contrôle.LONG_ADD, T_Dossier_Contrôle.LONG_ADD2, T_Dossier_Contrôle.COUPE, T_Dossier_Contrôle.POS,  "
    Sql = Sql & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.FA, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR], T_Dossier_Contrôle.PRECO, T_Dossier_Contrôle.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR], T_Dossier_Contrôle.POS2, T_Dossier_Contrôle.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Contrôle.FA2, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2], T_Dossier_Contrôle.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR2], T_Dossier_Contrôle.PRECO2, T_Dossier_Contrôle.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR2], T_Dossier_Contrôle.PRECOG, T_Dossier_Contrôle.Option,  "
    Sql = Sql & "T_Dossier_Contrôle.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Dossier_Contrôle.Id;"
    
    Sql = "SELECT T_Dossier_Contrôle.Onglet "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "GROUP BY T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.Id_IndiceProjet "
    Sql = Sql & "HAVING T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & ";"

    
    
    
Set Rs = Con.OpenRecordSet(Sql)

    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Max = NbEregistrement
    Me.ProgressBar1Caption.Caption = " Exporter liste Dossier de Contrôle :"
  Me.Fab.Sheets(1).Select
  
 IncrmentServer FormBarGrah, ""
    Rs.Requery
    Ionglet = 0
    Do While Rs.EOF = False
    IncremanteBarGrah Me
    Ionglet = Ionglet + 1
        Sql = "SELECT  T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.ACTIVER, T_Dossier_Contrôle.LIAI,  "
    Sql = Sql & "T_Dossier_Contrôle.DESIGNATION, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
    Sql = Sql & "T_Dossier_Contrôle.TEINT2, T_Dossier_Contrôle.ISO, T_Dossier_Contrôle.LONG, T_Dossier_Contrôle.[LONG CP],  "
    Sql = Sql & "T_Dossier_Contrôle.LONG_ADD, T_Dossier_Contrôle.LONG_ADD2, T_Dossier_Contrôle.COUPE, T_Dossier_Contrôle.POS,  "
    Sql = Sql & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.FA, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR], T_Dossier_Contrôle.PRECO, T_Dossier_Contrôle.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR], T_Dossier_Contrôle.POS2, T_Dossier_Contrôle.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Contrôle.FA2, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2], T_Dossier_Contrôle.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR2], T_Dossier_Contrôle.PRECO2, T_Dossier_Contrôle.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR2], T_Dossier_Contrôle.PRECOG, T_Dossier_Contrôle.Option,  "
    Sql = Sql & "T_Dossier_Contrôle.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & " and T_Dossier_Contrôle.Onglet ='" & Rs!Onglet & "' "
    Sql = Sql & "ORDER BY T_Dossier_Contrôle.Id;"
   
   Set RsOnglet = Con.OpenRecordSet(Sql)
    
    If Ionglet > 1 Then
        Me.Cont.Sheets.Add After:=Me.Cont.Sheets.Count
    End If
          Debug.Print Me.Cont.Sheets(Ionglet).Name
     Me.Cont.Sheets(Ionglet).Name = Replace("" & Rs!Onglet, "-", "_")
      Debug.Print Me.Cont.Sheets(Ionglet).Name
      Copy_Rs_Spreadsheet Me, Cont.Sheets(Ionglet), RsOnglet, "Con_" & Ionglet, Me, Me.ProgressBar1Caption.Caption
      

        'ExporterRecordsetExcel RsOnglet, MyWorkbook, Trim("Cont_" & Rs!Onglet), IdIndiceProjet, True, True, "CONT_"
'        Rs.Filter = ""
         If Rs.EOF = True Then Exit Do
        Rs.MoveNext
    Loop
    
End If
Me.Cont.Sheets(1).Select
'Set 'MySeet = IsertSheet(MyWorkbook, "Prix Fils", True)
''MySeet.Application.Visible = True
'MySeet.Delete
'Set 'MySeet = IsertSheet(MyWorkbook, "Appro_Habillage", True)
'MySeet.Delete
'Set 'MySeet = IsertSheet(MyWorkbook, "Appro Connectique", True)
'MySeet.Delete
'Set 'MySeet = Nothing
'***********************************************************************************************************************
'*                                      Exporte RAPPORT DE_CONTRÔLE_FILAIRE.                                            *
'ExporteXlsRAPPORT_DE_CONTROLE_FILAIRE IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte Fiche de Contrôle.                                            *
'ExporteXlsFiche_de_Controle IdIndiceProjet

'***********************************************************************************************************************
'*                                      Supprime le fichier Excel s'il existe                                          *
'If Fso.FileExists(Xls & ".xls") Then Fso.DeleteFile Xls & ".xls"
'Set Fso = Nothing

'***********************************************************************************************************************
'*                                      Enregistre le fichier & referme Excel.                                         *
Err.Clear
'MyWorkbook.Worksheets(1).Select
'
'MyWorkbook.Application.DisplayAlerts = False
'MyWorkbook.SaveAs Xls, ReadOnlyRecommended:=True
'If NotSaveRacourci = False Then
' If IdFils <> 0 Then
'        sql = "SELECT RqCartouche.* "
'        sql = sql & "FROM RqCartouche "
'        sql = sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
'        Set Rs2 = Con.OpenRecordSet(sql)
'         PathPl2 = PathArchive(PathArchiveAutocad, "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
'       Racourci "" & PathPl2, Xls & "", "XLS"
'    End If
'
'End If
If Err Then
MyErr = Err.Description
     FunError 11, "", MyErr
    If IsServeur = False Then
       MsgBox MyErr
       Resume Next
    End If
End If
    
On Error GoTo 0
'MyWorkbook.Close False
'Set MyWorkbook = Nothing
'MyExcel.Quit
'
'Set MyExcel = Nothing
'***********************************************************************************************************************
 Me.ProgressBar1.Value = 0
 Me.ProgressBar1Caption = ""
IncrmentServer FormBarGrah, ""
NewUserForm2 = True
Sql = "UPDATE T_Job SET T_Job.IdExcel = 0 "
Sql = Sql & "WHERE T_Job.Job=" & Command & ";"
If IsServeur = True Then Con.Execute Sql


 

Set Myme = Me
Me.Caption = MeCapTion
If BooolBloque = True Then
    Command3.Visible = False
    Command1.Visible = False
    Command2.Visible = False
End If
IdProjet = IdIndiceProjet

CreatListe
 LstMaj
 IClaseurOnglet = 13
MyClient = Client
Nouveau = NouveauF
NotSortie = True

Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Command3.Enabled = True
Me.Command4.Enabled = True



NoMacro2 = False
NoMacro5 = False
NoMacro3 = False
NoMacro1Change = False
'Me.Combo1.Width= Comp.Cells(Comp.ActiveCell.Row, NumCollonne("comppath")).Width
On Error Resume Next
For I = 1 To Crit.Range("a1").CurrentRegion.Columns.Count
'    ColecBool
    If Crit.Cells(1, I).NumberFormat = "Yes/No" Then
        ColecBool.Add True, Crit.Cells(1, I)
    Else
        ColecBool.Add False, Crit.Cells(1, I)
    End If
Next
For I = 1 To Conn.Range("a1").CurrentRegion.Columns.Count
'    ColecBool
    If Conn.Cells(1, I).NumberFormat = "Yes/No" Then
        ColecBool.Add True, Conn.Cells(1, I)
    Else
        ColecBool.Add False, Conn.Cells(1, I)
    End If
Next
For I = 1 To Fil.Range("a1").CurrentRegion.Columns.Count
'    ColecBool
    If Fil.Cells(1, I).NumberFormat = "Yes/No" Then
        ColecBool.Add True, Fil.Cells(1, I)
    Else
        ColecBool.Add False, Fil.Cells(1, I)
    End If
Next
For I = 1 To Comp.Range("a1").CurrentRegion.Columns.Count
'    ColecBool
    If Comp.Cells(1, I).NumberFormat = "Yes/No" Then
        ColecBool.Add True, Comp.Cells(1, I)
    Else
        ColecBool.Add False, Comp.Cells(1, I)
    End If
Next
For I = 1 To Notas.Range("a1").CurrentRegion.Columns.Count
'    ColecBool
    If Notas.Cells(1, I).NumberFormat = "Yes/No" Then
        ColecBool.Add True, Notas.Cells(1, I)
    Else
        ColecBool.Add False, Notas.Cells(1, I)
    End If
Next


'ColecBool
Me.Visible = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = NotSortie
bool_Activate = False
End Sub


Private Sub Crit_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
'NumCollonne("CritACTIVER")

On Error Resume Next
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Static SaveRow As Long
Dim Sql As String
Dim Rs As Recordset
If (NoMacro5 = True) Or (NoMacro5Select = True) Or (Row = 1) Then GoTo Fin
Row = Me.Crit.ActiveCell.Row
Col = Me.Crit.ActiveCell.Column

If SaveRow = 0 Then SaveRow = Row
If Row = 1 Then GoTo Fin
NoMacro5 = True
ValideCritaire Crit, SaveRow, Row
boolActu = False


SaveRow = Row
NoMacro5 = False
Fin:
End Sub

Private Sub Crit_Click(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Crit_Change EventInfo
End Sub

Private Sub Crit_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
If NoMacro5Select = True Then GoTo Fin
 NoMacro5Select = True
Crit_Change EventInfo
Static Row As Long

Dim aa
Row = Crit.ActiveCell.Row
If Row = 0 Then Row = 1
If Row > 1 And IfValidationOk = True Then
On Error Resume Next

If Crit.Cells(Row, NumCollonne("CritCODE_CRITERE")) <> "" Then
If Crit.Cells(Row, NumCollonne("CritACTIVER")) = True Then
aa = ""
    aa = CollecCrieres(Crit.Cells(Row, NumCollonne("CritCODE_CRITERE")))
    If Err Then
    Err.Clear
        CollecCrieres.Add Crit.Cells(Row, NumCollonne("CritCODE_CRITERE")), Crit.Cells(Row, NumCollonne("CritCODE_CRITERE"))
    Else
        MsgBox " Le code Code Critères : " & Crit.Cells(Row, NumCollonne("CritCODE_CRITERE")) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
            Me.Crit.Cells(Row, NumCollonne("CritCODE_CRITERE")).Select
            Me.Crit.SetFocus
            NoMacro5Select = False
             On Error GoTo 0
        GoTo Fin
    End If
 End If
End If

If Crit.Cells(Row, NumCollonne("CritCRITERES")) <> "" Then
If Crit.Cells(Row, NumCollonne("CritACTIVER")) = True Then
aa = ""
    aa = CollecCrieresCode(Crit.Cells(Row, NumCollonne("CritCRITERES")))
    If Err Then
        Err.Clear
        CollecCrieresCode.Add Crit.Cells(Row, NumCollonne("CritCRITERES")), Crit.Cells(Row, NumCollonne("CritCRITERES"))
    Else
        MsgBox " Le Critères : " & Crit.Cells(Row, NumCollonne("CritCRITERES")) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
             Me.Crit.Cells(Row, NumCollonne("CritCRITERES")).Select
            Me.Crit.SetFocus
            NoMacro5Select = False
             On Error GoTo 0
        GoTo Fin
        GoTo Fin
    End If
    End If
End If
If Crit.Cells(Row, NumCollonne("CritDESIGNATION")) <> "" Then
If Crit.Cells(Row, NumCollonne("CritACTIVER")) = True Then
aa = ""
    aa = CollecCrieresDesigne(Crit.Cells(Row, NumCollonne("CritDESIGNATION")))
    If Err Then
        Err.Clear
        CollecCrieresDesigne.Add Crit.Cells(Row, NumCollonne("CritDESIGNATION")), Crit.Cells(Row, NumCollonne("CritDESIGNATION"))
    Else
        MsgBox " Le code Designation Critères : " & Crit.Cells(Row, NumCollonne("CritDESIGNATION")) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
             Me.Crit.Cells(Row, NumCollonne("CritDESIGNATION")).Select
            Me.Crit.SetFocus
            NoMacro5Select = False
             On Error GoTo 0
        GoTo Fin
             On Error GoTo 0
    End If
   End If
End If
'    Set CollecCrieres = New Collection
'Set CollecCrieresCode = New Collection
'Set CollecCrieresDesigne = New Collection
End If
Row = Row
NoMacro5Select = False
Fin:
End Sub

Private Sub Noeuds_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Row = Noeuds.ActiveCell.Row
If Row = 1 Then GoTo Fin
If NoMacro6 = True Then GoTo Fin
boolActu = False

NoMacro6 = True
Set MyRange = Nothing
 Set MyRange = Me.Noeuds.Range("a1").CurrentRegion
   ConverOuiNon MyRange, Row
If Trim("" & Noeuds.Cells(Row, 1)) <> "" Then
'    Noeuds.Cells(Row, 4) = NoeuName(Row)
'    Else
'        If Trim("" & Noeuds.Cells(Row, 8)) <> "" Then
'           Noeuds.Cells(Row, 4) = NoeuName(Row)
'        Else
'            If Trim("" & Noeuds.Cells(Row, 9)) <> "" Then
'                Noeuds.Cells(Row, 4) = NoeuName(Row)
'            Else
'                If Trim("" & Noeuds.Cells(Row, 10)) <> "" Then
                    Noeuds.Cells(Row, 4) = NoeuName(Row)
DoEvents
'                End If
'        End If
'    End If
End If

NoMacro6 = False
Fin:

End Sub

Function NoeuName2(Row As Long)
Dim Txt As String
Dim Ofset As Long
Dim nbTour As Long
Dim NbTord As Long
Dim txtColone As Long
txtColone = 2
Txt = "AA"
Ofset = 0
nbTour = 0
NbTord = 0
Reprise:

For I = 0 To Row - 2
aa = Mid(Txt, Len(Txt) - Ofset, 1)

    aa = Chr(Asc("A") + (1 * (I - (26 * nbTour))))

Mid(Txt, Len(Txt) - Ofset, 1) = aa


If Asc(Mid(aa, 1, 1)) < 65 Or Asc(Mid(aa, 1, 1)) > 90 Then

Mid(Txt, Len(Txt) - Ofset, 1) = "A"


    Ofset = Ofset + 1
    nbTour = nbTour + 1
    Mid(Txt, Len(Txt) - Ofset, 1) = Chr(Asc(Mid(Txt, Len(Txt) - Ofset, 1)) + 1)
    If Asc(Mid(Txt, 1, 1)) < 65 Or Asc(Mid(Txt, 1, 1)) > 90 Then
 Mid(Txt, 1, 1) = "A"
    Txt = Txt & "A"
    

End If
Ofset = 0
   
End If


Next

NoeuName2 = Txt
End Function

Private Sub Noeuds_SelectionChanging(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim aa As String
Dim MyTxt As String
On Error Resume Next
If boolSelctChange = False Then
 Me.Tag = ""
If EventInfo.Range.Row > 1 Then
Me.Tag = EventInfo.Range.Row
Me.Fleche_Droite.Value = Abs(Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 2))
TORON_P.Value = Abs(Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 3))
    Me.Longueur = CStr(Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 5))
    Long_C = CStr(Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 6))
    Me.NOUED = Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 4)

    If Trim("" & Noeuds.Cells(EventInfo.Range.Row, Num)) <> "" Then
            Me.Hab.ListIndex = MyCollectionHab("N" & Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 7))
    Else
        If Trim("" & Noeuds.Cells(EventInfo.Range.Row, 8)) <> "" Then
           Me.RSA.ListIndex = MyCollectionRSA("N" & Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 8))
        Else
            If Trim("" & Noeuds.Cells(EventInfo.Range.Row, 9)) <> "" Then
                Me.PSA.ListIndex = MyCollectionPSA("N" & Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 9))
            Else
                If Trim("" & Noeuds.Cells(EventInfo.Range.Row, 10)) <> "" Then
                   Me.ENC.ListIndex = MyCollectionENC("N" & Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 10))
                  Else
                    Me.Hab.ListIndex = 0

                End If
        End If
    End If
End If


   Me.Activer.Value = Abs(Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 1))
   Me.DIAMETRE = Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 11)
   Me.CLASSE_T = Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 12)
    Me.txtOption = Noeuds.ActiveSheet.Cells(EventInfo.Range.Row, 13)

End If
End If
DoEvents
'On Error GoTo 0
End Sub
Sub CreatListe()
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT T_Regle_Comp_Hab.ENCELADE, T_Regle_Comp_Hab.PSA, "
Sql = Sql & "T_Regle_Comp_Hab.RSA, T_Regle_Comp_Hab.libellé, T_Regle_Comp_Hab.Numéro "
Sql = Sql & "FROM T_Regle_Comp_Hab "
Sql = Sql & "ORDER BY T_Regle_Comp_Hab.libellé;"
I = 0
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
I = I + 1
Rs.MoveNext
Wend
ReDim MyTableENC(I, 4)
ReDim MyTablePSA(I, 4)
ReDim MyTableRSA(I, 4)
ReDim MyTableHab(I, 4)
ReDim MyTableHab(I, 4)

ReDim MyTableHab(I, 4)
Rs.Requery
I = 0
Me.ENC.Clear
Me.PSA.Clear
Me.RSA.Clear
Me.Hab.Clear

Me.ENC.AddItem ""
Me.PSA.AddItem ""
Me.RSA.AddItem ""
Me.Hab.AddItem ""

Set MyCollectionENC = Nothing
Set MyCollectionPSA = Nothing
Set MyCollectionRSA = Nothing
Set MyCollectionHab = Nothing
Set MyCollectionLienHab = Nothing

Set MyCollectionENC = New Collection
Set MyCollectionPSA = New Collection
Set MyCollectionRSA = New Collection
Set MyCollectionHab = New Collection
Set MyCollectionLienHab = New Collection




  For I = 0 To UBound(MyTablePSA)
       For I2 = 1 To 4
            MyTableENC(I, I2) = "N"
            MyTablePSA(I, I2) = "N"
            MyTableRSA(I, I2) = "N"
            MyTableHab(I, I2) = "N"
          
        Next
  Next
  I = 0
  Rs.Requery
 
   
    
    
While Rs.EOF = False

    I = I + 1
    txtMyCollectionLienHab = ""
    
    
    If Trim("" & Rs!ENCELADE) <> "" Then
        Me.ENC.AddItem Rs!ENCELADE
    End If
        txtMyCollectionLienHab = txtMyCollectionLienHab & "N" & Rs!ENCELADE & ";"
'    Else
'        txtMyCollectionLienHab = txtMyCollectionLienHab & "N;"
''    End If
    If Trim("" & Rs!PSA) <> "" Then
        Me.PSA.AddItem Rs!PSA
    End If
       txtMyCollectionLienHab = txtMyCollectionLienHab & "N" & Rs!PSA & ";"
'     Else
'        txtMyCollectionLienHab = txtMyCollectionLienHab & "N;"
''
'    End If
    If Trim("" & Rs!RSA) <> "" Then
        Me.RSA.AddItem Rs!RSA
        
'     Else
'        txtMyCollectionLienHab = txtMyCollectionLienHab & "N;"
    End If
 txtMyCollectionLienHab = txtMyCollectionLienHab & "N" & Rs!RSA & ";"
    
    
    If Trim("" & Rs!libellé) <> "" Then
        Me.Hab.AddItem Rs!libellé
        
'    Else
'        txtMyCollectionLienHab = txtMyCollectionLienHab & "N;"
    End If
      txtMyCollectionLienHab = txtMyCollectionLienHab & "N" & Rs!libellé & ";"
     MyCollectionLienHab.Add txtMyCollectionLienHab, "N" & Rs!libellé
    Rs.MoveNext
Wend

Set Rs = Con.CloseRecordSet(Rs)
End Sub

 Sub LstMaj()

Set MyCollectionENC = Nothing
Set MyCollectionPSA = Nothing
Set MyCollectionRSA = Nothing
Set MyCollectionHab = Nothing


Set MyCollectionENC = New Collection
Set MyCollectionPSA = New Collection
Set MyCollectionRSA = New Collection
Set MyCollectionHab = New Collection

NoMaj = True
DoEvents
For I = 0 To Me.ENC.ListCount - 1
        MyCollectionENC.Add I, "N" & Trim(Me.ENC.List(I))
        
Next
For I = 0 To Me.PSA.ListCount - 1
       MyCollectionPSA.Add I, "N" & Trim(Me.PSA.List(I))
Next
For I = 0 To Me.RSA.ListCount - 1
        MyCollectionRSA.Add I, "N" & Trim(Me.RSA.List(I))
Next

For I = 0 To Me.Hab.ListCount - 1
      MyCollectionHab.Add I, "N" & Me.Hab.List(I)
Next
For I = 1 To MyCollectionLienHab.Count
    zz = Split(MyCollectionLienHab(I), ";")
    For I2 = 0 To 3
        If zz(I2) <> "N" Then
            MyTableHab(MyCollectionHab(zz(3)), I2 + 1) = Trim(zz(I2))
            
        
        End If
    Next
Next
For I = 1 To UBound(MyTableHab)
    If MyTableHab(I, 1) <> "N" Then
        MyTableENC(MyCollectionENC(MyTableHab(I, 1)), 1) = MyTableHab(I, 1)
        MyTableENC(MyCollectionENC(MyTableHab(I, 1)), 4) = MyTableHab(I, 4)
        
        If MyTableHab(I, 2) <> "N" Then
            MyTableENC(MyCollectionENC(MyTableHab(I, 1)), 2) = MyTableHab(I, 2)
        End If
        
        If MyTableHab(I, 3) <> "N" Then
            MyTableENC(MyCollectionENC(MyTableHab(I, 1)), 3) = MyTableHab(I, 3)
        End If
        
    End If
    
    If MyTableHab(I, 2) <> "N" Then
        MyTablePSA(MyCollectionPSA(MyTableHab(I, 2)), 2) = MyTableHab(I, 2)
        MyTablePSA(MyCollectionPSA(MyTableHab(I, 2)), 4) = MyTableHab(I, 4)
        
        If MyTableHab(I, 1) <> "N" Then
             MyTablePSA(MyCollectionPSA(MyTableHab(I, 2)), 1) = MyTableHab(I, 1)
        End If
        
        If MyTableHab(I, 3) <> "N" Then
            MyTablePSA(MyCollectionPSA(MyTableHab(I, 2)), 3) = MyTableHab(I, 3)
        End If
    End If
     If MyTableHab(I, 3) <> "N" Then
        MyTableRSA(MyCollectionRSA(MyTableHab(I, 3)), 3) = MyTableHab(I, 3)
        MyTableRSA(MyCollectionRSA(MyTableHab(I, 3)), 4) = MyTableHab(I, 4)
        
        If MyTableHab(I, 1) <> "N" Then
             MyTableRSA(MyCollectionRSA(MyTableHab(I, 3)), 1) = MyTableHab(I, 1)
        End If
        
        If MyTableHab(I, 2) <> "N" Then
            MyTableRSA(MyCollectionRSA(MyTableHab(I, 3)), 2) = MyTableHab(I, 2)
        End If
    End If
Next

End Sub

Sub ConverOuiNon(MyRange, Index)
On Error Resume Next
For I = 1 To MyRange.Columns.Count
   If ColecBool(MyRange(1, I)) = True Then
        MyRange(Index, I).NumberFormat = "Yes/No"
        MyRange(Index, I).Value = MyRange(Index, I).Value
        If Not IsNumeric(MyRange(Index, I).Value) Then
            If UCase(Left(MyRange(Index, I).Value, 1)) = "N" Then
                MyRange(Index, I).Value = 0
                DoEvents
               
            Else
                MyRange(Index, I).Value = 1
                DoEvents
               
            End If
        End If
        If Trim("" & MyRange(Index, I)) <> "" Then
            MyRange(Index, I).Value = Abs(MyRange(Index, I).Value) * -1
         End If
   End If
     
Next
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.StopMaco.Visible = True Then
    txtMacro = txtMacro & Space(5) & "SSTab1.Tab =" & SSTab1.Tab & vbCrLf
End If
        Mygrid = CollectionMenu(SSTab1.Caption)
'

End Sub

Private Sub StopMaco_Click()
Dim Sql As String
If MacroName <> "" Then
    Me.txtMacro = Me.txtMacro & "End Sub " & vbCrLf
    Sql = "UPDATE T_Macro SET T_Macro.Sub = '" & Replace(Me.txtMacro, "'", "''") & "'  "
    Sql = Sql & "WHERE T_Macro.Formulaire='" & Me.Name & "'  "
    Sql = Sql & "AND T_Macro.Macro='" & Replace(MacroName, "'", "''") & "' ;"
    Con.Execute Sql
End If
MacroName = ""
 Me.txtMacro = ""
StopMaco.Visible = False
End Sub

Private Sub SupMacro_Click()
frmMacco.charger Me.Name, "SUP"
Unload frmMacco
End Sub

'Sub continuer(Optional import As Boolean)
'Dim pathTmpXls As String
'Dim sql As String
'Dim Rs As Recordset
'Me.Visible = True
'If import = True Then
'        pathTmpXls = UserForm2.Caption
'        UserForm2_boolExcute = UserForm2.boolExcute
'       Unload UserForm2
'        If UserForm2_boolExcute = False Then
'            Me.Enabled = True
'
'              Exit Sub
'        End If
'          MsgAutoCad = "Vos données ont bien été enregistrées, toustefois :" & vbCrLf & vbCrLf
'        ImporteXls pathTmpXls, CLng(Me.txt3.Tag), Edition:=True
'    End If
' If boolAutoCAD = False And IsCilent = False Then
'
'
'
'    MsgBox MsgAutoCad & "Vous ne possédez pas de licence AutoCad." & vbCrLf & "Vous ne pouvez pas reporter vos modifications" & vbCrLf & "sur vos différents plans. "
'Else
' Planche_Clous.Chargement CLng(Me.txt3.Tag)
'
'
'
'    sql = "SELECT T_Path.PathVar FROM T_Path WHERE T_Path.NameVar='PathOutils';"
'    Set Rs = Con.OpenRecordSet(sql)
'    If Rs.EOF = False Then
'        RepPlacheClous = "" & Rs!PathVar
'    End If
'Set Rs = Con.CloseRecordSet(Rs)
'    RepPlacheClous = RepPlacheClous & "\" & Planche_Clous.PlanchClous
'PlanchClous = Planche_Clous.PlanchClous
'Planche_Clous_boolAnnuler = Planche_Clous.boolAnnuler
'
'
'Unload Planche_Clous
'    If Planche_Clous_boolAnnuler = True Then
'        Me.Enabled = True
'        Exit Sub
'    End If
'    If IsCilent = False Then
'
'        subDessinerPlan Me.txt3.Tag
'        subDessinerOutil Me.txt3.Tag
'
'
'        MsgBox "Fin du traitement" & vbCrLf & NbError & " Erreur(s) détectée(s) !"
'    Else
'        MsgBox "Votre demande a été prise en compte vous pouvez suivre l'évolution de votre travail dans la fenêtre (Liste des JOB)"
'    End If
' End If
'NbError = 0
' Noquite = False
'
'Unload Me
'
'End Sub
Function ValideCritaire(sheet As Object, SaveRow As Long, Row As Long) As Boolean
Set MyRange = Nothing
 Set MyRange = sheet.Range("a1").CurrentRegion
   ConverOuiNon MyRange, Row
     sheet.Cells(SaveRow, NumCollonne("CritCODE_CRITERE")) = UCase(Trim("'" & sheet.Cells(SaveRow, NumCollonne("CritCODE_CRITERE"))))
     sheet.Cells(SaveRow, NumCollonne("CritCRITERES")) = UCase(Trim("'" & sheet.Cells(SaveRow, NumCollonne("CritCRITERES"))))

    If (Trim("" & sheet.Cells(SaveRow, NumCollonne("CritCODE_CRITERE"))) <> "") And (Trim("" & sheet.Cells(SaveRow, NumCollonne("CritCRITERES"))) = "") And (SaveRow <> Row) Then
        MsgBox "Le champ CRITERES est obligatoire", vbQuestion, "AutoCâble: Critères"
        msg = "?"
       sheet.Cells(SaveRow, NumCollonne("CritCRITERES")).Select
       Crit.SetFocus
    End If
    If (Trim("" & sheet.Cells(SaveRow, NumCollonne("CritCODE_CRITERE"))) = "") And (Trim("" & sheet.Cells(SaveRow, NumCollonne("CritCRITERES"))) <> "") And (SaveRow <> Row) Then
        MsgBox "Le champ CODE CRITERES est obligatoire", vbQuestion, "AutoCâble: Critères"
        msg = "?"
       sheet.Cells(SaveRow, NumCollonne("CritCODE_CRITERE")).Select
       Crit.SetFocus
    End If
 Set MyRange = Nothing
End Function
Function ValideFils(sheet)
Dim Row As Long
Static SaveRow As Long
Dim Col As Long
Dim MyRange
Dim Rs As Recordset
Dim Sql As String
Dim LibCode_APP As String
'Dim TrouveConnecteur() As Boolean
Dim boolReprise As Boolean
Static Col3 As Long
Row = Me.Fil.ActiveCell.Row
Col = Me.Fil.ActiveCell.Column
If SaveRow = 0 Then SaveRow = 1
If (NoMacro2 = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro2 = True
RepriseCritaire:

   ConverOuiNon sheet.Range("a1").CurrentRegion, Row
   If sheet.Cells(Row, NumCollonne("filsACTIVER")) = True Then 'ACTIVER
If (Col = NumCollonne("filsFa")) Or Col = NumCollonne("filsPOS-OUT") Or Col = NumCollonne("filsPOS") Or Col = NumCollonne("filsREF CONNECTEUR") Or Col = NumCollonne("filsREF CONNECTEUR_FOUR") Then
  Col = NumCollonne("filsApp")
End If
If Col = NumCollonne("filsFa2") Or Col = NumCollonne("filsPOS-OUT2") Or Col = NumCollonne("filsPOS2") Or Col = NumCollonne("filsREF CONNECTEUR2") Or Col = NumCollonne("filsREF CONNECTEUR_FOUR2") Then
  Col = NumCollonne("filsApp2")
End If
If (Trim("" & sheet.Cells(Row, NumCollonne("filsApp"))) <> "") And (Trim("" & sheet.Cells(Row, NumCollonne("filsApp2"))) <> "") Then
Reprise:
        
'Je  cherche si le code APP existe dans la liste des connecteurs.

        I = ChercheXls(UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp")))), Me.Conn.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))

       If I <> 0 Then 'il existe ?
        sheet.Cells(Row, NumCollonne("filsFA")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConN°")).Value
        sheet.Cells(Row, NumCollonne("filsPOS")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConPOS")).Value
         sheet.Cells(Row, NumCollonne("filsPOS-OUT")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConPOS-OUT")).Value
        sheet.Cells(Row, NumCollonne("filsRef Connecteur")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConCONNECTEUR")).Value
        sheet.Cells(Row, NumCollonne("filsRef Connecteur_Four")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConRefConnecteurFour")).Value
                        'OUI
                aa = Split(Me.Conn.Cells(I, NumCollonne("ConOPTION")) & ";", ";")
'                J'affect les Option du connecteur à la laison

                For Iaa = 0 To UBound(aa)
                     If InStr(1, sheet.Cells(Row, NumCollonne("filsOPTION")) & ";", "" & Trim(aa(Iaa)) & ";") = 0 Then
                         sheet.Cells(Row, NumCollonne("filsOPTION")) = "" & sheet.Cells(Row, NumCollonne("filsOPTION")) & ";" & Trim(aa(Iaa))


                     End If
                Next
        Else
            'Non
            msg = "Le connecteur : " & UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp")))) & " introuvable"
                    MsgBox msg, vbExclamation, "AutoCâble: Tableau de fils"
'                    Me.Fil.RangeChrCollonne("filsApp") & NumCollonne("filsApp"))
                     sheet.Cells(Row, NumCollonne("filsApp")).Select
                      sheet.SetFocus
                      GoTo Termine
        End If
        
'Je  cherche si le code APP2 existe dans la liste des connecteurs.
             I = ChercheXls(UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp2")))), Me.Conn.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))

        If I <> 0 Then
                 sheet.Cells(Row, NumCollonne("filsFA2")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConN°")).Value
                 sheet.Cells(Row, NumCollonne("filsPOS2")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConPOS")).Value
                 sheet.Cells(Row, NumCollonne("filsPOS-OUT2")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConPOS-OUT")).Value
                 '
                  sheet.Cells(Row, NumCollonne("filsRef Connecteur2")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConCONNECTEUR")).Value
        sheet.Cells(Row, NumCollonne("filsRef Connecteur_Four2")).Value = "'" & Me.Conn.Cells(I, NumCollonne("ConRefConnecteurFour")).Value
                'OUI
             aa = Split("" & Me.Conn.Cells(I, NumCollonne("ConOPTION")) & ";", ";")
'              J'affect les Option du connecteur à la laison
                For Iaa = 0 To UBound(aa)
                     If InStr(1, sheet.Cells(Row, NumCollonne("filsOPTION")) & ";", "" & Trim(aa(Iaa)) & ";") = 0 Then

                         sheet.Cells(Row, NumCollonne("filsOPTION")) = "" & sheet.Cells(Row, NumCollonne("filsOPTION")) & ";" & Trim(aa(Iaa))
                        If InStr(1, sheet.Cells(Row, NumCollonne("filsOPTION")) & ";", "ALL;") <> 0 Then
                            sheet.Cells(Row, NumCollonne("filsOPTION")) = Replace(sheet.Cells(Row, NumCollonne("filsOPTION")) & ";", "ALL;", "")
                        End If


                     End If
                Next
        Else
                'Non
            msg = "Le connecteur : " & UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp2")))) & " introuvable"
                    MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                     sheet.Cells(Row, Col).Select
                      sheet.SetFocus
                      GoTo Termine
        End If
        If Left(sheet.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then sheet.Cells(Row, NumCollonne("filsOPTION")) = Right(sheet.Cells(Row, NumCollonne("filsOPTION")), Len(sheet.Cells(Row, NumCollonne("filsOPTION"))) - 1)
        If Right(sheet.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then sheet.Cells(Row, NumCollonne("filsOPTION")) = Left(sheet.Cells(Row, NumCollonne("filsOPTION")), Len(sheet.Cells(Row, NumCollonne("filsOPTION"))) - 1)
        If UBound(Split(sheet.Cells(Row, NumCollonne("filsOPTION")) & ";", ";")) > 2 Then

            MsgBox "Vous ne pouvez pas saisir plus de deux options." & vbCrLf & vbCrLf & sheet.Cells(Row, NumCollonne("filsOPTION")), vbQuestion, "AutoCâble: Tableau de fils"
          msg = "?"
          sheet.Cells(Row, NumCollonne("filsOPTION")) = ""
          If boolReprise = False Then
            boolReprise = True
            GoTo Reprise
            End If
        End If
         sheet.Cells(Row, NumCollonne("filsOPTION")) = UCase("" & sheet.Cells(Row, NumCollonne("filsOPTION")))
' Je test l'Options ALL
        If InStr(1, UCase(";" & sheet.Cells(Row, NumCollonne("filsOPTION"))) & ";", ";ALL;") <> 0 Then
            sheet.Cells(Row, NumCollonne("filsOPTION")) = "ALL"
        End If
        
        sheet.Cells(Row, NumCollonne("filsOPTION")) = Replace(sheet.Cells(Row, NumCollonne("filsOPTION")), ";;", "")
        If Right(sheet.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then sheet.Cells(Row, NumCollonne("filsOPTION")) = Left(sheet.Cells(Row, NumCollonne("filsOPTION")), Len(sheet.Cells(Row, NumCollonne("filsOPTION"))) - 1)
          If Left(sheet.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then sheet.Cells(Row, NumCollonne("filsOPTION")) = Right(sheet.Cells(Row, NumCollonne("filsOPTION")), Len(sheet.Cells(Row, NumCollonne("filsOPTION"))) - 1)

    
      If Trim("" & sheet.Cells(Row, NumCollonne("filsOPTION"))) = "" Then
        sheet.Cells(Row, NumCollonne("filsOPTION")) = "ALL" ' MsgBox "Vous devez saisir un code critère."


      End If
     End If
'     If (Trim("" & sheet.Cells(Row, NumCollonne("filsApp"))) <> "") And (Trim("" & sheet.Cells(Row, NumCollonne("filsApp2"))) <> "") And (Trim(UCase("" & sheet.Cells(Row, NumCollonne("filsOPTION")))) = "ALL") Then
'
''        Set Myrange = Me.Conn.ActiveSheet.Range("F1").CurrentRegion
''        Set Myrange = Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count))
'
'
'        I = ChercheXls(UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp")))), Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))
'
'        If I <> 0 Then
'            If Trim(Me.Conn.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
'                sheet.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Conn.Cells(I, NumCollonne("ConOPTION"))
'           End If
'
'        End If
'        If UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsOPTION")))) = "ALL" Then
'             I = ChercheXls(UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp2")))), Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))
'
'        If I <> 0 Then
'            If Trim(Me.Conn.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
'                sheet.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Conn.Cells(I, NumCollonne("ConOPTION"))
'           End If
'
'        End If
'        End If
     

'     End If

If InStr(1, UCase(";" & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) & ";", ";ALL;") <> 0 Then
    sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = "ALL"
End If

If (Trim("" & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) <> "") And (Trim("" & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) <> "ALL") Then
        sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = UCase(sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")))
'    Set Myrange = Me.Crit.ActiveSheet.Range("a1").CurrentRegion
'    Set Myrange = Me.Crit.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count))

    aa = Split(sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) & ";", ";")
    If UBound(aa) = 3 Then
        MsgBox "Vous ne pouvez pas saisir plus de deux critére." & vbCrLf & vbCrLf & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")), vbQuestion, "AutoCâble: Tableau de fils"
         sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = ""
         GoTo RepriseCritaire


    Else
        zz = ""
        For Iaa = 0 To UBound(aa) - 1
            I = ChercheXls(UCase(Trim("" & aa(Iaa))), Me.Crit.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Me.Crit.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))
            If I <> 0 Then
                If Me.Crit.ActiveSheet.Cells(I, 1) = 0 Then GoTo ErrorCritere2
                If InStr(1, ";" & zz & ";", ";" & aa(Iaa) & ";") = 0 Then
                    zz = "" & zz & Trim(aa(Iaa)) & ";"
                End If
            Else
ErrorCritere2:
                MsgBox "Code Critère: " & aa(Iaa) & " introuvable ou Inactif", vbExclamation
            End If
        Next
        sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = zz

        If Right(sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")), 1) = ";" Then sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = Left(sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")), Len(sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) - 1)
'         Sheet.Cells(Row, 28).Value = Sheet.Cells(Row, NumCollonne("filsOPTION")).Value
         If Trim("" & (sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")))) = "" Then GoTo RepriseCritaire
    End If
End If
If (Trim("" & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) <> "") Then
    sheet.Cells(Row, NumCollonne("filsOPTION")).Value = sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")).Value
End If

'
'
'If (Trim("" & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) <> "") Then
'    sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = UCase(sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")))
'
'
'    If (Trim("" & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) <> "ALL") Then
'               aa = Split(sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) & ";", ";")
'        For Iaa = 0 To UBound(aa) - 1
'            I = ChercheXls(UCase(Trim("" & aa(Iaa))), Me.Crit.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Me.Crit.Range("a1").CurrentRegion.Rows.Count)))
'            If UCase(aa(Iaa)) <> "ALL" Then
'                If I = 0 Then
'                    msg = "CODE CRITERE : " & UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsOPTION")))) & " introuvable"
'                    MsgBox msg, vbExclamation, "AutoCâble: Tableau de fils"
'                    sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = ""
'                    GoTo Termine
'                Else
'                    sheet.Cells(Row, NumCollonne("filsOPTION")) = UCase(Replace(sheet.Cells(Row, NumCollonne("filsOPTION")) & ";", aa(Iaa) & ";", ""))
'                    If Right(sheet.Cells(Row, NumCollonne("filsOPTION")).Value, 1) = ";" Then sheet.Cells(Row, NumCollonne("filsOPTION")) = Left(sheet.Cells(Row, NumCollonne("filsOPTION")), Len(sheet.Cells(Row, NumCollonne("filsOPTION"))) - 1)
'
'                End If
'            End If
'        Next
'        If Trim("" & sheet.Cells(Row, NumCollonne("filsOPTION"))) = "" Then sheet.Cells(Row, NumCollonne("filsOPTION")) = "ALL"
'    End If
'     If (Trim("" & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) <> "") Then
'        sheet.Cells(Row, NumCollonne("filsOPTION")) = UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))))
'     End If
'End If

'If InStr(1, UCase(sheet.Cells(Row, NumCollonne("filsOPTION"))) & ";", "ALL;") <> 0 And Len(Trim(sheet.Cells(Row, NumCollonne("filsOPTION"))) & ";") > Len("ALL;") Then sheet.Cells(Row, NumCollonne("filsOPTION")) = Replace(sheet.Cells(Row, NumCollonne("filsOPTION")), "ALL", "")
'
'sheet.Cells(Row, NumCollonne("filsOPTION")) = Replace(sheet.Cells(Row, NumCollonne("filsOPTION")), ";;", "")
'
'If Right(sheet.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then sheet.Cells(Row, NumCollonne("filsOPTION")) = Left(sheet.Cells(Row, NumCollonne("filsOPTION")), Len(sheet.Cells(Row, NumCollonne("filsOPTION"))) - 1)
''Set Myrange = Me.Conn.ActiveSheet.Range("a1").CurrentRegion
''        Set Myrange = Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count))
'
'
'        I = ChercheXls(UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp")))), Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.Range("a1").CurrentRegion.Rows.Count)))
'
'        If I <> 0 Then
'            If Trim(Me.Conn.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
'                sheet.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Conn.Cells(I, NumCollonne("ConOPTION"))
'           End If
'
'        End If
'        If UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsOPTION")))) = "ALL" Then
'             I = ChercheXls(UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp2")))), Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))
'
'        If I <> 0 Then
'            If Trim(Me.Conn.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
'                sheet.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Conn.Cells(I, NumCollonne("ConOPTION"))
'           End If
'
'        End If
'        End If
    
End If

    If Trim("" & sheet.Cells(Row, 2)) <> "" Then
    If (Row > 1) And (Row = 2) Then
        If Trim("" & sheet.Cells(Row, NumCollonne("filsFIL"))) <> Col3 Then
            sheet.Cells(Row, NumCollonne("filsFIL")) = 1
            Col3 = 1
        End If
    Else
        If (Row > 1) And (Row <> 2) Then
            If Trim("" & sheet.Cells(Row, NumCollonne("filsFIL"))) <> Col3 Then
            Col3 = Row - 1
            sheet.Cells(Row, NumCollonne("filsFIL")) = Col3
        End If
    End If

 End If


' If Trim("" & Sheet.Cells(Row, 17)) = "" Then
'        Msg = "le champ  est obligatoire"
'        MsgBox Msg, vbExclamation
'        Sheet.Cells(Row - 1, 17).Select
' Else
'    If Trim("" & Sheet.Cells(Row, 24)) = "" Then
'        Msg = "le champ est obligatoire"
'        MsgBox Msg, vbExclamation
'        Sheet.Cells(Row - 1, 24).Select
'    End If
'End If
'NumCollonne
'If (Col = NumCollonne("filsApp")) Or (Col = NumCollonne("filsApp2")) Then
'
'        If Trim("" & Sheet.Cells(SaveRow, Col)) <> "" Then
'
'            NoMacro2 = True
'
'            I = ChercheXls(UCase(Trim("" & Sheet.Cells(Row, Col))), Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))
'
'            If I <> 0 Then
'            If (Col = NumCollonne("filsApp")) Then
'               Sheet.Cells(Row, NumCollonne("filsFA")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConN°"))))
'                Sheet.Cells(Row, NumCollonne("filsPOS-OUT")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConPOS-OUT"))))
'                Sheet.Cells(Row, NumCollonne("filsPOS")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConPOS"))))
'                Sheet.Cells(Row, NumCollonne("filsREF CONNECTEUR")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConCONNECTEUR"))))
'                 Sheet.Cells(Row, NumCollonne("filsREF CONNECTEUR_FOUR")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConREFCONNECTEURFOUR"))))
'
'            Else
'                Sheet.Cells(Row, NumCollonne("filsFA2")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConN°"))))
'                Sheet.Cells(Row, NumCollonne("filsPOS-OUT2")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConPOS-OUT"))))
'                Sheet.Cells(Row, NumCollonne("filsPOS2")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConPOS"))))
'                 Sheet.Cells(Row, NumCollonne("filsREF CONNECTEUR2")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConCONNECTEUR"))))
'                 Sheet.Cells(Row, NumCollonne("filsREF CONNECTEUR_FOUR2")) = UCase(Trim("'" & Me.Conn.ActiveSheet.Cells(I, NumCollonne("ConREFCONNECTEURFOUR"))))
'            End If
'
'                Else
'                    If (Col = NumCollonne("filsApp")) Then
'                        Sheet.Cells(Row, NumCollonne("filsFA")) = "0"
'                        Sheet.Cells(Row, NumCollonne("filsPOS-OUT")) = ""
'                    Else
'                        Sheet.Cells(Row, NumCollonne("filsFA2")) = "0"
'                        Sheet.Cells(Row, NumCollonne("filsPOS-OUT2")) = ""
'                    End If
'                    msg = "Le connecteur : " & UCase(Trim("" & Sheet.Cells(Row, Col))) & " introuvable"
'                    MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
'                     Sheet.Cells(Row, Col).Select
'                      Sheet.SetFocus
'                End If
'
'
'
'
'
'
'
'                End If
'            End If






            Col3 = 0
        End If
   If (Trim("" & sheet.Cells(Row, NumCollonne("filsApp"))) <> "") Then


'        Set Myrange = Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count))


'        I = ChercheXls(UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp")))), Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))

'        If I <> 0 Then
'            If Trim(Me.Conn.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
'                Myapp1 = "" & Me.Conn.Cells(I, NumCollonne("ConOPTION"))
'           End If
'
'        End If

'             I = ChercheXls(UCase(Trim("" & sheet.Cells(Row, NumCollonne("filsApp2")))), Me.Conn.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))

'        If I <> 0 Then
'            If Trim(Me.Conn.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
'               Myapp2 = "" & Me.Conn.Cells(I, NumCollonne("ConOPTION"))
'           End If
'    If Trim("" & Myapp1) = "" Then Myapp1 = Myapp2
'     If Trim(Myapp2) = "" Then Myapp2 = Myapp1
'        If ("" & Myapp1 <> Myapp2) And (UCase(Myapp1) <> "ALL") And (UCase(Myapp2) <> "ALL") Then
'        MsgBox "Une liaison ne peut pas pointer sur deux options différentes : " & Myapp1 & " & " & Myapp2, vbQuestion, "AutoCâble: Tableau de fils"
'        sheet.Cells(Row, NumCollonne("filsOPTION")) = ""
'        Fil.SetFocus
'        msg = "?"
'        End If
'        End If
If UCase(sheet.Cells(Row, NumCollonne("filsOPTION"))) = "TOUS" Then
    sheet.Cells(Row, NumCollonne("filsOPTION")) = "ALL"
End If
If (Trim("" & sheet.Cells(Row, NumCollonne("filsOPTION"))) <> "") And (Trim("" & sheet.Cells(Row, NumCollonne("filsOPTION"))) <> "ALL") Then
        sheet.Cells(Row, NumCollonne("filsOPTION")) = UCase(sheet.Cells(Row, NumCollonne("filsOPTION")))
'    Set Myrange = Me.Crit.ActiveSheet.Range("a1").CurrentRegion
'    Set Myrange = Me.Crit.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Me.Conn.ActiveSheet.Range("a1").CurrentRegion.Rows.Count))

    aa = Split(sheet.Cells(Row, NumCollonne("filsOPTION")) & ";", ";")
    If UBound(aa) = 3 Then
        MsgBox "Vous ne pouvez pas saisir plus de deux critére." & vbCrLf & vbCrLf & sheet.Cells(Row, 28), vbQuestion, "AutoCâble: Tableau de fils"
        sheet.Cells(Row, NumCollonne("filsOPTION")) = "ALL"
         sheet.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = ""
         GoTo RepriseCritaire


    Else
        zz = ""
        For Iaa = 0 To UBound(aa) - 1
            I = ChercheXls(UCase(Trim("" & aa(Iaa))), Me.Crit.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Me.Crit.ActiveSheet.Range("a1").CurrentRegion.Rows.Count)))
            If I <> 0 Then
                If Me.Crit.ActiveSheet.Cells(I, 1) = 0 Then GoTo ErrorCritere
                If InStr(1, ";" & zz & ";", ";" & aa(Iaa) & ";") = 0 Then
                    zz = "" & zz & Trim(aa(Iaa)) & ";"
                End If
            Else
ErrorCritere:
                MsgBox "Code Critère: " & aa(Iaa) & " introuvable ou Inactif", vbExclamation
            End If
        Next
        sheet.Cells(Row, NumCollonne("filsOPTION")) = zz

        If Right(sheet.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then sheet.Cells(Row, NumCollonne("filsOPTION")) = Left(sheet.Cells(Row, NumCollonne("filsOPTION")), Len(sheet.Cells(Row, NumCollonne("filsOPTION"))) - 1)
'         Sheet.Cells(Row, 28).Value = Sheet.Cells(Row, NumCollonne("filsOPTION")).Value
         If Trim("" & (sheet.Cells(Row, NumCollonne("filsOPTION")))) = "" Then GoTo RepriseCritaire
    End If
End If
End If
Termine:
SaveRow = Row
    NoMacro2 = False
Fin:
End Function
