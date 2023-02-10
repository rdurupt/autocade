VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{0002E550-0000-0000-C000-000000000046}#1.1#0"; "OWC10.DLL"
Begin VB.Form UserForm2 
   ClientHeight    =   11685
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   18240
   Icon            =   "AutoCable.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11685
   ScaleWidth      =   18240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMacro 
      Height          =   285
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   45
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
      TabIndex        =   44
      ToolTipText     =   "Reprendre le traitement"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   16200
      TabIndex        =   4
      Top             =   -70
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   525
         Left            =   45
         Picture         =   "AutoCable.frx":0BD4
         Stretch         =   -1  'True
         Top             =   150
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Annuler"
      Height          =   435
      Left            =   12480
      TabIndex        =   3
      Top             =   11160
      Width           =   2130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Valider"
      Height          =   435
      Left            =   8360
      TabIndex        =   2
      Top             =   11160
      Width           =   2130
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualiser / Valider"
      Height          =   435
      Left            =   4240
      TabIndex        =   1
      Top             =   11160
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualiser"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   11160
      Width           =   2130
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10245
      Left            =   -120
      TabIndex        =   5
      Top             =   480
      Width           =   17535
      _ExtentX        =   30930
      _ExtentY        =   18071
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      Tab             =   9
      TabsPerRow      =   11
      TabHeight       =   520
      TabCaption(0)   =   "Critères"
      TabPicture(0)   =   "AutoCable.frx":10C96
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Spreadsheet5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Connecteurs"
      TabPicture(1)   =   "AutoCable.frx":10CB2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "Spreadsheet1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tableau de fils"
      TabPicture(2)   =   "AutoCable.frx":10CCE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).Control(1)=   "Spreadsheet2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Composants"
      TabPicture(3)   =   "AutoCable.frx":10CEA
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Spreadsheet3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Notas"
      TabPicture(4)   =   "AutoCable.frx":10D06
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Spreadsheet4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Noeuds"
      TabPicture(5)   =   "AutoCable.frx":10D22
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label10"
      Tab(5).Control(1)=   "Label9"
      Tab(5).Control(2)=   "Label8"
      Tab(5).Control(3)=   "Label7"
      Tab(5).Control(4)=   "NOUED"
      Tab(5).Control(5)=   "Label6"
      Tab(5).Control(6)=   "Label5"
      Tab(5).Control(7)=   "Label4"
      Tab(5).Control(8)=   "Label3"
      Tab(5).Control(9)=   "Label2"
      Tab(5).Control(10)=   "Label1"
      Tab(5).Control(11)=   "Spreadsheet6"
      Tab(5).Control(12)=   "Command8"
      Tab(5).Control(13)=   "txtOption"
      Tab(5).Control(14)=   "Fleche_Droite"
      Tab(5).Control(15)=   "Long_C"
      Tab(5).Control(16)=   "TORON_P"
      Tab(5).Control(17)=   "CLASSE_T"
      Tab(5).Control(18)=   "DIAMETRE"
      Tab(5).Control(19)=   "ACTIVER"
      Tab(5).Control(20)=   "Command7"
      Tab(5).Control(21)=   "Command6"
      Tab(5).Control(22)=   "Command5"
      Tab(5).Control(23)=   "ENC"
      Tab(5).Control(24)=   "PSA"
      Tab(5).Control(25)=   "RSA"
      Tab(5).Control(26)=   "Hab"
      Tab(5).Control(27)=   "Longueur"
      Tab(5).ControlCount=   28
      TabCaption(6)   =   "Nomenclature Connecteur"
      TabPicture(6)   =   "AutoCable.frx":10D3E
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Spreadsheet7"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Nomenclature Fils"
      TabPicture(7)   =   "AutoCable.frx":10D5A
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Spreadsheet8"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Nomenclature Habillage"
      TabPicture(8)   =   "AutoCable.frx":10D76
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Spreadsheet9"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Nomenclatures"
      TabPicture(9)   =   "AutoCable.frx":10D92
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "Spreadsheet10"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Tab 10"
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      Begin OWC10.Spreadsheet Spreadsheet10 
         Height          =   9720
         Left            =   120
         OleObjectBlob   =   "AutoCable.frx":10DAE
         TabIndex        =   46
         Top             =   480
         Width           =   17220
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -69810
         Picture         =   "AutoCable.frx":11B1A
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   6
         Top             =   845
         Width           =   225
      End
      Begin OWC.Spreadsheet Spreadsheet5 
         Height          =   9720
         Left            =   -74760
         TabIndex        =   19
         Top             =   480
         Width           =   16650
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":11BA0
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
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
      Begin OWC.Spreadsheet Spreadsheet1 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   43
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":123F6
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
         Left            =   -66840
         TabIndex        =   27
         Top             =   780
         Width           =   1215
      End
      Begin VB.ComboBox Hab 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         ItemData        =   "AutoCable.frx":12B31
         Left            =   -70560
         List            =   "AutoCable.frx":12B33
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   1260
         Width           =   2535
      End
      Begin VB.ComboBox RSA 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -66840
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   1260
         Width           =   1215
      End
      Begin VB.ComboBox PSA 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -64320
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   1260
         Width           =   1455
      End
      Begin VB.ComboBox ENC 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         ItemData        =   "AutoCable.frx":12B35
         Left            =   -61800
         List            =   "AutoCable.frx":12B37
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   1260
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   -74880
         Picture         =   "AutoCable.frx":12B39
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ajouter"
         Top             =   780
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   -73800
         Picture         =   "AutoCable.frx":133AF
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Supprimer"
         Top             =   780
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   -74340
         Picture         =   "AutoCable.frx":13BA9
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Modifier"
         Top             =   780
         Width           =   495
      End
      Begin VB.CheckBox ACTIVER 
         Alignment       =   1  'Right Justify
         Caption         =   "ACTIVER"
         Height          =   315
         Left            =   -72960
         TabIndex        =   17
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox DIAMETRE 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -61800
         TabIndex        =   16
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox CLASSE_T 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -59400
         TabIndex        =   15
         Top             =   780
         Width           =   1815
      End
      Begin VB.CheckBox TORON_P 
         Alignment       =   1  'Right Justify
         Caption         =   "TORON/P"
         Height          =   315
         Left            =   -72960
         TabIndex        =   13
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox Long_C 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -64320
         TabIndex        =   12
         Top             =   780
         Width           =   1455
      End
      Begin VB.CheckBox Fleche_Droite 
         Alignment       =   1  'Right Justify
         Caption         =   "Fleche D"
         Height          =   315
         Left            =   -74400
         TabIndex        =   11
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox txtOption 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -59400
         TabIndex        =   10
         Top             =   1260
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   -57600
         Picture         =   "AutoCable.frx":14357
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1260
         Width           =   315
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FontTransparent =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -69810
         Picture         =   "AutoCable.frx":15199
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   7
         Top             =   845
         Width           =   225
      End
      Begin OWC.Spreadsheet Spreadsheet2 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   17175
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":1521F
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
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
      Begin OWC.Spreadsheet Spreadsheet6 
         Height          =   8490
         Left            =   -74760
         TabIndex        =   14
         Top             =   1620
         Width           =   16995
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":159AD
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
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
      Begin OWC.Spreadsheet Spreadsheet3 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   17175
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":16407
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
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
      Begin OWC.Spreadsheet Spreadsheet4 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   17175
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":16B9B
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
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
      Begin OWC.Spreadsheet Spreadsheet7 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   16650
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":1732A
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
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
      Begin OWC.Spreadsheet Spreadsheet8 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   16650
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":17D84
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
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
      Begin OWC.Spreadsheet Spreadsheet9 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   17175
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":187DE
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
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
         Left            =   -71760
         TabIndex        =   42
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "LONGUEUR"
         Height          =   315
         Left            =   -67920
         TabIndex        =   41
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DESIGN.HAB."
         Height          =   315
         Left            =   -71760
         TabIndex        =   40
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CODE.RSA."
         Height          =   315
         Left            =   -67920
         TabIndex        =   39
         Top             =   1260
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CODE.PSA."
         Height          =   315
         Left            =   -65520
         TabIndex        =   38
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CODE.ENC."
         Height          =   255
         Left            =   -62760
         TabIndex        =   37
         Top             =   1260
         Width           =   870
      End
      Begin VB.Label NOUED 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -70560
         TabIndex        =   36
         Top             =   780
         Width           =   2535
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DIAMETRE"
         Height          =   315
         Left            =   -62760
         TabIndex        =   35
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CLASSE_T"
         Height          =   315
         Left            =   -60240
         TabIndex        =   34
         Top             =   780
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "LONGUEUR/C"
         Height          =   315
         Left            =   -65520
         TabIndex        =   33
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "OPTION"
         Height          =   315
         Left            =   -60240
         TabIndex        =   32
         Top             =   1260
         Width           =   615
      End
   End
   Begin VB.Menu Outils 
      Caption         =   "Outils"
      Begin VB.Menu Masquer_Colonnes 
         Caption         =   "Masquer Colonnes"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu Maro 
      Caption         =   "Maro"
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
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Dim OngletName As String
Dim CollecApp As Collection
Dim Mygrid As String
Dim CollectionMenu As Collection

Sub Charger_Colection(grid, Lib As String)
Dim Myrange
Dim C As Long
Dim I As Long
Dim Txt As String
Dim Adress
Set Myrange = grid.Range("a1").CurrentRegion
For C = 1 To Myrange.Columns.Count
    NumCollonne.Add C, Lib & Trim("" & Myrange(1, C).Value)
    Adress = Myrange(1, C).Address
    Txt = ""
    For I = 1 To Len(Adress)
        If Not IsNumeric(Mid(Adress, I, 1)) Then
            Txt = Txt & Mid(Adress, I, 1)
        End If
    Next
    ChrCollonne.Add Txt, Lib & Trim("" & Myrange(1, C).Value)
    
Next
End Sub









Private Sub Command3_Click()
If boolActu = False Then
    MsgBox "Il est impossible de valide l'étude si un test de d'actualisation na pas été effectué."
    Exit Sub
End If
If Trim(msg) <> "" Then
    MsgBox "Il est impossible de valide l'étude si le test de validation présente des erreurs."
    Exit Sub
End If

Dim MyExcel As EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim MyRange2
Set MyExcel = New EXCEL.Application
'MyExcel.Visible = True
MyExcel.DisplayAlerts = False
Dim MyTim
Dim BoolErr As Boolean
BoolErr = False
Dim Fso As New FileSystemObject
   
    
'MyExcel.Visible = True
'If Nouveau = False Then
'    Set MyWorkbook = MyExcel.Workbooks.Open(Me.Caption)
'Else
    If Fso.FileExists(Me.Caption) Then Fso.DeleteFile (Me.Caption)
    DoEvents
    Set MyWorkbook = MyExcel.Workbooks.Add
    On Error Resume Next
    MyWorkbook.SaveAs Replace(Me.Caption, "Rév.:", "")
    If Err Then
        BoolErr = True
        MsgBox Err.Description
        Err.Clear
        On Error GoTo 0
        GoTo Fin
    End If
    MyWorkbook.Close
    MyTim = Now
    While DateDiff("s", MyTim, Now) < 1
        DoEvents
    Wend
    MyExcel.DisplayAlerts = False
    Set MyWorkbook = MyExcel.Workbooks.Open(Me.Caption)
'    MyExcel.Visible = True
'End If
'MyExcel.Visible = True
'MyWorkbook.Application.Visible = True
IsertSheet MyWorkbook, "Nomenclature Habillage"
   
 Set Myrange = MyWorkbook.Worksheets("Nomenclature Habillage").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Nomenclature Habillage").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Nomenclature Habillage").Select
MyWorkbook.Worksheets("Nomenclature Habillage").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet9
Set MyRange2 = Me.Spreadsheet9.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Nomenclature Habillage").Range("a1").Select
MyWorkbook.Worksheets("Nomenclature Habillage").Paste
ReplaceNull MyWorkbook.Worksheets("Nomenclature Habillage"), "©", Chr(10)

'MyWorkbook.Application.Visible = True
IsertSheet MyWorkbook, "Nomenclature Fils"
MyWorkbook.Worksheets("Nomenclature Fils").Select
Set Myrange = MyWorkbook.Worksheets("Nomenclature Fils").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Nomenclature Fils").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Nomenclature Fils").Select

MyWorkbook.Worksheets("Nomenclature Fils").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet8
Set MyRange2 = Me.Spreadsheet8.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Nomenclature Fils").Range("a1").Select
MyWorkbook.Worksheets("Nomenclature Fils").Paste
ReplaceNull MyWorkbook.Worksheets("Nomenclature Fils"), "©", Chr(10)

   

IsertSheet MyWorkbook, "Nomenclature Connecteur"

Set Myrange = MyWorkbook.Worksheets("Nomenclature Connecteur").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Nomenclature Connecteur").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Nomenclature Connecteur").Select
MyWorkbook.Worksheets("Nomenclature Connecteur").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet7
Set MyRange2 = Me.Spreadsheet7.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Nomenclature Connecteur").Range("a1").Select
MyWorkbook.Worksheets("Nomenclature Connecteur").Paste
ReplaceNull MyWorkbook.Worksheets("Nomenclature Connecteur"), "©", Chr(10)

IsertSheet MyWorkbook, "NOEUDS"
Set Myrange = MyWorkbook.Worksheets("NOEUDS").Range("a1").CurrentRegion
MyWorkbook.Worksheets("NOEUDS").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("NOEUDS").Select
MyWorkbook.Worksheets("NOEUDS").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet6
Set MyRange2 = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("NOEUDS").Paste
ReplaceNull MyWorkbook.Worksheets("NOEUDS"), "©", Chr(10)



IsertSheet MyWorkbook, "Critères"
Set Myrange = MyWorkbook.Worksheets("Critères").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Critères").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Critères").Select
MyWorkbook.Worksheets("Critères").Range("a1").Select
'MyWorkbook.Application.Visible = True
'Me.Spreadsheet5.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet5
Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range(MyRange2(1, 1).Address & ":" & MyRange2(MyRange2.Rows.Count, 4).Address)
MyRange2.Copy
MyWorkbook.Worksheets("Critères").Paste
ReplaceNull MyWorkbook.Worksheets("Critères"), "©", Chr(10)
'MyWorkbook.Application.Visible = True
IsertSheet MyWorkbook, "Notas"
Set Myrange = MyWorkbook.Worksheets("Notas").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Notas").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Notas").Select
MyWorkbook.Worksheets("Notas").Range("a1").Select
'Me.Spreadsheet4.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet4
Set MyRange2 = Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Notas").Paste
ReplaceNull MyWorkbook.Worksheets("Notas"), "©", Chr(10)

MyWorkbook.Worksheets("Notas").Range("a1").Select

IsertSheet MyWorkbook, "Composants"
Set Myrange = MyWorkbook.Worksheets("Composants").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Composants").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Composants").Select
MyWorkbook.Worksheets("Composants").Range("a1").Select
'Me.Spreadsheet3.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet3
Set MyRange2 = Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Composants").Paste
ReplaceNull MyWorkbook.Worksheets("Composants"), "©", Chr(10)
MyWorkbook.Worksheets("Composants").Range("a1").Select

IsertSheet MyWorkbook, "Ligne_Tableau_fils"
Set Myrange = MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range(Myrange(12, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Ligne_Tableau_fils").Select
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").Select
'Me.Spreadsheet2.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet2
Set MyRange2 = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Ligne_Tableau_fils").Paste
ReplaceNull MyWorkbook.Worksheets("Ligne_Tableau_fils"), "©", Chr(10)




IsertSheet MyWorkbook, "Connecteurs"
MyWorkbook.Worksheets("Connecteurs").Range("a1").Select
Set Myrange = MyWorkbook.Worksheets("Connecteurs").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Connecteurs").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Connecteurs").Select
MyWorkbook.Worksheets("Connecteurs").Range("a1").Select
'Me.Spreadsheet1.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet1
Set MyRange2 = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy

'IsertSheet MyWorkbook, "Connecteurs"

MyWorkbook.Worksheets("Connecteurs").Paste
ReplaceNull MyWorkbook.Worksheets("Connecteurs"), "©", Chr(10)

MyWorkbook.Worksheets("Connecteurs").Range("a1").Select













 MyWorkbook.Save
Fin:
 
 
 Set MyRange2 = Nothing
 Set Myrange = Nothing
 MyWorkbook.Close False
 Set MyWorkbook = Nothing
 MyExcel.Quit
 Set MyExcel = Nothing
 If BoolErr = False Then boolExcute = True
 NotSortie = False
 boolActu = False
Me.Hide
End Sub



Private Sub Command4_Click()

MenuShow = True
 boolExcute = False
 NotSortie = False
 boolActu = False
Me.Hide
End Sub

Private Sub Command1_Click()
DoEvents
Dim Myrange
Dim sql As String
Set CollecApp = Nothing
Set CollecApp = New Collection
Set CollecCrieres = Nothing
Set CollecCrieresCode = Nothing
Set CollecCrieresDesigne = Nothing

Set CollecCrieres = New Collection
Set CollecCrieresCode = New Collection
Set CollecCrieresDesigne = New Collection

Me.Spreadsheet5.Cells(1, 1).Select
sql = "DELETE Ajout_LIAISON_CONNECTEURS.* FROM Ajout_LIAISON_CONNECTEURS;"
Con.Exequte sql
sql = "DELETE Ajout_LIAISON.* FROM Ajout_LIAISON;"
Con.Exequte sql

msg = ""
DoEvents
IfValidationOk = True
RazFiltreEditExcel Me.Spreadsheet5
Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
SSTab1.Tab = 0
Me.Spreadsheet5.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
IfValidationOk = True
Me.Spreadsheet5.Cells(I, 1).Select
ConverOuiNon Myrange, I
If msg <> "" Then
    IfValidationOk = False
'    Me.Spreadsheet5.ActiveSheet.AutoFilter = True
    Exit Sub
End If
    Me.Spreadsheet5.Cells(I, 1).Value = Me.Spreadsheet5.Cells(I, 1).Value
If msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next I
RazFiltreEditExcel Me.Spreadsheet1
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
OngletName = "Connecteur"
SSTab1.Tab = 1
DoEvents
Me.Spreadsheet1.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
IfValidationOk = True
Me.Spreadsheet1.Cells(I, 1).Select
ConverOuiNon Myrange, I
If msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
    Me.Spreadsheet1.Cells(I, 1).Value = Me.Spreadsheet1.Cells(I, 1).Value
If msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next I
RazFiltreEditExcel Me.Spreadsheet2
Set Myrange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion

SSTab1.Tab = 2
DoEvents
Me.Spreadsheet2.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
Me.Spreadsheet2.Cells(I, NumCollonne("filsApp")).Select
ConverOuiNon Myrange, I
IfValidationOk = True
    Me.Spreadsheet2.Cells(I, NumCollonne("filsApp")).Value = UCase("'" & Me.Spreadsheet2.Cells(I, NumCollonne("filsApp")).Value)
If msg <> "" Then
'Me.Spreadsheet2.ActiveSheet.AutoFilter = True
    IfValidationOk = False
    Exit Sub
End If
Me.Spreadsheet2.Cells(I, NumCollonne("filsApp2")).Select
    Me.Spreadsheet2.Cells(I, NumCollonne("filsApp2")).Value = Me.Spreadsheet2.Cells(I, NumCollonne("filsApp2")).Value
    If msg <> "" Then
'    Me.Spreadsheet2.ActiveSheet.AutoFilterMode = True

        IfValidationOk = False
        Exit Sub
    End If
DoEvents
Next I
'Me.Spreadsheet5.ActiveSheet.AutoFilter = False
SSTab1.Tab = 3
RazFiltreEditExcel Me.Spreadsheet3
Set Myrange = Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion

Me.Spreadsheet3.Cells(1, 2).Select
For I = 2 To Myrange.Rows.Count
Me.Spreadsheet3.Cells(I, 2).Select
ConverOuiNon Myrange, I
Me.Spreadsheet3.Cells(I, 3) = 0
IfValidationOk = True
 If msg <> "" Then Exit Sub
DoEvents
Next I

SSTab1.Tab = 4
RazFiltreEditExcel Me.Spreadsheet4
Set Myrange = Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion
SSTab1.Tab = 4
Me.Spreadsheet4.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
Me.Spreadsheet4.Cells(I, 3).Select
ConverOuiNon Myrange, I
Me.Spreadsheet4.Cells(I, 3) = I - 1
IfValidationOk = True
 
DoEvents
Next I
RazFiltreEditExcel Me.Spreadsheet6
Set Myrange = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion
SSTab1.Tab = 5
Me.Spreadsheet6.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
IfValidationOk = True
ConverOuiNon Myrange, I
Me.Spreadsheet6.Cells(I, 1).Select

DoEvents
If msg <> "" Then
'Me.Spreadsheet6.ActiveSheet.AutoFilter = True
    IfValidationOk = False
    Exit Sub
End If
'If Spreadsheet6.Cells(i, 4) = "x" Then
'    MsgBox ""
'End If
    Command7_Click
If msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next I


'Me.Spreadsheet6.ActiveSheet.AutoFilter = True
If MyErr = True Then
    LoadLiasons.charger MyClient
    Unload LoadLiasons
End If
MyErr = False
    IfValidationOk = False
    boolActu = True

End Sub
Private Sub Command2_Click()
Command1_Click
If msg <> "" Then Exit Sub
Command3_Click
End Sub

Private Sub Command5_Click()
Set Myrange = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion

'If Me.Tag = "" Then
    If Trim("" & Me.Hab) <> "" Then
        boolSelctChange = True
        Myrange(Myrange.Rows.Count + 1, 1).Select
        Myrange(Myrange.Rows.Count + 1, 2) = Me.Fleche_Droite.Value
        Myrange(Myrange.Rows.Count + 1, 3) = Me.TORON_P.Value
        Myrange(Myrange.Rows.Count + 1, 1) = Me.Activer.Value
        Myrange(Myrange.Rows.Count + 1, 5) = Val(Replace("" & Me.Longueur, ",", "."))
        Myrange(Myrange.Rows.Count + 1, 6) = Val(Replace("" & Me.Long_C, ",", "."))
        Myrange(Myrange.Rows.Count + 1, 7) = "'" & Me.Hab
        Myrange(Myrange.Rows.Count + 1, 8) = "'" & Me.RSA
        Myrange(Myrange.Rows.Count + 1, 9) = "'" & Me.PSA
        Myrange(Myrange.Rows.Count + 1, 10) = "'" & Me.ENC
         
        Myrange(Myrange.Rows.Count + 1, 11) = Val(Replace("" & Me.DIAMETRE, ",", "."))
        Myrange(Myrange.Rows.Count + 1, 12) = "'" & Me.CLASSE_T
        Myrange(Myrange.Rows.Count + 1, 13) = "'" & Me.txtOption
        boolSelctChange = False
    End If
'Else
'     If Trim("" & Me.Hab) <> "" Then
'     boolSelctChange = True
'         Me.Spreadsheet6.ActiveSheet.Cells(Me.Tag, 1).InsertRows
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
    Me.Spreadsheet6.ActiveSheet.Rows(Val(Me.Tag)).DeleteRows
    Me.Tag = ""
    Me.Hab.ListIndex = 0
    Longueur = ""
    Me.Tag = ""
    Me.NOUED = ""
    boolSelctChange = False
End If
End Sub

Private Sub Command7_Click()
Set Myrange = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion

If Me.Tag = "" Then
    Command5_Click
Else
'    If Trim("" & Me.Hab) <> "" Then
        boolSelctChange = True
        Myrange(Val(Me.Tag), 1).Select
        
         Myrange(Val(Me.Tag), 2) = Fleche_Droite.Value
          Myrange(Val(Me.Tag), 3) = TORON_P.Value
         Myrange(Val(Me.Tag), 1) = Me.Activer.Value
        Myrange(Val(Me.Tag), 5) = Val(Replace("" & Me.Longueur, ",", "."))
         Myrange(Val(Me.Tag), 6) = Val(Replace("" & Me.Long_C, ",", "."))
        Myrange(Val(Me.Tag), 7) = "'" & Me.Hab
        Myrange(Val(Me.Tag), 8) = "'" & Me.RSA
        Myrange(Val(Me.Tag), 9) = "'" & Me.PSA
        Myrange(Val(Me.Tag), 10) = "'" & Me.ENC
        Myrange(Val(Me.Tag), 11) = Val(Replace("" & Me.DIAMETRE, ",", "."))
        Myrange(Val(Me.Tag), 12) = "'" & Me.CLASSE_T
        If InStr("" & Me.txtOption, "TOUS") <> 0 Then
            If Len("" & Me.txtOption) > Len("TOUS;") Then
                Myrange(Val(Me.Tag), 13) = "'" & Replace(Me.txtOption, "TOUS", "")
                Myrange(Val(Me.Tag), 13) = Replace(Myrange(Val(Me.Tag), 13), ";;", ";")
                If Left("" & Myrange(Val(Me.Tag), 13), 1) = ";" Then Myrange(Val(Me.Tag), 13) = Right(Myrange(Val(Me.Tag), 13), Len(Myrange(Val(Me.Tag), 13)) - 1)
                
            Else
                Myrange(Val(Me.Tag), 13) = "'" & Me.txtOption
            End If
        Else
         Myrange(Val(Me.Tag), 13) = "'" & Me.txtOption
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
Me.txtOption = FrmSelectCriteres.Chargement(Spreadsheet5, Me.txtOption)
Unload FrmSelectCriteres
End Sub

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
   Dim sc, m
   Set sc = CreateObject("ScriptControl")
   sc.Language = "VBScript"
   sc.AddObject "Me", Me, True
   ' Ajoute un module.
   Set m = sc.Modules.Add("Module1")
   ' Ajoute du code au module.
   m.AddCode Me.txtMacro
   ' Exécute le script.
   m.Run MacroName
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
    txtMacro = txtMacro & Space(5) & Mygrid & ".sheets(" & Chr(34) & Controls(Mygrid).ActiveSheet.Name & Chr(34) & ").Select" & vbCrLf
     txtMacro = txtMacro & Space(5) & Mygrid & ".sheets(" & Chr(34) & Controls(Mygrid).ActiveSheet.Name & Chr(34) & ").Range(" & Chr(34) & Controls(Mygrid).Selection.Address & Chr(34) & ").select" & vbCrLf
      txtMacro = txtMacro & Space(5) & Mygrid & ".Selection.ColumnWidth = 0" & vbCrLf
    Else
    txtMacro = txtMacro & Space(5) & Mygrid & ".Range(" & Chr(34) & Controls(Mygrid).Selection.Address & Chr(34) & ").ColumnWidth = 0" & vbCrLf
    End If
End If

    Controls(Mygrid).Selection.ColumnWidth = 0
End Sub

Private Sub NewMacro_Click()
Dim sql As String
Dim Rs As Recordset
Dim Trouve As Boolean
MacroName = ""
Reprise:
    MacroName = InputBox("Entrez le nom de la macro", "Nouvelle Macro", MacroName)
    If Trim("" & MacroName) <> "" Then
    MacroName = Replace(MacroName, " ", "_")
    sql = "SELECT T_Macro.Formulaire, T_Macro.Macro "
    sql = sql & "FROM T_Macro "
    sql = sql & "WHERE T_Macro.Formulaire='" & Me.Name & "'  "
    sql = sql & "AND T_Macro.Macro='" & Replace(MacroName, "'", "''") & "';"
    Set Rs = Con.OpenRecordSet(sql)
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
    sql = "INSERT INTO T_Macro ( Formulaire, Macro ) "
    sql = sql & "VALUES ( '" & Me.Name & "' , '" & Replace(Replace(MacroName, "'", "''"), Chr(34), Chr(34) & Chr(34)) & "' );"
Con.Exequte sql
    End If

        Me.txtMacro.Text = "Sub " & MacroName & "()" & vbCrLf
        StopMaco.Visible = True
    End If


End Sub

Private Sub Picture1_Click()
frmEditClip.charger Me.Spreadsheet2, NumCollonne
End Sub

Private Sub Picture2_Click()
frmEditBouchon.charger Me.Spreadsheet1, NumCollonne
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

Private Sub Spreadsheet1_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveRow As Long

Dim LibCode_APP As String
Dim sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet1.ActiveCell.Row
Col = Me.Spreadsheet1.ActiveCell.Column

If (NoMacro1Change = True Or NoMacro1Select = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro1Change = True
Set Myrange = Me.Spreadsheet1.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
   
   Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION")) = UCase("" & Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION")))
   Me.Spreadsheet1.Cells(Row, NumCollonne("conCONNECTEUR")) = UCase("'" & Me.Spreadsheet1.Cells(Row, NumCollonne("conCONNECTEUR")))
    Me.Spreadsheet1.Cells(Row, NumCollonne("conCODE_APP")) = UCase("'" & Me.Spreadsheet1.Cells(Row, NumCollonne("conCODE_APP")))
If Trim("" & Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION"))) <> "" Then
    If UCase(Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION"))) <> "TOUS" Then
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Myrange.Rows.Count))
        Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
         Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range(ChrCollonne("CritACTIVER") & "1", ChrCollonne("CritACTIVER") & CStr(MyRange2.Rows.Count))
        aa = Split(UCase(Trim("" & Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION")))) & ";", ";")
        For Iaa = 0 To UBound(aa) - 1
        I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))), MyRange2, True)
        
        If I = 0 Then
            
           
            msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbExclamation
             Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION")) = Replace(Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION")) & ";", aa(Iaa) & ";", "")
             If Right(Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION")), 1) = ";" Then Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION")) = Left(Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION")), Len(Me.Spreadsheet1.Cells(Row, NumCollonne("conOPTION"))))
                 Spreadsheet1.SetFocus
        End If
        Next
     Set Myrange = Nothing
     End If
     
End If




    
        If Trim("" & Me.Spreadsheet1.Cells(Row, NumCollonne("conCONNECTEUR"))) <> "" Then
            Me.Spreadsheet1.Cells(Row, NumCollonne("conN°")) = Row - 1
        End If
        
    
   
   
 NoMacro1Change = False
    Col3 = 0
    SaveRow = Row
    
Fin:
DoEvents
End Sub



Private Sub Spreadsheet1_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Static NoMacro As Boolean
Static SaveRow As Long
Dim sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet1.ActiveCell.Row
Col = Me.Spreadsheet1.ActiveCell.Column

If (NoMacro = False) And (Row > 1) Then
NoMacro = True
NoMacro1Change = True
 
    If Row > 1 Then
        If Trim("" & Me.Spreadsheet1.Cells(Row, NumCollonne("ConCONNECTEUR"))) <> "" Then
            Me.Spreadsheet1.Cells(Row, NumCollonne("ConN°")) = Row - 1
        End If
        If Trim("" & Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app"))) <> "" Then
            sql = "SELECT LIAISON_CONNECTEURS.LIB FROM LIAISON_CONNECTEURS "
            sql = sql & "WHERE LIAISON_CONNECTEURS.CLIENT='" & MyReplace(MyClient) & "' "
            sql = sql & "AND LIAISON_CONNECTEURS.LIAISON='" & MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app")))) & "';"
            Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = False Then
                Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app") - 1) = Trim("'" & Rs!Lib)
            Else
'                If IfValidationOk = False Then
''                    If MsgBox("Le code App : " & DecodeCode_APP(Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app"))) & " n'existe pas" & vbCrLf & "Voulez-vous le créer", vbQuestion + vbYesNo, "Liaison Connecteur :") = vbYes Then
''                        LibCode_APP = InputBox("Entrez la désignation du code APP : " & DecodeCode_APP(Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app"))), "Ajout d'un code App")
'''                        If Trim(LibCode_APP) <> "" Then
'''                            Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app") - 1) = LibCode_APP
'''                            sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
'''                            sql = sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app"))))) & "', '" & UCase(MyReplace(Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app") - 1))) & "' );"
'''                            Con.Exequte sql
'''                        End If
''                    End If
'                Else
                   sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON "
                   sql = sql & "FROM Ajout_LIAISON_CONNECTEURS "
                   sql = sql & "WHERE Ajout_LIAISON_CONNECTEURS.LIAISON='" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app"))))) & "' "
                   sql = sql & "AND Ajout_LIAISON_CONNECTEURS.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(sql)
                    If Rs.EOF = True Then
                        sql = "INSERT INTO Ajout_LIAISON_CONNECTEURS ( LIAISON, LIB,Job ) "
                        sql = sql & "values ( '" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app"))))) & "', '" & MyReplace(Me.Spreadsheet1.Cells(Row, NumCollonne("ConCode_app") - 1)) & "'," & NmJob & ");"
                        Con.Exequte sql
                        MyErr = True
                    End If
'                End If
            End If
            Set Rs = Con.CloseRecordSet(Rs)
        
        End If
    End If
   
   
    NoMacro1Change = False
    NoMacro = False
    Col3 = 0
End If
End Sub

Private Sub Spreadsheet2_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Static SaveRow As Long
Dim Col As Long
Dim Myrange
Dim Rs As Recordset
Dim sql As String
Dim LibCode_APP As String
'Dim TrouveConnecteur() As Boolean
Dim boolReprise As Boolean
Static Col3 As Long
Row = Me.Spreadsheet2.ActiveCell.Row
Col = Me.Spreadsheet2.ActiveCell.Column
If SaveRow = 0 Then SaveRow = 1
If (NoMacro2 = True) Or (Row = 1) Then GoTo Fin
RepriseCritaire:
boolActu = False
NoMacro2 = True
 Set Myrange = Me.Spreadsheet2.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
If (Col = NumCollonne("filsFa")) Or Col = NumCollonne("filsPOS-OUT") Or Col = NumCollonne("filsPOS") Or Col = NumCollonne("filsREF CONNECTEUR") Or Col = NumCollonne("filsREF CONNECTEUR_FOUR") Then
  Col = NumCollonne("filsApp")
End If
If Col = NumCollonne("filsFa2") Or Col = NumCollonne("filsPOS-OUT2") Or Col = NumCollonne("filsPOS2") Or Col = NumCollonne("filsREF CONNECTEUR2") Or Col = NumCollonne("filsREF CONNECTEUR_FOUR2") Then
  Col = NumCollonne("filsApp2")
End If
   Debug.Print NumCollonne("filsApp")
If (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp"))) <> "") And (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp2"))) <> "") Then
'   If Trim("" & Me.Spreadsheet2.Cells(Row,NumCollonne("filsApp"))) = "186.AA" Then
'   MsgBox "186.AA"
'End If
' If Trim("" & Me.Spreadsheet2.Cells(Row, 20)) = "186.AA" Then
'   MsgBox "186.AA"
'End If
Reprise:
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Myrange.Rows.Count))


        I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp")))))

       If I <> 0 Then
                aa = Split(Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION")) & ";", ";")
'                ReDim TrouveConnecteur(UBound(aa))
                
                For Iaa = 0 To UBound(aa)
                     If InStr(1, Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";", "" & Trim(aa(Iaa)) & ";") = 0 Then
                         Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";" & Trim(aa(Iaa))
                        
                   
                     End If
                Next
        End If
        
             I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp2")))))

        If I <> 0 Then
             aa = Split("" & Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION")) & ";", ";")
'             ic = UBound(TrouveConnecteur)
'             ReDim Preserve TrouveConnecteur(ic + UBound(aa))
                For Iaa = 0 To UBound(aa)
                     If InStr(1, Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";", "" & Trim(aa(Iaa)) & ";") = 0 Then
                       
                         Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";" & Trim(aa(Iaa))
                        If InStr(1, Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";", "TOUS;") <> 0 Then
                            Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Replace(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";", "TOUS;", "")
                        End If
                       
                          
                     End If
                Next

       
        End If
        If Left(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Right(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), Len(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) - 1)
        If Right(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Left(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), Len(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) - 1)
        If UBound(Split(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";", ";")) > 2 Then
        
            MsgBox "Vous ne pouvez pas saisir plus de deux options." & vbCrLf & vbCrLf & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), vbQuestion, "AutoCâble: Tableau de fils"
          msg = "?"
          Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = ""
          If boolReprise = False Then
            boolReprise = True
            GoTo Reprise
            End If
        End If
        If InStr(1, UCase(";" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))), "TOUS;") <> 0 Then
            If Len("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) > Len("TOUS") Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Replace(UCase("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))), "TOUS", "")
                
            
        End If
        Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Replace(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), ";;", "")
        If Right(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Left(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), Len(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) - 1)
          If Left(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Right(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), Len(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) - 1)
          
     Set Myrange = Nothing
      If Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) = "" Then
        Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "TOUS" ' MsgBox "Vous devez saisir un code critère."
     
     
      End If
     End If
     If (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp"))) <> "") And (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp2"))) <> "") And (Trim(UCase("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")))) = "TOUS") Then
   
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("F1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Myrange.Rows.Count))


        I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp")))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))
           End If

        End If
        If UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")))) = "TOUS" Then
             I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp2")))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))
           End If

        End If
        End If
     Set Myrange = Nothing
     
     End If

 If (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))) <> "") Then
 Me.Spreadsheet2.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = UCase(Me.Spreadsheet2.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")))
Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES"))))
    If (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) <> "TOUS") Then
 Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = UCase(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")))
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        aa = Split(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";", ";")
        For Iaa = 0 To UBound(aa) - 1
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Myrange.Rows.Count))
        
        
        I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))))
        If UCase(aa(Iaa)) <> "TOUS" Then
        If I = 0 Then
            
           
            msg = "CODE CRITERE : " & UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")))) & " introuvable"
             MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
            Me.Spreadsheet2.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = ""
           
             Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Replace(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";", aa(Iaa) & ";", "")
                 If Right(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")).Value, 1) = ";" Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Left(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), Len(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) - 1)
        End If
        End If
        Next
        If Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) = "" Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "TOUS"
     Set Myrange = Nothing
     End If
 
      End If

If InStr(1, UCase(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) & ";", "TOUS;") <> 0 And Len(Trim(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) & ";") > Len("TOUS;") Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Replace(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), "TOUS", "")

Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Replace(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), ";;", "")

If Right(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Left(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), Len(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) - 1)
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Myrange.Rows.Count))


        I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp")))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))
           End If

        End If
        If UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")))) = "TOUS" Then
             I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp2")))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "" & Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))
           End If

        End If
        End If
     Set Myrange = Nothing


    If Trim("" & Me.Spreadsheet2.Cells(Row, 2)) <> "" Then
    If (Row > 1) And (Row = 2) Then
        If Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsFIL"))) <> Col3 Then
            Me.Spreadsheet2.Cells(Row, NumCollonne("filsFIL")) = 1
            Col3 = 1
        End If
    Else
        If (Row > 1) And (Row <> 2) Then
            If Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsFIL"))) <> Col3 Then
            Col3 = Row - 1
            Me.Spreadsheet2.Cells(Row, NumCollonne("filsFIL")) = Col3
        End If
    End If

 End If


' If Trim("" & Me.Spreadsheet2.Cells(Row, 17)) = "" Then
'        Msg = "le champ REFC/L est obligatoire"
'        MsgBox Msg, vbExclamation
'        Me.Spreadsheet2.Cells(Row - 1, 17).Select
' Else
'    If Trim("" & Me.Spreadsheet2.Cells(Row, 24)) = "" Then
'        Msg = "le champ REFC/L2 est obligatoire"
'        MsgBox Msg, vbExclamation
'        Me.Spreadsheet2.Cells(Row - 1, 24).Select
'    End If
'End If
'NumCollonne
If (Col = NumCollonne("filsApp")) Or (Col = NumCollonne("filsApp2")) Then
   
        If Trim("" & Me.Spreadsheet2.Cells(SaveRow, Col)) <> "" Then
            Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
            Set Myrange = Me.Spreadsheet1.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Myrange.Rows.Count))
            NoMacro2 = True
            
            I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))))
        
            If I <> 0 Then
            If (Col = NumCollonne("filsApp")) Then
               Me.Spreadsheet2.Cells(Row, NumCollonne("filsFA")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConN°"))))
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsPOS-OUT")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConPOS-OUT"))))
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsPOS")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConPOS"))))
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsREF CONNECTEUR")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConCONNECTEUR"))))
                 Me.Spreadsheet2.Cells(Row, NumCollonne("filsREF CONNECTEUR_FOUR")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConREFCONNECTEURFOUR"))))

            Else
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsFA2")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConN°"))))
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsPOS-OUT2")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConPOS-OUT"))))
                Me.Spreadsheet2.Cells(Row, NumCollonne("filsPOS2")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConPOS"))))
                 Me.Spreadsheet2.Cells(Row, NumCollonne("filsREF CONNECTEUR2")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConCONNECTEUR"))))
                 Me.Spreadsheet2.Cells(Row, NumCollonne("filsREF CONNECTEUR_FOUR2")) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, NumCollonne("ConREFCONNECTEURFOUR"))))
            End If

                Else
                    If (Col = NumCollonne("filsApp")) Then
                        Me.Spreadsheet2.Cells(Row, NumCollonne("filsFA")) = "0"
                        Me.Spreadsheet2.Cells(Row, NumCollonne("filsPOS-OUT")) = ""
                    Else
                        Me.Spreadsheet2.Cells(Row, NumCollonne("filsFA2")) = "0"
                        Me.Spreadsheet2.Cells(Row, NumCollonne("filsPOS-OUT2")) = ""
                    End If
                    msg = "Le connecteur : " & UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))) & " introuvable"
                    MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                     Me.Spreadsheet2.Cells(Row, Col).Select
                      Me.Spreadsheet2.SetFocus
                End If
            
            
            
            
                
                
                    
                End If
            End If
        
        
       
        
           
            
            Col3 = 0
        End If
   If (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp"))) <> "") Then

 Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range(ChrCollonne("ConCODE_APP") & "1", ChrCollonne("ConCODE_APP") & CStr(Myrange.Rows.Count))


        I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp")))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
                Myapp1 = "" & Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))
           End If

        End If
       
             I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsApp2")))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))) <> "" Then
               Myapp2 = "" & Me.Spreadsheet1.Cells(I, NumCollonne("ConOPTION"))
           End If
    If Trim("" & Myapp1) = "" Then Myapp1 = Myapp2
     If Trim(Myapp2) = "" Then Myapp2 = Myapp1
        If ("" & Myapp1 <> Myapp2) And (UCase(Myapp1) <> "TOUS") And (UCase(Myapp2) <> "TOUS") Then
        MsgBox "Une liaison ne peut pas pointer sur deux options différentes : " & Myapp1 & " & " & Myapp2, vbQuestion, "AutoCâble: Tableau de fils"
        Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = ""
        Spreadsheet2.SetFocus
        msg = "?"
        End If
        End If
        
If (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) <> "") And (Trim("" & Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) <> "TOUS") Then
        Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = UCase(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")))
    Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
    Set Myrange = Me.Spreadsheet5.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Myrange.Rows.Count))

    aa = Split(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) & ";", ";")
    If UBound(aa) = 3 Then
        MsgBox "Vous ne pouvez pas saisir plus de deux critére." & vbCrLf & vbCrLf & Me.Spreadsheet2.Cells(Row, 28), vbQuestion, "AutoCâble: Tableau de fils"
        Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = "TOUS"
         Me.Spreadsheet2.Cells(Row, NumCollonne("filsCRITÈRES SPÉCIFIQUES")) = ""
         GoTo RepriseCritaire
         
         
    Else
        zz = ""
        For Iaa = 0 To UBound(aa) - 1
            I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))))
            If I <> 0 Then
                If Me.Spreadsheet5.ActiveSheet.Cells(I, 1) = 0 Then GoTo ErrorCritere
                If InStr(1, ";" & zz & ";", ";" & aa(Iaa) & ";") = 0 Then
                    zz = "" & zz & Trim(aa(Iaa)) & ";"
                End If
            Else
ErrorCritere:
                MsgBox "Code Critère: " & aa(Iaa) & " introuvable ou Inactif", vbExclamation
            End If
        Next
        Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = zz
         
        If Right(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), 1) = ";" Then Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")) = Left(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")), Len(Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION"))) - 1)
'         Me.Spreadsheet2.Cells(Row, 28).Value = Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")).Value
         If Trim("" & (Me.Spreadsheet2.Cells(Row, NumCollonne("filsOPTION")))) = "" Then GoTo RepriseCritaire
    End If
End If
Set Myrange = Nothing
End If
SaveRow = Row
    NoMacro2 = False
Fin:
End Sub

Private Sub Spreadsheet2_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Static SaveRow As Long
Dim Row As Long
Dim sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet2.ActiveCell.Row
If (NoMacro2 = True) Or (Row = 1) Then GoTo Fin
If SaveRow = 0 Then SaveRow = Me.Spreadsheet2.ActiveCell.Row
 If Trim("" & Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsLIAI"))) <> "" Then
                sql = "SELECT LIAISON.LIB FROM LIAISON "
                sql = sql & "WHERE LIAISON.CLIENT='" & MyReplace(MyClient) & "' "
                sql = sql & "AND LIAISON.LIAISON='" & MyReplace(Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsLIAI"))) & "';"
                Set Rs = Con.OpenRecordSet(sql)
                If Rs.EOF = False Then
                Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsDESIGNATION")) = Trim("'" & Rs!Lib)
                Else
                    If IfValidationOk = False Then
                        If SaveRow <> Row And SaveRow <> 1 Then
'                            If MsgBox("La liaison : " & Me.Spreadsheet2.Cells(SaveRow,NumCollonne("filsLIAI")) & " n'existe pas" & vbCrLf & "Voulez-vous la créer", vbYesNo + vbQuestion, "AutoCâble: Tableau de fils") = vbYes Then
'                                LibCode_APP = InputBox("Entrez la désignation de la liaison : " & Me.Spreadsheet2.Cells(SaveRow, 1), "Ajout de liaison")
''                                If Trim(LibCode_APP) <> "" Then
''                                    Me.Spreadsheet2.Cells(SaveRow,NumCollonne("filsDESIGNATION")) = LibCode_APP
''                                    sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
''                                    sql = sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(Me.Spreadsheet2.Cells(SaveRow,NumCollonne("filsLIAI")))) & "', '" & UCase(MyReplace("" & LibCode_APP)) & "' );"
''                                    Con.Exequte sql
''                                End If
'                            End If
                        End If
                        Else
                        
                   sql = "SELECT Ajout_LIAISON.LIAISON "
                   sql = sql & "FROM Ajout_LIAISON "
                   sql = sql & "WHERE Ajout_LIAISON.LIAISON='" & UCase(MyReplace(Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsLIAI")))) & "' "
                   sql = sql & "AND Ajout_LIAISON.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(sql)
                    If Rs.EOF = True Then
                        sql = "INSERT INTO Ajout_LIAISON ( LIAISON, LIB,Job ) "
                        sql = sql & "values ( '" & UCase(MyReplace(Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsLIAI")))) & "', '" & MyReplace(Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsDESIGNATION"))) & "'," & NmJob & ");"
                        Con.Exequte sql
                        MyErr = True
                    End If
                    End If
                End If
                Set Rs = Con.CloseRecordSet(Rs)
            If Trim("" & Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsLIAI"))) <> "" And SaveRow <> Row Then
                        If UCase(Trim("" & Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsACTIVER")))) <> 0 Then
                        If Len(Trim("" & Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsApp")))) = 0 Then
                        
                            msg = "Le code APP ne peut être Nul"
                            MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                             Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsApp")).Select
                              Me.Spreadsheet2.SetFocus
                               Row = SaveRow
                              GoTo Fin
                          End If
                         If Len(Trim("" & Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsApp2")))) = 0 Then
                        
                            msg = "Le code APP ne peut être Nul"
                            MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                             Me.Spreadsheet2.Cells(SaveRow, NumCollonne("filsApp2")).Select
                              Me.Spreadsheet2.SetFocus
                               Row = SaveRow
                          End If
                        End If
                    End If
            End If
Fin:
SaveRow = Row
End Sub

Private Sub Spreadsheet3_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveRow As Long
Dim LibCode_APP As String
Dim sql As String
Dim Rs As Recordset
Dim BoolOui As Boolean
If SaveRow = 0 Then SaveRow = 1


Row = Me.Spreadsheet3.ActiveCell.Row
Col = Me.Spreadsheet3.ActiveCell.Column
If (NoMacro3 = True) Or (Row = 1) Then GoTo Fin

If SaveRow = 0 Then SaveRow = Row
If Row = 1 Then GoTo Fin
boolActu = False
NoMacro3 = True
 Set Myrange = Me.Spreadsheet3.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
   If Trim("" & Me.Spreadsheet3.Cells(Row, 2)) <> "" Then Me.Spreadsheet3.Cells(Row, 3) = Row - 1
   If Trim("" & Me.Spreadsheet3.Cells(Row, 5)) <> "" Then
    Me.Spreadsheet3.Cells(Row, 5) = UCase(Me.Spreadsheet3.Cells(Row, 5))
    Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Myrange.Rows.Count))
 aa = Split(Me.Spreadsheet3.Cells(Row, 5) & ";", ";")
                For Iaa = 0 To UBound(aa) - 1

        I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))))

        If I = 0 Then
                msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbQuestion, "AutoCâble: Composants"
                     Me.Spreadsheet3.Cells(Row, 5) = Replace(Me.Spreadsheet3.Cells(Row, 5) & ";", "" & Trim(aa(Iaa)) & ";", "")
                Spreadsheet3.SetFocus
        End If
        Next
        If Right(Me.Spreadsheet3.Cells(Row, 5), 1) = ";" Then Me.Spreadsheet3.Cells(Row, 5) = Left(Me.Spreadsheet3.Cells(Row, 5), Len(Me.Spreadsheet3.Cells(Row, 5)) - 1)
   End If
   If Col > NumCollonne("CompCODE_APP_LIER") Then
    For I = NumCollonne("CompCODE_APP_LIER") + 1 To NbFinOuiNon
        If Me.Spreadsheet3.Cells(Row, I) = 1 And Me.Spreadsheet3.Cells(Row, Col) = 1 Then
            If I <> Col Then
                MsgBox "Vous ne pouvez pas sélectionner plusieurs répertoires.", vbQuestion, "AutoCâble: Composants"
                Me.Spreadsheet3.Cells(Row, Col) = 0
               Exit For
            End If
        End If
    Next I
End If
If Col = NumCollonne("CompCODE_APP_LIER") Then

    If ChercheXls(Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion, Spreadsheet3.ActiveCell.Value) = 0 Then
        MsgBox "Le code App : " & Spreadsheet3.ActiveCell.Value & " n'existe pas dans la liste des connecteur.", vbExclamation
        Spreadsheet3.ActiveCell.Clear
    End If
End If
   NoMacro3 = False
Fin:
End Sub

Private Sub Spreadsheet3_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveCol As Long
Static SaveRow As Long
Dim LibCode_APP As String
Dim sql As String
Dim Rs As Recordset
Dim BoolOui As Boolean
If SaveRow = 0 Then SaveRow = 1
If SaveCol = 0 Then SaveCol = 1
If NoMacro3 = True Then GoTo Fin

Row = Me.Spreadsheet3.ActiveCell.Row
Col = Me.Spreadsheet3.ActiveCell.Column

If Row = 1 Then GoTo Fin
NoMacro3 = True
 Set Myrange = Me.Spreadsheet1.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
   If Trim("" & Me.Spreadsheet3.Cells(Row, 2)) <> "" Then Me.Spreadsheet3.Cells(Row, 3) = Row - 1
    
If Col > NumCollonne("CompCODE_APP_LIER") Then
    For I = NumCollonne("CompCODE_APP_LIER") + 1 To NbFinOuiNon
        If Me.Spreadsheet3.Cells(Row, I) = 1 And Me.Spreadsheet3.Cells(Row, Col) = 1 Then
            If I <> Col Then
                MsgBox "Vous ne pouvez pas sélectionner plusieurs répertoires.", vbQuestion, "AutoCâble: Composants"
                Me.Spreadsheet3.Cells(Row, Col) = 0
               Exit For
            End If
        End If
    Next I
End If
BoolOui = False
If (SaveRow <> 1) And (SaveRow <> Row) And (Trim("" & Me.Spreadsheet3.Cells(SaveRow, 1)) <> "") And Col > NumCollonne("CompCODE_APP_LIER") Then
 For I = NumCollonne("CompCODE_APP_LIER") + 1 To NbFinOuiNon
    If Val(Me.Spreadsheet3.Cells(SaveRow, I)) = 1 Then
        BoolOui = True
        
        Exit For
    End If
    
    Next I
  If BoolOui = False And Me.Spreadsheet3.Cells(SaveRow, 1) = 1 Then
    MsgBox "Vous devez sélectionner un répertoire.", vbQuestion, "AutoCâble: Composants"
    Me.Spreadsheet3.Cells(SaveRow, NumCollonne("CompCODE_APP_LIER") + 1).Select
     Me.Spreadsheet3.SetFocus
     Row = SaveRow
    msg = "?"
  End If
   
End If
SaveRow = Row
SaveCol = Col
NoMacro3 = False
Fin:
End Sub

Private Sub Spreadsheet4_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Dim sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet4.ActiveCell.Row
Col = Me.Spreadsheet4.ActiveCell.Column
If Row = 1 Then GoTo Fin
If (NoMacro4 = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro4 = True
 Set Myrange = Me.Spreadsheet4.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
If Col > 2 Then
   If Trim("" & Me.Spreadsheet4.Cells(Row, 2)) <> "" Then Me.Spreadsheet4.Cells(Row, 3) = Row - 1
    
    If Trim("" & Me.Spreadsheet4.Cells(Row, 4)) <> "" Then
    Me.Spreadsheet4.Cells(Row, 4) = UCase(Me.Spreadsheet4.Cells(Row, 4))
    Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range(ChrCollonne("CritCODE_CRITERE") & "1", ChrCollonne("CritCODE_CRITERE") & CStr(Myrange.Rows.Count))
 aa = Split(Me.Spreadsheet4.Cells(Row, 4) & ";", ";")
                For Iaa = 0 To UBound(aa) - 1

        I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))))

        If I = 0 Then
                msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbQuestion, "AutoCâble: Notas"
                     Me.Spreadsheet4.Cells(Row, 4) = Replace(Me.Spreadsheet4.Cells(Row, 4) & ";", "" & Trim(aa(Iaa)) & ";", "")
                Spreadsheet4.SetFocus
        End If
        Next
        If Right(Me.Spreadsheet4.Cells(Row, 4), 1) = ";" Then Me.Spreadsheet4.Cells(Row, 4) = Left(Me.Spreadsheet4.Cells(Row, 4), Len(Me.Spreadsheet4.Cells(Row, 4)) - 1)
   End If
End If
NoMacro4 = False

Fin:
End Sub

Private Sub Form_Activate()
Dim MyExcel As EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
If BooolBloque = True Then
    Command3.Visible = False
    Command1.Visible = False
    Command2.Visible = False
End If
If bool_Activate = True Then GoTo Fin
Set CollectionMenu = Nothing
Set CollectionMenu = New Collection
CollectionMenu.Add "Spreadsheet5", "Critères"
CollectionMenu.Add "Spreadsheet1", "Connecteurs"
CollectionMenu.Add "Spreadsheet2", "Tableau de fils"
CollectionMenu.Add "Spreadsheet3", "Composants"
CollectionMenu.Add "Spreadsheet4", "Notas"
CollectionMenu.Add "Spreadsheet6", "Noeuds"
CollectionMenu.Add "Spreadsheet7", "Nomenclature Connecteur"
CollectionMenu.Add "Spreadsheet8", "Nomenclature Fils"
CollectionMenu.Add "Spreadsheet9", "Nomenclature Habillage"
CollectionMenu.Add "Spreadsheet10", "Nomenclatures"

'
'CollectionMenu.Add "Spreadsheet1", "Connecteurs"
'CollectionMenu.Add "Spreadsheet1", "Connecteurs"
'CollectionMenu.Add "Spreadsheet1", "Connecteurs"
'CollectionMenu.Add "Spreadsheet1", "Connecteurs"

 Mygrid = "Spreadsheet5"
Set NumCollonne = Nothing
Set NumCollonne = New Collection
Set ChrCollonne = Nothing
Set ChrCollonne = New Collection
bool_Activate = True
Dim Myrange As EXCEL.Range
Set MyExcel = New EXCEL.Application
MyExcel.DisplayAlerts = False
NotSortie = True
'MyExcel.Visible = True
Set a = Me.Spreadsheet1.Cells(2, 2)
If Trim(Me.Caption) = "" Then Exit Sub
If Nouveau = False Then

    Set MyWorkbook = MyExcel.Workbooks.Open(Me.Caption)
Else

    Set MyWorkbook = MyExcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
End If

Set Myrange = MyWorkbook.Sheets("Critères").Range("a1").CurrentRegion
 Myrange.Replace Chr(10), "©"
'For I = 1 To Myrange.Count
'    Myrange(I) = Replace(Myrange(I), Chr(10), "©")
'Next
Myrange.Copy
Me.Spreadsheet5.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet5.ActiveSheet, "Crit"
Set Myrange = MyWorkbook.Sheets("Connecteurs").Range("a1").CurrentRegion
Myrange.Replace Chr(10), "©"
Myrange.Copy
Me.Spreadsheet1.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet1.ActiveSheet, "Con"
'Me.Spreadsheet1.Sheet.Add

Set Myrange = MyWorkbook.Sheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
'For I = 1 To Myrange.Count
'    Myrange(I) = Replace(Myrange(I), Chr(10), "©")
'Next
 Myrange.Replace Chr(10), "©"

Myrange.Copy
Me.Spreadsheet2.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet2.ActiveSheet, "Fils"


Set Myrange = MyWorkbook.Sheets("Composants").Range("a1").CurrentRegion
'For I = 1 To Myrange.Count
'    Myrange(I) = Replace(Myrange(I), Chr(10), "©")
'Next
 Myrange.Replace Chr(10), "©"
NbFinOuiNon = Myrange.Columns.Count
Myrange.Copy
Me.Spreadsheet3.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet3.ActiveSheet, "Comp"

Set Myrange = MyWorkbook.Sheets("Notas").Range("a1").CurrentRegion
 Myrange.Replace Chr(10), "©"
'For I = 1 To Myrange.Count
'    Myrange(I) = Replace(Myrange(I), Chr(10), "©")
'Next
Myrange.Copy
Me.Spreadsheet4.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet4.ActiveSheet, "Not"
Set Myrange = MyWorkbook.Sheets("NOEUDS").Range("a1").CurrentRegion
'For I = 1 To Myrange.Count
'    Myrange(I) = Replace(Myrange(I), Chr(10), "©")
'Next
 Myrange.Replace Chr(10), "©"
Myrange.Copy
Me.Spreadsheet6.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet6.ActiveSheet, "Noeu"
Set Myrange = MyWorkbook.Sheets("Nomenclature Connecteur").Range("a1").CurrentRegion
'For I = 1 To Myrange.Count
'    Myrange(I) = Replace(Myrange(I), Chr(10), "©")
'Next
 Myrange.Replace Chr(10), "©"
Myrange.Copy
Me.Spreadsheet7.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet7.ActiveSheet, "NomCon"
Set Myrange = MyWorkbook.Sheets("Nomenclature Fils").Range("a1").CurrentRegion
'For I = 1 To Myrange.Count
'    Myrange(I) = Replace(Myrange(I), Chr(10), "©")
'Next
 Myrange.Replace Chr(10), "©"
Myrange.Copy
Me.Spreadsheet8.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet8.ActiveSheet, "NomFil"
Set Myrange = MyWorkbook.Sheets("Nomenclature Habillage").Range("a1").CurrentRegion
'For I = 1 To Myrange.Count
'    Myrange(I) = Replace(Myrange(I), Chr(10), "©")
'Next
 Myrange.Replace Chr(10), "©"
Myrange.Copy
Me.Spreadsheet9.ActiveSheet.Range("a1").Paste
Charger_Colection Me.Spreadsheet9.ActiveSheet, "NimHab"
MyExcel.AlertBeforeOverwriting = False
Myrange.Application.Visible = True
'Me.Spreadsheet10.Sheets(1).Name = "Nomenclature"
'Me.Spreadsheet10.Sheets.Add After:=1
'Me.Spreadsheet10.Sheets(2).Name = "Nomenclature Finale"
Set Myrange = MyWorkbook.Sheets("Nomenclature").Range("a1").CurrentRegion 'Selection.CurrentRegion.Select
Myrange.Replace Chr(10), "©"
Myrange.Copy
'Me.Spreadsheet10.Sheets(1).Cells.Locked = False
Me.Spreadsheet10.Sheets(1).Range("a1").Paste
On Error Resume Next
Me.Spreadsheet10.Sheets(1).AutoFilterMode = True
'Me.Spreadsheet10.Sheets(1).Cells.Locked = True
Set Myrange = MyWorkbook.Sheets("Nomenclature Finale").Range("a1").CurrentRegion 'Selection.CurrentRegion.Select
Myrange.Replace Chr(10), "©"
Myrange.Copy
'Me.Spreadsheet10.Sheets(2).Cells.Locked = False
Me.Spreadsheet10.Sheets(2).Range("a1").Paste
Me.Spreadsheet10.Sheets(2).AutoFilterMode = True
On Error GoTo 0
'Me.Spreadsheet10.Sheets(2).Cells.Locked = True
Set Myrange = Nothing
MyWorkbook.Close False
Set MyWorkbook = Nothing

MyExcel.Quit
Set MyExcel = Nothing


'Me.Spreadsheet1.ActiveSheet.Panes(1).VisibleRange = False
Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet4.ActiveSheet.Range("a1").Select
Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet5.ActiveSheet.Range("a1").Select
Me.Spreadsheet5.Columns(NumCollonne("CritActiver")).NumberFormat = "Yes/No"

Me.Spreadsheet6.Columns(NumCollonne("NoeuActiver")).NumberFormat = "Yes/No"
Me.Spreadsheet6.Columns(NumCollonne("NoeuFleche_Droite")).NumberFormat = "Yes/No"
Me.Spreadsheet6.Columns(NumCollonne("NoeuTORON_PRINCIPAL")).NumberFormat = "Yes/No"
Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet6.ActiveSheet.Range("a1").Select

Me.Spreadsheet3.ActiveSheet.Range("a1").Select
Me.Spreadsheet2.ActiveSheet.Range("a1").Select
Me.Spreadsheet2.Columns(NumCollonne("FilsActiver")).NumberFormat = "Yes/No"
Me.Spreadsheet4.Columns(NumCollonne("NotActiver")).NumberFormat = "Yes/No"
Me.Spreadsheet1.ActiveSheet.Range("a1").Select
Me.Spreadsheet1.Columns(NumCollonne("ConActiver")).NumberFormat = "Yes/No"
Me.Spreadsheet1.Columns(NumCollonne("ConO/N")).NumberFormat = "Yes/No"
Me.Spreadsheet3.Columns(NumCollonne("CompACTIVER")).NumberFormat = "Yes/No"
For I = NumCollonne("CompCODE_APP_LIER") + 1 To Me.Spreadsheet3.Range("A1").CurrentRegion.Columns.Count
    Me.Spreadsheet3.Columns(I).NumberFormat = "Yes/No"
 Next I
    
DoEvents
LstMaj
Fin:
End Sub
Public Sub Chargement(fichier As String, Client As String, Id As Long, Optional NouveauF As Boolean)
Dim Rs As Recordset
Dim sql As String
Dim txtMyCollectionLienHab As String
IdProjet = Id
sql = "SELECT T_Regle_Comp_Hab.ENCELADE, T_Regle_Comp_Hab.PSA, "
sql = sql & "T_Regle_Comp_Hab.RSA, T_Regle_Comp_Hab.libellé, T_Regle_Comp_Hab.Numéro "
sql = sql & "FROM T_Regle_Comp_Hab "
sql = sql & "ORDER BY T_Regle_Comp_Hab.libellé;"
I = 0
Set Rs = Con.OpenRecordSet(sql)
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
MyClient = Client
Nouveau = NouveauF
Me.Caption = fichier
Me.Show vbModal
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = NotSortie
bool_Activate = False
End Sub


Private Sub Spreadsheet5_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
'NumCollonne("CritACTIVER")


Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Static SaveRow As Long
Dim sql As String
Dim Rs As Recordset
If (NoMacro5 = True) Or (NoMacro5Select = True) Or (Row = 1) Then GoTo Fin
Row = Me.Spreadsheet5.ActiveCell.Row
Col = Me.Spreadsheet5.ActiveCell.Column

If SaveRow = 0 Then SaveRow = Row
If Row = 1 Then GoTo Fin
NoMacro5 = True

boolActu = False

 Set Myrange = Me.Spreadsheet5.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
     Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCODE_CRITERE")) = UCase(Trim("" & Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCODE_CRITERE"))))
     Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCRITERES")) = UCase(Trim("" & Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCRITERES"))))
     
    If (Trim("" & Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCODE_CRITERE"))) <> "") And (Trim("" & Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCRITERES"))) = "") And (SaveRow <> Row) Then
        MsgBox "Le champ CRITERES est obligatoire", vbQuestion, "AutoCâble: Critères"
        msg = "?"
       Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCRITERES")).Select
       Spreadsheet5.SetFocus
    End If
    If (Trim("" & Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCODE_CRITERE"))) = "") And (Trim("" & Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCRITERES"))) <> "") And (SaveRow <> Row) Then
        MsgBox "Le champ CODE CRITERES est obligatoire", vbQuestion, "AutoCâble: Critères"
        msg = "?"
       Me.Spreadsheet5.Cells(SaveRow, NumCollonne("CritCODE_CRITERE")).Select
       Spreadsheet5.SetFocus
    End If
SaveRow = Row
NoMacro5 = False
Fin:
End Sub

Private Sub Spreadsheet5_Click(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Spreadsheet5_Change EventInfo
End Sub

Private Sub Spreadsheet5_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
If NoMacro5Select = True Then GoTo Fin
 NoMacro5Select = True
Spreadsheet5_Change EventInfo
Static Row As Long

Dim aa
Row = Spreadsheet5.ActiveCell.Row
If Row = 0 Then Row = 1
If Row > 1 And IfValidationOk = True Then
On Error Resume Next

If Spreadsheet5.Cells(Row, NumCollonne("CritCODE_CRITERE")) <> "" Then
If Spreadsheet5.Cells(Row, NumCollonne("CritACTIVER")) = 1 Then
aa = ""
    aa = CollecCrieres(Spreadsheet5.Cells(Row, NumCollonne("CritCODE_CRITERE")))
    If Err Then
    Err.Clear
        CollecCrieres.Add Spreadsheet5.Cells(Row, NumCollonne("CritCODE_CRITERE")), Spreadsheet5.Cells(Row, NumCollonne("CritCODE_CRITERE"))
    Else
        MsgBox " Le code Code Critères : " & Spreadsheet5.Cells(Row, NumCollonne("CritCODE_CRITERE")) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
            Me.Spreadsheet5.Cells(Row, NumCollonne("CritCODE_CRITERE")).Select
            Me.Spreadsheet5.SetFocus
            NoMacro5Select = False
             On Error GoTo 0
        GoTo Fin
    End If
 End If
End If

If Spreadsheet5.Cells(Row, NumCollonne("CritCRITERES")) <> "" Then
If Spreadsheet5.Cells(Row, NumCollonne("CritACTIVER")) = 1 Then
aa = ""
    aa = CollecCrieresCode(Spreadsheet5.Cells(Row, NumCollonne("CritCRITERES")))
    If Err Then
        Err.Clear
        CollecCrieresCode.Add Spreadsheet5.Cells(Row, NumCollonne("CritCRITERES")), Spreadsheet5.Cells(Row, NumCollonne("CritCRITERES"))
    Else
        MsgBox " Le Critères : " & Spreadsheet5.Cells(Row, NumCollonne("CritCRITERES")) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
             Me.Spreadsheet5.Cells(Row, NumCollonne("CritCRITERES")).Select
            Me.Spreadsheet5.SetFocus
            NoMacro5Select = False
             On Error GoTo 0
        GoTo Fin
        GoTo Fin
    End If
    End If
End If
If Spreadsheet5.Cells(Row, NumCollonne("CritDESIGNATION")) <> "" Then
If Spreadsheet5.Cells(Row, NumCollonne("CritACTIVER")) = 1 Then
aa = ""
    aa = CollecCrieresDesigne(Spreadsheet5.Cells(Row, NumCollonne("CritDESIGNATION")))
    If Err Then
        Err.Clear
        CollecCrieresDesigne.Add Spreadsheet5.Cells(Row, NumCollonne("CritDESIGNATION")), Spreadsheet5.Cells(Row, NumCollonne("CritDESIGNATION"))
    Else
        MsgBox " Le code Designation Critères : " & Spreadsheet5.Cells(Row, NumCollonne("CritDESIGNATION")) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
             Me.Spreadsheet5.Cells(Row, NumCollonne("CritDESIGNATION")).Select
            Me.Spreadsheet5.SetFocus
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

Private Sub Spreadsheet6_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Row = Spreadsheet6.ActiveCell.Row
If Row = 1 Then GoTo Fin
If NoMacro6 = True Then GoTo Fin
boolActu = False

NoMacro6 = True
 Set Myrange = Me.Spreadsheet6.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
If Trim("" & Spreadsheet6.Cells(Row, 1)) <> "" Then
'    Spreadsheet6.Cells(Row, 4) = NoeuName(Row)
'    Else
'        If Trim("" & Spreadsheet6.Cells(Row, 8)) <> "" Then
'           Spreadsheet6.Cells(Row, 4) = NoeuName(Row)
'        Else
'            If Trim("" & Spreadsheet6.Cells(Row, 9)) <> "" Then
'                Spreadsheet6.Cells(Row, 4) = NoeuName(Row)
'            Else
'                If Trim("" & Spreadsheet6.Cells(Row, 10)) <> "" Then
                    Spreadsheet6.Cells(Row, 4) = NoeuName(Row)
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
Dim NbTour As Long
Dim NbTord As Long
Dim txtColone As Long
txtColone = 2
Txt = "AA"
Ofset = 0
NbTour = 0
NbTord = 0
Reprise:

For I = 0 To Row - 2
aa = Mid(Txt, Len(Txt) - Ofset, 1)

    aa = Chr(Asc("A") + (1 * (I - (26 * NbTour))))

Mid(Txt, Len(Txt) - Ofset, 1) = aa


If Asc(Mid(aa, 1, 1)) < 65 Or Asc(Mid(aa, 1, 1)) > 90 Then

Mid(Txt, Len(Txt) - Ofset, 1) = "A"


    Ofset = Ofset + 1
    NbTour = NbTour + 1
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

Private Sub Spreadsheet6_SelectionChanging(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim aa As String
Dim MyTxt As String
On Error Resume Next
If boolSelctChange = False Then
 Me.Tag = ""
If EventInfo.Range.Row > 1 Then
Me.Tag = EventInfo.Range.Row
Me.Fleche_Droite.Value = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 2)
TORON_P.Value = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 3)
    Me.Longueur = CStr(Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 5))
    Long_C = CStr(Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 6))
    Me.NOUED = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 4)
    
    If Trim("" & Spreadsheet6.Cells(EventInfo.Range.Row, 7)) <> "" Then
            Me.Hab.ListIndex = MyCollectionHab("N" & Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 7))
    Else
        If Trim("" & Spreadsheet6.Cells(EventInfo.Range.Row, 8)) <> "" Then
           Me.RSA.ListIndex = MyCollectionRSA("N" & Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 8))
        Else
            If Trim("" & Spreadsheet6.Cells(EventInfo.Range.Row, 9)) <> "" Then
                Me.PSA.ListIndex = MyCollectionPSA("N" & Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 9))
            Else
                If Trim("" & Spreadsheet6.Cells(EventInfo.Range.Row, 10)) <> "" Then
                   Me.ENC.ListIndex = MyCollectionENC("N" & Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 10))
                  Else
                    Me.Hab.ListIndex = 0
                   
                End If
        End If
    End If
End If
    
   
   Me.Activer.Value = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 1)
   Me.DIAMETRE = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 11)
   Me.CLASSE_T = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 12)
    Me.txtOption = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 13)
   
End If
End If
DoEvents
On Error GoTo 0
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

Sub ConverOuiNon(Myrange, Index)
For I = 1 To Myrange.Columns.Count
   If Myrange(Index, I).NumberFormat = "Yes/No" Then
  
        If Not IsNumeric(Myrange(Index, I).Value) Then
            If UCase(Left(Myrange(Index, I).Value, 1)) = "N" Then
                Myrange(Index, I).Value = 0
                DoEvents
               
            Else
                Myrange(Index, I).Value = 1
                DoEvents
               
            End If
        End If
      
   End If
    
Next
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.StopMaco.Visible = True Then
    txtMacro = txtMacro & Space(5) & "SSTab1.Tab =" & SSTab1.Tab & vbCrLf
End If
        Mygrid = CollectionMenu(SSTab1.Caption)
    

End Sub

Private Sub StopMaco_Click()
Dim sql As String
If MacroName <> "" Then
    Me.txtMacro = Me.txtMacro & "End Sub " & vbCrLf
    sql = "UPDATE T_Macro SET T_Macro.Sub = '" & Replace(Me.txtMacro, "'", "''") & "'  "
    sql = sql & "WHERE T_Macro.Formulaire='" & Me.Name & "'  "
    sql = sql & "AND T_Macro.Macro='" & Replace(MacroName, "'", "''") & "' ;"
    Con.Exequte sql
End If
MacroName = ""
 Me.txtMacro = ""
StopMaco.Visible = False
End Sub

Private Sub SupMacro_Click()
frmMacco.charger Me.Name, "SUP"
Unload frmMacco
End Sub
