VERSION 5.00
Begin VB.Form FrmNomChoix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choix Nomenclature :"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Valider"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sélection du type d'opération : "
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton Option1 
         Caption         =   "Préparer la liste d'approvisinnement finale"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nomenclature par Code Appareil"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Prénomenclature par Code Appareil"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Préparer nomenclature par Code Appareil"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
   End
End
Attribute VB_Name = "FrmNomChoix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Valide As Boolean
Public PreparNomk As Integer

Private Sub Command1_Click()
If Me.Option1(0).Value = True Then PreparNomk = 0
If Me.Option1(1).Value = True Then PreparNomk = 1
If Me.Option1(2).Value = True Then PreparNomk = 2
If Me.Option1(3).Value = True Then PreparNomk = 3

Valide = True
Me.Hide
End Sub

Private Sub Command2_Click()
Valide = False
'frmAutocâble.DesEnabledMenu
Me.Hide
End Sub

