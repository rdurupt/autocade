VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ShershFihier 
   Caption         =   "Form1"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   285
      Left            =   10080
      Picture         =   "ShershFihier.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   285
      Left            =   10080
      Picture         =   "ShershFihier.frx":00A5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.TextBox Ficher 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "html"
      Filter          =   "*.html|*.*"
      FilterIndex     =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Liste des laisons filaires :"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Liste des connecteurs :"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1725
   End
End
Attribute VB_Name = "ShershFihier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
On Error GoTo Fin
Me.CommonDialog1.CancelError = True
Me.CommonDialog1.ShowOpen
Me.Ficher.Text = Me.CommonDialog1.Filename
Fin:
Err.Clear
On Error GoTo 0
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

Public Function chargement() As String
Me.Show vbModal
chargement = Me.Ficher.Text
End Function

