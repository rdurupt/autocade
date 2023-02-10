VERSION 5.00
Begin VB.Form MenuMagasin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Magasin"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   Icon            =   "MenuMagasin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Retour"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Prix"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command4 
         Caption         =   "Ex&port Habillage:"
         Height          =   615
         Left            =   2280
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Import &Habillage:"
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export Câbles:"
         Height          =   615
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Import Câbles"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "MenuMagasin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
XlsPrix = "CablePrix"
subImport
End Sub

Private Sub Command2_Click()
XlsPrix = "CablePrix"
subExport
End Sub

Private Sub Command3_Click()
XlsPrix = "HabillagePrix"
subImport
End Sub


Private Sub Command4_Click()
XlsPrix = "HabillagePrix"
subExport

End Sub

Private Sub Command5_Click()
Unload Me
End Sub
