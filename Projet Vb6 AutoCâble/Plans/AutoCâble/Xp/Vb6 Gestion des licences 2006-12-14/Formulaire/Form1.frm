VERSION 5.00
Object = "{0002E550-0000-0000-C000-000000000046}#1.1#0"; "OWC10.DLL"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OWC10.Spreadsheet Spreadsheet1 
      Height          =   6855
      Left            =   240
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   360
      Width           =   9735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Spreadsheet1.Worksheets.Add
End Sub
