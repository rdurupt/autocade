VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HELP"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Text1 = Me.Text1 & vbCrLf & App.Path & "\" & App.EXEName & ".EXE -install"
Me.Text1 = Me.Text1 & vbCrLf & App.Path & "\" & App.EXEName & ".EXE -uninstall"
Me.Text1 = Me.Text1 & vbCrLf & App.Path & "\" & App.EXEName & ".EXE -debug"
'C:\MCT\v1\ServeurEuxia\MCT_Serveur_Euxia.exe -install
'C:\MCT\v1\ServeurEuxia\MCT_Serveur_Euxia.exe -uninstall
'C:\MCT\v1\ServeurEuxia\MCT_Serveur_Euxia.exe -debug

End Sub

