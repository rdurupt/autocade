VERSION 5.00
Begin VB.Form Restart 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Restart.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Reprendre le traitement"
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Restart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
SetTopMostWindow Me, False
Me.WindowState = 1
 Unload Me
End Sub

Private Sub Form_Activate()

    SetTopMostWindow Me, True

End Sub

Private Sub Form_Load()
DoEvents
End Sub

Private Sub Timer1_Timer()
DoEvents
End Sub
