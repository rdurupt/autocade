VERSION 5.00
Begin VB.Form frmLstComposants 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   285
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4665
   ControlBox      =   0   'False
   Icon            =   "frmLstComposants.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmLstComposants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyRange


Public Sub chargement(lst, MyCell)
Dim I As Long
Set MyRange = MyCell

'Me.Visible = False
For I = 0 To UBound(lst)
    Me.Combo1.AddItem lst(I)
Next
Me.Show
End Sub
Public Sub Rensienge(MyCell, x As Long, y As Long)


Set MyRange = MyCell
GetCursorPos PosCursor

Me.Top = 0 '(PosCursor.X)
Me.Left = 0 '(PosCursor.Y)
Me.Visible = True

End Sub

Private Sub Combo1_Click()
MyRange = Me.Combo1.List(Me.Combo1.ListIndex)
SetTopMostWindow Me, False
Me.Hide

End Sub

Private Sub Form_Activate()
SetTopMostWindow Me, True
End Sub

