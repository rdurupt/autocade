VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CARTOUCHE_RENAULT 
   Caption         =   "CARTOUCHE  RENAULT:"
   ClientHeight    =   6420
   ClientLeft      =   2670
   ClientTop       =   330
   ClientWidth     =   6945
   OleObjectBlob   =   "CARTOUCHE_RENAULT.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "CARTOUCHE_RENAULT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim NoClose As Boolean

Private Sub Precedant_Click()

Unload Me
End Sub

Private Sub UserForm_Activate()
VarPreced = True
Me.txt6 = CartoucheEncelade.txt14
Me.txt2 = CartoucheEncelade.txt5
Me.txt12 = CartoucheEncelade.txt0
Me.txt13 = CartoucheEncelade.txt13
Me.txt9 = CartoucheEncelade.txt17
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NoClose

End Sub

Private Sub Valider_Click()
If ValideChampsTexte(Me, 14) = False Then Exit Sub
NoClose = False
NbContolClient = 14
VarPreced = False





Me.Hide
End Sub

