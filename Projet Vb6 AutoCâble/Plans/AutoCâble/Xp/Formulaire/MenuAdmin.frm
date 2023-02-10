VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuAdmin 
   Caption         =   "Menu Adim :"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   OleObjectBlob   =   "MenuAdmin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MenuAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoFermer As Boolean
Private Sub CommandButton1_Click()
ModifierUser.Show
End Sub

Private Sub CommandButton10_Click()
Set FormBarGrah = Me
MousePointer = fmMousePointerHourGlass
UserForm6.chargement
Unload UserForm6
MousePointer = fmMousePointerDefault
End Sub

Private Sub CommandButton2_Click()
NoFermer = False
Me.Hide
End Sub

Private Sub CommandButton3_Click()
UserForm3.Show
End Sub

Private Sub CommandButton4_Click()
UserForm1.Charger txt1, " ", "Equipement:", " "
Unload UserForm1
End Sub

Private Sub CommandButton5_Click()
UserForm1.Charger txt1, " ", "Vagues:", " "
Unload UserForm1
End Sub

Private Sub CommandButton6_Click()
UserForm1.Charger txt1, " ", "Ensemble:", " "
Unload UserForm1
End Sub

Private Sub CommandButton7_Click()
RepSystem.Show
End Sub

Private Sub CommandButton8_Click()
UserForm4.Charge Me, "IdStatus=3 or VerifieDate= Null"
Unload UserForm4
End Sub

Private Sub CommandButton9_Click()
UserForm5.Charge Me
Unload UserForm5

End Sub

Private Sub UserForm_Activate()
NoFermer = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NoFermer
End Sub

