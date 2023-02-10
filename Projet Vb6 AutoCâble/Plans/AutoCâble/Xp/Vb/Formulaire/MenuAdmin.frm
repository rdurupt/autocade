VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuAdmin 
   Caption         =   "Menu Adim :"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   Icon            =   "MenuAdmin.dsx":0000
   OleObjectBlob   =   "MenuAdmin.dsx":08CA
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MenuAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoFermer As Boolean
Private Sub CommandButton1_Click()
ModifierUser.Show vbModal
End Sub

Private Sub CommandButton10_Click()
Set FormBarGrah = Me
MousePointer = fmMousePointerHourGlass
UserForm6.Chargement
Unload UserForm6
MousePointer = fmMousePointerDefault
End Sub

Private Sub CommandButton11_Click()
UtilitairesListes.Show vbModal
End Sub

Private Sub CommandButton12_Click()
    FrmHabillage.Chargement
End Sub

Private Sub CommandButton13_Click()
Liste_Projets.Show vbModal
End Sub

Private Sub CommandButton2_Click()
NoFermer = False
Me.Hide
End Sub

Private Sub CommandButton3_Click()
UserForm3.Show vbModal
End Sub

Private Sub CommandButton4_Click()

UserForm1.Charger txt1, " ", "Equipement:", " "
Unload UserForm1
txt1 = ""
End Sub

Private Sub CommandButton5_Click()
UserForm1.Charger txt1, " ", "Vagues:", " "
Unload UserForm1
txt1 = ""
End Sub

Private Sub CommandButton6_Click()
UserForm1.Charger txt1, " ", "Ensemble:", " "
Unload UserForm1
txt1 = ""
End Sub

Private Sub CommandButton7_Click()
 RepSystem.Show vbModal
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
txt1 = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NoFermer
End Sub

