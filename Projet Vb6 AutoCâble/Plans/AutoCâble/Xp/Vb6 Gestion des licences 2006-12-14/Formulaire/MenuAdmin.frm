VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MenuAdmin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu Admin :"
   ClientHeight    =   10845
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   6165
   Icon            =   "MenuAdmin.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "MenuAdmin.dsx":08CA
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MenuAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoFermer As Boolean
Private Sub CommandButton1_Click()
EditUser.Show vbModal
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

Private Sub CommandButton14_Click()
FrmEtats.Show vbModal
End Sub

Private Sub CommandButton15_Click()
MenuSys.Show vbModal
End Sub

Private Sub CommandButton16_Click()
frmPOP3.Show vbModal
End Sub

Private Sub CommandButton17_Click()
FrmMesageDroits.Show vbModal
End Sub

Private Sub CommandButton18_Click()
EditGroupe.Show vbModal
End Sub

Private Sub CommandButton2_Click()
NoFermer = False
Admin = False
Me.Hide
End Sub

Private Sub CommandButton3_Click()
UserForm3.Show vbModal
End Sub

Private Sub CommandButton4_Click()

UserForm1.charger txt1, " ", "Equipement:", " "
Unload UserForm1
txt1 = ""
End Sub

Private Sub CommandButton5_Click()
UserForm1.charger txt1, " ", "Vagues:", " "
Unload UserForm1
txt1 = ""
End Sub

Private Sub CommandButton6_Click()
UserForm1.charger txt1, " ", "Ensemble:", " "
Unload UserForm1
txt1 = ""
End Sub

Private Sub CommandButton7_Click()
 RepSystem.Show vbModal
End Sub

Private Sub CommandButton8_Click()
UserForm4.Charge Me, "IdStatus=3 or (VerifieDate= Null and IdStatus<>4)"
Unload UserForm4
End Sub

Private Sub CommandButton9_Click()
UserForm5.Charge Me, "IdStatus=4"
Unload UserForm5

End Sub

Private Sub UserForm_Initialize()
 NoFermer = True
Admin = True
txt1 = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NoFermer
End Sub

