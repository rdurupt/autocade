VERSION 5.00
Object = "{50299AB4-73AA-4780-89B3-BF90895272BB}#1.0#0"; "ChercheOcx.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RechercheOxx2.RecherAutocable RecherAutocable1 
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
      Database        =   "\\Autocable\Autocable Access\AutoCable.mdb"
      Filtre          =   "VerifieDate<> null and Archiver=false and IdStatus<4"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'Me.RecherAutocable1.Charge "VerifieDate<> null and Archiver=false and IdStatus<4 ", "\\Autocable\Autocable Access\AutoCable.mdb"
End Sub

Private Sub RecherAutocable1_Actueliser()

End Sub

Private Sub RecherAutocable1_Change()

End Sub

Private Sub RecherAutocable1_Action(Tableau_Valeur As Variant, Annuler As Variant)

End Sub
