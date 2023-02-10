VERSION 5.00
Begin VB.Form frmRegistreLicence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enregistre Licence:"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Annuller"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Rec 
      Caption         =   "&Enregistrer"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox PasWord 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Serial 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2235
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   840
      Picture         =   "frmRegistreLicence.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label User 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Pass Word:"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "N° Licence:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2235
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "User:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label Societe 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Société:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmRegistreLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Societe.Caption = FiledLicence.General.Societe
Me.User = FiledLicence.Record(FiledLicence.Count - 1).UserName

End Sub

Private Sub Rec_Click()


If Me.Serial <> FiledLicence.Record(FiledLicence.Count - 1).Serial Then
    MsgBox "Le N° de Licence ne corespond pas à la valeur attendue", vbExclamation
    Me.Serial = ""
    Me.Serial.SetFocus
    Exit Sub
End If
If Me.PasWord <> FiledLicence.Record(FiledLicence.Count - 1).PassWord Then
     MsgBox "Le Pass Word ne corespond pas à la valeur attendue", vbExclamation
    Me.PasWord = ""
    Me.PasWord.SetFocus
    Exit Sub
End If
    FiledLicence.General.Enregistre = "Yes"
    Unload Me
End Sub
