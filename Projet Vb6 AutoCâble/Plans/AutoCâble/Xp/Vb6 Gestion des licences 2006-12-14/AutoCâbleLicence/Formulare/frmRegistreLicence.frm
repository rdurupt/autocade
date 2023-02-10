VERSION 5.00
Begin VB.Form frmRegistreLicence 
   Caption         =   "Editeur de Licence AutoCâble:"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3930
   ControlBox      =   0   'False
   Icon            =   "frmRegistreLicence.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "InFo Serveur DB"
      Height          =   1455
      Left            =   120
      TabIndex        =   24
      Top             =   7320
      Width           =   3615
      Begin VB.TextBox PassWordDb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox UserDb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "PassWord"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "User"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CheckBox ChkReg 
      Alignment       =   1  'Right Justify
      Caption         =   "Préenregistré la licence"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type de Licence"
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   5760
      Width           =   3615
      Begin VB.TextBox NbJeton 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Jeton"
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "TOUS"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Nb Jetons"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.TextBox Prix 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CheckBox ChkLicence 
      Alignment       =   1  'Right Justify
      Caption         =   "Possibilité d'Achetter la Licence"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   4284
      Width           =   3255
   End
   Begin VB.TextBox DateF 
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   3770
      Width           =   1935
   End
   Begin VB.TextBox DateD 
      BackColor       =   &H00FFFFC0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   3
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   3256
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   9120
   End
   Begin VB.TextBox User 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1714
      Width           =   1935
   End
   Begin VB.TextBox Socete 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Annuller"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   8880
      Width           =   1095
   End
   Begin VB.CommandButton Rec 
      Caption         =   "Enregistrer"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "€"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   5385
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Prix Licence:"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Date Fin:"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3770
      Width           =   1215
   End
   Begin VB.Label Serial 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2228
      Width           =   1935
   End
   Begin VB.Label PasWord 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2742
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Date Début:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3256
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   840
      Picture         =   "frmRegistreLicence.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Pass Word:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2742
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "N° Licence:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2228
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "User:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1714
      Width           =   1215
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

Private Sub ChkLicence_Click()
If Me.Prix.Enabled = True Then
    Me.Prix.Enabled = False
    Me.ChkReg.Enabled = False
    Me.ChkReg.Value = False
Else
    Me.Prix.Enabled = True
    Me.ChkReg.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Timer1.Interval = 50
Me.Timer1.Enabled = True

End Sub


Private Sub Option1_Click(Index As Integer)
Me.NbJeton.Enabled = Me.Option1(1).Value
End Sub

Private Sub Rec_Click()


If Me.Socete = "" Then
    MsgBox "Le Nom de La Société est obligatiore", vbExclamation
    Me.Socete = ""
    Me.Socete.SetFocus
    Exit Sub
End If
If Me.User = "" Then
    MsgBox "Le Nom du User est obligatiore", vbExclamation
    Me.User = ""
    Me.User.SetFocus
    Exit Sub
End If
If Me.DateD = "" Then
    MsgBox "La date de début est obligatiore", vbExclamation
    Me.DateD = ""
    Me.DateD.SetFocus
    Exit Sub
End If
If Not IsDate(Me.DateD) Then
    MsgBox "Vous devez entrer une Date au Format DD/MM/YYYY", vbExclamation
    Me.DateD = ""
    Me.DateD.SetFocus
    Exit Sub
End If
If Me.DateF = "" Then
    MsgBox "La date de Fin est obligatiore", vbExclamation
    Me.DateF = ""
    Me.DateF.SetFocus
    Exit Sub
End If
If Not IsDate(Me.DateF) Then
    MsgBox "Vous devez entrer une Date au Format DD/MM/YYYY", vbExclamation
    Me.DateF = ""
    Me.DateF.SetFocus
    Exit Sub
End If
If Me.Option1(1).Value = True And Me.NbJeton = "" Then
    MsgBox "Vous devez saisir le nombre de jentons accordé au client", vbExclamation
    Me.NbJeton = ""
    Me.NbJeton.SetFocus
    Exit Sub
End If
    FiledLicence.Count = 1
    ReDim FiledLicence.Record(FiledLicence.Count - 1)
    If Me.ChkLicence = 1 Then
        FiledLicence.General.AficheFrm = "Yes"
    Else
        FiledLicence.General.AficheFrm = "No"
    End If
    FiledLicence.General.DateDeb = Me.DateD.Text
    FiledLicence.General.DateFin = Me.DateF.Text
    FiledLicence.General.Societe = Me.Socete.Text
    FiledLicence.General.Tous = "Yes"
    If Me.Option1(0).Value = True Then
         FiledLicence.General.NbJeton = "0"
    Else
        FiledLicence.General.NbJeton = Me.NbJeton.Text
    End If
    FiledLicence.General.NbJetonActif = "0"
    If Me.ChkLicence.Value = 0 Then
        PrixV = "0"
    Else
        PrixV = Me.Prix & " €"
    End If
    If Me.ChkReg.Value = 1 Then
        FiledLicence.General.Enregistre = "Yes"
    Else
        FiledLicence.General.Enregistre = "No"
   End If
   PassDb.PassWordDb = Trim("" & Me.PassWordDb.Text)
   PassDb.UserDb = Trim("" & Me.UserDb.Text)
    FiledLicence.Record(FiledLicence.Count - 1).Serial = Me.Serial.Caption
    FiledLicence.Record(FiledLicence.Count - 1).PassWord = Me.PasWord.Caption
    FiledLicence.Record(FiledLicence.Count - 1).Useur = Me.User.Text
    CodageX.EcrirLicence App.Path & "\Licence\" & Me.Socete.Text & "\" & Me.User.Text
    MsgBox "Création de la Dll effectué."
End Sub

Private Sub Timer1_Timer()
Dim txt1 As String
Dim txt2 As String
CodageX.DefinSerialPass Left(Me.User & Me.Socete & Space(255), 8), txt1, txt2
Me.Serial = txt1
Me.PasWord = txt2
End Sub

