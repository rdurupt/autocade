VERSION 5.00
Begin VB.Form frmPOP3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cofuguration Serveur SMTP:"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Activer 
      Alignment       =   1  'Right Justify
      Caption         =   "Serveur Activé"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   615
      Left            =   2520
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enregistrer"
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CheckBox Authentification 
      Alignment       =   1  'Right Justify
      Caption         =   "Authentification"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox PassWord 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Messagerie 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Utilisatuer 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Port 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox SMTP 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Mot de passe"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Messagerie"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Utilisateur"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Sur le Port"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Server SMTP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmPOP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim sql As String
Dim Rs As Recordset
Dim Port2 As Integer
If Trim("" & SMTP) = "" Then
    MsgBox "Vous devez obligatoirement saisir le nom du server."
    SMTP.SetFocus
    Exit Sub
End If

If Val(Port) = 0 Then
    MsgBox "Vous devez obligatoirement saisir le N° du Port au farmat numérique."
    Port.SetFocus
    Exit Sub
Else
    Port2 = Val(Port)
    If Port2 <> Val(Port) Then
        MsgBox "Vous devez obligatoirement saisir le N° du Port au farmat numérique (Sans Virgule)."
        Port.SetFocus
        Exit Sub
    End If
End If
If Trim("" & Utilisatuer) = "" Then
    MsgBox "Vous devez obligatoirement saisir le nom de l'Utilisatuer."
    Utilisatuer.SetFocus
    Exit Sub
End If

If Trim("" & Messagerie) = "" Then
    MsgBox "Vous devez obligatoirement saisir l'address de Messagerie."
    Messagerie.SetFocus
    Exit Sub
End If
If Trim("" & PassWord) = "" Then
    MsgBox "Vous devez obligatoirement saisir le Mot de passe."
    PassWord.SetFocus
    Exit Sub
End If



sql = "SELECT T_Serveur_Smtp.* FROM T_Serveur_Smtp;"
Set Rs = Con.OpenRecordSet(sql)
Rs!SMTP = SMTP
Rs!Port = Port
Rs!Utilisatuer = Utilisatuer
Rs!Messagerie = Messagerie
Rs!PassWord = PassWord
If Authentification.Value = 1 Then
    Rs!Authentification = True
Else
   Rs!Authentification = False
End If

If Activer.Value = 1 Then
    Rs!Activer = True
Else
    Rs!Activer = False
End If
Rs.Update
Set Rs = Con.CloseRecordSet(Rs)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim sql As String
Dim Rs As Recordset
sql = "SELECT T_Serveur_Smtp.* FROM T_Serveur_Smtp;"
Set Rs = Con.OpenRecordSet(sql)
Me.Tag = Rs!Id
Me.SMTP = Rs!SMTP
Port = Rs!Port
Utilisatuer = Rs!Utilisatuer
Messagerie = Rs!Messagerie
PassWord = Rs!PassWord
If Rs!Authentification = True Then
    Authentification.Value = 1
End If
If Rs!Activer = True Then
    Activer.Value = 1
End If
Set Rs = Con.CloseRecordSet(Rs)
End Sub

