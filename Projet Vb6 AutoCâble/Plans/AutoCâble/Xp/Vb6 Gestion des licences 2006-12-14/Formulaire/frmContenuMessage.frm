VERSION 5.00
Begin VB.Form frmContenuMessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enregistrer"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Body 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox Sujet 
      Height          =   285
      Left            =   120
      MaxLength       =   254
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Message"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sujet"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmContenuMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub chargement(Id As Long)
Dim Rs As Recordset
Dim Sql As String
Me.Tag = Id
Sql = "SELECT T_Message_Mail.* FROM T_Message_Mail "
Sql = Sql & "WHERE T_Message_Mail.Id=" & Id & ";"
Set Rs = Con.OpenRecordSet(Sql)
Me.Caption = "Message : " & Rs!Routine
Me.Sujet = Trim("" & Rs!Sujet)
Me.Body = Trim("" & Rs!Body)
Set Rs = Con.CloseRecordSet(Rs)
Me.Show vbModal
End Sub

Private Sub Command1_Click()
Dim Sql As String

Sql = "UPDATE T_Message_Mail SET "
Sql = Sql & "T_Message_Mail.Sujet = '" & MyReplace(Sujet) & "', "
Sql = Sql & "T_Message_Mail.Body = '" & MyReplace(Body) & "' "
Sql = Sql & "WHERE T_Message_Mail.Id=" & Me.Tag & ";"
Con.Execute Sql
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub
