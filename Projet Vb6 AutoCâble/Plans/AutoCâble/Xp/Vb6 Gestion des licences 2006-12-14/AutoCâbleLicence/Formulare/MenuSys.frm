VERSION 5.00
Begin VB.Form MenuSys 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renommer barre de Menu :"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "MenuSys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enregistrer"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   8040
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6855
      LargeChange     =   3220
      Left            =   4080
      SmallChange     =   322
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3735
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   322
         Left            =   0
         ScaleHeight     =   3150
         ScaleMode       =   0  'User
         ScaleWidth      =   3660
         TabIndex        =   2
         Top             =   0
         Width           =   3655
         Begin VB.TextBox Boutons 
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   3615
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Boutons"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "MenuSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L As Long

Private Sub Command1_Click()
Dim Sql As String
Dim I As Long
For I = 0 To L - 1
    If Trim("" & Boutons(I)) = "" Then
        MsgBox "Valeur obligatoir", vbCritical
        Boutons(I).SetFocus
        Exit Sub
    End If
    Sql = "UPDATE T_Boutons SET T_Boutons.Bouton = '" & MyReplace(Boutons(I)) & "' WHERE T_Boutons.Id=" & Boutons(I).Tag & ";"
    Con.Execute Sql
Next
MajDroitsFrm Id_Users
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT T_Boutons.* FROM T_Boutons ORDER BY T_Boutons.Ordre;"
Set Rs = Con.OpenRecordSet(Sql)

While Rs.EOF = False
If L <> 0 Then
    Load Boutons(L)
    Boutons(L).Top = Boutons(L - 1).Top + Boutons(L).Height
    Me.Picture1.Height = 322 + (322 * L)
End If
Boutons(L).Visible = True
Boutons(L) = Trim("" & Rs!Bouton)
Boutons(L).Tag = Rs!Id
L = L + 1
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
If L > 23 Then VScroll1.Visible = True
End Sub

Private Sub Form_Load()
L = 0
End Sub

Private Sub VScroll1_Change()
Me.Picture1.Top = VScroll1 * -1

End Sub
