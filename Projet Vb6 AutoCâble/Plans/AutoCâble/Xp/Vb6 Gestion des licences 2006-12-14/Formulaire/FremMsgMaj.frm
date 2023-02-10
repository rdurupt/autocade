VERSION 5.00
Begin VB.Form FremMsgMaj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message aux utilisateurs d'AutoCâble (Lire impérativement)"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "FremMsgMaj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LstFichier 
      Columns         =   2
      Height          =   6360
      Left            =   240
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   6720
      Width           =   1695
   End
End
Attribute VB_Name = "FremMsgMaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Desactive As String
Dim Sql As String
Dim I As Long
Dim boolMsg As Boolean
Desactive = "False"
For I = 0 To Me.LstFichier.ListCount - 1
    If Me.LstFichier.Selected(I) = True Then
        If boolMsg = False Then
            boolMsg = True
            If MsgBox("Voulez vous désactiver les fichiers de votre compte: " & UserName & " " & Machine, vbYesNo) = vbYes Then
                Desactive = "True"
            End If
        End If
        Sql = "UPDATE Document SET Document.PlusAficher = " & Desactive & " "
        Sql = Sql & "WHERE Document.Machine='" & MyReplace(Machine) & "' and Document.UserName='" & MyReplace(UserName) & "' "
        Sql = Sql & "AND Document.Documment='" & MyReplace(Me.LstFichier.List(I)) & "';"
        Con.Execute Sql
        MyExecute App.Path & "\RepDoc\" & Me.LstFichier.List(I)
    End If
Next
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT Document.* FROM Document "
Sql = Sql & "WHERE  Document.Machine='" & MyReplace(Machine) & "'and  Document.UserName='" & MyReplace(UserName) & "' "
Sql = Sql & " AND Document.PlusAficher=False;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Me.LstFichier.AddItem "" & Rs!Documment
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
End Sub
