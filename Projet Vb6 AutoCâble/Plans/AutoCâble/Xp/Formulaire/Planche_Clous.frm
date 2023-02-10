VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Planche_Clous 
   Caption         =   "Choix de la planche à clous :"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   OleObjectBlob   =   "Planche_Clous.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Planche_Clous"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public boolAnnuler As Boolean
Dim boolCloseForm As Boolean

Private Sub CommandButton1_Click()
boolAnnuler = False
If Trim(PlanchClous.Text) = "" Then
    MsgBox "Vous devez sélectionner une planche à clous", vbExclamation
    Me.PlanchClous.SetFocus
    Exit Sub
End If
boolCloseForm = False
Me.Hide
End Sub

Private Sub CommandButton2_Click()
boolAnnuler = True
boolCloseForm = False
Me.Hide
End Sub

Private Sub UserForm_Activate()
Dim sql As String
Dim MyPath As String
Dim Rs As Recordset
Dim MyFichier As String
Set TableauPath = funPath
PlanchClous.Clear
MyPath = TableauPath.Item("PathOutils") & "\"
If Left(MyPath, 2) <> "\\" Then MyPath = TableauPath.Item("PathServer") & MyPath & "\"



If Trim(MyPath) <> "" Then
MyFichier = Dir(MyPath & "*.dwg")
PlanchClous.AddItem ""
While MyFichier <> ""
PlanchClous.AddItem MyFichier
    MyFichier = Dir
 Wend
End If
boolCloseForm = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = boolCloseForm
End Sub
