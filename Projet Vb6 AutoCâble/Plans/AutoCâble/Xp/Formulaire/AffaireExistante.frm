VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AffaireExistante 
   Caption         =   "Affaire Existante"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   OleObjectBlob   =   "AffaireExistante.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AffaireExistante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Annuler As Boolean
Private Sub CommandButton1_Click()
Dim sql As String
Dim Rs As Recordset
CherchPices.Charge Me, " LiAutoCadSave <>  Null and NbErr=0", False, True
Unload CherchPices
sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"

Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
    If Rs!Pere > 0 Then Me.Tag = Rs!Pere
Else
    If Rs.EOF = True Then
    sql = "SELECT Archive_T_indiceProjet.Pere FROM Archive_T_indiceProjet "
    sql = sql & "WHERE Archive_T_indiceProjet.Id=" & Me.txt3.Tag & ";"

    Set Rs = Con.OpenRecordSet(sql)
        If Rs!Pere > 0 Then Me.Tag = Rs!Pere
    End If
End If
End Sub

Private Sub CommandButton2_Click()
Dim Piece As Long
Dim pathTmpXls As String
Dim sql As String
Dim Rs As Recordset


Dim Fso As New FileSystemObject
Set FormBarGrah = Me
If Trim("" & Me.txt3.Tag) = "" Then
    MsgBox "Le champ Pièce est obligatoire"
    CommandButton1_Click
    Exit Sub
End If
Annuler = False
Me.Hide
End Sub

Private Sub CommandButton3_Click()
Me.Hide
End Sub

Private Sub Label35_Click()

End Sub

Private Sub UserForm_Activate()
Annuler = True
End Sub

