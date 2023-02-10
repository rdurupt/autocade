VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Menu principal"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   OleObjectBlob   =   "Menu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Noquite As Boolean

Private Sub CommandButton1_Click()
Me.Frame4.Enabled = False
boolCreationPlan = True
ImportXLS.Show
Unload ImportXLS
Me.Frame4.Enabled = True
End Sub

Private Sub CommandButton2_Click()
NoClose = False
Unload Me
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

End Sub

Private Sub CommandButton3_Click()
ImportXLS.Show
End Sub

Private Sub CommandButton4_Click()
Me.Frame4.Enabled = False
ExportXls.Show
Me.Frame4.Enabled = True
End Sub

Private Sub CommandButton5_Click()
Me.Frame4.Enabled = False
ImportDwg.Show
Me.Frame4.Enabled = True
End Sub

Private Sub CommandButton6_Click()
MenuShow = False
Me.Frame4.Enabled = False
EDITER.Show
Unload EDITER
Me.Frame4.Enabled = True
If MenuShow = True Then Me.Show
MenuShow = False

End Sub

Private Sub CommandButton7_Click()
Me.Frame4.Enabled = False
boolCreationPlan = True
EDITER.Charger "Modifier un plan :"
Me.Frame4.Enabled = True
Unload EDITER

End Sub

Private Sub Frame4_Click()

End Sub

Private Sub ProgressBar1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

End Sub

Private Sub UserForm_Activate()
NoClose = True
DbNumPlan = CherCheInFihier("Bdnumero")
 db = CherCheInFihier("BdAutocable")

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NoClose
End Sub

Private Sub UserForm_Terminate()

boolExec = False
End Sub
'Sub Modifier()
'Dim Sql As String
'Dim Rs As Recordset
'Con.OpenConnetion db
'Sql = "SELECT T_indiceProjet.Id, T_indiceProjet.IdProjet, T_indiceProjet.Indice, T_indiceProjet.Description, T_indiceProjet.Li, T_indiceProjet.IdStatus, T_indiceProjet.IdApprobateur, T_indiceProjet.AutoCadSaveAs, T_indiceProjet.AutoCadSave "
'Sql = Sql & "FROM T_indiceProjet "
'Sql = Sql & "WHERE T_indiceProjet.Id=" & EDITER.lstIndice.List(EDITER.lstIndice.ListIndex, 1) & ";"
'Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
'Sql = "INSERT INTO T_indiceProjet ( IdProjet, Indice, Description, Li, IdStatus, IdApprobateur, AutoCadSave ) "
'Sql = Sql & "VALUES(" & Rs!IdProjet & ", '" & Rs!Indice & "','" & Rs!Description & "','" & LI & "', 2," & Rs!IdApprobateur & ",'" & Rs!AutoCadSaveAs & "') "
'Sql = Sql & "WHERE T_indiceProjet.Id=37;"
'
'
'End If
'Con.CloseConnection
'End Sub

