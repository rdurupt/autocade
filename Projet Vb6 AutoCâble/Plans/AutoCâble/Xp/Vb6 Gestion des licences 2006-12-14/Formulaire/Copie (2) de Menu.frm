VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Menu principal"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   OleObjectBlob   =   "Menu.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean

Private Sub CommandButton1_Click()
Set FormBarGrah = Me
  Me.Frame4.Enabled = False
MenuShow = True

boolCreationPlan = True
SubCreer
Me.Frame4.Enabled = True

End Sub

Private Sub CommandButton10_Click()
Utilitaire

End Sub

Private Sub CommandButton11_Click()
Set FormBarGrah = Me
NomenclatureOk = False

MenuShow = True
Me.Frame4.Enabled = False
subExporter
    Me.Frame4.Enabled = True
End Sub

Private Sub CommandButton12_Click()
subImport
End Sub

Private Sub CommandButton13_Click()
subExport
End Sub

Private Sub CommandButton2_Click()

NoClose = False
 AutoApp.Quit
 Set AutoApp = Nothing
End
End Sub



Private Sub CommandButton3_Click()
ImportXls.Show vbModal
Unload ImportXls
End Sub

Private Sub CommandButton4_Click()
Set FormBarGrah = Me
Me.Frame4.Enabled = False
SubExportXls
Me.Frame4.Enabled = True
End Sub

Private Sub CommandButton5_Click()
subVerifierEtude

End Sub

Private Sub CommandButton6_Click()
Set FormBarGrah = Me


MenuShow = True
Me.Frame4.Enabled = False
subEDITER
    Me.Frame4.Enabled = True



End Sub

Private Sub CommandButton7_Click()
Set FormBarGrah = Me
subModifierCartouche

End Sub

Private Sub CommandButton8_Click()

subUtilisateur

End Sub

Private Sub CommandButton9_Click()
subApprobation

End Sub

Private Sub UserForm_Activate()


NoClose = True
  
End Sub

Private Sub UserForm_Initialize()
'ChDir "C:\Program Files\AutoCAD 2002 Fra\"
NomenclatureOk = True
Set AutoApp = New AutoCAD.AcadApplication
'AutoApp = GetObject("", AutoCAD.AcadApplication)
AutoApp.Documents(0).Close False

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
'
'Sql = "SELECT T_indiceProjet.Id, T_indiceProjet.IdProjet, T_indiceProjet.Indice, T_indiceProjet.Description, T_indiceProjet.Li, T_indiceProjet.IdStatus, T_indiceProjet.IdApprobateur, T_indiceProjet.PlAutoCadSave, T_indiceProjet.AutoCadSave "
'Sql = Sql & "FROM T_indiceProjet "
'Sql = Sql & "WHERE T_indiceProjet.Id=" & EDITER.lstIndice.List(EDITER.lstIndice.ListIndex, 1) & ";"
'Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
'Sql = "INSERT INTO T_indiceProjet ( IdProjet, Indice, Description, Li, IdStatus, IdApprobateur, AutoCadSave ) "
'Sql = Sql & "VALUES(" & Rs!IdProjet & ", '" & Rs!Indice & "','" & Rs!Description & "','" & LI & "', 2," & Rs!IdApprobateur & ",'" & Rs!PlAutoCadSave & "') "
'Sql = Sql & "WHERE T_indiceProjet.Id=37;"
'
'
'End If
'
'End Sub

