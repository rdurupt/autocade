VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu principal"
   ClientHeight    =   7470
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11625
   Icon            =   "Menu.dsx":0000
   MaxButton       =   0   'False
   OleObjectBlob   =   "Menu.dsx":0A7A
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean

Private Sub CommandButton1_Click()

boolCreationPlan = True
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton1") = True Then SubCreer CommandButton1.Caption
Me.Frame4.Enabled = True
MenuShow = False

End Sub

Private Sub CommandButton10_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton10") = True Then Utilitaire
Me.Frame4.Enabled = True
MenuShow = False

End Sub

Private Sub CommandButton11_Click()
NomenclatureOk = False
NotSaveRacourci = True
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton11") = True Then subExporter
Me.Frame4.Enabled = True
MenuShow = False

    
End Sub

Private Sub CommandButton12_Click()
XlsPrix = "CablePrix"
subImport
End Sub

Private Sub CommandButton13_Click()
XlsPrix = "CablePrix"
subExport
End Sub

Private Sub CommandButton14_Click()
If boolAutoCAD = False Then
    MsgBox "Vous ne possédez pas de licence AutoCad." & vbCrLf & "Vous ne pouvez pas effectuer ce test."
Else
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton14") = True Then LireRepEval
Me.Frame4.Enabled = True
MenuShow = False

    
End If
End Sub

Private Sub CommandButton15_Click()
XlsPrix = "HabillagePrix"
subImport
End Sub

Private Sub CommandButton16_Click()
XlsPrix = "HabillagePrix"
subExport

End Sub

Private Sub CommandButton17_Click()
subExporterSynthese
End Sub

Private Sub CommandButton18_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton18") = True Then AddAtrib
Me.Frame4.Enabled = True
MenuShow = False

End Sub

Private Sub CommandButton19_Click()
 Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton19") = True Then subTestWord
Me.Frame4.Enabled = True
MenuShow = False

End Sub

Private Sub CommandButton2_Click()

CodageX.DcrJenton
LoadDb
Con.Execute "DELETE [Utilise_Par].* FROM [Utilise_Par] WHERE [Utilise_Par].Machine='" & MyReplace(Machine) & "' and [Utilise_Par].User='" & MyReplace(UserName) & "';"


'Con.CloseConnection

NoClose = False
If boolAutoCAD = True Then
 AutoApp.Quit
 End If
 Set AutoApp = Nothing
 
Unload Me
On Error Resume Next
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

Private Sub CommandButton20_Click()
'Set FormBarGrah = Me
'MenuShow = True
'Me.Frame4.Enabled = False
'If GestionDesDroit(CommandButton20.Caption) = True Then
'    FrmIndice.Show vbmodal
'    Unload FrmIndice
'End If
'Me.Frame4.Enabled = True
'MenuShow = False
LstConecteur
'PreparationNomenclatuer 745
'frmEditClip.Show vbmodal
'GenairEtiquette2 745, "BVM", CheckBox1.Value, CheckBox2.Value
  boolExec = True
'LoadDb
'Generer_NomenclatuerFinal 745
'frm_Etiquettes_Serie.Show vbModal
boolExec = False

End Sub

Private Sub CommandButton21_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton21") = True Then subJob
Me.Frame4.Enabled = True
MenuShow = False


End Sub

Private Sub CommandButton22_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton22") = True Then subGenerateur
Me.Frame4.Enabled = True
MenuShow = False


End Sub


Private Sub CommandButton26_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton26") = True Then subSuperUtilisateur
Me.Frame4.Enabled = True
MenuShow = False
End Sub

Private Sub CommandButton27_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton27") = True Then MenuMagasin.Show vbModal

Me.Frame4.Enabled = True
MenuShow = False
End Sub

Private Sub CommandButton28_Click()
FrmSynthese.Show vbModal
'
'Dim Rs As Recordset
'Set Rs = Con.OpenRecordSet("SELECT T_Serveur_Smtp.* FROM T_Serveur_Smtp;")
'
'
'MailEnvoi Rs!SMTP, Rs!Authentification, Rs!Utilisatuer, Rs!PassWord, Rs!Port, 15, Rs!Messagerie, "robert.durupt@encelade.fr", "robert.durupt@encelade.fr", "aaaa", "Test Sur Une Autre Mahine", ""

End Sub

Private Sub CommandButton29_Click()
Set FormBarGrah = Me

MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton29") = True Then
    
    subEDITER CommandButton29.Caption, True
End If
Me.Frame4.Enabled = True

MenuShow = False
BooolBloque = False
End Sub

Private Sub CommandButton30_Click()
Set FormBarGrah = Me
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton30") = True Then subGenerateurEtiquette
Me.Frame4.Enabled = True
MenuShow = False



End Sub

Private Sub CommandButton31_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton31") = True Then Synthese
Me.Frame4.Enabled = True
MenuShow = False
End Sub

Private Sub CommandButton5_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton5") = True Then subVerifierEtude
Me.Frame4.Enabled = True
MenuShow = False


End Sub

Private Sub CommandButton6_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton6") = True Then subEDITER CommandButton1.Caption, False
Me.Frame4.Enabled = True
MenuShow = False
End Sub

Private Sub CommandButton7_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton7") = True Then subModifierCartouche
Me.Frame4.Enabled = True
MenuShow = False


End Sub

Private Sub CommandButton8_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton8") = True Then subUtilisateur
Me.Frame4.Enabled = True
MenuShow = False


End Sub

Private Sub CommandButton9_Click()
Set FormBarGrah = Me
MenuShow = True
Me.Frame4.Enabled = False
If GestionDesDroit("CommandButton9") = True Then subApprobation
Me.Frame4.Enabled = True
MenuShow = False



End Sub

Private Sub UserForm_Activate()
Dim MyControl As New Collection
Dim Rs As Recordset
Dim I As Long
Dim Sql As String
For I = 0 To Me.Controls.Count - 1
    MyControl.Add I, Me.Controls(I).Name
Next

Sql = "UPDATE T_indiceProjet SET T_indiceProjet.UserName = Null "
Sql = Sql & "WHERE T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "';"
Con.Execute Sql
'UPDATE T_indiceProjet SET T_indiceProjet.UserName = Null
'WHERE (((T_indiceProjet.Id)=94));

Set Rs = Con.OpenRecordSet("SELECT T_Boutons.Bouton, T_Boutons.Name FROM T_Boutons where T_Boutons.ContonTotal=false ;")
While Rs.EOF = False
    Me.Controls(MyControl(Rs!Name)).Caption = Trim("" & Rs!Bouton)
    Rs.MoveNext
Wend

Unload Modifier
NotSaveRacourci = True
Bool_Fichier_Li = False
NoClose = True
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
DeIconify Me
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
'ChDir "C:\Program Files\AutoCAD 2002 Fra\"
NomenclatureOk = True
If IsCilent = False Then
    boolAutoCAD = True
    If MsgBox("Voulez vous ouvrir une licence AUTOCAD.", vbQuestion + vbYesNo) = vbYes Then
       SetAutocad
        If Err = 0 Then
            AutoApp.Visible = True
            AutoApp.Documents(0).Close False
            DoEvents
            
        Else
            MsgBox "Plus de licence Autocad disponible", vbInformation, "AutoCâble  licence :"
            boolAutoCAD = False
        End If
    Else
        boolAutoCAD = False
    End If
End If


End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
 Dim hProcess As Long
    If Not Visible Then
        Select Case x \ Screen.TwipsPerPixelX
            Case WM_MOUSEMOVE
            Case WM_LBUTTONDOWN
            Case WM_LBUTTONUP
            Case WM_LBUTTONDBLCLK
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
                DeIconify Me
'                Me.StartUpPosition = vbStartUpScreen
            Case WM_RBUTTONDOWN
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
'                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
            Case WM_RBUTTONDBLCLK
        End Select
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NoClose

End Sub

Private Sub UserForm_Resize()
'If WindowState = vbMinimized Then
'        Iconify Me, MainTitle
'        DoEvents
'        Exit Sub
'    End If
End Sub

Private Sub UserForm_Terminate()

boolExec = False
funCloseConnextion
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

