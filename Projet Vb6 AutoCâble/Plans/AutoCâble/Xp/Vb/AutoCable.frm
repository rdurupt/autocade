VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form UserForm2 
   ClientHeight    =   11685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18240
   Icon            =   "AutoCable.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11685
   ScaleWidth      =   18240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   16200
      TabIndex        =   26
      Top             =   -70
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   525
         Left            =   45
         Picture         =   "AutoCable.frx":08CA
         Stretch         =   -1  'True
         Top             =   150
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Annuler"
      Height          =   435
      Left            =   12480
      TabIndex        =   3
      Top             =   11160
      Width           =   2130
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Valider"
      Height          =   435
      Left            =   8360
      TabIndex        =   2
      Top             =   11160
      Width           =   2130
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualiser / Valider"
      Height          =   435
      Left            =   4240
      TabIndex        =   1
      Top             =   11160
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualiser"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   11160
      Width           =   2130
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10245
      Left            =   0
      TabIndex        =   4
      Top             =   410
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   18071
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Critères"
      TabPicture(0)   =   "AutoCable.frx":4C71
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Spreadsheet5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Connecteurs"
      TabPicture(1)   =   "AutoCable.frx":4C8D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Spreadsheet1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tableau de fils"
      TabPicture(2)   =   "AutoCable.frx":4CA9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Spreadsheet2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Composants"
      TabPicture(3)   =   "AutoCable.frx":4CC5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Spreadsheet3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Notas"
      TabPicture(4)   =   "AutoCable.frx":4CE1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Spreadsheet4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Noeuds"
      TabPicture(5)   =   "AutoCable.frx":4CFD
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Fleche_Droite"
      Tab(5).Control(1)=   "Long_C"
      Tab(5).Control(2)=   "TORON_P"
      Tab(5).Control(3)=   "Spreadsheet6"
      Tab(5).Control(4)=   "CLASSE_T"
      Tab(5).Control(5)=   "DIAMETRE"
      Tab(5).Control(6)=   "ACTIVER"
      Tab(5).Control(7)=   "Command7"
      Tab(5).Control(8)=   "Command6"
      Tab(5).Control(9)=   "Command5"
      Tab(5).Control(10)=   "ENC"
      Tab(5).Control(11)=   "PSA"
      Tab(5).Control(12)=   "RSA"
      Tab(5).Control(13)=   "Hab"
      Tab(5).Control(14)=   "Longueur"
      Tab(5).Control(15)=   "Label9"
      Tab(5).Control(16)=   "Label8"
      Tab(5).Control(17)=   "Label7"
      Tab(5).Control(18)=   "NOUED"
      Tab(5).Control(19)=   "Label6"
      Tab(5).Control(20)=   "Label5"
      Tab(5).Control(21)=   "Label4"
      Tab(5).Control(22)=   "Label3"
      Tab(5).Control(23)=   "Label2"
      Tab(5).Control(24)=   "Label1"
      Tab(5).ControlCount=   25
      Begin VB.CheckBox Fleche_Droite 
         Alignment       =   1  'Right Justify
         Caption         =   "Fleche D"
         Height          =   255
         Left            =   -74400
         TabIndex        =   35
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Long_C 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -64400
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox TORON_P 
         Alignment       =   1  'Right Justify
         Caption         =   "TORON/P"
         Height          =   255
         Left            =   -72960
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin OWC.Spreadsheet Spreadsheet6 
         Height          =   8490
         Left            =   -74760
         TabIndex        =   10
         Top             =   1320
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":4D19
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin VB.TextBox CLASSE_T 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -59520
         TabIndex        =   30
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox DIAMETRE 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -62065
         TabIndex        =   28
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox ACTIVER 
         Alignment       =   1  'Right Justify
         Caption         =   "ACTIVER"
         Height          =   255
         Left            =   -72960
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin OWC.Spreadsheet Spreadsheet3 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   6
         Top             =   345
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":56F1
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin OWC.Spreadsheet Spreadsheet5 
         Height          =   9720
         Left            =   120
         TabIndex        =   9
         Top             =   345
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":5E03
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   -74340
         Picture         =   "AutoCable.frx":65D7
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Modifier"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   -73800
         Picture         =   "AutoCable.frx":6D85
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Supprimer"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   -74880
         Picture         =   "AutoCable.frx":757F
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ajouter"
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox ENC 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         ItemData        =   "AutoCable.frx":7DF5
         Left            =   -59520
         List            =   "AutoCable.frx":7DF7
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox PSA 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -62785
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox RSA 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -65793
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox Hab 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         ItemData        =   "AutoCable.frx":7DF9
         Left            =   -70136
         List            =   "AutoCable.frx":7DFB
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Longueur 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -67080
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin OWC.Spreadsheet Spreadsheet4 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   5
         Top             =   345
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":7DFD
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin OWC.Spreadsheet Spreadsheet2 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   7
         Top             =   345
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":850A
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin OWC.Spreadsheet Spreadsheet1 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   8
         Top             =   345
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":8C16
         DataType        =   "HTMLDATA"
         AutoFit         =   -1  'True
         DisplayColHeaders=   -1  'True
         DisplayGridlines=   -1  'True
         DisplayHorizontalScrollBar=   -1  'True
         DisplayRowHeaders=   -1  'True
         DisplayTitleBar =   -1  'True
         DisplayToolbar  =   -1  'True
         DisplayVerticalScrollBar=   -1  'True
         EnableAutoCalculate=   -1  'True
         EnableEvents    =   -1  'True
         MoveAfterReturn =   -1  'True
         MoveAfterReturnDirection=   0
         RightToLeft     =   0   'False
         ViewableRange   =   "1:65536"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "LONGUEUR/C"
         Height          =   195
         Left            =   -65730
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CLASSE_T"
         Height          =   195
         Left            =   -60470
         TabIndex        =   31
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DIAMETRE"
         Height          =   195
         Left            =   -63045
         TabIndex        =   29
         Top             =   480
         Width           =   840
      End
      Begin VB.Label NOUED 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -70810
         TabIndex        =   17
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CODE.ENC."
         Height          =   195
         Left            =   -60861
         TabIndex        =   16
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CODE.PSA."
         Height          =   195
         Left            =   -64109
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CODE.RSA."
         Height          =   195
         Left            =   -67132
         TabIndex        =   14
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DESIGN.HAB."
         Height          =   195
         Left            =   -71640
         TabIndex        =   13
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "LONGUEUR"
         Height          =   195
         Left            =   -68135
         TabIndex        =   12
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOEUDS"
         Height          =   195
         Left            =   -71640
         TabIndex        =   11
         Top             =   480
         Width           =   690
      End
   End
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim NoMacro As Boolean
Dim Nouveau As Boolean
Public boolExcute As Boolean
Dim NotSortie As Boolean
Dim MyClient As String
Dim Msg As String
Dim MyErr As Boolean
Dim IfValidationOk As Boolean
Dim NbFinOuiNon As Long
Dim NoMacro1 As Boolean
Dim NoMacro2 As Boolean
Dim NoMacro3 As Boolean
Dim NoMacro4 As Boolean
Dim NoMacro5 As Boolean
Dim NoMacro6 As Boolean
Dim boolSelctChange As Boolean
Dim boolMajListe As Boolean
Dim MyTableENC() As String
Dim MyTablePSA() As String
Dim MyTableRSA() As String
Dim MyTableHab() As String
Dim NoMaj As Boolean
Dim MyCollectionENC As New Collection
Dim MyCollectionPSA As New Collection
Dim MyCollectionRSA As New Collection
Dim MyCollectionHab As New Collection
Dim MyCollectionLienHab As New Collection
Dim bool_Activate As Boolean
Dim boolActu As Boolean







Private Sub Command3_Click()
If boolActu = False Then
    MsgBox "Il est impossible de valide l'étude si un test de d'actualisation na pas été effectué."
    Exit Sub
End If
If Trim(Msg) <> "" Then
    MsgBox "Il est impossible de valide l'étude si le test de validation présente des erreurs."
    Exit Sub
End If

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim MyRange2
Set MyEcel = New EXCEL.Application
Dim MyTim
Dim BoolErr As Boolean
BoolErr = False
Dim Fso As New FileSystemObject
   
    
'MyEcel.Visible = True
If Nouveau = False Then
    Set MyWorkbook = MyEcel.Workbooks.Open(Me.Caption)
Else
    If Fso.FileExists(Me.Caption) Then Fso.DeleteFile (Me.Caption)
    DoEvents
    Set MyWorkbook = MyEcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
    On Error Resume Next
    MyWorkbook.SaveAs Replace(Me.Caption, "Rév.:", "")
    If Err Then
        BoolErr = True
        MsgBox Err.Description
        Err.Clear
        On Error GoTo 0
        GoTo Fin
    End If
    MyWorkbook.Close
    MyTim = Now
    While DateDiff("s", MyTim, Now) < 1
        DoEvents
    Wend
    Set MyWorkbook = MyEcel.Workbooks.Open(Me.Caption)
'    MyEcel.Visible = True
End If
'MyEcel.Visible = True
'MyWorkbook.Application.Visible = True
Set Myrange = MyWorkbook.Worksheets("NOEUDS").Range("a1").CurrentRegion
MyWorkbook.Worksheets("NOEUDS").Range(Myrange(2, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("NOEUDS").Select
MyWorkbook.Worksheets("NOEUDS").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet6
Set MyRange2 = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("NOEUDS").Paste





Set Myrange = MyWorkbook.Worksheets("Critères").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Critères").Range(Myrange(2, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Critères").Select
MyWorkbook.Worksheets("Critères").Range("a1").Select
'Me.Spreadsheet5.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet5
Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Critères").Paste

Set Myrange = MyWorkbook.Worksheets("Notas").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Notas").Range(Myrange(2, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Notas").Select
MyWorkbook.Worksheets("Notas").Range("a1").Select
'Me.Spreadsheet4.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet4
Set MyRange2 = Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Notas").Paste
MyWorkbook.Worksheets("Notas").Range("a1").Select


Set Myrange = MyWorkbook.Worksheets("Composants").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Composants").Range(Myrange(2, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Composants").Select
MyWorkbook.Worksheets("Composants").Range("a1").Select
'Me.Spreadsheet3.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet3
Set MyRange2 = Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Composants").Paste
MyWorkbook.Worksheets("Composants").Range("a1").Select


Set Myrange = MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range(Myrange(2, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Ligne_Tableau_fils").Select
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").Select
'Me.Spreadsheet2.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet2
Set MyRange2 = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Ligne_Tableau_fils").Paste






MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").Select
Set Myrange = MyWorkbook.Worksheets("Connecteurs").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Connecteurs").Range(Myrange(2, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Connecteurs").Select
MyWorkbook.Worksheets("Connecteurs").Range("a1").Select
'Me.Spreadsheet1.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet1
Set MyRange2 = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy


MyWorkbook.Worksheets("Connecteurs").Paste





MyWorkbook.Worksheets("Connecteurs").Range("a1").Select













 MyWorkbook.Save
Fin:
 
 
 Set MyRange2 = Nothing
 Set Myrange = Nothing
 MyWorkbook.Close False
 Set MyWorkbook = Nothing
 MyEcel.Quit
 Set MyEcel = Nothing
 If BoolErr = False Then boolExcute = True
 NotSortie = False
 boolActu = False
Me.Hide
End Sub



Private Sub Command4_Click()

MenuShow = True
 boolExcute = False
 NotSortie = False
 boolActu = False
Me.Hide
End Sub

Private Sub Command1_Click()
DoEvents
Dim Myrange
Dim Sql As String
Me.Spreadsheet5.Cells(1, 1).Select
Sql = "DELETE Ajout_LIAISON_CONNECTEURS.* FROM Ajout_LIAISON_CONNECTEURS;"
Con.Exequte Sql
Sql = "DELETE Ajout_LIAISON.* FROM Ajout_LIAISON;"
Con.Exequte Sql

Msg = ""
DoEvents
IfValidationOk = True
RazFiltreEditExcel Me.Spreadsheet1
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
For i = 2 To Myrange.Rows.Count
IfValidationOk = True
Me.Spreadsheet1.Cells(i, 1).Select
ConverOuiNon Myrange, i
If Msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
    Me.Spreadsheet1.Cells(i, 1).Value = Me.Spreadsheet1.Cells(i, 1).Value
If Msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next i
RazFiltreEditExcel Me.Spreadsheet2
Set Myrange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion


For i = 2 To Myrange.Rows.Count
Me.Spreadsheet2.Cells(i, 15).Select
ConverOuiNon Myrange, i
IfValidationOk = True
    Me.Spreadsheet2.Cells(i, 15).Value = UCase("'" & Me.Spreadsheet2.Cells(i, 15).Value)
If Msg <> "" Then
'Me.Spreadsheet2.ActiveSheet.AutoFilter = True
    IfValidationOk = False
    Exit Sub
End If
Me.Spreadsheet2.Cells(i, 20).Select
    Me.Spreadsheet2.Cells(i, 20).Value = Me.Spreadsheet2.Cells(i, 20).Value
    If Msg <> "" Then
'    Me.Spreadsheet2.ActiveSheet.AutoFilterMode = True

        IfValidationOk = False
        Exit Sub
    End If
DoEvents
Next i
'Me.Spreadsheet5.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet5
Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion


For i = 2 To Myrange.Rows.Count
IfValidationOk = True
Me.Spreadsheet5.Cells(i, 1).Select
ConverOuiNon Myrange, i
If Msg <> "" Then
    IfValidationOk = False
'    Me.Spreadsheet5.ActiveSheet.AutoFilter = True
    Exit Sub
End If
    Me.Spreadsheet5.Cells(i, 1).Value = Me.Spreadsheet5.Cells(i, 1).Value
If Msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next i
RazFiltreEditExcel Me.Spreadsheet6
Set Myrange = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion


For i = 2 To Myrange.Rows.Count
IfValidationOk = True
ConverOuiNon Myrange, i
Me.Spreadsheet6.Cells(i, 1).Select

DoEvents
If Msg <> "" Then
'Me.Spreadsheet6.ActiveSheet.AutoFilter = True
    IfValidationOk = False
    Exit Sub
End If
'If Spreadsheet6.Cells(i, 4) = "x" Then
'    MsgBox ""
'End If
    Command7_Click
If Msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next i

RazFiltreEditExcel Me.Spreadsheet4
Set Myrange = Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion


For i = 2 To Myrange.Rows.Count
Me.Spreadsheet4.Cells(i, 3).Select
ConverOuiNon Myrange, i
Me.Spreadsheet4.Cells(i, 3) = i - 1
IfValidationOk = True
 
DoEvents
Next i

RazFiltreEditExcel Me.Spreadsheet3
Set Myrange = Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion


For i = 2 To Myrange.Rows.Count
Me.Spreadsheet3.Cells(i, 2).Select
ConverOuiNon Myrange, i
Me.Spreadsheet3.Cells(i, 3) = 0
IfValidationOk = True
 
DoEvents
Next i
'Me.Spreadsheet6.ActiveSheet.AutoFilter = True
If MyErr = True Then
    LoadLiasons.Charger MyClient
End If
MyErr = False
    IfValidationOk = False
    boolActu = True

End Sub
Private Sub Command2_Click()
Command1_Click
If Msg <> "" Then Exit Sub
Command3_Click
End Sub

Private Sub Command5_Click()
Set Myrange = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion

'If Me.Tag = "" Then
    If Trim("" & Me.Hab) <> "" Then
        boolSelctChange = True
        Myrange(Myrange.Rows.Count + 1, 1).Select
        Myrange(Myrange.Rows.Count + 1, 2) = Me.Fleche_Droite.Value
        Myrange(Myrange.Rows.Count + 1, 3) = Me.TORON_P.Value
        Myrange(Myrange.Rows.Count + 1, 1) = Me.ACTIVER.Value
        Myrange(Myrange.Rows.Count + 1, 5) = Val(Replace("" & Me.Longueur, ",", "."))
        Myrange(Myrange.Rows.Count + 1, 6) = Val(Replace("" & Me.Long_C, ",", "."))
        Myrange(Myrange.Rows.Count + 1, 7) = "'" & Me.Hab
        Myrange(Myrange.Rows.Count + 1, 8) = "'" & Me.RSA
        Myrange(Myrange.Rows.Count + 1, 9) = "'" & Me.PSA
        Myrange(Myrange.Rows.Count + 1, 10) = "'" & Me.ENC
         
        Myrange(Myrange.Rows.Count + 1, 11) = Val(Replace("" & Me.DIAMETRE, ",", "."))
        Myrange(Myrange.Rows.Count + 1, 12) = "'" & Me.CLASSE_T
        boolSelctChange = False
    End If
'Else
'     If Trim("" & Me.Hab) <> "" Then
'     boolSelctChange = True
'         Me.Spreadsheet6.ActiveSheet.Cells(Me.Tag, 1).InsertRows
'        MyRange(Me.Tag, 1).Select
'        MyRange(Me.Tag, 1) = Me.Fleche_Droite.Value
'         MyRange(Me.Tag, 2) = Me.TORON_P.Value
'         MyRange(Me.Tag, 3) = Me.ACTIVER.Value
'        MyRange(Me.Tag, 5) = Val(Replace("" & Me.Longueur, ",", "."))
'         MyRange(Me.Tag, 6) = Val(Replace("" & Me.Long_C, ",", "."))
'        MyRange(Me.Tag, 7) = "" & Me.Hab
'        MyRange(Me.Tag, 8) = "" & Me.RSA
'        MyRange(Me.Tag, 9) = "" & Me.PSA
'        MyRange(Me.Tag, 10) = "" & Me.ENC
'         MyRange(Me.Tag, 11) = Val(Replace("" & Me.DIAMETRE, ",", "."))
'        MyRange(Me.Tag, 12) = "" & Me.CLASSE_T
'        boolSelctChange = False
'    End If
'
'End If
Me.Fleche_Droite.Value = 0
Me.ACTIVER.Value = 0
TORON_P.Value = 0
Me.Long_C = ""
 Me.DIAMETRE = ""
Me.CLASSE_T = ""
Me.NOUED = ""
  Longueur = ""
Me.Hab.ListIndex = 0
Me.Tag = ""
End Sub

Private Sub Command6_Click()
If Me.Tag <> "" Then
    boolSelctChange = True
    Me.Spreadsheet6.ActiveSheet.Rows(Val(Me.Tag)).DeleteRows
    Me.Tag = ""
    Me.Hab.ListIndex = 0
    Longueur = ""
    Me.Tag = ""
    Me.NOUED = ""
    boolSelctChange = False
End If
End Sub

Private Sub Command7_Click()
Set Myrange = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion

If Me.Tag = "" Then
    Command5_Click
Else
'    If Trim("" & Me.Hab) <> "" Then
        boolSelctChange = True
        Myrange(Val(Me.Tag), 1).Select
        
         Myrange(Val(Me.Tag), 2) = Fleche_Droite.Value
          Myrange(Val(Me.Tag), 3) = TORON_P.Value
         Myrange(Val(Me.Tag), 1) = Me.ACTIVER.Value
        Myrange(Val(Me.Tag), 5) = Val(Replace("" & Me.Longueur, ",", "."))
         Myrange(Val(Me.Tag), 6) = Val(Replace("" & Me.Long_C, ",", "."))
        Myrange(Val(Me.Tag), 7) = "'" & Me.Hab
        Myrange(Val(Me.Tag), 8) = "'" & Me.RSA
        Myrange(Val(Me.Tag), 9) = "'" & Me.PSA
        Myrange(Val(Me.Tag), 10) = "'" & Me.ENC
        Myrange(Val(Me.Tag), 11) = Val(Replace("" & Me.DIAMETRE, ",", "."))
        Myrange(Val(Me.Tag), 12) = "'" & Me.CLASSE_T
        
        boolSelctChange = False
'    End If
End If
Me.ACTIVER.Value = 0
TORON_P.Value = 0
Me.Long_C = ""
 Me.DIAMETRE = ""
Me.CLASSE_T = ""
Me.NOUED = ""
  Longueur = ""
Me.Hab.ListIndex = 0
Me.Fleche_Droite.Value = 0
Me.Tag = ""
End Sub



Private Sub DIAMETRE_LostFocus()
If MyFormat("dbl", DIAMETRE, "DIAMETRE") = False Then Exit Sub
End Sub

Private Sub ENC_Click()
If boolMajListe = False Then
    boolMajListe = True
'     Me.ENC.ListIndex = MyCollectionENC(MyTableENC(MyCollectionENC("N" & Me.ENC.Text), 1))
     Me.PSA.ListIndex = MyCollectionPSA(MyTableENC(MyCollectionENC("N" & Trim(Me.ENC.Text)), 2))
    Me.RSA.ListIndex = MyCollectionRSA(MyTableENC(MyCollectionENC("N" & Trim(Me.ENC.Text)), 3))
    Me.Hab.ListIndex = MyCollectionHab(MyTableENC(MyCollectionENC("N" & Trim(Me.ENC.Text)), 4))
    boolMajListe = False
End If


End Sub

Private Sub ENC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    Me.ENC.ListIndex = MyCollectionENC(MyTableENC(MyCollectionENC("N" & Me.ENC.Text), 1))
On Error GoTo 0
End If
End Sub

Private Sub Form_Initialize()
boolActu = False
End Sub

Private Sub Hab_Click()
If boolMajListe = False Then
    boolMajListe = True
    
     Me.ENC.ListIndex = MyCollectionENC(MyTableHab(MyCollectionHab("N" & Trim(Me.Hab.Text)), 1))
     Me.PSA.ListIndex = MyCollectionPSA(MyTableHab(MyCollectionHab("N" & Trim(Me.Hab.Text)), 2))
    Me.RSA.ListIndex = MyCollectionRSA(MyTableHab(MyCollectionHab("N" & Trim(Me.Hab.Text)), 3))
'    Me.Hab.ListIndex = MyCollectionHab(MyTableHab(MyCollectionHab("N" & Me.Hab.Text), 4))
    boolMajListe = False
End If

End Sub



Private Sub Hab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
 Me.Hab.ListIndex = MyCollectionHab(MyTableHab(MyCollectionHab("N" & Trim(Me.Hab.Text)), 4))
  On Error GoTo 0
 End If
End Sub

Private Sub Long_C_Change()
If MyFormat("dbl", Long_C, "Longueur cumuler") = False Then Exit Sub

End Sub

Private Sub Longueur_LostFocus()
If MyFormat("dbl", Longueur, "Longueur") = False Then Exit Sub
End Sub

Private Sub PSA_Click()
If boolMajListe = False Then
    boolMajListe = True
     Me.ENC.ListIndex = MyCollectionENC(MyTablePSA(MyCollectionPSA("N" & Trim(Me.PSA.Text)), 1))
'     Me.PSA.ListIndex = MyCollectionPSA(MyTablePSA(MyCollectionPSA("N" & Me.PSA.Text), 2))
    Me.RSA.ListIndex = MyCollectionRSA(MyTablePSA(MyCollectionPSA("N" & Trim(Me.PSA.Text)), 3))
    Me.Hab.ListIndex = MyCollectionHab(MyTablePSA(MyCollectionPSA("N" & Trim(Me.PSA.Text)), 4))
    boolMajListe = False
End If
End Sub

Private Sub PSA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
 Me.PSA.ListIndex = MyCollectionPSA(MyTablePSA(MyCollectionPSA("N" & Trim(Me.PSA.Text)), 2))
    On Error GoTo 0
 End If
End Sub

Private Sub RSA_Click()
If boolMajListe = False Then
    boolMajListe = True
   
     Me.ENC.ListIndex = MyCollectionENC(MyTableRSA(MyCollectionRSA("N" & Trim(Me.RSA.Text)), 1))
     Me.PSA.ListIndex = MyCollectionPSA(MyTableRSA(MyCollectionRSA("N" & Trim(Me.RSA.Text)), 2))
'    Me.RSA.ListIndex = MyCollectionRSA(MyTableRSA(MyCollectionRSA("N" & Me.RSA.Text), 3))
    Me.Hab.ListIndex = MyCollectionHab(MyTableRSA(MyCollectionRSA("N" & Trim(Me.RSA.Text)), 4))
    boolMajListe = False
End If


End Sub

Private Sub RSA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
On Error Resume Next
 Me.RSA.ListIndex = MyCollectionRSA(MyTableRSA(MyCollectionRSA("N" & Trim(Me.RSA.Text)), 3))
  On Error GoTo 0
 End If
End Sub

Private Sub Spreadsheet1_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet1.ActiveCell.Row
Col = Me.Spreadsheet1.ActiveCell.Column

If (NoMacro1 = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro1 = True
Set Myrange = Me.Spreadsheet1.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
   Me.Spreadsheet1.Cells(Row, 11) = UCase("" & Me.Spreadsheet1.Cells(Row, 11))
   Me.Spreadsheet1.Cells(Row, 2) = UCase("'" & Me.Spreadsheet1.Cells(Row, 2))
    Me.Spreadsheet1.Cells(Row, 5) = UCase("'" & Me.Spreadsheet1.Cells(Row, 5))
If Trim("" & Me.Spreadsheet1.Cells(Row, 11)) <> "" Then
    If UCase(Me.Spreadsheet1.Cells(Row, 11)) <> "TOUS" Then
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("B1", "B" & CStr(Myrange.Rows.Count))
        Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
         Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range("A1", "A" & CStr(MyRange2.Rows.Count))
        
        i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet1.Cells(Row, 11))), MyRange2, True)
        
        If i = 0 Then
            
           
            Msg = "CODE CRITERE : " & UCase(Trim("" & Me.Spreadsheet1.Cells(Row, 11))) & " introuvable"
            MsgBox Msg, vbQuestion
             Me.Spreadsheet1.Cells(Row, 11) = ""
                    
        End If
     Set Myrange = Nothing
     End If
     
End If




    
        If Trim("" & Me.Spreadsheet1.Cells(Row, 2)) <> "" Then
            Me.Spreadsheet1.Cells(Row, 6) = Row - 1
        End If
        If Trim("" & Me.Spreadsheet1.Cells(Row, 5)) <> "" Then
        
            Sql = "SELECT LIAISON_CONNECTEURS.LIB FROM LIAISON_CONNECTEURS "
            Sql = Sql & "WHERE LIAISON_CONNECTEURS.CLIENT='" & MyReplace(MyClient) & "' "
            Sql = Sql & "AND LIAISON_CONNECTEURS.LIAISON='" & MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 5))) & "';"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = False Then
                Me.Spreadsheet1.Cells(Row, 4) = Trim("'" & Rs!Lib)
            Else
                If IfValidationOk = False Then
'                    If MsgBox("Le code App : " & DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 54)) & " n'existe pas" & vbCrLf & "Voulez-vous le créer", vbQuestion + vbYesNo, "Liaison Connecteur :") = vbYes Then
                        LibCode_APP = InputBox("Entrez la désignation du code APP : " & DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 5)), "Ajout d'un code App")
                        If Trim(LibCode_APP) <> "" Then
                            Me.Spreadsheet1.Cells(Row, 4) = LibCode_APP
                            Sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
                            Sql = Sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 5)))) & "', '" & UCase(MyReplace(Me.Spreadsheet1.Cells(Row, 4))) & "' );"
                            Con.Exequte Sql
                        End If
'                    End If
                Else
                   Sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON "
                   Sql = Sql & "FROM Ajout_LIAISON_CONNECTEURS "
                   Sql = Sql & "WHERE Ajout_LIAISON_CONNECTEURS.LIAISON='" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 5)))) & "' "
                   Sql = Sql & "AND Ajout_LIAISON_CONNECTEURS.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(Sql)
                    If Rs.EOF = True Then
                        Sql = "INSERT INTO Ajout_LIAISON_CONNECTEURS ( LIAISON, LIB,Job ) "
                        Sql = Sql & "values ( '" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 5)))) & "', '" & MyReplace(Me.Spreadsheet1.Cells(Row, 4)) & "'," & NmJob & ");"
                        Con.Exequte Sql
                        MyErr = True
                    End If
                End If
            End If
            Set Rs = Con.CloseRecordSet(Rs)
        
        End If
    
   
   
    
    NoMacro1 = False
    Col3 = 0
Fin:
End Sub

Private Sub Spreadsheet1_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Static SaveRow As Long
Dim Row As Long
Row = Me.Spreadsheet1.ActiveCell.Row
If Row = 1 Then GoTo Fin
If SaveRow = 0 Then SaveRow = Row
If NoMacro1 = True Then GoTo Fin
    NoMacro1 = True
    
    If IfValidationOk = True Then 'NEANT
           
            If (Me.Spreadsheet1.Cells(Row, 2)) <> "" And Me.Spreadsheet1.Cells(Row, 5) = "" Then
            If (Me.Spreadsheet1.Cells(Row, 1)) <> 0 Then
                Me.Spreadsheet1.Cells(Row, 5).Select
                
                MsgBox "Vous devez saisir le Code Appareil", vbCritical, "Code Appareil Connecteur"
                Msg = "?"
                End If
            End If
        Else
            If (SaveRow <> Row) And (Me.Spreadsheet1.Cells(SaveRow, 2)) <> "" And Me.Spreadsheet1.Cells(SaveRow, 5) = "" Then
            If (UCase(Me.Spreadsheet1.Cells(SaveRow, 1))) <> 0 Then
                Me.Spreadsheet1.Cells(SaveRow, 5).Select
                
                MsgBox "Vous devez saisir le Code Appareil", vbExclamation, "Code Appareil Connecteur"
                Msg = "?"
            Else
             Me.Spreadsheet1.Cells(SaveRow, 2) = UCase(Me.Spreadsheet1.Cells(SaveRow, 2))
            End If
            End If
    End If

SaveRow = Row


NoMacro1 = False
Fin:
End Sub

Private Sub Spreadsheet2_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim Myrange
Dim Rs As Recordset
Dim Sql As String
Dim LibCode_APP As String

Static Col3 As Long
Row = Me.Spreadsheet2.ActiveCell.Row
Col = Me.Spreadsheet2.ActiveCell.Column

If (NoMacro2 = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro2 = True
 Set Myrange = Me.Spreadsheet2.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
If (Trim("" & Me.Spreadsheet2.Cells(Row, 15)) <> "") And (Trim("" & Me.Spreadsheet2.Cells(Row, 20)) <> "") And (Trim("" & Me.Spreadsheet2.Cells(Row, 23)) = "") Then
   
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("E1", "E" & CStr(Myrange.Rows.Count))


        i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 15))))

        If i <> 0 Then

            Me.Spreadsheet2.Cells(Row, 23) = "" & Me.Spreadsheet1.Cells(i, 11)
           

        End If
        If Trim("" & Me.Spreadsheet2.Cells(Row, 23)) = "" Then
             i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 20))))

        If i <> 0 Then

            Me.Spreadsheet2.Cells(Row, 23) = "" & Me.Spreadsheet1.Cells(i, 11)
           

        End If
        End If
     Set Myrange = Nothing
      If Trim("" & Me.Spreadsheet2.Cells(Row, 23)) = "" Then
      MsgBox "Vous devez saisir un code critère."
      Msg = "?"
     
      End If
     End If
     If (Trim("" & Me.Spreadsheet2.Cells(Row, 15)) <> "") And (Trim("" & Me.Spreadsheet2.Cells(Row, 20)) <> "") And (Trim(UCase("" & Me.Spreadsheet2.Cells(Row, 23))) = "TOUS") Then
   
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("e1", "e" & CStr(Myrange.Rows.Count))


        i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 15))))

        If i <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(i, 11)) <> "" Then
                Me.Spreadsheet2.Cells(Row, 23) = "" & Me.Spreadsheet1.Cells(i, 11)
           End If

        End If
        If UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 23))) = "TOUS" Then
             i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 20))))

        If i <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(i, 11)) <> "" Then
                Me.Spreadsheet2.Cells(Row, 23) = "" & Me.Spreadsheet1.Cells(i, 11)
           End If

        End If
        End If
     Set Myrange = Nothing
     
     End If

 If (Trim("" & Me.Spreadsheet2.Cells(Row, 23)) <> "") Then
Me.Spreadsheet2.Cells(Row, 23) = UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 23)))
    If (Trim("" & Me.Spreadsheet2.Cells(Row, 23)) <> "TOUS") Then
 Me.Spreadsheet2.Cells(Row, 23) = UCase(Me.Spreadsheet2.Cells(Row, 23))
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("B1", "B" & CStr(Myrange.Rows.Count))
        
        
        i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 23))))
        
        If i = 0 Then
            
           
            Msg = "CODE CRITERE : " & UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 23))) & " introuvable"
            MsgBox Msg, vbQuestion
             Me.Spreadsheet2.Cells(Row, 23) = ""
                    
        End If
     Set Myrange = Nothing
     End If
 
      End If





Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("E1", "E" & CStr(Myrange.Rows.Count))


        i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 15))))

        If i <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(i, 11)) <> "" Then
                Me.Spreadsheet2.Cells(Row, 23) = "" & Me.Spreadsheet1.Cells(i, 11)
           End If

        End If
        If UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 23))) = "TOUS" Then
             i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 20))))

        If i <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(i, 11)) <> "" Then
                Me.Spreadsheet2.Cells(Row, 23) = "" & Me.Spreadsheet1.Cells(i, 11)
           End If

        End If
        End If
     Set Myrange = Nothing


    If Trim("" & Me.Spreadsheet2.Cells(Row, 2)) <> "" Then
    If (Row > 1) And (Row = 2) Then
        If Trim("" & Me.Spreadsheet2.Cells(Row, 4)) <> Col3 Then
            Me.Spreadsheet2.Cells(Row, 4) = 1
            Col3 = 1
        End If
    Else
        If (Row > 1) And (Row <> 2) Then
            If Trim("" & Me.Spreadsheet2.Cells(Row, 4)) <> Col3 Then
            Col3 = Row - 1
            Me.Spreadsheet2.Cells(Row, 4) = Col3
        End If
    End If

 End If

If (Col = 15) Or (Col = 20) Then
   
        If Trim("" & Me.Spreadsheet2.Cells(Row, Col)) <> "" Then
            Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
            Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("E1", "E" & CStr(Myrange.Rows.Count))
            NoMacro2 = True
            
            i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))))
        
            If i <> 0 Then
               
                Me.Spreadsheet2.Cells(Row, Col - 1) = UCase(Trim("'" & Myrange(i, 2)))
                Me.Spreadsheet2.Cells(Row, Col - 2) = UCase(Trim("'" & Myrange(i, 4)))
                Me.Spreadsheet2.Cells(Row, Col - 3) = UCase(Trim("'" & Myrange(i, 3)))
                Else
                    Me.Spreadsheet2.Cells(Row, Col - 1) = "0"
                    Me.Spreadsheet2.Cells(Row, Col - 2) = ""
                    
                    Msg = "Le connecteur : " & UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))) & " introuvable"
                    MsgBox Msg, vbQuestion
                     Me.Spreadsheet2.Cells(Row - 1, Col).Select
                End If
            
            
            
            
                Else
                    If Trim("" & Me.Spreadsheet2.Cells(Row, 2)) <> "" Then
                    If UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 1))) <> 0 Then
                        Msg = "Le code APP ne peut être Nul"
                        MsgBox Msg, vbExclamation, "Ligne_Tableau_fils"
                         Me.Spreadsheet2.Cells(Row - 1, Col).Select
                    End If
                    End If
                End If
            End If
        
        
       
        
            If Trim("" & Me.Spreadsheet2.Cells(Row, 2)) <> "" Then
                Sql = "SELECT LIAISON.LIB FROM LIAISON "
                Sql = Sql & "WHERE LIAISON.CLIENT='" & MyReplace(MyClient) & "' "
                Sql = Sql & "AND LIAISON.LIAISON='" & MyReplace(Me.Spreadsheet2.Cells(Row, 2)) & "';"
                Set Rs = Con.OpenRecordSet(Sql)
                If Rs.EOF = False Then
                Me.Spreadsheet2.Cells(Row, 3) = Trim("'" & Rs!Lib)
                Else
                    If IfValidationOk = False Then
                        If MsgBox("La liaison : " & Me.Spreadsheet2.Cells(Row, 2) & " n'existe pas" & vbCrLf & "Voulez-vous la créer", vbQuestion + vbYesNo, "Liaison Fils :") = vbYes Then
                            LibCode_APP = InputBox("Entrez la désignation de la liaison : " & Me.Spreadsheet2.Cells(Row, 1), "Ajout de liaison")
                            If Trim(LibCode_APP) <> "" Then
                                Me.Spreadsheet2.Cells(Row, 3) = LibCode_APP
                                Sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
                                Sql = Sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(Me.Spreadsheet2.Cells(Row, 2))) & "', '" & UCase(MyReplace(Me.Spreadsheet2.Cells(Row, 3))) & "' );"
                                Con.Exequte Sql
                            End If
                        End If
                        Else
                        
                   Sql = "SELECT Ajout_LIAISON.LIAISON "
                   Sql = Sql & "FROM Ajout_LIAISON "
                   Sql = Sql & "WHERE Ajout_LIAISON.LIAISON='" & UCase(MyReplace(Me.Spreadsheet2.Cells(Row, 2))) & "' "
                   Sql = Sql & "AND Ajout_LIAISON.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(Sql)
                    If Rs.EOF = True Then
                        Sql = "INSERT INTO Ajout_LIAISON ( LIAISON, LIB,Job ) "
                        Sql = Sql & "values ( '" & UCase(MyReplace(Me.Spreadsheet2.Cells(Row, 2))) & "', '" & MyReplace(Me.Spreadsheet2.Cells(Row, 3)) & "'," & NmJob & ");"
                        Con.Exequte Sql
                        MyErr = True
                    End If
                    End If
                End If
                Set Rs = Con.CloseRecordSet(Rs)
            
            End If
            
            
            Col3 = 0
        End If
   If (Trim("" & Me.Spreadsheet2.Cells(Row, 23)) <> "") Then

 Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("D1", "D" & CStr(Myrange.Rows.Count))


        i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 15))))

        If i <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(i, 11)) <> "" Then
                Myapp1 = "" & Me.Spreadsheet1.Cells(i, 11)
           End If

        End If
       
             i = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 20))))

        If i <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(i, 11)) <> "" Then
               Myapp2 = "" & Me.Spreadsheet1.Cells(i, 11)
           End If
    If Trim("" & Myapp1) = "" Then Myapp1 = Myapp2
     If Trim(Myapp2) = "" Then Myapp2 = Myapp1
        If ("" & Myapp1 <> Myapp2) And (UCase(Myapp1) <> "TOUS") And (UCase(Myapp2) <> "TOUS") Then
        MsgBox "Une liaison ne peut pas pointer sur deux options différentes : " & Myapp1 & " & " & Myapp2
        Me.Spreadsheet2.Cells(Row, 23) = ""
        Msg = "?"
        End If
        End If
     Set Myrange = Nothing
End If

    NoMacro2 = False
Fin:
End Sub

Private Sub Spreadsheet3_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveRow As Long
Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Dim BoolOui As Boolean
If SaveRow = 0 Then SaveRow = 1


Row = Me.Spreadsheet3.ActiveCell.Row
Col = Me.Spreadsheet3.ActiveCell.Column
If (NoMacro3 = True) Or (Row = 1) Then GoTo Fin

If SaveRow = 0 Then SaveRow = Row
If Row = 1 Then GoTo Fin
boolActu = False
NoMacro3 = True
 Set Myrange = Me.Spreadsheet3.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
   If Trim("" & Me.Spreadsheet3.Cells(Row, 2)) <> "" Then Me.Spreadsheet3.Cells(Row, 3) = Row - 1
   NoMacro3 = False
Fin:
End Sub

Private Sub Spreadsheet3_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveRow As Long
Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Dim BoolOui As Boolean
If SaveRow = 0 Then SaveRow = 1
If NoMacro3 = True Then GoTo Fin

Row = Me.Spreadsheet3.ActiveCell.Row
Col = Me.Spreadsheet3.ActiveCell.Column
If SaveRow = 0 Then SaveRow = Row
If Row = 1 Then GoTo Fin
NoMacro3 = True
 Set Myrange = Me.Spreadsheet1.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
   If Trim("" & Me.Spreadsheet3.Cells(Row, 2)) <> "" Then Me.Spreadsheet3.Cells(Row, 3) = Row - 1
    

If Col > 4 Then
    For i = 5 To NbFinOuiNon
        If Me.Spreadsheet3.Cells(Row, i) = 1 Then
            If i <> Col Then
                MsgBox "Vous ne pouvez pas sélectionner plusieurs répertoires.", vbExclamation
                Me.Spreadsheet3.Cells(Row, Col) = 0
            End If
        End If
    Next i
End If
BoolOui = False
If (SaveRow <> 1) And (SaveRow <> Row) And (Trim("" & Me.Spreadsheet3.Cells(SaveRow, 1)) <> "") Then
 For i = 5 To NbFinOuiNon
    If Val(Me.Spreadsheet3.Cells(SaveRow, i)) = 1 Then
        BoolOui = True
        Exit For
    End If
    
    Next i
  If BoolOui = False Then
    MsgBox "Vous devez sélectionner un répertoire.", vbExclamation
    Me.Spreadsheet3.Cells(SaveRow, 5).Select
  End If
   
End If
SaveRow = Row
NoMacro3 = False
Fin:
End Sub

Private Sub Spreadsheet4_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet4.ActiveCell.Row
Col = Me.Spreadsheet4.ActiveCell.Column
If Row = 1 Then GoTo Fin
If (NoMacro4 = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro4 = True
 Set Myrange = Me.Spreadsheet4.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
If Col = 2 Then
   If Trim("" & Me.Spreadsheet4.Cells(Row, 2)) <> "" Then Me.Spreadsheet4.Cells(Row, 3) = Row - 1
    
End If
NoMacro4 = False

Fin:
End Sub

Private Sub Form_Activate()
Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet

If bool_Activate = True Then GoTo Fin
bool_Activate = True
Dim Myrange As EXCEL.Range
Set MyEcel = New EXCEL.Application
NotSortie = True
'MyEcel.Visible = True
Set a = Me.Spreadsheet1.Cells(2, 2)
If Trim(Me.Caption) = "" Then Exit Sub
If Nouveau = False Then
    Set MyWorkbook = MyEcel.Workbooks.Open(Me.Caption)
Else

    Set MyWorkbook = MyEcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
End If

Set Myrange = MyWorkbook.Sheets("Critères").Range("a1").CurrentRegion

Myrange.Copy
Me.Spreadsheet5.ActiveSheet.Range("a1").Paste
Set Myrange = MyWorkbook.Sheets("Connecteurs").Range("a1").CurrentRegion
Myrange.Copy
Me.Spreadsheet1.ActiveSheet.Range("a1").Paste

Set Myrange = MyWorkbook.Sheets("Ligne_Tableau_fils").Range("a1").CurrentRegion

Myrange.Copy
Me.Spreadsheet2.ActiveSheet.Range("a1").Paste
Set Myrange = MyWorkbook.Sheets("Connecteurs").Range("a1").CurrentRegion
Myrange.Copy
Me.Spreadsheet1.ActiveSheet.Range("a1").Paste

Set Myrange = MyWorkbook.Sheets("Composants").Range("a1").CurrentRegion
NbFinOuiNon = Myrange.Columns.Count
Myrange.Copy
Me.Spreadsheet3.ActiveSheet.Range("a1").Paste

Set Myrange = MyWorkbook.Sheets("Notas").Range("a1").CurrentRegion
Myrange.Copy
Me.Spreadsheet4.ActiveSheet.Range("a1").Paste

Set Myrange = MyWorkbook.Sheets("NOEUDS").Range("a1").CurrentRegion
Myrange.Copy
Me.Spreadsheet6.ActiveSheet.Range("a1").Paste

MyEcel.AlertBeforeOverwriting = False

Set Myrange = Nothing
MyWorkbook.Close False
Set MyWorkbook = Nothing

MyEcel.Quit
Set MyExcel = Nothing


'Me.Spreadsheet1.ActiveSheet.Panes(1).VisibleRange = False
Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet4.ActiveSheet.Range("a1").Select
Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet5.ActiveSheet.Range("a1").Select
Me.Spreadsheet5.Columns(1).NumberFormat = "Yes/No"

Me.Spreadsheet6.Columns(1).NumberFormat = "Yes/No"
Me.Spreadsheet6.Columns(2).NumberFormat = "Yes/No"
Me.Spreadsheet6.Columns(3).NumberFormat = "Yes/No"
Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet6.ActiveSheet.Range("a1").Select

Me.Spreadsheet3.ActiveSheet.Range("a1").Select
Me.Spreadsheet2.ActiveSheet.Range("a1").Select
Me.Spreadsheet2.Columns(1).NumberFormat = "Yes/No"
Me.Spreadsheet4.Columns(1).NumberFormat = "Yes/No"
Me.Spreadsheet1.ActiveSheet.Range("a1").Select
Me.Spreadsheet1.Columns(1).NumberFormat = "Yes/No"
Me.Spreadsheet1.Columns(3).NumberFormat = "Yes/No"
Me.Spreadsheet3.Columns(1).NumberFormat = "Yes/No"
For i = 5 To 305
    Me.Spreadsheet3.Columns(i).NumberFormat = "Yes/No"
 Next i
    
DoEvents
LstMaj
Fin:
End Sub
Public Sub Chargement(Fichier As String, Client As String, Optional NouveauF As Boolean)
Dim Rs As Recordset
Dim Sql As String
Dim txtMyCollectionLienHab As String
Sql = "SELECT T_Regle_Comp_Hab.ENCELADE, T_Regle_Comp_Hab.PSA, "
Sql = Sql & "T_Regle_Comp_Hab.RSA, T_Regle_Comp_Hab.libellé, T_Regle_Comp_Hab.Numéro "
Sql = Sql & "FROM T_Regle_Comp_Hab "
Sql = Sql & "ORDER BY T_Regle_Comp_Hab.libellé;"
i = 0
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
i = i + 1
Rs.MoveNext
Wend
ReDim MyTableENC(i, 4)
ReDim MyTablePSA(i, 4)
ReDim MyTableRSA(i, 4)
ReDim MyTableHab(i, 4)
ReDim MyTableHab(i, 4)

ReDim MyTableHab(i, 4)
Rs.Requery
i = 0
Me.ENC.AddItem ""
Me.PSA.AddItem ""
Me.RSA.AddItem ""
Me.Hab.AddItem ""

Set MyCollectionENC = Nothing
Set MyCollectionPSA = Nothing
Set MyCollectionRSA = Nothing
Set MyCollectionHab = Nothing
Set MyCollectionLienHab = Nothing

Set MyCollectionENC = New Collection
Set MyCollectionPSA = New Collection
Set MyCollectionRSA = New Collection
Set MyCollectionHab = New Collection
Set MyCollectionLienHab = New Collection




  For i = 0 To UBound(MyTablePSA)
       For i2 = 1 To 4
            MyTableENC(i, i2) = "N"
            MyTablePSA(i, i2) = "N"
            MyTableRSA(i, i2) = "N"
            MyTableHab(i, i2) = "N"
          
        Next
  Next
  i = 0
While Rs.EOF = False

    i = i + 1
    txtMyCollectionLienHab = ""
    
    
    If Trim("" & Rs!ENCELADE) <> "" Then
        Me.ENC.AddItem Rs!ENCELADE
    End If
        txtMyCollectionLienHab = txtMyCollectionLienHab & "N" & Rs!ENCELADE & ";"
'    Else
'        txtMyCollectionLienHab = txtMyCollectionLienHab & "N;"
''    End If
    If Trim("" & Rs!PSA) <> "" Then
        Me.PSA.AddItem Rs!PSA
    End If
       txtMyCollectionLienHab = txtMyCollectionLienHab & "N" & Rs!PSA & ";"
'     Else
'        txtMyCollectionLienHab = txtMyCollectionLienHab & "N;"
''
'    End If
    If Trim("" & Rs!RSA) <> "" Then
        Me.RSA.AddItem Rs!RSA
        
'     Else
'        txtMyCollectionLienHab = txtMyCollectionLienHab & "N;"
    End If
 txtMyCollectionLienHab = txtMyCollectionLienHab & "N" & Rs!RSA & ";"
    
    
    If Trim("" & Rs!libellé) <> "" Then
        Me.Hab.AddItem Rs!libellé
        
'    Else
'        txtMyCollectionLienHab = txtMyCollectionLienHab & "N;"
    End If
      txtMyCollectionLienHab = txtMyCollectionLienHab & "N" & Rs!libellé & ";"
     MyCollectionLienHab.Add txtMyCollectionLienHab, "N" & Rs!libellé
    Rs.MoveNext
Wend

Set Rs = Con.CloseRecordSet(Rs)
MyClient = Client
Nouveau = NouveauF
Me.Caption = Fichier
Me.Show vbModal
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = NotSortie
bool_Activate = False
End Sub

Private Sub Spreadsheet5_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)


Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Static SaveRow As Long
Dim Sql As String
Dim Rs As Recordset

Row = Me.Spreadsheet5.ActiveCell.Row
Col = Me.Spreadsheet5.ActiveCell.Column
If SaveRow = 0 Then SaveRow = Row
If (NoMacro5 = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro5 = True
 Set Myrange = Me.Spreadsheet5.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
     Me.Spreadsheet5.Cells(SaveRow, 2) = UCase(Trim("" & Me.Spreadsheet5.Cells(SaveRow, 2)))
     Me.Spreadsheet5.Cells(SaveRow, 3) = UCase(Trim("" & Me.Spreadsheet5.Cells(SaveRow, 3)))
     
    If (Trim("" & Me.Spreadsheet5.Cells(SaveRow, 2)) <> "") And (Trim("" & Me.Spreadsheet5.Cells(SaveRow, 3)) = "") And (SaveRow <> Row) Then
        MsgBox "Le champ CRITERES est obligatoire", vbExclamation
        Msg = "?"
       Me.Spreadsheet5.Cells(SaveRow, 3).Select
    End If
    If (Trim("" & Me.Spreadsheet5.Cells(SaveRow, 2)) = "") And (Trim("" & Me.Spreadsheet5.Cells(SaveRow, 3)) <> "") And (SaveRow <> Row) Then
        MsgBox "Le champ CODE CRITERES est obligatoire", vbExclamation
        Msg = "?"
       Me.Spreadsheet5.Cells(SaveRow, 2).Select
    End If
SaveRow = Row
NoMacro5 = False
Fin:
End Sub

Private Sub Spreadsheet5_Click(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Spreadsheet5_Change EventInfo
End Sub

Private Sub Spreadsheet5_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Spreadsheet5_Change EventInfo
End Sub

Private Sub Spreadsheet6_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Row = Spreadsheet6.ActiveCell.Row
If Row = 1 Then GoTo Fin
If NoMacro6 = True Then GoTo Fin
boolActu = False

NoMacro6 = True
 Set Myrange = Me.Spreadsheet6.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
If Trim("" & Spreadsheet6.Cells(Row, 1)) <> "" Then
'    Spreadsheet6.Cells(Row, 4) = NoeuName(Row)
'    Else
'        If Trim("" & Spreadsheet6.Cells(Row, 8)) <> "" Then
'           Spreadsheet6.Cells(Row, 4) = NoeuName(Row)
'        Else
'            If Trim("" & Spreadsheet6.Cells(Row, 9)) <> "" Then
'                Spreadsheet6.Cells(Row, 4) = NoeuName(Row)
'            Else
'                If Trim("" & Spreadsheet6.Cells(Row, 10)) <> "" Then
                    Spreadsheet6.Cells(Row, 4) = NoeuName(Row)
DoEvents
'                End If
'        End If
'    End If
End If

NoMacro6 = False
Fin:

End Sub

Function NoeuName2(Row As Long)
Dim txt As String
Dim Ofset As Long
Dim NbTour As Long
Dim NbTord As Long
Dim txtColone As Long
txtColone = 2
txt = "AA"
Ofset = 0
NbTour = 0
NbTord = 0
Reprise:

For i = 0 To Row - 2
aa = Mid(txt, Len(txt) - Ofset, 1)

    aa = Chr(Asc("A") + (1 * (i - (26 * NbTour))))

Mid(txt, Len(txt) - Ofset, 1) = aa


If Asc(Mid(aa, 1, 1)) < 65 Or Asc(Mid(aa, 1, 1)) > 90 Then

Mid(txt, Len(txt) - Ofset, 1) = "A"


    Ofset = Ofset + 1
    NbTour = NbTour + 1
    Mid(txt, Len(txt) - Ofset, 1) = Chr(Asc(Mid(txt, Len(txt) - Ofset, 1)) + 1)
    If Asc(Mid(txt, 1, 1)) < 65 Or Asc(Mid(txt, 1, 1)) > 90 Then
 Mid(txt, 1, 1) = "A"
    txt = txt & "A"
    

End If
Ofset = 0
   
End If


Next

NoeuName2 = txt
End Function

Private Sub Spreadsheet6_SelectionChanging(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim aa As String
Dim MyTxt As String
On Error Resume Next
If boolSelctChange = False Then
 Me.Tag = ""
If EventInfo.Range.Row > 1 Then
Me.Tag = EventInfo.Range.Row
Me.Fleche_Droite.Value = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 2)
TORON_P.Value = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 3)
    Me.Longueur = CStr(Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 5))
    Long_C = CStr(Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 6))
    Me.NOUED = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 4)
    
    If Trim("" & Spreadsheet6.Cells(EventInfo.Range.Row, 7)) <> "" Then
            Me.Hab.ListIndex = MyCollectionHab("N" & Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 7))
    Else
        If Trim("" & Spreadsheet6.Cells(EventInfo.Range.Row, 8)) <> "" Then
           Me.RSA.ListIndex = MyCollectionRSA("N" & Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 8))
        Else
            If Trim("" & Spreadsheet6.Cells(EventInfo.Range.Row, 9)) <> "" Then
                Me.PSA.ListIndex = MyCollectionPSA("N" & Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 9))
            Else
                If Trim("" & Spreadsheet6.Cells(EventInfo.Range.Row, 10)) <> "" Then
                   Me.ENC.ListIndex = MyCollectionENC("N" & Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 10))
                  Else
                    Me.Hab.ListIndex = 0
                   
                End If
        End If
    End If
End If
    
   
   Me.ACTIVER.Value = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 1)
   Me.DIAMETRE = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 11)
   Me.CLASSE_T = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 12)
   
End If
End If
DoEvents
On Error GoTo 0
End Sub
Sub LstMaj()
Set MyCollectionENC = Nothing
Set MyCollectionPSA = Nothing
Set MyCollectionRSA = Nothing
Set MyCollectionHab = Nothing


Set MyCollectionENC = New Collection
Set MyCollectionPSA = New Collection
Set MyCollectionRSA = New Collection
Set MyCollectionHab = New Collection

NoMaj = True
DoEvents
For i = 0 To Me.ENC.ListCount - 1
        MyCollectionENC.Add i, "N" & Trim(Me.ENC.List(i))
        
Next
For i = 0 To Me.PSA.ListCount - 1
       MyCollectionPSA.Add i, "N" & Trim(Me.PSA.List(i))
Next
For i = 0 To Me.RSA.ListCount - 1
        MyCollectionRSA.Add i, "N" & Trim(Me.RSA.List(i))
Next

For i = 0 To Me.Hab.ListCount - 1
      MyCollectionHab.Add i, "N" & Me.Hab.List(i)
Next
For i = 1 To MyCollectionLienHab.Count
    zz = Split(MyCollectionLienHab(i), ";")
    For i2 = 0 To 3
        If zz(i2) <> "N" Then
            MyTableHab(MyCollectionHab(zz(3)), i2 + 1) = Trim(zz(i2))
            
        
        End If
    Next
Next
For i = 1 To UBound(MyTableHab)
    If MyTableHab(i, 1) <> "N" Then
        MyTableENC(MyCollectionENC(MyTableHab(i, 1)), 1) = MyTableHab(i, 1)
        MyTableENC(MyCollectionENC(MyTableHab(i, 1)), 4) = MyTableHab(i, 4)
        
        If MyTableHab(i, 2) <> "N" Then
            MyTableENC(MyCollectionENC(MyTableHab(i, 1)), 2) = MyTableHab(i, 2)
        End If
        
        If MyTableHab(i, 3) <> "N" Then
            MyTableENC(MyCollectionENC(MyTableHab(i, 1)), 3) = MyTableHab(i, 3)
        End If
        
    End If
    
    If MyTableHab(i, 2) <> "N" Then
        MyTablePSA(MyCollectionPSA(MyTableHab(i, 2)), 2) = MyTableHab(i, 2)
        MyTablePSA(MyCollectionPSA(MyTableHab(i, 2)), 4) = MyTableHab(i, 4)
        
        If MyTableHab(i, 1) <> "N" Then
             MyTablePSA(MyCollectionPSA(MyTableHab(i, 2)), 1) = MyTableHab(i, 1)
        End If
        
        If MyTableHab(i, 3) <> "N" Then
            MyTablePSA(MyCollectionPSA(MyTableHab(i, 2)), 3) = MyTableHab(i, 3)
        End If
    End If
     If MyTableHab(i, 3) <> "N" Then
        MyTableRSA(MyCollectionRSA(MyTableHab(i, 3)), 3) = MyTableHab(i, 3)
        MyTableRSA(MyCollectionRSA(MyTableHab(i, 3)), 4) = MyTableHab(i, 4)
        
        If MyTableHab(i, 1) <> "N" Then
             MyTableRSA(MyCollectionRSA(MyTableHab(i, 3)), 1) = MyTableHab(i, 1)
        End If
        
        If MyTableHab(i, 2) <> "N" Then
            MyTableRSA(MyCollectionRSA(MyTableHab(i, 3)), 2) = MyTableHab(i, 2)
        End If
    End If
Next

End Sub

Sub ConverOuiNon(Myrange, Index)
For i = 1 To Myrange.Columns.Count
   If Myrange(Index, i).NumberFormat = "Yes/No" Then
  
        If Not IsNumeric(Myrange(Index, i).Value) Then
            If UCase(Left(Myrange(Index, i).Value, 1)) = "N" Then
                Myrange(Index, i).Value = 0
                DoEvents
               
            Else
                Myrange(Index, i).Value = 1
                DoEvents
               
            End If
        End If
      
   End If
    
Next
End Sub
