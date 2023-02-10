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
      Left            =   -120
      TabIndex        =   4
      Top             =   480
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   18071
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   1
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "Critères"
      TabPicture(0)   =   "AutoCable.frx":1098C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Spreadsheet5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Connecteurs"
      TabPicture(1)   =   "AutoCable.frx":109A8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Spreadsheet1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tableau de fils"
      TabPicture(2)   =   "AutoCable.frx":109C4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Spreadsheet2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Composants"
      TabPicture(3)   =   "AutoCable.frx":109E0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Spreadsheet3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Notas"
      TabPicture(4)   =   "AutoCable.frx":109FC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Spreadsheet4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Noeuds"
      TabPicture(5)   =   "AutoCable.frx":10A18
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command8"
      Tab(5).Control(1)=   "txtOption"
      Tab(5).Control(2)=   "Fleche_Droite"
      Tab(5).Control(3)=   "Long_C"
      Tab(5).Control(4)=   "TORON_P"
      Tab(5).Control(5)=   "Spreadsheet6"
      Tab(5).Control(6)=   "CLASSE_T"
      Tab(5).Control(7)=   "DIAMETRE"
      Tab(5).Control(8)=   "ACTIVER"
      Tab(5).Control(9)=   "Command7"
      Tab(5).Control(10)=   "Command6"
      Tab(5).Control(11)=   "Command5"
      Tab(5).Control(12)=   "ENC"
      Tab(5).Control(13)=   "PSA"
      Tab(5).Control(14)=   "RSA"
      Tab(5).Control(15)=   "Hab"
      Tab(5).Control(16)=   "Longueur"
      Tab(5).Control(17)=   "Label10"
      Tab(5).Control(18)=   "Label9"
      Tab(5).Control(19)=   "Label8"
      Tab(5).Control(20)=   "Label7"
      Tab(5).Control(21)=   "NOUED"
      Tab(5).Control(22)=   "Label6"
      Tab(5).Control(23)=   "Label5"
      Tab(5).Control(24)=   "Label4"
      Tab(5).Control(25)=   "Label3"
      Tab(5).Control(26)=   "Label2"
      Tab(5).Control(27)=   "Label1"
      Tab(5).ControlCount=   28
      TabCaption(6)   =   "Nomenclature Connecteur"
      TabPicture(6)   =   "AutoCable.frx":10A34
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Spreadsheet7"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Nomenclature Fils"
      TabPicture(7)   =   "AutoCable.frx":10A50
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Spreadsheet8"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Nomenclature Habillage"
      TabPicture(8)   =   "AutoCable.frx":10A6C
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Spreadsheet9"
      Tab(8).ControlCount=   1
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   -57600
         Picture         =   "AutoCable.frx":10A88
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1260
         Width           =   315
      End
      Begin VB.TextBox txtOption 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -59400
         TabIndex        =   37
         Top             =   1260
         Width           =   1815
      End
      Begin VB.CheckBox Fleche_Droite 
         Alignment       =   1  'Right Justify
         Caption         =   "Fleche D"
         Height          =   315
         Left            =   -74400
         TabIndex        =   35
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox Long_C 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -64320
         TabIndex        =   33
         Top             =   780
         Width           =   1455
      End
      Begin VB.CheckBox TORON_P 
         Alignment       =   1  'Right Justify
         Caption         =   "TORON/P"
         Height          =   315
         Left            =   -72960
         TabIndex        =   32
         Top             =   1260
         Width           =   1095
      End
      Begin OWC.Spreadsheet Spreadsheet6 
         Height          =   8490
         Left            =   -74760
         TabIndex        =   10
         Top             =   1620
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":118CA
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
         Left            =   -59400
         TabIndex        =   30
         Top             =   780
         Width           =   1815
      End
      Begin VB.TextBox DIAMETRE 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -61800
         TabIndex        =   28
         Top             =   780
         Width           =   1455
      End
      Begin VB.CheckBox ACTIVER 
         Alignment       =   1  'Right Justify
         Caption         =   "ACTIVER"
         Height          =   315
         Left            =   -72960
         TabIndex        =   27
         Top             =   780
         Width           =   1095
      End
      Begin OWC.Spreadsheet Spreadsheet3 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":12324
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
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":12AB8
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
         Picture         =   "AutoCable.frx":1330E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Modifier"
         Top             =   780
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   -73800
         Picture         =   "AutoCable.frx":13ABC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Supprimer"
         Top             =   780
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   -74880
         Picture         =   "AutoCable.frx":142B6
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ajouter"
         Top             =   780
         Width           =   495
      End
      Begin VB.ComboBox ENC 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         ItemData        =   "AutoCable.frx":14B2C
         Left            =   -61800
         List            =   "AutoCable.frx":14B2E
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   1260
         Width           =   1455
      End
      Begin VB.ComboBox PSA 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -64320
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   1260
         Width           =   1455
      End
      Begin VB.ComboBox RSA 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -66840
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   1260
         Width           =   1215
      End
      Begin VB.ComboBox Hab 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         ItemData        =   "AutoCable.frx":14B30
         Left            =   -70560
         List            =   "AutoCable.frx":14B32
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   1260
         Width           =   2535
      End
      Begin VB.TextBox Longueur 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -66840
         TabIndex        =   18
         Top             =   780
         Width           =   1215
      End
      Begin OWC.Spreadsheet Spreadsheet4 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":14B34
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
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":152C3
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
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":15A51
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
      Begin OWC.Spreadsheet Spreadsheet7 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   39
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":16265
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
      Begin OWC.Spreadsheet Spreadsheet8 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":16CBF
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
      Begin OWC.Spreadsheet Spreadsheet9 
         Height          =   9720
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   17220
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable.frx":17719
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "OPTION"
         Height          =   315
         Left            =   -60240
         TabIndex        =   36
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "LONGUEUR/C"
         Height          =   315
         Left            =   -65520
         TabIndex        =   34
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CLASSE_T"
         Height          =   315
         Left            =   -60240
         TabIndex        =   31
         Top             =   780
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DIAMETRE"
         Height          =   315
         Left            =   -62760
         TabIndex        =   29
         Top             =   780
         Width           =   840
      End
      Begin VB.Label NOUED 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   315
         Left            =   -70560
         TabIndex        =   17
         Top             =   780
         Width           =   2535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CODE.ENC."
         Height          =   255
         Left            =   -62760
         TabIndex        =   16
         Top             =   1260
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CODE.PSA."
         Height          =   315
         Left            =   -65520
         TabIndex        =   15
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CODE.RSA."
         Height          =   315
         Left            =   -67920
         TabIndex        =   14
         Top             =   1260
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DESIGN.HAB."
         Height          =   315
         Left            =   -71760
         TabIndex        =   13
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "LONGUEUR"
         Height          =   315
         Left            =   -67920
         TabIndex        =   12
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NOEUDS"
         Height          =   315
         Left            =   -71760
         TabIndex        =   11
         Top             =   780
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
Dim msg As String
Dim MyErr As Boolean
Dim IfValidationOk As Boolean
Dim NbFinOuiNon As Long
Dim NoMacro1Change As Boolean
Dim NoMacro1Select As Boolean
Dim NoMacro2 As Boolean
Dim NoMacro3 As Boolean
Dim NoMacro4 As Boolean
Dim NoMacro5 As Boolean
Dim NoMacro5Select As Boolean
Dim CollecCrieres As Collection
Dim CollecCrieresCode As Collection
Dim CollecCrieresDesigne As Collection

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
Dim IdProjet As Long
Dim OngletName As String
Dim CollecApp As Collection








Private Sub Command3_Click()
If boolActu = False Then
    MsgBox "Il est impossible de valide l'étude si un test de d'actualisation na pas été effectué."
    Exit Sub
End If
If Trim(msg) <> "" Then
    MsgBox "Il est impossible de valide l'étude si le test de validation présente des erreurs."
    Exit Sub
End If

Dim MyExcel As EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim MyRange2
Set MyExcel = New EXCEL.Application
MyExcel.DisplayAlerts = False
Dim MyTim
Dim BoolErr As Boolean
BoolErr = False
Dim Fso As New FileSystemObject
   
    
'MyExcel.Visible = True
'If Nouveau = False Then
'    Set MyWorkbook = MyExcel.Workbooks.Open(Me.Caption)
'Else
    If Fso.FileExists(Me.Caption) Then Fso.DeleteFile (Me.Caption)
    DoEvents
    Set MyWorkbook = MyExcel.Workbooks.Add
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
    MyExcel.DisplayAlerts = False
    Set MyWorkbook = MyExcel.Workbooks.Open(Me.Caption)
'    MyExcel.Visible = True
'End If
'MyExcel.Visible = True
'MyWorkbook.Application.Visible = True
IsertSheet MyWorkbook, "Nomenclature Habillage"
   
 Set Myrange = MyWorkbook.Worksheets("Nomenclature Habillage").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Nomenclature Habillage").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Nomenclature Habillage").Select
MyWorkbook.Worksheets("Nomenclature Habillage").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet9
Set MyRange2 = Me.Spreadsheet9.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Nomenclature Habillage").Range("a1").Select
MyWorkbook.Worksheets("Nomenclature Habillage").Paste
ReplaceNull MyWorkbook.Worksheets("Nomenclature Habillage"), "©", Chr(10)

'MyWorkbook.Application.Visible = True
IsertSheet MyWorkbook, "Nomenclature Fils"
MyWorkbook.Worksheets("Nomenclature Fils").Select
Set Myrange = MyWorkbook.Worksheets("Nomenclature Fils").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Nomenclature Fils").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Nomenclature Fils").Select

MyWorkbook.Worksheets("Nomenclature Fils").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet8
Set MyRange2 = Me.Spreadsheet8.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Nomenclature Fils").Range("a1").Select
MyWorkbook.Worksheets("Nomenclature Fils").Paste
ReplaceNull MyWorkbook.Worksheets("Nomenclature Fils"), "©", Chr(10)

   

IsertSheet MyWorkbook, "Nomenclature Connecteur"

Set Myrange = MyWorkbook.Worksheets("Nomenclature Connecteur").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Nomenclature Connecteur").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Nomenclature Connecteur").Select
MyWorkbook.Worksheets("Nomenclature Connecteur").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet7
Set MyRange2 = Me.Spreadsheet7.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Nomenclature Connecteur").Range("a1").Select
MyWorkbook.Worksheets("Nomenclature Connecteur").Paste
ReplaceNull MyWorkbook.Worksheets("Nomenclature Connecteur"), "©", Chr(10)

IsertSheet MyWorkbook, "NOEUDS"
Set Myrange = MyWorkbook.Worksheets("NOEUDS").Range("a1").CurrentRegion
MyWorkbook.Worksheets("NOEUDS").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("NOEUDS").Select
MyWorkbook.Worksheets("NOEUDS").Range("a1").Select
'Me.Spreadsheet6.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet6
Set MyRange2 = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("NOEUDS").Paste




IsertSheet MyWorkbook, "Critères"
Set Myrange = MyWorkbook.Worksheets("Critères").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Critères").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Critères").Select
MyWorkbook.Worksheets("Critères").Range("a1").Select
'MyWorkbook.Application.Visible = True
'Me.Spreadsheet5.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet5
Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range(MyRange2(1, 1).Address & ":" & MyRange2(MyRange2.Rows.Count, 4).Address)
MyRange2.Copy
MyWorkbook.Worksheets("Critères").Paste
'MyWorkbook.Application.Visible = True
IsertSheet MyWorkbook, "Notas"
Set Myrange = MyWorkbook.Worksheets("Notas").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Notas").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Notas").Select
MyWorkbook.Worksheets("Notas").Range("a1").Select
'Me.Spreadsheet4.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet4
Set MyRange2 = Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Notas").Paste
MyWorkbook.Worksheets("Notas").Range("a1").Select

IsertSheet MyWorkbook, "Composants"
Set Myrange = MyWorkbook.Worksheets("Composants").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Composants").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Composants").Select
MyWorkbook.Worksheets("Composants").Range("a1").Select
'Me.Spreadsheet3.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet3
Set MyRange2 = Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Composants").Paste
MyWorkbook.Worksheets("Composants").Range("a1").Select

IsertSheet MyWorkbook, "Ligne_Tableau_fils"
Set Myrange = MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range(Myrange(12, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Ligne_Tableau_fils").Select
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").Select
'Me.Spreadsheet2.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet2
Set MyRange2 = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Ligne_Tableau_fils").Paste





IsertSheet MyWorkbook, "Connecteurs"
MyWorkbook.Worksheets("Connecteurs").Range("a1").Select
Set Myrange = MyWorkbook.Worksheets("Connecteurs").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Connecteurs").Range(Myrange(1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Connecteurs").Select
MyWorkbook.Worksheets("Connecteurs").Range("a1").Select
'Me.Spreadsheet1.ActiveSheet.AutoFilter = False
RazFiltreEditExcel Me.Spreadsheet1
Set MyRange2 = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy

'IsertSheet MyWorkbook, "Connecteurs"

MyWorkbook.Worksheets("Connecteurs").Paste





MyWorkbook.Worksheets("Connecteurs").Range("a1").Select













 MyWorkbook.Save
Fin:
 
 
 Set MyRange2 = Nothing
 Set Myrange = Nothing
 MyWorkbook.Close False
 Set MyWorkbook = Nothing
 MyExcel.Quit
 Set MyExcel = Nothing
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
Set CollecApp = Nothing
Set CollecApp = New Collection
Set CollecCrieres = Nothing
Set CollecCrieresCode = Nothing
Set CollecCrieresDesigne = Nothing

Set CollecCrieres = New Collection
Set CollecCrieresCode = New Collection
Set CollecCrieresDesigne = New Collection

Me.Spreadsheet5.Cells(1, 1).Select
Sql = "DELETE Ajout_LIAISON_CONNECTEURS.* FROM Ajout_LIAISON_CONNECTEURS;"
Con.Exequte Sql
Sql = "DELETE Ajout_LIAISON.* FROM Ajout_LIAISON;"
Con.Exequte Sql

msg = ""
DoEvents
IfValidationOk = True
RazFiltreEditExcel Me.Spreadsheet5
Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
SSTab1.Tab = 0
Me.Spreadsheet5.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
IfValidationOk = True
Me.Spreadsheet5.Cells(I, 1).Select
ConverOuiNon Myrange, I
If msg <> "" Then
    IfValidationOk = False
'    Me.Spreadsheet5.ActiveSheet.AutoFilter = True
    Exit Sub
End If
    Me.Spreadsheet5.Cells(I, 1).Value = Me.Spreadsheet5.Cells(I, 1).Value
If msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next I
RazFiltreEditExcel Me.Spreadsheet1
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
OngletName = "Connecteur"
SSTab1.Tab = 1
DoEvents
Me.Spreadsheet1.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
IfValidationOk = True
Me.Spreadsheet1.Cells(I, 1).Select
ConverOuiNon Myrange, I
If msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
    Me.Spreadsheet1.Cells(I, 1).Value = Me.Spreadsheet1.Cells(I, 1).Value
If msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next I
RazFiltreEditExcel Me.Spreadsheet2
Set Myrange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion

SSTab1.Tab = 2
DoEvents
Me.Spreadsheet2.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
Me.Spreadsheet2.Cells(I, 15).Select
ConverOuiNon Myrange, I
IfValidationOk = True
    Me.Spreadsheet2.Cells(I, 15).Value = UCase("'" & Me.Spreadsheet2.Cells(I, 15).Value)
If msg <> "" Then
'Me.Spreadsheet2.ActiveSheet.AutoFilter = True
    IfValidationOk = False
    Exit Sub
End If
Me.Spreadsheet2.Cells(I, 25).Select
    Me.Spreadsheet2.Cells(I, 25).Value = Me.Spreadsheet2.Cells(I, 25).Value
    If msg <> "" Then
'    Me.Spreadsheet2.ActiveSheet.AutoFilterMode = True

        IfValidationOk = False
        Exit Sub
    End If
DoEvents
Next I
'Me.Spreadsheet5.ActiveSheet.AutoFilter = False
SSTab1.Tab = 3
RazFiltreEditExcel Me.Spreadsheet3
Set Myrange = Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion

Me.Spreadsheet3.Cells(1, 2).Select
For I = 2 To Myrange.Rows.Count
Me.Spreadsheet3.Cells(I, 2).Select
ConverOuiNon Myrange, I
Me.Spreadsheet3.Cells(I, 3) = 0
IfValidationOk = True
 If msg <> "" Then Exit Sub
DoEvents
Next I

SSTab1.Tab = 4
RazFiltreEditExcel Me.Spreadsheet4
Set Myrange = Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion
SSTab1.Tab = 4
Me.Spreadsheet4.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
Me.Spreadsheet4.Cells(I, 3).Select
ConverOuiNon Myrange, I
Me.Spreadsheet4.Cells(I, 3) = I - 1
IfValidationOk = True
 
DoEvents
Next I
RazFiltreEditExcel Me.Spreadsheet6
Set Myrange = Me.Spreadsheet6.ActiveSheet.Range("a1").CurrentRegion
SSTab1.Tab = 5
Me.Spreadsheet6.Cells(1, 1).Select
For I = 2 To Myrange.Rows.Count
IfValidationOk = True
ConverOuiNon Myrange, I
Me.Spreadsheet6.Cells(I, 1).Select

DoEvents
If msg <> "" Then
'Me.Spreadsheet6.ActiveSheet.AutoFilter = True
    IfValidationOk = False
    Exit Sub
End If
'If Spreadsheet6.Cells(i, 4) = "x" Then
'    MsgBox ""
'End If
    Command7_Click
If msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
DoEvents
Next I


'Me.Spreadsheet6.ActiveSheet.AutoFilter = True
If MyErr = True Then
    LoadLiasons.Charger MyClient
    Unload LoadLiasons
End If
MyErr = False
    IfValidationOk = False
    boolActu = True

End Sub
Private Sub Command2_Click()
Command1_Click
If msg <> "" Then Exit Sub
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
        Myrange(Myrange.Rows.Count + 1, 13) = "'" & Me.txtOption
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
Me.txtOption = ""
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
        If InStr("" & Me.txtOption, "TOUS") <> 0 Then
            If Len("" & Me.txtOption) > Len("TOUS;") Then
                Myrange(Val(Me.Tag), 13) = "'" & Replace(Me.txtOption, "TOUS", "")
                Myrange(Val(Me.Tag), 13) = Replace(Myrange(Val(Me.Tag), 13), ";;", ";")
                If Left("" & Myrange(Val(Me.Tag), 13), 1) = ";" Then Myrange(Val(Me.Tag), 13) = Right(Myrange(Val(Me.Tag), 13), Len(Myrange(Val(Me.Tag), 13)) - 1)
                
            Else
                Myrange(Val(Me.Tag), 13) = "'" & Me.txtOption
            End If
        Else
         Myrange(Val(Me.Tag), 13) = "'" & Me.txtOption
        End If
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
Me.txtOption = ""
End Sub



Private Sub Command8_Click()
Me.txtOption = FrmSelectCriteres.Chargement(Spreadsheet5, Me.txtOption)
Unload FrmSelectCriteres
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
Static SaveRow As Long

Dim LibCode_APP As String
Dim Sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet1.ActiveCell.Row
Col = Me.Spreadsheet1.ActiveCell.Column

If (NoMacro1Change = True Or NoMacro1Select = True) Or (Row = 1) Then GoTo Fin
boolActu = False
NoMacro1Change = True
Set Myrange = Me.Spreadsheet1.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
   Me.Spreadsheet1.Cells(Row, 12) = UCase("" & Me.Spreadsheet1.Cells(Row, 12))
   Me.Spreadsheet1.Cells(Row, 2) = UCase("'" & Me.Spreadsheet1.Cells(Row, 2))
    Me.Spreadsheet1.Cells(Row, 6) = UCase("'" & Me.Spreadsheet1.Cells(Row, 6))
If Trim("" & Me.Spreadsheet1.Cells(Row, 12)) <> "" Then
    If UCase(Me.Spreadsheet1.Cells(Row, 12)) <> "TOUS" Then
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("B1", "B" & CStr(Myrange.Rows.Count))
        Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
         Set MyRange2 = Me.Spreadsheet5.ActiveSheet.Range("A1", "A" & CStr(MyRange2.Rows.Count))
        aa = Split(UCase(Trim("" & Me.Spreadsheet1.Cells(Row, 12))) & ";", ";")
        For Iaa = 0 To UBound(aa) - 1
        I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))), MyRange2, True)
        
        If I = 0 Then
            
           
            msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbExclamation
             Me.Spreadsheet1.Cells(Row, 12) = Replace(Me.Spreadsheet1.Cells(Row, 12) & ";", aa(Iaa) & ";", "")
             If Right(Me.Spreadsheet1.Cells(Row, 12), 1) = ";" Then Me.Spreadsheet1.Cells(Row, 12) = Left(Me.Spreadsheet1.Cells(Row, 12), Len(Me.Spreadsheet1.Cells(Row, 12)))
                 Spreadsheet1.SetFocus
        End If
        Next
     Set Myrange = Nothing
     End If
     
End If




    
        If Trim("" & Me.Spreadsheet1.Cells(Row, 2)) <> "" Then
            Me.Spreadsheet1.Cells(Row, 7) = Row - 1
        End If
        
    
   
   
 NoMacro1Change = False
    Col3 = 0
    SaveRow = Row
    
Fin:
DoEvents
End Sub

Private Sub Spreadsheet1_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Static SaveRow As Long
Dim Row As Long
Dim Sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet1.ActiveCell.Row
DoEvents
DoEvents

If Row = 1 Then GoTo Fin
If SaveRow = 0 Then SaveRow = Row
If NoMacro1Select = True Then GoTo Fin
    NoMacro1Select = True
    
    If IfValidationOk = True Then 'NEANT
           
            If (Me.Spreadsheet1.Cells(Row, 2)) <> "" And Me.Spreadsheet1.Cells(Row, 6) = "" Then
            If (Me.Spreadsheet1.Cells(Row, 1)) <> 0 Then
                Me.Spreadsheet1.Cells(Row, 6).Select
                
                MsgBox "Vous devez saisir le Code Appareil", vbQuestion, "AutoCâble: Connecteurs"
                msg = "?"
                Spreadsheet1.SetFocus
                End If
            End If
        Else
            If (SaveRow <> Row) And (Me.Spreadsheet1.Cells(SaveRow, 2)) <> "" And Me.Spreadsheet1.Cells(SaveRow, 5) = "" Then
            If (UCase(Me.Spreadsheet1.Cells(SaveRow, 1))) <> 0 Then
                Me.Spreadsheet1.Cells(SaveRow, 5).Select
                
                MsgBox "Vous devez saisir le Code Appareil", vbQuestion, "AutoCâble: Connecteurs"
                msg = "?"
                Spreadsheet1.SetFocus
            Else
             Me.Spreadsheet1.Cells(SaveRow, 2) = UCase(Me.Spreadsheet1.Cells(SaveRow, 2))
            End If
            End If
    End If
    If IfValidationOk = True Then
    If SaveRow <> Row Then
        On Error Resume Next
        If Me.Spreadsheet1.Cells(Row, 1) = 1 Then
        aa = ""
        aa = CollecApp(Me.Spreadsheet1.Cells(Row, 6))
        Debug.Print Me.Spreadsheet1.Cells(Row, 6) & " L" & SaveRow & "c6"
        If Err = 0 Then
        
           MsgBox " Le code appareil : " & CollecApp(Me.Spreadsheet1.Cells(Row, 6)) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Connecteurs"
            msg = "?"
            Me.Spreadsheet1.Cells(Row, 6).Select
            Me.Spreadsheet1.SetFocus
        Else
        Err.Clear
            CollecApp.Add Me.Spreadsheet1.Cells(Row, 6), Me.Spreadsheet1.Cells(Row, 6)
           
        End If
        End If
        On Error GoTo 0
      End If
    End If
If Trim("" & Me.Spreadsheet1.Cells(SaveRow, 6)) <> "" Then
        
            Sql = "SELECT LIAISON_CONNECTEURS.LIB FROM LIAISON_CONNECTEURS "
            Sql = Sql & "WHERE LIAISON_CONNECTEURS.CLIENT='" & MyReplace(MyClient) & "' "
            Sql = Sql & "AND LIAISON_CONNECTEURS.LIAISON='" & MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(SaveRow, 6))) & "';"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = False Then
                Me.Spreadsheet1.Cells(SaveRow, 5) = Trim("'" & Rs!Lib)
            Else
                If IfValidationOk = False Then
                    If SaveRow <> Row Then
                    DoEvents
'                        If MsgBox("Le code App : " & DecodeCode_APP(Me.Spreadsheet1.Cells(SaveRow, 6)) & " n'existe pas" & vbCrLf & "Voulez-vous le créer", vbYesNo + vbQuestion, "AutoCâble: Connecteurs") = vbYes Then
'                            LibCode_APP = InputBox("Entrez la désignation du code APP : " & DecodeCode_APP(Me.Spreadsheet1.Cells(SaveRow, 6)), "Ajout d'un code App")
''                            If Trim(LibCode_APP) <> "" Then
''                                Me.Spreadsheet1.Cells(SaveRow, 5) = LibCode_APP
''                                sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
''                                sql = sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(SaveRow, 6)))) & "', '" & UCase(MyReplace("" & LibCode_APP)) & "' );"
''                                Con.Exequte sql
''                            End If
'                        End If
                        
                   End If
                Else
                   Sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON "
                   Sql = Sql & "FROM Ajout_LIAISON_CONNECTEURS "
                   Sql = Sql & "WHERE Ajout_LIAISON_CONNECTEURS.LIAISON='" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(SaveRow, 6)))) & "' "
                   Sql = Sql & "AND Ajout_LIAISON_CONNECTEURS.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(Sql)
                    If Rs.EOF = True Then
                        Sql = "INSERT INTO Ajout_LIAISON_CONNECTEURS ( LIAISON, LIB,Job ) "
                        Sql = Sql & "values ( '" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(SaveRow, 6)))) & "', '" & MyReplace(Me.Spreadsheet1.Cells(SaveRow, 5)) & "'," & NmJob & ");"
                        Con.Exequte Sql
                        MyErr = True
                    End If
                End If
            End If
            Set Rs = Con.CloseRecordSet(Rs)
        SaveRow = Row



        End If
NoMacro1Select = False
Fin:
End Sub

Private Sub Spreadsheet2_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Static SaveRow As Long
Dim Col As Long
Dim Myrange
Dim Rs As Recordset
Dim Sql As String
Dim LibCode_APP As String
'Dim TrouveConnecteur() As Boolean
Dim boolReprise As Boolean
Static Col3 As Long
Row = Me.Spreadsheet2.ActiveCell.Row
Col = Me.Spreadsheet2.ActiveCell.Column
If SaveRow = 0 Then SaveRow = 1
If (NoMacro2 = True) Or (Row = 1) Then GoTo Fin
RepriseCritaire:
boolActu = False
NoMacro2 = True
 Set Myrange = Me.Spreadsheet2.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
If (Trim("" & Me.Spreadsheet2.Cells(Row, 15)) <> "") And (Trim("" & Me.Spreadsheet2.Cells(Row, 25)) <> "") Then
'   If Trim("" & Me.Spreadsheet2.Cells(Row, 15)) = "186.AA" Then
'   MsgBox "186.AA"
'End If
' If Trim("" & Me.Spreadsheet2.Cells(Row, 20)) = "186.AA" Then
'   MsgBox "186.AA"
'End If
Reprise:
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("F1", "F" & CStr(Myrange.Rows.Count))


        I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 15))))

        If I <> 0 Then
                aa = Split(Me.Spreadsheet1.Cells(I, 12) & ";", ";")
'                ReDim TrouveConnecteur(UBound(aa))
                
                For Iaa = 0 To UBound(aa)
                     If InStr(1, Me.Spreadsheet2.Cells(Row, 33) & ";", "" & Trim(aa(Iaa)) & ";") = 0 Then
                         Me.Spreadsheet2.Cells(Row, 33) = "" & Me.Spreadsheet2.Cells(Row, 33) & ";" & Trim(aa(Iaa))
                        
                   
                     End If
                Next
        End If
        
             I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 25))))

        If I <> 0 Then
             aa = Split("" & Me.Spreadsheet1.Cells(I, 12) & ";", ";")
'             ic = UBound(TrouveConnecteur)
'             ReDim Preserve TrouveConnecteur(ic + UBound(aa))
                For Iaa = 0 To UBound(aa)
                     If InStr(1, Me.Spreadsheet2.Cells(Row, 33) & ";", "" & Trim(aa(Iaa)) & ";") = 0 Then
                       
                         Me.Spreadsheet2.Cells(Row, 33) = "" & Me.Spreadsheet2.Cells(Row, 33) & ";" & Trim(aa(Iaa))
                        If InStr(1, Me.Spreadsheet2.Cells(Row, 33) & ";", "TOUS;") <> 0 Then
                            Me.Spreadsheet2.Cells(Row, 33) = Replace(Me.Spreadsheet2.Cells(Row, 33) & ";", "TOUS;", "")
                        End If
                       
                          
                     End If
                Next

       
        End If
        If Left(Me.Spreadsheet2.Cells(Row, 33), 1) = ";" Then Me.Spreadsheet2.Cells(Row, 33) = Right(Me.Spreadsheet2.Cells(Row, 33), Len(Me.Spreadsheet2.Cells(Row, 33)) - 1)
        If Right(Me.Spreadsheet2.Cells(Row, 33), 1) = ";" Then Me.Spreadsheet2.Cells(Row, 33) = Left(Me.Spreadsheet2.Cells(Row, 33), Len(Me.Spreadsheet2.Cells(Row, 33)) - 1)
        If UBound(Split(Me.Spreadsheet2.Cells(Row, 33) & ";", ";")) > 2 Then
        
            MsgBox "Vous ne pouvez pas saisir plus de deux options." & vbCrLf & vbCrLf & Me.Spreadsheet2.Cells(Row, 33), vbQuestion, "AutoCâble: Tableau de fils"
          msg = "?"
          Me.Spreadsheet2.Cells(Row, 33) = ""
          If boolReprise = False Then
            boolReprise = True
            GoTo Reprise
            End If
        End If
        If InStr(1, UCase(";" & Me.Spreadsheet2.Cells(Row, 33)), "TOUS;") <> 0 Then
            If Len("" & Me.Spreadsheet2.Cells(Row, 33)) > Len("TOUS") Then Me.Spreadsheet2.Cells(Row, 33) = Replace(UCase("" & Me.Spreadsheet2.Cells(Row, 33)), "TOUS", "")
                
            
        End If
        Me.Spreadsheet2.Cells(Row, 33) = Replace(Me.Spreadsheet2.Cells(Row, 33), ";;", "")
        If Right(Me.Spreadsheet2.Cells(Row, 33), 1) = ";" Then Me.Spreadsheet2.Cells(Row, 33) = Left(Me.Spreadsheet2.Cells(Row, 33), Len(Me.Spreadsheet2.Cells(Row, 33)) - 1)
          If Left(Me.Spreadsheet2.Cells(Row, 33), 1) = ";" Then Me.Spreadsheet2.Cells(Row, 33) = Right(Me.Spreadsheet2.Cells(Row, 33), Len(Me.Spreadsheet2.Cells(Row, 33)) - 1)
          
     Set Myrange = Nothing
      If Trim("" & Me.Spreadsheet2.Cells(Row, 33)) = "" Then
        Me.Spreadsheet2.Cells(Row, 33) = "TOUS" ' MsgBox "Vous devez saisir un code critère."
     
     
      End If
     End If
     If (Trim("" & Me.Spreadsheet2.Cells(Row, 15)) <> "") And (Trim("" & Me.Spreadsheet2.Cells(Row, 25)) <> "") And (Trim(UCase("" & Me.Spreadsheet2.Cells(Row, 33))) = "TOUS") Then
   
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("F1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("F1", "F" & CStr(Myrange.Rows.Count))


        I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 15))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, 12)) <> "" Then
                Me.Spreadsheet2.Cells(Row, 33) = "" & Me.Spreadsheet1.Cells(I, 12)
           End If

        End If
        If UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 33))) = "TOUS" Then
             I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 25))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, 12)) <> "" Then
                Me.Spreadsheet2.Cells(Row, 33) = "" & Me.Spreadsheet1.Cells(I, 12)
           End If

        End If
        End If
     Set Myrange = Nothing
     
     End If

 If (Trim("" & Me.Spreadsheet2.Cells(Row, 34)) <> "") Then
Me.Spreadsheet2.Cells(Row, 33) = UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 34)))
    If (Trim("" & Me.Spreadsheet2.Cells(Row, 33)) <> "TOUS") Then
 Me.Spreadsheet2.Cells(Row, 33) = UCase(Me.Spreadsheet2.Cells(Row, 33))
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        aa = Split(Me.Spreadsheet2.Cells(Row, 33) & ";", ";")
        For Iaa = 0 To UBound(aa) - 1
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("B1", "B" & CStr(Myrange.Rows.Count))
        
        
        I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))))
        If UCase(aa(Iaa)) <> "TOUS" Then
        If I = 0 Then
            
           
            msg = "CODE CRITERE : " & UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 33))) & " introuvable"
             MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
            Me.Spreadsheet2.Cells(Row, 34) = ""
           
             Me.Spreadsheet2.Cells(Row, 33) = Replace(Me.Spreadsheet2.Cells(Row, 33) & ";", aa(Iaa) & ";", "")
                 If Right(Me.Spreadsheet2.Cells(Row, 33).Value, 1) = ";" Then Me.Spreadsheet2.Cells(Row, 33) = Left(Me.Spreadsheet2.Cells(Row, 33), Len(Me.Spreadsheet2.Cells(Row, 33)) - 1)
        End If
        End If
        Next
        If Trim("" & Me.Spreadsheet2.Cells(Row, 33)) = "" Then Me.Spreadsheet2.Cells(Row, 33) = "TOUS"
     Set Myrange = Nothing
     End If
 
      End If

If InStr(1, UCase(Me.Spreadsheet2.Cells(Row, 33)) & ";", "TOUS;") <> 0 And Len(Trim(Me.Spreadsheet2.Cells(Row, 33)) & ";") > Len("TOUS;") Then Me.Spreadsheet2.Cells(Row, 33) = Replace(Me.Spreadsheet2.Cells(Row, 33), "TOUS", "")

Me.Spreadsheet2.Cells(Row, 33) = Replace(Me.Spreadsheet2.Cells(Row, 33), ";;", "")

If Right(Me.Spreadsheet2.Cells(Row, 33), 1) = ";" Then Me.Spreadsheet2.Cells(Row, 33) = Left(Me.Spreadsheet2.Cells(Row, 33), Len(Me.Spreadsheet2.Cells(Row, 33)) - 1)
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("F1", "F" & CStr(Myrange.Rows.Count))


        I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 15))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, 12)) <> "" Then
                Me.Spreadsheet2.Cells(Row, 33) = "" & Me.Spreadsheet1.Cells(I, 12)
           End If

        End If
        If UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 33))) = "TOUS" Then
             I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 25))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, 12)) <> "" Then
                Me.Spreadsheet2.Cells(Row, 33) = "" & Me.Spreadsheet1.Cells(I, 12)
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
If (Col = 14) Or (Col = 21) Then
  Col = Col + 1
End If
' If Trim("" & Me.Spreadsheet2.Cells(Row, 17)) = "" Then
'        Msg = "le champ REFC/L est obligatoire"
'        MsgBox Msg, vbExclamation
'        Me.Spreadsheet2.Cells(Row - 1, 17).Select
' Else
'    If Trim("" & Me.Spreadsheet2.Cells(Row, 24)) = "" Then
'        Msg = "le champ REFC/L2 est obligatoire"
'        MsgBox Msg, vbExclamation
'        Me.Spreadsheet2.Cells(Row - 1, 24).Select
'    End If
'End If
If (Col = 15) Or (Col = 25) Then
   
        If Trim("" & Me.Spreadsheet2.Cells(SaveRow, Col)) <> "" Then
            Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
            Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("F1", "F" & CStr(Myrange.Rows.Count))
            NoMacro2 = True
            
            I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))))
        
            If I <> 0 Then
               
                Me.Spreadsheet2.Cells(Row, Col - 1) = UCase(Trim("'" & Myrange(I, 2)))
                Me.Spreadsheet2.Cells(Row, Col - 2) = UCase(Trim("'" & Myrange(I, 4)))
                Me.Spreadsheet2.Cells(Row, Col - 3) = UCase(Trim("'" & Myrange(I, 3)))
                Me.Spreadsheet2.Cells(Row, Col + 2) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, 2)))
                If Trim("" & Me.Spreadsheet2.Cells(Row, Col + 2)) = "" Then
                    Me.Spreadsheet2.Cells(Row, Col + 2) = UCase(Trim("'" & Me.Spreadsheet1.ActiveSheet.Cells(I, 3)))
                 End If
                Else
                    Me.Spreadsheet2.Cells(Row, Col - 1) = "0"
                    Me.Spreadsheet2.Cells(Row, Col - 2) = ""
                    
                    msg = "Le connecteur : " & UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))) & " introuvable"
                    MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                     Me.Spreadsheet2.Cells(SaveRow, Col).Select
                      Me.Spreadsheet2.SetFocus
                End If
            
            
            
            
                
                
                    
                End If
            End If
        
        
       
        
           
            
            Col3 = 0
        End If
   If (Trim("" & Me.Spreadsheet2.Cells(Row, 27)) <> "") Then

 Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("F1", "F" & CStr(Myrange.Rows.Count))


        I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 15))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, 12)) <> "" Then
                Myapp1 = "" & Me.Spreadsheet1.Cells(I, 12)
           End If

        End If
       
             I = ChercheXls(Myrange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 25))))

        If I <> 0 Then
            If Trim(Me.Spreadsheet1.Cells(I, 12)) <> "" Then
               Myapp2 = "" & Me.Spreadsheet1.Cells(I, 12)
           End If
    If Trim("" & Myapp1) = "" Then Myapp1 = Myapp2
     If Trim(Myapp2) = "" Then Myapp2 = Myapp1
        If ("" & Myapp1 <> Myapp2) And (UCase(Myapp1) <> "TOUS") And (UCase(Myapp2) <> "TOUS") Then
        MsgBox "Une liaison ne peut pas pointer sur deux options différentes : " & Myapp1 & " & " & Myapp2, vbQuestion, "AutoCâble: Tableau de fils"
        Me.Spreadsheet2.Cells(Row, 33) = ""
        Spreadsheet2.SetFocus
        msg = "?"
        End If
        End If
        
If (Trim("" & Me.Spreadsheet2.Cells(Row, 28)) <> "") Then
        Me.Spreadsheet2.Cells(Row, 28) = UCase(Me.Spreadsheet2.Cells(Row, 28))
    Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
    Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("B1", "B" & CStr(Myrange.Rows.Count))

    aa = Split(Me.Spreadsheet2.Cells(Row, 28) & ";", ";")
    If UBound(aa) = 3 Then
        MsgBox "Vous ne pouvez pas saisir plus de deux critére." & vbCrLf & vbCrLf & Me.Spreadsheet2.Cells(Row, 28), vbQuestion, "AutoCâble: Tableau de fils"
        Me.Spreadsheet2.Cells(Row, 33) = "TOUS"
         Me.Spreadsheet2.Cells(Row, 34) = ""
         GoTo RepriseCritaire
         
         
    Else
        zz = ""
        For Iaa = 0 To UBound(aa) - 1
            I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))))
            If I <> 0 Then
                If Me.Spreadsheet5.ActiveSheet.Cells(I, 1) = 0 Then GoTo ErrorCritere
                If InStr(1, ";" & zz & ";", ";" & aa(Iaa) & ";") = 0 Then
                    zz = "" & zz & Trim(aa(Iaa)) & ";"
                End If
            Else
ErrorCritere:
                MsgBox "Code Critère: " & aa(Iaa) & " introuvable ou Inactif", vbExclamation
            End If
        Next
        Me.Spreadsheet2.Cells(Row, 33) = zz
         
        If Right(Me.Spreadsheet2.Cells(Row, 33), 1) = ";" Then Me.Spreadsheet2.Cells(Row, 33) = Left(Me.Spreadsheet2.Cells(Row, 33), Len(Me.Spreadsheet2.Cells(Row, 33)) - 1)
         Me.Spreadsheet2.Cells(Row, 28).Value = Me.Spreadsheet2.Cells(Row, 33).Value
         If Trim("" & (Me.Spreadsheet2.Cells(Row, 33))) = "" Then GoTo RepriseCritaire
    End If
End If
Set Myrange = Nothing
End If
SaveRow = Row
    NoMacro2 = False
Fin:
End Sub

Private Sub Spreadsheet2_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Static SaveRow As Long
Dim Row As Long
Dim Sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet2.ActiveCell.Row
If (NoMacro2 = True) Or (Row = 1) Then GoTo Fin
If SaveRow = 0 Then SaveRow = Me.Spreadsheet2.ActiveCell.Row
 If Trim("" & Me.Spreadsheet2.Cells(SaveRow, 2)) <> "" Then
                Sql = "SELECT LIAISON.LIB FROM LIAISON "
                Sql = Sql & "WHERE LIAISON.CLIENT='" & MyReplace(MyClient) & "' "
                Sql = Sql & "AND LIAISON.LIAISON='" & MyReplace(Me.Spreadsheet2.Cells(SaveRow, 2)) & "';"
                Set Rs = Con.OpenRecordSet(Sql)
                If Rs.EOF = False Then
                Me.Spreadsheet2.Cells(SaveRow, 3) = Trim("'" & Rs!Lib)
                Else
                    If IfValidationOk = False Then
                        If SaveRow <> Row And SaveRow <> 1 Then
'                            If MsgBox("La liaison : " & Me.Spreadsheet2.Cells(SaveRow, 2) & " n'existe pas" & vbCrLf & "Voulez-vous la créer", vbYesNo + vbQuestion, "AutoCâble: Tableau de fils") = vbYes Then
'                                LibCode_APP = InputBox("Entrez la désignation de la liaison : " & Me.Spreadsheet2.Cells(SaveRow, 1), "Ajout de liaison")
''                                If Trim(LibCode_APP) <> "" Then
''                                    Me.Spreadsheet2.Cells(SaveRow, 3) = LibCode_APP
''                                    sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
''                                    sql = sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(Me.Spreadsheet2.Cells(SaveRow, 2))) & "', '" & UCase(MyReplace("" & LibCode_APP)) & "' );"
''                                    Con.Exequte sql
''                                End If
'                            End If
                        End If
                        Else
                        
                   Sql = "SELECT Ajout_LIAISON.LIAISON "
                   Sql = Sql & "FROM Ajout_LIAISON "
                   Sql = Sql & "WHERE Ajout_LIAISON.LIAISON='" & UCase(MyReplace(Me.Spreadsheet2.Cells(SaveRow, 2))) & "' "
                   Sql = Sql & "AND Ajout_LIAISON.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(Sql)
                    If Rs.EOF = True Then
                        Sql = "INSERT INTO Ajout_LIAISON ( LIAISON, LIB,Job ) "
                        Sql = Sql & "values ( '" & UCase(MyReplace(Me.Spreadsheet2.Cells(SaveRow, 2))) & "', '" & MyReplace(Me.Spreadsheet2.Cells(SaveRow, 3)) & "'," & NmJob & ");"
                        Con.Exequte Sql
                        MyErr = True
                    End If
                    End If
                End If
                Set Rs = Con.CloseRecordSet(Rs)
            If Trim("" & Me.Spreadsheet2.Cells(SaveRow, 2)) <> "" And SaveRow <> Row Then
                        If UCase(Trim("" & Me.Spreadsheet2.Cells(SaveRow, 1))) <> 0 Then
                        If Len(Trim("" & Me.Spreadsheet2.Cells(SaveRow, 15))) = 0 Then
                        
                            msg = "Le code APP ne peut être Nul"
                            MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                             Me.Spreadsheet2.Cells(SaveRow, 15).Select
                              Me.Spreadsheet2.SetFocus
                               Row = SaveRow
                              GoTo Fin
                          End If
                         If Len(Trim("" & Me.Spreadsheet2.Cells(SaveRow, 22))) = 0 Then
                        
                            msg = "Le code APP ne peut être Nul"
                            MsgBox msg, vbQuestion, "AutoCâble: Tableau de fils"
                             Me.Spreadsheet2.Cells(SaveRow, 22).Select
                              Me.Spreadsheet2.SetFocus
                               Row = SaveRow
                          End If
                        End If
                    End If
            End If
Fin:
SaveRow = Row
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
   If Trim("" & Me.Spreadsheet3.Cells(Row, 5)) <> "" Then
    Me.Spreadsheet3.Cells(Row, 5) = UCase(Me.Spreadsheet3.Cells(Row, 5))
    Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("B1", "B" & CStr(Myrange.Rows.Count))
 aa = Split(Me.Spreadsheet3.Cells(Row, 5) & ";", ";")
                For Iaa = 0 To UBound(aa) - 1

        I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))))

        If I = 0 Then
                msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbQuestion, "AutoCâble: Composants"
                     Me.Spreadsheet3.Cells(Row, 5) = Replace(Me.Spreadsheet3.Cells(Row, 5) & ";", "" & Trim(aa(Iaa)) & ";", "")
                Spreadsheet3.SetFocus
        End If
        Next
        If Right(Me.Spreadsheet3.Cells(Row, 5), 1) = ";" Then Me.Spreadsheet3.Cells(Row, 5) = Left(Me.Spreadsheet3.Cells(Row, 5), Len(Me.Spreadsheet3.Cells(Row, 5)) - 1)
   End If
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
    

If Col > 5 Then
    For I = 5 To NbFinOuiNon
        If Me.Spreadsheet3.Cells(Row, I) = 1 Then
            If I <> Col Then
                MsgBox "Vous ne pouvez pas sélectionner plusieurs répertoires.", vbQuestion, "AutoCâble: Composants"
                Me.Spreadsheet3.Cells(Row, Col) = 0
            End If
        End If
    Next I
End If
BoolOui = False
If (SaveRow <> 1) And (SaveRow <> Row) And (Trim("" & Me.Spreadsheet3.Cells(SaveRow, 1)) <> "") Then
 For I = 6 To NbFinOuiNon
    If Val(Me.Spreadsheet3.Cells(SaveRow, I)) = 1 Then
        BoolOui = True
        Exit For
    End If
    
    Next I
  If BoolOui = False And Me.Spreadsheet3.Cells(SaveRow, 1) = 1 Then
    MsgBox "Vous devez sélectionner un répertoire.", vbQuestion, "AutoCâble: Composants"
    Me.Spreadsheet3.Cells(SaveRow, 6).Select
     Me.Spreadsheet3.SetFocus
    msg = "?"
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
If Col > 2 Then
   If Trim("" & Me.Spreadsheet4.Cells(Row, 2)) <> "" Then Me.Spreadsheet4.Cells(Row, 3) = Row - 1
    
    If Trim("" & Me.Spreadsheet4.Cells(Row, 4)) <> "" Then
    Me.Spreadsheet4.Cells(Row, 4) = UCase(Me.Spreadsheet4.Cells(Row, 4))
    Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("a1").CurrentRegion
        Set Myrange = Me.Spreadsheet5.ActiveSheet.Range("B1", "B" & CStr(Myrange.Rows.Count))
 aa = Split(Me.Spreadsheet4.Cells(Row, 4) & ";", ";")
                For Iaa = 0 To UBound(aa) - 1

        I = ChercheXls(Myrange, UCase(Trim("" & aa(Iaa))))

        If I = 0 Then
                msg = "CODE CRITERE : " & UCase(Trim("" & aa(Iaa))) & " introuvable"
            MsgBox msg, vbQuestion, "AutoCâble: Notas"
                     Me.Spreadsheet4.Cells(Row, 4) = Replace(Me.Spreadsheet4.Cells(Row, 4) & ";", "" & Trim(aa(Iaa)) & ";", "")
                Spreadsheet4.SetFocus
        End If
        Next
        If Right(Me.Spreadsheet4.Cells(Row, 4), 1) = ";" Then Me.Spreadsheet4.Cells(Row, 4) = Left(Me.Spreadsheet4.Cells(Row, 4), Len(Me.Spreadsheet4.Cells(Row, 4)) - 1)
   End If
End If
NoMacro4 = False

Fin:
End Sub

Private Sub Form_Activate()
Dim MyExcel As EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
If BooolBloque = True Then
    Command3.Visible = False
    Command1.Visible = False
    Command2.Visible = False
End If
If bool_Activate = True Then GoTo Fin
bool_Activate = True
Dim Myrange As EXCEL.Range
Set MyExcel = New EXCEL.Application
MyExcel.DisplayAlerts = False
NotSortie = True
'MyExcel.Visible = True
Set a = Me.Spreadsheet1.Cells(2, 2)
If Trim(Me.Caption) = "" Then Exit Sub
If Nouveau = False Then

    Set MyWorkbook = MyExcel.Workbooks.Open(Me.Caption)
Else

    Set MyWorkbook = MyExcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
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

Set Myrange = MyWorkbook.Sheets("Nomenclature Connecteur").Range("a1").CurrentRegion
Myrange.Copy
Me.Spreadsheet7.ActiveSheet.Range("a1").Paste

Set Myrange = MyWorkbook.Sheets("Nomenclature Fils").Range("a1").CurrentRegion
Myrange.Copy
Me.Spreadsheet8.ActiveSheet.Range("a1").Paste

Set Myrange = MyWorkbook.Sheets("Nomenclature Habillage").Range("a1").CurrentRegion
Myrange.Copy
Me.Spreadsheet9.ActiveSheet.Range("a1").Paste

MyExcel.AlertBeforeOverwriting = False

Set Myrange = Nothing
MyWorkbook.Close False
Set MyWorkbook = Nothing

MyExcel.Quit
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
Me.Spreadsheet1.Columns(4).NumberFormat = "Yes/No"
Me.Spreadsheet3.Columns(1).NumberFormat = "Yes/No"
For I = 6 To 600
    Me.Spreadsheet3.Columns(I).NumberFormat = "Yes/No"
 Next I
    
DoEvents
LstMaj
Fin:
End Sub
Public Sub Chargement(fichier As String, Client As String, Id As Long, Optional NouveauF As Boolean)
Dim Rs As Recordset
Dim Sql As String
Dim txtMyCollectionLienHab As String
IdProjet = Id
Sql = "SELECT T_Regle_Comp_Hab.ENCELADE, T_Regle_Comp_Hab.PSA, "
Sql = Sql & "T_Regle_Comp_Hab.RSA, T_Regle_Comp_Hab.libellé, T_Regle_Comp_Hab.Numéro "
Sql = Sql & "FROM T_Regle_Comp_Hab "
Sql = Sql & "ORDER BY T_Regle_Comp_Hab.libellé;"
I = 0
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
I = I + 1
Rs.MoveNext
Wend
ReDim MyTableENC(I, 4)
ReDim MyTablePSA(I, 4)
ReDim MyTableRSA(I, 4)
ReDim MyTableHab(I, 4)
ReDim MyTableHab(I, 4)

ReDim MyTableHab(I, 4)
Rs.Requery
I = 0
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




  For I = 0 To UBound(MyTablePSA)
       For I2 = 1 To 4
            MyTableENC(I, I2) = "N"
            MyTablePSA(I, I2) = "N"
            MyTableRSA(I, I2) = "N"
            MyTableHab(I, I2) = "N"
          
        Next
  Next
  I = 0
While Rs.EOF = False

    I = I + 1
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
Me.Caption = fichier
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
If (NoMacro5 = True) Or (NoMacro5Select = True) Or (Row = 1) Then GoTo Fin
NoMacro5 = True
Row = Me.Spreadsheet5.ActiveCell.Row
Col = Me.Spreadsheet5.ActiveCell.Column
If SaveRow = 0 Then SaveRow = Row
boolActu = False

 Set Myrange = Me.Spreadsheet5.Range("a1").CurrentRegion
   ConverOuiNon Myrange, Row
     Me.Spreadsheet5.Cells(SaveRow, 2) = UCase(Trim("" & Me.Spreadsheet5.Cells(SaveRow, 2)))
     Me.Spreadsheet5.Cells(SaveRow, 3) = UCase(Trim("" & Me.Spreadsheet5.Cells(SaveRow, 3)))
     
    If (Trim("" & Me.Spreadsheet5.Cells(SaveRow, 2)) <> "") And (Trim("" & Me.Spreadsheet5.Cells(SaveRow, 3)) = "") And (SaveRow <> Row) Then
        MsgBox "Le champ CRITERES est obligatoire", vbQuestion, "AutoCâble: Critères"
        msg = "?"
       Me.Spreadsheet5.Cells(SaveRow, 3).Select
       Spreadsheet5.SetFocus
    End If
    If (Trim("" & Me.Spreadsheet5.Cells(SaveRow, 2)) = "") And (Trim("" & Me.Spreadsheet5.Cells(SaveRow, 3)) <> "") And (SaveRow <> Row) Then
        MsgBox "Le champ CODE CRITERES est obligatoire", vbQuestion, "AutoCâble: Critères"
        msg = "?"
       Me.Spreadsheet5.Cells(SaveRow, 2).Select
       Spreadsheet5.SetFocus
    End If
SaveRow = Row
NoMacro5 = False
Fin:
End Sub

Private Sub Spreadsheet5_Click(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Spreadsheet5_Change EventInfo
End Sub

Private Sub Spreadsheet5_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
If NoMacro5Select = True Then GoTo Fin
 NoMacro5Select = True
Spreadsheet5_Change EventInfo
Static Row As Long

Dim aa
Row = Spreadsheet5.ActiveCell.Row
If Row = 0 Then Row = 1
If Row > 1 And IfValidationOk = True Then
On Error Resume Next

If Spreadsheet5.Cells(Row, 2) <> "" Then
If Spreadsheet5.Cells(Row, 1) = 1 Then
aa = ""
    aa = CollecCrieres(Spreadsheet5.Cells(Row, 2))
    If Err Then
    Err.Clear
        CollecCrieres.Add Spreadsheet5.Cells(Row, 2), Spreadsheet5.Cells(Row, 2)
    Else
        MsgBox " Le code Code Critères : " & Spreadsheet5.Cells(Row, 2) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
            Me.Spreadsheet5.Cells(Row, 2).Select
            Me.Spreadsheet5.SetFocus
            NoMacro5Select = False
             On Error GoTo 0
        GoTo Fin
    End If
 End If
End If

If Spreadsheet5.Cells(Row, 3) <> "" Then
If Spreadsheet5.Cells(Row, 1) = 1 Then
aa = ""
    aa = CollecCrieresCode(Spreadsheet5.Cells(Row, 3))
    If Err Then
        Err.Clear
        CollecCrieresCode.Add Spreadsheet5.Cells(Row, 3), Spreadsheet5.Cells(Row, 3)
    Else
        MsgBox " Le Critères : " & Spreadsheet5.Cells(Row, 3) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
             Me.Spreadsheet5.Cells(Row, 3).Select
            Me.Spreadsheet5.SetFocus
            NoMacro5Select = False
             On Error GoTo 0
        GoTo Fin
        GoTo Fin
    End If
    End If
End If
If Spreadsheet5.Cells(Row, 4) <> "" Then
If Spreadsheet5.Cells(Row, 1) = 1 Then
aa = ""
    aa = CollecCrieresDesigne(Spreadsheet5.Cells(Row, 4))
    If Err Then
        Err.Clear
        CollecCrieresDesigne.Add Spreadsheet5.Cells(Row, 4), Spreadsheet5.Cells(Row, 4)
    Else
        MsgBox " Le code Designation Critères : " & Spreadsheet5.Cells(Row, 4) & " existe déjà.", vbExclamation, "Erreur sur l'onglet Critères"
            msg = "?"
             Me.Spreadsheet5.Cells(Row, 4).Select
            Me.Spreadsheet5.SetFocus
            NoMacro5Select = False
             On Error GoTo 0
        GoTo Fin
             On Error GoTo 0
    End If
   End If
End If
'    Set CollecCrieres = New Collection
'Set CollecCrieresCode = New Collection
'Set CollecCrieresDesigne = New Collection
End If
Row = Row
NoMacro5Select = False
Fin:
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
Dim Txt As String
Dim Ofset As Long
Dim NbTour As Long
Dim NbTord As Long
Dim txtColone As Long
txtColone = 2
Txt = "AA"
Ofset = 0
NbTour = 0
NbTord = 0
Reprise:

For I = 0 To Row - 2
aa = Mid(Txt, Len(Txt) - Ofset, 1)

    aa = Chr(Asc("A") + (1 * (I - (26 * NbTour))))

Mid(Txt, Len(Txt) - Ofset, 1) = aa


If Asc(Mid(aa, 1, 1)) < 65 Or Asc(Mid(aa, 1, 1)) > 90 Then

Mid(Txt, Len(Txt) - Ofset, 1) = "A"


    Ofset = Ofset + 1
    NbTour = NbTour + 1
    Mid(Txt, Len(Txt) - Ofset, 1) = Chr(Asc(Mid(Txt, Len(Txt) - Ofset, 1)) + 1)
    If Asc(Mid(Txt, 1, 1)) < 65 Or Asc(Mid(Txt, 1, 1)) > 90 Then
 Mid(Txt, 1, 1) = "A"
    Txt = Txt & "A"
    

End If
Ofset = 0
   
End If


Next

NoeuName2 = Txt
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
    Me.txtOption = Spreadsheet6.ActiveSheet.Cells(EventInfo.Range.Row, 13)
   
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
For I = 0 To Me.ENC.ListCount - 1
        MyCollectionENC.Add I, "N" & Trim(Me.ENC.List(I))
        
Next
For I = 0 To Me.PSA.ListCount - 1
       MyCollectionPSA.Add I, "N" & Trim(Me.PSA.List(I))
Next
For I = 0 To Me.RSA.ListCount - 1
        MyCollectionRSA.Add I, "N" & Trim(Me.RSA.List(I))
Next

For I = 0 To Me.Hab.ListCount - 1
      MyCollectionHab.Add I, "N" & Me.Hab.List(I)
Next
For I = 1 To MyCollectionLienHab.Count
    zz = Split(MyCollectionLienHab(I), ";")
    For I2 = 0 To 3
        If zz(I2) <> "N" Then
            MyTableHab(MyCollectionHab(zz(3)), I2 + 1) = Trim(zz(I2))
            
        
        End If
    Next
Next
For I = 1 To UBound(MyTableHab)
    If MyTableHab(I, 1) <> "N" Then
        MyTableENC(MyCollectionENC(MyTableHab(I, 1)), 1) = MyTableHab(I, 1)
        MyTableENC(MyCollectionENC(MyTableHab(I, 1)), 4) = MyTableHab(I, 4)
        
        If MyTableHab(I, 2) <> "N" Then
            MyTableENC(MyCollectionENC(MyTableHab(I, 1)), 2) = MyTableHab(I, 2)
        End If
        
        If MyTableHab(I, 3) <> "N" Then
            MyTableENC(MyCollectionENC(MyTableHab(I, 1)), 3) = MyTableHab(I, 3)
        End If
        
    End If
    
    If MyTableHab(I, 2) <> "N" Then
        MyTablePSA(MyCollectionPSA(MyTableHab(I, 2)), 2) = MyTableHab(I, 2)
        MyTablePSA(MyCollectionPSA(MyTableHab(I, 2)), 4) = MyTableHab(I, 4)
        
        If MyTableHab(I, 1) <> "N" Then
             MyTablePSA(MyCollectionPSA(MyTableHab(I, 2)), 1) = MyTableHab(I, 1)
        End If
        
        If MyTableHab(I, 3) <> "N" Then
            MyTablePSA(MyCollectionPSA(MyTableHab(I, 2)), 3) = MyTableHab(I, 3)
        End If
    End If
     If MyTableHab(I, 3) <> "N" Then
        MyTableRSA(MyCollectionRSA(MyTableHab(I, 3)), 3) = MyTableHab(I, 3)
        MyTableRSA(MyCollectionRSA(MyTableHab(I, 3)), 4) = MyTableHab(I, 4)
        
        If MyTableHab(I, 1) <> "N" Then
             MyTableRSA(MyCollectionRSA(MyTableHab(I, 3)), 1) = MyTableHab(I, 1)
        End If
        
        If MyTableHab(I, 2) <> "N" Then
            MyTableRSA(MyCollectionRSA(MyTableHab(I, 3)), 2) = MyTableHab(I, 2)
        End If
    End If
Next

End Sub

Sub ConverOuiNon(Myrange, Index)
For I = 1 To Myrange.Columns.Count
   If Myrange(Index, I).NumberFormat = "Yes/No" Then
  
        If Not IsNumeric(Myrange(Index, I).Value) Then
            If UCase(Left(Myrange(Index, I).Value, 1)) = "N" Then
                Myrange(Index, I).Value = 0
                DoEvents
               
            Else
                Myrange(Index, I).Value = 1
                DoEvents
               
            End If
        End If
      
   End If
    
Next
End Sub

