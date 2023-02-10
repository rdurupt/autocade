VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ModuleListes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liste des Modules :"
   ClientHeight    =   11775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15315
   Icon            =   "Module.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11775
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   13080
      TabIndex        =   5
      Top             =   0
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "Module.frx":08CA
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   12720
      TabIndex        =   3
      Top             =   10920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enregistrer"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   10920
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   11520
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   9585
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   15315
      HTMLURL         =   ""
      HTMLData        =   $"Module.frx":173C
      DataType        =   "HTMLDATA"
      AutoFit         =   0   'False
      DisplayColHeaders=   -1  'True
      DisplayGridlines=   -1  'True
      DisplayHorizontalScrollBar=   0   'False
      DisplayRowHeaders=   -1  'True
      DisplayTitleBar =   0   'False
      DisplayToolbar  =   -1  'True
      DisplayVerticalScrollBar=   -1  'True
      EnableAutoCalculate=   -1  'True
      EnableEvents    =   0   'False
      MoveAfterReturn =   -1  'True
      MoveAfterReturnDirection=   0
      RightToLeft     =   0   'False
      ViewableRange   =   "1:1600"
   End
   Begin VB.Label ProgressBar1Caption 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   11520
      Width           =   1575
   End
End
Attribute VB_Name = "ModuleListes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MyRange
Dim Rangecount As Long
Dim Sql As String
Dim Rs As Recordset
If MsgBox("Voulez vous enregistre les modification.", vbQuestion + vbYesNo) = vbNo Then Exit Sub
Me.Command1.Enabled = False
Me.Command1.Enabled = False
Sql = "UPDATE Module SET Module.Sup = True;"
Con.Execute Sql
'Rangecount = MyRange.Rows.Count
RazFiltreEditExcel Me.Spreadsheet1
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion

Rangecount = Rangecount + MyRange.Rows.Count
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = Rangecount
For I = 2 To MyRange.Rows.Count
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
DoEvents
    Sql = "select Module.NameBouton from Module WHERE "
    Sql = Sql & "Module.NameBouton='" & MyReplace("" & MyRange(I, 1)) & "' "
     Set Rs = Con.OpenRecordSet(Sql)
     If Rs.EOF = False Then
        Sql = "UPDATE Module SET Module.NameBouton ='" & MyReplace("" & MyRange(I, 1)) & "',  "
        Sql = Sql & "Module.Utilitaire  ='" & MyReplace("" & MyRange(I, 2)) & "',  "
        Sql = Sql & "Module.Sup = False WHERE "
        Sql = Sql & "Module.NameBouton='" & MyReplace("" & MyRange(I, 1)) & "' "
        
        Con.Execute Sql
     Else
        Sql = "INSERT INTO Module ( NameBouton, Utilitaire ) "
        Sql = Sql & "VALUES( '" & MyReplace("" & MyRange(I, 1)) & "', "
        Sql = Sql & "'" & MyReplace("" & MyRange(I, 2)) & "') ;"
'        Sql = Sql & "'" & MyReplace("" & MyRange(i, 3)) & "');"
        Con.Execute Sql
     End If
Next I
Con.Execute "DELETE Module.*, Module.Sup FROM Module WHERE Module.Sup=True;"
MajDroitsFrm Id_Users
Noquite = False
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim Rs As Recordset
Dim Row As Long



Set Rs = Con.OpenRecordSet("SELECT Module.NameBouton, Module.Utilitaire FROM Module ORDER BY Module.NameBouton;")
Me.ProgressBar1Caption.Caption = " Utilitaire:"
Me.ProgressBar1.Value = 0
DoEvents
Row = 0
While Rs.EOF = False
Row = Row + 1
    Rs.MoveNext
Wend
Me.ProgressBar1.Max = Row + 1
Rs.Requery
Row = 1
While Rs.EOF = False
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
    Row = Row + 1
    For I = 0 To Rs.Fields.Count - 1
     Spreadsheet1.Cells(Row, I + 1) = "" & Rs.Fields(I).Value
     Next I
     
Rs.MoveNext
Wend
'Set Rs = Con.OpenRecordSet("SELECT RqLiaisonFils.*FROM RqLiaisonFils;")
'
Me.ProgressBar1Caption.Caption = ""
Me.ProgressBar1.Value = 0
'DoEvents
'Row = 0
'Row = 0
'While Rs.EOF = False
'Row = Row + 1
'    Rs.MoveNext
'Wend
''FormBarGrah.ProgressBar1.Max = Row + 1
'Rs.Requery
'Row = 1
'While Rs.EOF = False
'IncremanteBarGrah FormBarGrah
'    Row = Row + 1
'    For i = 0 To Rs.Fields.Count - 1
'    DoEvents
'     Spreadsheet2.Cells(Row, i + 1) = "'" & Rs.Fields(i).Value
'     Next i
'
'Rs.MoveNext
'Wend
'FormBarGrah.ProgressBar1Caption.Caption = ""
'FormBarGrah.ProgressBar1.Value = 0
MousePointer = fmMousePointerDefault


End Sub
