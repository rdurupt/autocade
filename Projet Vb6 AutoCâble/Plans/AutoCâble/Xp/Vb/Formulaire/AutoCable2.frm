VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form UserForm6 
   Caption         =   "Codes Liaisons:"
   ClientHeight    =   12960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14910
   Icon            =   "AutoCable2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12960
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   12720
      TabIndex        =   6
      Top             =   0
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "AutoCable2.frx":08CA
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "Annuler"
      Height          =   435
      Left            =   12480
      TabIndex        =   4
      Top             =   12120
      Width           =   2130
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "Enregistre"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   12120
      Width           =   2130
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11085
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   19553
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Connecteurs"
      TabPicture(0)   =   "AutoCable2.frx":4C71
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Spreadsheet1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tableau de fils"
      TabPicture(1)   =   "AutoCable2.frx":4C8D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Spreadsheet2"
      Tab(1).ControlCount=   1
      Begin OWC.Spreadsheet Spreadsheet2 
         Height          =   10530
         Left            =   -74880
         TabIndex        =   2
         Top             =   345
         Width           =   14745
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable2.frx":4CA9
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
         Height          =   10530
         Left            =   120
         TabIndex        =   1
         Top             =   345
         Width           =   14745
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable2.frx":57D0
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
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   12720
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean

Private Sub CommandButton1_Click()
Noquite = False
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Dim Myrange
Dim Rangecount As Long
Dim Sql As String
Dim Rs As Recordset
If MsgBox("Voulez vous enregistre les modification.", vbQuestion + vbYesNo) = vbNo Then Exit Sub
Me.CommandButton1.Enabled = False
Me.CommandButton2.Enabled = False
Me.SSTab1.Enabled = False
Sql = "UPDATE LIAISON SET LIAISON.Sup = True;"
Con.Exequte Sql
Sql = "UPDATE LIAISON_CONNECTEURS SET LIAISON_CONNECTEURS.Sup = True;"
Con.Exequte Sql
RazFiltreEditExcel Me.Spreadsheet2
Set Myrange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion

Rangecount = Myrange.Rows.Count
RazFiltreEditExcel Me.Spreadsheet1
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion

Rangecount = Rangecount + Myrange.Rows.Count
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = Rangecount
For i = 2 To Myrange.Rows.Count
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
DoEvents
    Sql = "select LIAISON_CONNECTEURS.LIAISON from LIAISON_CONNECTEURS WHERE "
    Sql = Sql & "LIAISON_CONNECTEURS.CLIENT='" & MyReplace("" & Myrange(i, 1)) & "' AND "
     Sql = Sql & "LIAISON_CONNECTEURS.LIAISON='" & MyReplace("" & Myrange(i, 2)) & "' "
     Set Rs = Con.OpenRecordSet(Sql)
     If Rs.EOF = False Then
        Sql = "UPDATE LIAISON_CONNECTEURS SET LIAISON_CONNECTEURS.LIB ='" & MyReplace("" & Myrange(i, 3)) & "',  "
        Sql = Sql & "LIAISON_CONNECTEURS.Sup = False WHERE "
        Sql = Sql & "LIAISON_CONNECTEURS.CLIENT='" & MyReplace("" & Myrange(i, 1)) & "' AND "
        Sql = Sql & "LIAISON_CONNECTEURS.LIAISON='" & MyReplace("" & Myrange(i, 2)) & "' "
        Con.Exequte Sql
     Else
        Sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
        Sql = Sql & "VALUES( '" & MyReplace("" & Myrange(i, 1)) & "', "
        Sql = Sql & "'" & MyReplace("" & Myrange(i, 2)) & "' ,"
        Sql = Sql & "'" & MyReplace("" & Myrange(i, 3)) & "');"
        Con.Exequte Sql
     End If
Next i

Set Myrange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion
Rangecount = Rangecount + Myrange.Rows.Count
For i = 2 To Myrange.Rows.Count
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
DoEvents
    Sql = "select LIAISON.LIAISON from LIAISON WHERE "
    Sql = Sql & "LIAISON.CLIENT='" & MyReplace("" & Myrange(i, 1)) & "' AND "
     Sql = Sql & "LIAISON.LIAISON='" & MyReplace("" & Myrange(i, 2)) & "' "
     Set Rs = Con.OpenRecordSet(Sql)
     If Rs.EOF = False Then
        Sql = "UPDATE LIAISON SET LIAISON.LIB ='" & MyReplace("" & Myrange(i, 3)) & "',  "
        Sql = Sql & "LIAISON.Sup = False WHERE "
        Sql = Sql & "LIAISON.CLIENT='" & MyReplace("" & Myrange(i, 1)) & "' AND "
        Sql = Sql & "LIAISON.LIAISON='" & MyReplace("" & Myrange(i, 2)) & "' "
        Con.Exequte Sql
     Else
        Sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
        Sql = Sql & "VALUES( '" & MyReplace("" & Myrange(i, 1)) & "', "
        Sql = Sql & "'" & MyReplace("" & Myrange(i, 2)) & "' ,"
        Sql = Sql & "'" & MyReplace("" & Myrange(i, 3)) & "');"
        Con.Exequte Sql
     End If
Next i
Con.Exequte "DELETE LIAISON.*, LIAISON.Sup FROM LIAISON WHERE LIAISON.Sup=True;"
Con.Exequte "DELETE LIAISON_CONNECTEURS.*, LIAISON_CONNECTEURS.Sup FROM LIAISON_CONNECTEURS WHERE LIAISON_CONNECTEURS.Sup=True;"

Noquite = False
Me.Hide
End Sub

Public Sub Chargement()
Dim Rs As Recordset
Dim Row As Long



Set Rs = Con.OpenRecordSet("SELECT RqLiaisonConnecteur.* FROM RqLiaisonConnecteur;")
FormBarGrah.ProgressBar1Caption.Caption = " Liaisons Connecteur:"
FormBarGrah.ProgressBar1.Value = 0
DoEvents
Row = 0
While Rs.EOF = False
Row = Row + 1
    Rs.MoveNext
Wend
FormBarGrah.ProgressBar1.Max = Row + 1
Rs.Requery
Row = 1
While Rs.EOF = False
IncremanteBarGrah FormBarGrah
    Row = Row + 1
    For i = 0 To Rs.Fields.Count - 1
     Spreadsheet1.Cells(Row, i + 1) = "'" & Rs.Fields(i).Value
     Next i
     
Rs.MoveNext
Wend
Set Rs = Con.OpenRecordSet("SELECT RqLiaisonFils.*FROM RqLiaisonFils;")

FormBarGrah.ProgressBar1Caption.Caption = " Liaisons Fils:"
FormBarGrah.ProgressBar1.Value = 0
DoEvents
Row = 0
Row = 0
While Rs.EOF = False
Row = Row + 1
    Rs.MoveNext
Wend
FormBarGrah.ProgressBar1.Max = Row + 1
Rs.Requery
Row = 1
While Rs.EOF = False
IncremanteBarGrah FormBarGrah
    Row = Row + 1
    For i = 0 To Rs.Fields.Count - 1
    DoEvents
     Spreadsheet2.Cells(Row, i + 1) = "'" & Rs.Fields(i).Value
     Next i
     
Rs.MoveNext
Wend
FormBarGrah.ProgressBar1Caption.Caption = ""
FormBarGrah.ProgressBar1.Value = 0
MousePointer = fmMousePointerDefault
Me.Show vbModal
End Sub

Private Sub UserForm_Activate()

 Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite

End Sub

