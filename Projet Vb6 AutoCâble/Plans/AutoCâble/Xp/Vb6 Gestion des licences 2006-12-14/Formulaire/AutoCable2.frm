VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form UserForm6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Codes Liaisons:"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14910
   Icon            =   "AutoCable2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandButton1 
      Caption         =   "&Annuler"
      Height          =   435
      Left            =   12600
      TabIndex        =   4
      Top             =   9000
      Width           =   2130
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "&Enregistre"
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Top             =   9000
      Width           =   2130
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   15743
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Connecteurs"
      TabPicture(0)   =   "AutoCable2.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Spreadsheet1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tableau de fils"
      TabPicture(1)   =   "AutoCable2.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Spreadsheet2"
      Tab(1).ControlCount=   1
      Begin OWC.Spreadsheet Spreadsheet2 
         Height          =   8925
         Left            =   -74880
         TabIndex        =   2
         Top             =   345
         Width           =   14745
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable2.frx":0902
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
         Height          =   8925
         Left            =   120
         TabIndex        =   1
         Top             =   345
         Width           =   14145
         HTMLURL         =   ""
         HTMLData        =   $"AutoCable2.frx":142A
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
      Top             =   9720
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
Dim MyRange
Dim Rangecount As Long
Dim Sql As String
Dim Rs As Recordset
If MsgBox("Voulez vous enregistre les modification.", vbQuestion + vbYesNo) = vbNo Then Exit Sub
Me.CommandButton1.Enabled = False
Me.CommandButton2.Enabled = False
Me.SSTab1.Enabled = False
Sql = "UPDATE LIAISON SET LIAISON.Sup = True;"
Con.Execute Sql
Sql = "UPDATE LIAISON_CONNECTEURS SET LIAISON_CONNECTEURS.Sup = True;"
Con.Execute Sql
RazFiltreEditExcel Me.Spreadsheet2
Set MyRange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion

Rangecount = MyRange.Rows.Count
RazFiltreEditExcel Me.Spreadsheet1
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion

Rangecount = Rangecount + MyRange.Rows.Count
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = Rangecount
For I = 2 To MyRange.Rows.Count
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
DoEvents
    Sql = "select LIAISON_CONNECTEURS.LIAISON from LIAISON_CONNECTEURS WHERE "
    Sql = Sql & "LIAISON_CONNECTEURS.CLIENT='" & MyReplace("" & MyRange(I, 1)) & "' AND "
     Sql = Sql & "LIAISON_CONNECTEURS.LIAISON='" & MyReplace("" & MyRange(I, 2)) & "' "
     Set Rs = Con.OpenRecordSet(Sql)
     If Rs.EOF = False Then
        Sql = "UPDATE LIAISON_CONNECTEURS SET LIAISON_CONNECTEURS.LIB ='" & MyReplace("" & MyRange(I, 3)) & "',  "
        Sql = Sql & "LIAISON_CONNECTEURS.Sup = False WHERE "
        Sql = Sql & "LIAISON_CONNECTEURS.CLIENT='" & MyReplace("" & MyRange(I, 1)) & "' AND "
        Sql = Sql & "LIAISON_CONNECTEURS.LIAISON='" & MyReplace("" & MyRange(I, 2)) & "' "
        Con.Execute Sql
     Else
        Sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
        Sql = Sql & "VALUES( '" & MyReplace("" & MyRange(I, 1)) & "', "
        Sql = Sql & "'" & MyReplace("" & MyRange(I, 2)) & "' ,"
        Sql = Sql & "'" & MyReplace("" & MyRange(I, 3)) & "');"
        Con.Execute Sql
     End If
Next I

Set MyRange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion
Rangecount = Rangecount + MyRange.Rows.Count
For I = 2 To MyRange.Rows.Count
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
DoEvents
    Sql = "select LIAISON.LIAISON from LIAISON WHERE "
    Sql = Sql & "LIAISON.CLIENT='" & MyReplace("" & MyRange(I, 1)) & "' AND "
     Sql = Sql & "LIAISON.LIAISON='" & MyReplace("" & MyRange(I, 2)) & "' "
     Set Rs = Con.OpenRecordSet(Sql)
     If Rs.EOF = False Then
        Sql = "UPDATE LIAISON SET LIAISON.LIB ='" & MyReplace("" & MyRange(I, 3)) & "',  "
        Sql = Sql & "LIAISON.Sup = False WHERE "
        Sql = Sql & "LIAISON.CLIENT='" & MyReplace("" & MyRange(I, 1)) & "' AND "
        Sql = Sql & "LIAISON.LIAISON='" & MyReplace("" & MyRange(I, 2)) & "' "
        Con.Execute Sql
     Else
        Sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
        Sql = Sql & "VALUES( '" & MyReplace("" & MyRange(I, 1)) & "', "
        Sql = Sql & "'" & MyReplace("" & MyRange(I, 2)) & "' ,"
        Sql = Sql & "'" & MyReplace("" & MyRange(I, 3)) & "');"
        Con.Execute Sql
     End If
Next I
Con.Execute "DELETE LIAISON.*, LIAISON.Sup FROM LIAISON WHERE LIAISON.Sup=True;"
Con.Execute "DELETE LIAISON_CONNECTEURS.*, LIAISON_CONNECTEURS.Sup FROM LIAISON_CONNECTEURS WHERE LIAISON_CONNECTEURS.Sup=True;"

Noquite = False
Me.Hide
End Sub

Public Sub chargement()
Dim Rs As Recordset
Dim Row As Long
Const sDelimiteur$ = vbTab
    Dim toto


Set Rs = Con.OpenRecordSet("SELECT RqLiaisonConnecteur.* FROM RqLiaisonConnecteur;")

FormBarGrah.StatusBar1.Panels(2) = " Liaisons Connecteur:"
FormBarGrah.ProgressBar1.Value = 0
DoEvents
Row = 0
'While Rs.EOF = False
'Row = Row + 1
'    Rs.MoveNext
'Wend
FormBarGrah.ProgressBar1.Max = 1
Rs.Requery
Row = 1
If Rs.EOF = False Then
    
    Debug.Print Asc(vbCrLf)
   
    toto = Rs.GetString(, , sDelimiteur$ & "'", "¤")
    
    toto = Replace(toto, vbCrLf, Chr(10))
    toto = Replace(toto, Chr(13), "")
'    toto = Replace(toto, Chr(10), "")
   
    toto = Replace(toto, "¤", vbCrLf)
    Spreadsheet1.ActiveSheet.Protection.Enabled = False
    Spreadsheet1.ActiveSheet.Range("A2").ParseText _
    toto, sDelimiteur$

End If

Set Rs = Con.CloseRecordSet(Rs)
'While Rs.EOF = False
'IncremanteBarGrah FormBarGrah
'    Row = Row + 1
'    For I = 0 To Rs.Fields.Count - 1
'     Spreadsheet1.Cells(Row, I + 1) = "'" & Rs.Fields(I).Value
'     Next I
'
'Rs.MoveNext
'Wend
Set Rs = Con.OpenRecordSet("SELECT RqLiaisonFils.*FROM RqLiaisonFils;")

FormBarGrah.StatusBar1.Panels(2) = " Liaisons Fils:"
FormBarGrah.ProgressBar1.Value = 0
DoEvents
Row = 0
Row = 0
'While Rs.EOF = False
'Row = Row + 1
'    Rs.MoveNext
'Wend
FormBarGrah.ProgressBar1.Max = 1
Rs.Requery
Row = 1

If Rs.EOF = False Then
    
    toto = Rs.GetString(, , sDelimiteur$ & "'", "¤")
    
    toto = Replace(toto, vbCrLf, Chr(10))
    toto = Replace(toto, Chr(13), "")
'    toto = Replace(toto, Chr(10), "")
   
    toto = Replace(toto, "¤", vbCrLf)
    Spreadsheet2.ActiveSheet.Protection.Enabled = False
    Spreadsheet2.ActiveSheet.Range("A2").ParseText _
    toto, sDelimiteur$

End If
'While Rs.EOF = False
'IncremanteBarGrah FormBarGrah
'    Row = Row + 1
'    For I = 0 To Rs.Fields.Count - 1
'    DoEvents
'     Spreadsheet2.Cells(Row, I + 1) = "'" & Rs.Fields(I).Value
'     Next I
'
'Rs.MoveNext
'Wend
Spreadsheet1.ActiveSheet.Range("A1").CurrentRegion.AutoFitColumns
Spreadsheet2.ActiveSheet.Range("A1").CurrentRegion.AutoFitColumns
FormBarGrah.StatusBar1.Panels(2) = ""
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

