VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form UtliserPar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Utilisé par:"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   15930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   15735
      HTMLURL         =   ""
      HTMLData        =   $"UtliserPar.frx":0000
      DataType        =   "HTMLDATA"
      AutoFit         =   0   'False
      DisplayColHeaders=   0   'False
      DisplayGridlines=   -1  'True
      DisplayHorizontalScrollBar=   0   'False
      DisplayRowHeaders=   0   'False
      DisplayTitleBar =   0   'False
      DisplayToolbar  =   0   'False
      DisplayVerticalScrollBar=   -1  'True
      EnableAutoCalculate=   -1  'True
      EnableEvents    =   -1  'True
      MoveAfterReturn =   -1  'True
      MoveAfterReturnDirection=   0
      RightToLeft     =   0   'False
      ViewableRange   =   "1:65536"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   7320
      Width           =   3495
   End
End
Attribute VB_Name = "UtliserPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As Recordset
Dim Sql As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim I As Long
Sql = "SELECT MyFrom.Job, T_indiceProjet.PI, T_indiceProjet.Li, T_indiceProjet.PL, T_indiceProjet.[OU],  "
Sql = Sql & "Utilise_Par.Machine, Utilise_Par.User "
Sql = Sql & "FROM (T_indiceProjet LEFT JOIN (SELECT T_Job.Id_Piece, T_Job.Job, T_Job.Status, T_Job.FinTraitement "
Sql = Sql & "FROM T_Job "
Sql = Sql & "WHERE T_Job.Job Is Not Null AND T_Job.FinTraitement=False) AS  "
Sql = Sql & "MyFrom ON T_indiceProjet.Id = MyFrom.Id_Piece) LEFT JOIN Utilise_Par  "
Sql = Sql & "ON T_indiceProjet.UserName = Utilise_Par.Machine "
Sql = Sql & "WHERE T_indiceProjet.UserName Is Not Null;"

Set Rs = Con.OpenRecordSet(Sql)
For I = 0 To Rs.Fields.Count - 1
 Spreadsheet1.ActiveSheet.Cells(1, 1 + I) = Rs(I).Name
Next
If Rs.EOF = False Then
    Const sDelimiteur$ = vbTab
    Debug.Print Asc(vbCrLf)
    Dim toto
    toto = Rs.GetString(, , sDelimiteur$, "¤")
     toto = Replace(toto, sDelimiteur$ & "-1" & sDelimiteur$, sDelimiteur$ & "OUI" & sDelimiteur$)
     toto = Replace(toto, sDelimiteur$ & "0" & sDelimiteur$, sDelimiteur$ & "NON" & sDelimiteur$)
    toto = Replace(toto, vbCrLf, " ")
    toto = Replace(toto, Chr(13), "")
    toto = Replace(toto, Chr(10), "")
    toto = Replace(toto, "\", "")
    toto = Replace(toto, "¤", vbCrLf)
    
     
    Spreadsheet1.ActiveSheet.Protection.Enabled = False
    Spreadsheet1.ActiveSheet.Range("A2").ParseText _
    toto, sDelimiteur$

End If
Set Rs = Con.CloseRecordSet(Rs)
 Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion.AutoFilter
End Sub
