VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form frmCopyCollerExcel 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5595
      HTMLURL         =   ""
      HTMLData        =   $"frmCopyCollerExcel.frx":0000
      DataType        =   "HTMLDATA"
      AutoFit         =   0   'False
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
Attribute VB_Name = "frmCopyCollerExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Charger(txt As String)
Me.Spreadsheet1.Range("A1").CurrentRegion.Clear
Const sDelimiteur$ = vbTab
    
    Spreadsheet1.ActiveSheet.Protection.Enabled = False
    Spreadsheet1.ActiveSheet.Range("A1").ParseText _
    txt, sDelimiteur$
    Me.Spreadsheet1.Range("A1").CurrentRegion.Copy
'  Me.Show vbModal
End Sub

