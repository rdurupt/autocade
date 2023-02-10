VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmHabillage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Habillage:"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18240
   Icon            =   "FrmHabillage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   18240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Actualiser"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Enregistrer"
      Height          =   375
      Left            =   8100
      TabIndex        =   2
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   13200
      TabIndex        =   1
      Top             =   8880
      Width           =   1455
   End
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18240
      HTMLURL         =   ""
      HTMLData        =   $"FrmHabillage.frx":08CA
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   -360
      TabIndex        =   4
      Top             =   9480
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmHabillage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoQuitte As Boolean
Dim Actualiser As Boolean


Private Sub Command1_Click()
NoQuitte = False
Unload Me


End Sub

Private Sub Command2_Click()
If Actualiser = False Then
MsgBox "Vous devez actualiser avant d'enregistrer vos données.", vbCritical
Exit Sub
End If
Dim Sql As String
Dim SqlInsert As String

Dim MyRange
RazFiltreEditExcel Me.Spreadsheet1
Sql = "DELETE T_Regle_Comp_Hab.* FROM T_Regle_Comp_Hab;"
Con.Execute Sql
SqlInsert = "INSERT INTO T_Regle_Comp_Hab ( libellé, ENCELADE, RSA, PSA ) Values ("

Set MyRange = Spreadsheet1.Range("A1").CurrentRegion
ProgressBar1.Value = 0
ProgressBar1.Max = MyRange.Rows.Count

For I = 2 To "" & MyRange.Rows.Count
ProgressBar1.Value = I

    Sql = "'" & MyReplace("" & MyRange(I, 1)) & "',"
    Sql = Sql & "'" & MyReplace("" & MyRange(I, 2)) & "',"
    Sql = Sql & "'" & MyReplace("" & MyRange(I, 3)) & "',"
    Sql = Sql & "'" & MyReplace("" & MyRange(I, 4)) & "');"
    Con.Execute SqlInsert & Sql
Next
NoQuitte = False
Unload Me

End Sub

Public Sub chargement()

Dim Sql As String
Dim Rs As Recordset
Dim I As Long

NoQuitte = True
Sql = "SELECT T_Regle_Comp_Hab.ENCELADE, T_Regle_Comp_Hab.PSA, T_Regle_Comp_Hab.RSA, T_Regle_Comp_Hab.libellé "
Sql = Sql & "FROM T_Regle_Comp_Hab "
Sql = Sql & "ORDER BY T_Regle_Comp_Hab.libellé;"
Set Rs = Con.OpenRecordSet(Sql)
I = 1

While Rs.EOF = False
I = I + 1
Spreadsheet1.Cells(I, 1) = "'" & Rs!libellé
Spreadsheet1.Cells(I, 2) = "'" & Rs!ENCELADE
Spreadsheet1.Cells(I, 3) = "'" & Rs!RSA
Spreadsheet1.Cells(I, 4) = "'" & Rs!PSA
    Rs.MoveNext
Wend
Actualiser = True
Me.Show vbModal
End Sub

Private Sub Command3_Click()
Dim MyRange
Dim Pose As Long

RazFiltreEditExcel Me.Spreadsheet1
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion

For I = 2 To MyRange.Rows.Count
 MyRange(I, 1).Select
    If SercheValDouble(Val(I), "a", Me.Spreadsheet1, "" & MyRange(I, 1)) = True Then
        MyRange(I, 1).Select
        MsgBox "Risque de doublon sur cet Habillage", vbExclamation
        Actualiser = False
        Exit Sub
    End If
    If SercheValDouble(Val(I), "b", Me.Spreadsheet1, "" & MyRange(I, 2)) = True Then
        MyRange(I, 2).Select
        MsgBox "Risque de doublon sur cet Reff. Enc.", vbExclamation
        Actualiser = False
    Exit Sub
    End If

    If SercheValDouble(Val(I), "c", Me.Spreadsheet1, "" & MyRange(I, 3)) = True Then
        MyRange(I, 3).Select
        MsgBox "Risque de doublon sur cet Reff. RSA", vbExclamation
        Actualiser = False
         Exit Sub
    End If
    If SercheValDouble(Val(I), "d", Me.Spreadsheet1, "" & MyRange(I, 4)) = True Then
        MyRange(I, 4).Select
        MsgBox "Risque de doublon sur cet Reff. PSA.", vbExclamation
        Actualiser = False
         Exit Sub
    End If

Next
MyRange(1, 1).Select
Actualiser = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = NoQuitte
End Sub

Private Sub Spreadsheet1_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Actualiser = False
End Sub
Function SercheValDouble(Pose As Long, Colonne As String, Spreadsheet As Object, Serche As String) As Boolean
Dim MyRange
Dim Trouve As Boolean
If Trim(Serche) = "" Then Exit Function
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range(Colonne & "1").CurrentRegion
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range(Colonne & "1:" & Colonne & MyRange.Rows.Count)
For I = 2 To MyRange.Rows.Count
    If UCase(Serche) = UCase(MyRange(I)) Then
        If Trouve = False Then
            Trouve = True
        Else
            If I > Pose Then
                SercheValDouble = True
                Exit Function
            End If
            
        End If
    End If
Next
End Function
