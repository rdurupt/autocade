VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmHabillage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Habillage:"
   ClientHeight    =   12945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18240
   Icon            =   "FrmHabillage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12945
   ScaleWidth      =   18240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Actualiser"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   12120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enregistrer"
      Height          =   375
      Left            =   8100
      TabIndex        =   2
      Top             =   12120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   13200
      TabIndex        =   1
      Top             =   12120
      Width           =   1455
   End
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   11640
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
      Top             =   12720
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

Dim Myrange
RazFiltreEditExcel Me.Spreadsheet1
Sql = "DELETE T_Regle_Comp_Hab.* FROM T_Regle_Comp_Hab;"
Con.Exequte Sql
SqlInsert = "INSERT INTO T_Regle_Comp_Hab ( libellé, ENCELADE, RSA, PSA ) Values ("

Set Myrange = Spreadsheet1.Range("A1").CurrentRegion
ProgressBar1.Value = 0
ProgressBar1.Max = Myrange.Rows.Count

For i = 2 To "" & Myrange.Rows.Count
ProgressBar1.Value = i

    Sql = "'" & MyReplace("" & Myrange(i, 1)) & "',"
    Sql = Sql & "'" & MyReplace("" & Myrange(i, 2)) & "',"
    Sql = Sql & "'" & MyReplace("" & Myrange(i, 3)) & "',"
    Sql = Sql & "'" & MyReplace("" & Myrange(i, 4)) & "');"
    Con.Exequte SqlInsert & Sql
Next
NoQuitte = False
Unload Me

End Sub

Public Sub Chargement()

Dim Sql As String
Dim Rs As Recordset
Dim i As Long

NoQuitte = True
Sql = "SELECT T_Regle_Comp_Hab.ENCELADE, T_Regle_Comp_Hab.PSA, T_Regle_Comp_Hab.RSA, T_Regle_Comp_Hab.libellé "
Sql = Sql & "FROM T_Regle_Comp_Hab "
Sql = Sql & "ORDER BY T_Regle_Comp_Hab.libellé;"
Set Rs = Con.OpenRecordSet(Sql)
i = 1

While Rs.EOF = False
i = i + 1
Spreadsheet1.Cells(i, 1) = "'" & Rs!libellé
Spreadsheet1.Cells(i, 2) = "'" & Rs!ENCELADE
Spreadsheet1.Cells(i, 3) = "'" & Rs!RSA
Spreadsheet1.Cells(i, 4) = "'" & Rs!PSA
    Rs.MoveNext
Wend
Actualiser = True
Me.Show vbModal
End Sub

Private Sub Command3_Click()
Dim Myrange
Dim Pose As Long

RazFiltreEditExcel Me.Spreadsheet1
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion

For i = 2 To Myrange.Rows.Count
 Myrange(i, 1).Select
    If SercheValDouble(Val(i), "a", Me.Spreadsheet1, "" & Myrange(i, 1)) = True Then
        Myrange(i, 1).Select
        MsgBox "Risque de doublon sur cet Habillage", vbExclamation
        Actualiser = False
        Exit Sub
    End If
    If SercheValDouble(Val(i), "b", Me.Spreadsheet1, "" & Myrange(i, 2)) = True Then
        Myrange(i, 2).Select
        MsgBox "Risque de doublon sur cet Reff. Enc.", vbExclamation
        Actualiser = False
    Exit Sub
    End If

    If SercheValDouble(Val(i), "c", Me.Spreadsheet1, "" & Myrange(i, 3)) = True Then
        Myrange(i, 3).Select
        MsgBox "Risque de doublon sur cet Reff. RSA", vbExclamation
        Actualiser = False
         Exit Sub
    End If
    If SercheValDouble(Val(i), "d", Me.Spreadsheet1, "" & Myrange(i, 4)) = True Then
        Myrange(i, 4).Select
        MsgBox "Risque de doublon sur cet Reff. PSA.", vbExclamation
        Actualiser = False
         Exit Sub
    End If

Next
Myrange(1, 1).Select
Actualiser = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = NoQuitte
End Sub

Private Sub Spreadsheet1_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Actualiser = False
End Sub
Function SercheValDouble(Pose As Long, Colonne As String, Spreadsheet As Object, Serche As String) As Boolean
Dim Myrange
Dim Trouve As Boolean
If Trim(Serche) = "" Then Exit Function
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range(Colonne & "1").CurrentRegion
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range(Colonne & "1:" & Colonne & Myrange.Rows.Count)
For i = 2 To Myrange.Rows.Count
    If UCase(Serche) = UCase(Myrange(i)) Then
        If Trouve = False Then
            Trouve = True
        Else
            If i > Pose Then
                SercheValDouble = True
                Exit Function
            End If
            
        End If
    End If
Next
End Function
