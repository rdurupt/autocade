VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14925
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
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
Dim MSG As String

Private Sub CommandButton1_Click()

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
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
    MyWorkbook.SaveAs Me.Caption
    If Err Then
        BoolErr = True
        MsgBox Err.Description
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
Set MyRange = MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range(MyRange(2, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address).Delete

Set MyRange = MyWorkbook.Worksheets("Connecteurs").Range("a1").CurrentRegion

MyWorkbook.Worksheets("Connecteurs").Range(MyRange(2, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address).Delete
Set MyRange2 = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion
 Me.Spreadsheet2.ActiveSheet.Range(MyRange2(2, 1).Address & ":" & MyRange2(MyRange2.Rows.Count + 1, MyRange2.Columns.Count).Address).Copy
 MyWorkbook.Worksheets("Ligne_Tableau_fils").Select
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a2").Select
 MyWorkbook.Worksheets("Ligne_Tableau_fils").Paste
 MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").Select
 
 Set MyRange2 = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
 Me.Spreadsheet1.ActiveSheet.Range(MyRange2(2, 1).Address & ":" & MyRange2(MyRange2.Rows.Count + 1, MyRange2.Columns.Count).Address).Copy
 MyWorkbook.Worksheets("Connecteurs").Select
MyWorkbook.Worksheets("Connecteurs").Range("a2").Select
 MyWorkbook.Worksheets("Connecteurs").Paste
  MyWorkbook.Worksheets("Connecteurs").Range("a1").Select
  MyWorkbook.Worksheets("Ligne_Tableau_fils").Select
  MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").Select

 MyWorkbook.Save
Fin:
 MyWorkbook.Close False
 MyEcel.Quit
 Set MyRange2 = Nothing
 Set MyRange = Nothing
 Set MyWorkbook = Nothing
 Set MyEcel = Nothing
 If BoolErr = False Then boolExcute = True
 NotSortie = False
Me.Hide
End Sub

Private Sub CommandButton2_Click()
MenuShow = True
 boolExcute = False
 NotSortie = False
Me.Hide
End Sub

Private Sub CommandButton3_Click()
Dim MyRange

Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion


For I = 2 To MyRange.Rows.Count
Me.Spreadsheet1.Cells(I, 1).Select
    Me.Spreadsheet1.Cells(I, 1).Value = Me.Spreadsheet1.Cells(I, 1).Value

DoEvents
Next I


Set MyRange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion


For I = 2 To MyRange.Rows.Count
Me.Spreadsheet2.Cells(I, 14).Select
    Me.Spreadsheet2.Cells(I, 14).Value = Me.Spreadsheet2.Cells(I, 14).Value
If MSG <> "" Then Exit Sub
Me.Spreadsheet2.Cells(I, 19).Select
    Me.Spreadsheet2.Cells(I, 19).Value = Me.Spreadsheet2.Cells(I, 19).Value
If MSG <> "" Then Exit Sub
DoEvents
Next I

End Sub

Private Sub CommandButton4_Click()
CommandButton3_Click
If MSG <> "" Then Exit Sub
CommandButton1_Click
End Sub

Private Sub Spreadsheet1_BeforeCommand(ByVal EventInfo As OWC.SpreadsheetEventInfo)

End Sub

Private Sub Spreadsheet1_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Row = Me.Spreadsheet1.ActiveCell.Row
Col = Me.Spreadsheet1.ActiveCell.Column
If Row > 1 Then
    If Trim("" & Me.Spreadsheet1.Cells(Row, 1)) <> "" Then
        Me.Spreadsheet1.Cells(Row, 5) = Row - 1
    End If
End If
End Sub

Private Sub Spreadsheet2_BeforeCommand(ByVal EventInfo As OWC.SpreadsheetEventInfo)

End Sub

Private Sub Spreadsheet2_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim MyRange
Dim Rs As Recordset
Dim Sql As String
Static Col3 As Long
Row = Me.Spreadsheet2.ActiveCell.Row
Col = Me.Spreadsheet2.ActiveCell.Column
If Trim("" & Me.Spreadsheet2.Cells(Row, 1)) <> "" Then
    If (Row > 1) And (Row = 2) Then
       If Trim("" & Me.Spreadsheet2.Cells(Row, 3)) <> Col3 Then
        Me.Spreadsheet2.Cells(Row, 3) = 1
        Col3 = 1
        End If
    Else
        If (Row > 1) And (Row <> 2) Then
               If Trim("" & Me.Spreadsheet2.Cells(Row, 3)) <> Col3 Then
               Col3 = Me.Spreadsheet2.Cells(Row - 1, 3) + 1
                Me.Spreadsheet2.Cells(Row, 3) = Me.Spreadsheet2.Cells(Row - 1, 3) + 1
               End If
        End If
    
    End If

End If
If NoMacro = False Then
If (Col = 14) Or (Col = 19) Then
    If Row > 1 Then
        If Trim("" & Me.Spreadsheet2.Cells(Row, Col)) <> "" Then
            Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
             Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("D1", "d" & CStr(MyRange.Rows.Count))
            NoMacro = True
           
          I = ChercheXls(MyRange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))))
     
                If I <> 0 Then
                MSG = ""
                    Me.Spreadsheet2.Cells(Row, Col - 1) = UCase(Trim("'" & MyRange(I, 2)))
                    Me.Spreadsheet2.Cells(Row, Col - 2) = UCase(Trim("'" & MyRange(I, 4)))
                    Me.Spreadsheet2.Cells(Row, Col - 3) = UCase(Trim("'" & MyRange(I, 3)))
                Else
                    Me.Spreadsheet2.Cells(Row, Col - 1) = "0"
                    Me.Spreadsheet2.Cells(Row, Col - 2) = ""

                MSG = "Le connecteur : " & UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))) & " introuvable"
                MsgBox MSG, vbQuestion
                End If

           
           
          NoMacro = False
        Else
            If Trim("" & Me.Spreadsheet2.Cells(Row, 1)) <> "" Then
                MSG = "Le code APP ne peut être Nul"
                MsgBox MSG, vbExclamation, "Ligne_Tableau_fils"

            End If
        End If
    End If


End If
End If
If (NoMacro = False) And (Row > 1) Then
NoMacro = True
    Con.OpenConnetion db
    If Trim("" & Me.Spreadsheet2.Cells(Row, 1)) <> "" Then
        Sql = "SELECT LIAISON.LIB FROM LIAISON "
        Sql = Sql & "WHERE LIAISON.CLIENT='" & MyReplace(MyClient) & "' "
        Sql = Sql & "AND LIAISON.LIAISON='" & MyReplace(Me.Spreadsheet2.Cells(Row, 1)) & "';"
        Set Rs = Con.OpenRecordSet(Sql)
        If Rs.EOF = False Then
             Me.Spreadsheet2.Cells(Row, 2) = Trim("'" & Rs!LIB)
        Else
             Me.Spreadsheet2.Cells(Row, 2) = ""
        End If
         Set Rs = Con.CloseRecordSet(Rs)
    Else
         Me.Spreadsheet2.Cells(Row, 2) = ""
    End If
    
    Con.CloseConnection
    NoMacro = False
    Col3 = 0
End If
End Sub

Private Sub UserForm_Activate()
Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Set MyEcel = New EXCEL.Application
NotSortie = True
'MyEcel.Visible = True
If Trim(Me.Caption) = "" Then Exit Sub
If Nouveau = False Then
    Set MyWorkbook = MyEcel.Workbooks.Open(Me.Caption)
Else

    Set MyWorkbook = MyEcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
End If
Set MyRange = MyWorkbook.Sheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
MyRange.Copy
Me.Spreadsheet2.ActiveSheet.Range("a1").Paste
Set MyRange = MyWorkbook.Sheets("Connecteurs").Range("a1").CurrentRegion
MyRange.Copy
Me.Spreadsheet1.ActiveSheet.Range("a1").Paste
MyEcel.AlertBeforeOverwriting = False
MyWorkbook.Close False
MyEcel.Quit
Set MyRange = Nothing
Set MyWorkbook = Nothing
Set MyExcel = Nothing
'Me.Spreadsheet1.ActiveSheet.Panes(1).VisibleRange = False
Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet2.ActiveSheet.Range("a1").Select
Me.Spreadsheet1.ActiveSheet.Range("a1").Select

DoEvents
End Sub
Public Sub Chargement(Fichier As String, Client As String, Optional NouveauF As Boolean)
MyClient = Client
Nouveau = NouveauF
Me.Caption = Fichier
Me.Show
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NotSortie
End Sub
