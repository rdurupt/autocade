VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Editeur Excel:"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14925
   OleObjectBlob   =   "UserForm2.dsx":0000
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
Dim Msg As String
Dim MyErr As Boolean
Dim IfValidationOk As Boolean
Dim NbFinOuiNon As Long



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

Set MyRange = MyWorkbook.Worksheets("Notas").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Notas").Range(MyRange(2, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Notas").Select
MyWorkbook.Worksheets("Notas").Range("a1").Select
Set MyRange2 = Me.Spreadsheet4.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Notas").Paste
MyWorkbook.Worksheets("Notas").Range("a1").Select


Set MyRange = MyWorkbook.Worksheets("Composants").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Composants").Range(MyRange(2, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Composants").Select
MyWorkbook.Worksheets("Composants").Range("a1").Select
Set MyRange2 = Me.Spreadsheet3.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Composants").Paste
MyWorkbook.Worksheets("Composants").Range("a1").Select


Set MyRange = MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range(MyRange(2, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Ligne_Tableau_fils").Select
MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").Select
Set MyRange2 = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Ligne_Tableau_fils").Paste


MyWorkbook.Worksheets("Ligne_Tableau_fils").Range("a1").Select
Set MyRange = MyWorkbook.Worksheets("Connecteurs").Range("a1").CurrentRegion
MyWorkbook.Worksheets("Connecteurs").Range(MyRange(2, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address).Delete
MyWorkbook.Worksheets("Connecteurs").Select
MyWorkbook.Worksheets("Connecteurs").Range("a1").Select
Set MyRange2 = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
MyRange2.Copy
MyWorkbook.Worksheets("Connecteurs").Paste
MyWorkbook.Worksheets("Connecteurs").Range("a1").Select













 MyWorkbook.Save
Fin:
 
 
 Set MyRange2 = Nothing
 Set MyRange = Nothing
 MyWorkbook.Close False
 Set MyWorkbook = Nothing
 MyEcel.Quit
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
Dim sql As String

sql = "DELETE Ajout_LIAISON_CONNECTEURS.* FROM Ajout_LIAISON_CONNECTEURS;"
Con.Exequte sql
sql = "DELETE Ajout_LIAISON.* FROM Ajout_LIAISON;"
Con.Exequte sql

Msg = ""
IfValidationOk = True
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion


For i = 2 To MyRange.Rows.Count
Me.Spreadsheet1.Cells(i, 1).Select
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


Set MyRange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion


For i = 2 To MyRange.Rows.Count
Me.Spreadsheet2.Cells(i, 14).Select
    Me.Spreadsheet2.Cells(i, 14).Value = Me.Spreadsheet2.Cells(i, 14).Value
If Msg <> "" Then
    IfValidationOk = False
    Exit Sub
End If
Me.Spreadsheet2.Cells(i, 19).Select
    Me.Spreadsheet2.Cells(i, 19).Value = Me.Spreadsheet2.Cells(i, 19).Value
    If Msg <> "" Then
        IfValidationOk = False
        Exit Sub
    End If
DoEvents
Next i
If MyErr = True Then
    LoadLiasons.Charger MyClient
End If
MyErr = False
    IfValidationOk = False
End Sub

Private Sub CommandButton4_Click()
CommandButton3_Click
If Msg <> "" Then Exit Sub
CommandButton1_Click
End Sub

Private Sub Spreadsheet1_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Static NoMacro As Boolean
Static SaveRow As Long
Dim sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet1.ActiveCell.Row
Col = Me.Spreadsheet1.ActiveCell.Column

If (NoMacro = False) And (Row > 1) Then
NoMacro = True
   
 
    If Row > 1 Then
        If Trim("" & Me.Spreadsheet1.Cells(Row, 1)) <> "" Then
            Me.Spreadsheet1.Cells(Row, 5) = Row - 1
        End If
        If Trim("" & Me.Spreadsheet1.Cells(Row, 4)) <> "" Then
            sql = "SELECT LIAISON_CONNECTEURS.LIB FROM LIAISON_CONNECTEURS "
            sql = sql & "WHERE LIAISON_CONNECTEURS.CLIENT='" & MyReplace(MyClient) & "' "
            sql = sql & "AND LIAISON_CONNECTEURS.LIAISON='" & MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 4))) & "';"
            Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = False Then
                Me.Spreadsheet1.Cells(Row, 3) = Trim("'" & Rs!LIB)
            Else
                If IfValidationOk = False Then
                    If MsgBox("Le code App : " & DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 4)) & " n'existe pas" & vbCrLf & "Voulez-vous le créer", vbQuestion + vbYesNo, "Liaison Connecteur :") = vbYes Then
                        LibCode_APP = InputBox("Entrez la désignation du code APP : " & DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 4)), "Ajout d'un code App")
                        If Trim(LibCode_APP) <> "" Then
                            Me.Spreadsheet1.Cells(Row, 3) = LibCode_APP
                            sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
                            sql = sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 4)))) & "', '" & UCase(MyReplace(Me.Spreadsheet1.Cells(Row, 3))) & "' );"
                            Con.Exequte sql
                        End If
                    End If
                Else
                   sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON "
                   sql = sql & "FROM Ajout_LIAISON_CONNECTEURS "
                   sql = sql & "WHERE Ajout_LIAISON_CONNECTEURS.LIAISON='" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 4)))) & "' "
                   sql = sql & "AND Ajout_LIAISON_CONNECTEURS.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(sql)
                    If Rs.EOF = True Then
                        sql = "INSERT INTO Ajout_LIAISON_CONNECTEURS ( LIAISON, LIB,Job ) "
                        sql = sql & "values ( '" & UCase(MyReplace(DecodeCode_APP(Me.Spreadsheet1.Cells(Row, 4)))) & "', '" & MyReplace(Me.Spreadsheet1.Cells(Row, 3)) & "'," & NmJob & ");"
                        Con.Exequte sql
                        MyErr = True
                    End If
                End If
            End If
            Set Rs = Con.CloseRecordSet(Rs)
        
        End If
    End If
   
   
    
    NoMacro = False
    Col3 = 0
End If
End Sub

Private Sub Spreadsheet1_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Static NoMacro As Boolean
Static SaveRow As Long
Dim Row As Long
Row = Me.Spreadsheet1.ActiveCell.Row

If Row = 1 Then Exit Sub
If SaveRow = 0 Then SaveRow = Row
If NoMacro = True Then Exit Sub
    NoMacro = True
    
    If IfValidationOk = True Then 'NEANT
           
            If (Me.Spreadsheet1.Cells(Row, 1)) <> "" And Me.Spreadsheet1.Cells(Row, 4) = "" Then
            If (Me.Spreadsheet1.Cells(Row, 1)) <> "NEANT" Then
                Me.Spreadsheet1.Cells(Row, 4).Select
                
                MsgBox "Vous deves saisir le Code Appareil", vbCritical, "Code Appareil Connecteur"
                Msg = "?"
                End If
            End If
        Else
            If (SaveRow <> Row) And (Me.Spreadsheet1.Cells(SaveRow, 1)) <> "" And Me.Spreadsheet1.Cells(SaveRow, 4) = "" Then
            If (Me.Spreadsheet1.Cells(SaveRow, 1)) <> "NEANT" Then
                Me.Spreadsheet1.Cells(SaveRow, 4).Select
                
                MsgBox "Vous devez saisir le Code Appareil", vbExclamation, "Code Appareil Connecteur"
                Msg = "?"
            End If
            End If
    End If

SaveRow = Row


NoMacro = False
End Sub

Private Sub Spreadsheet2_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim MyRange
Dim Rs As Recordset
Dim sql As String
Dim LibCode_APP As String

Static Col3 As Long
Row = Me.Spreadsheet2.ActiveCell.Row
Col = Me.Spreadsheet2.ActiveCell.Column

If NoMacro = False Then
NoMacro = True
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


If (Col = 14) Or (Col = 19) Then
    If Row > 1 Then
        If Trim("" & Me.Spreadsheet2.Cells(Row, Col)) <> "" Then
            Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
            Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("D1", "d" & CStr(MyRange.Rows.Count))
            NoMacro = True
            
            i = ChercheXls(MyRange, UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))))
        
            If i <> 0 Then
                Msg = ""
                Me.Spreadsheet2.Cells(Row, Col - 1) = UCase(Trim("'" & MyRange(i, 2)))
                Me.Spreadsheet2.Cells(Row, Col - 2) = UCase(Trim("'" & MyRange(i, 4)))
                Me.Spreadsheet2.Cells(Row, Col - 3) = UCase(Trim("'" & MyRange(i, 3)))
                Else
                    Me.Spreadsheet2.Cells(Row, Col - 1) = "0"
                    Me.Spreadsheet2.Cells(Row, Col - 2) = ""
                    
                    Msg = "Le connecteur : " & UCase(Trim("" & Me.Spreadsheet2.Cells(Row, Col))) & " introuvable"
                    MsgBox Msg, vbQuestion
                End If
            
            
            
            
                Else
                    If Trim("" & Me.Spreadsheet2.Cells(Row, 1)) <> "" Then
                    If UCase(Trim("" & Me.Spreadsheet2.Cells(Row, 1))) <> "SUPPRIMER" Then
                        Msg = "Le code APP ne peut être Nul"
                        MsgBox Msg, vbExclamation, "Ligne_Tableau_fils"
                    End If
                    End If
                End If
            End If
        
        
        End If
   
        If (Row > 1) Then
        
            If Trim("" & Me.Spreadsheet2.Cells(Row, 1)) <> "" Then
                sql = "SELECT LIAISON.LIB FROM LIAISON "
                sql = sql & "WHERE LIAISON.CLIENT='" & MyReplace(MyClient) & "' "
                sql = sql & "AND LIAISON.LIAISON='" & MyReplace(Me.Spreadsheet2.Cells(Row, 1)) & "';"
                Set Rs = Con.OpenRecordSet(sql)
                If Rs.EOF = False Then
                Me.Spreadsheet2.Cells(Row, 2) = Trim("'" & Rs!LIB)
                Else
                    If IfValidationOk = False Then
                        If MsgBox("La liaison : " & Me.Spreadsheet2.Cells(Row, 1) & " n'existe pas" & vbCrLf & "Voulez-vous la créer", vbQuestion + vbYesNo, "Liaison Fils :") = vbYes Then
                            LibCode_APP = InputBox("Entrez la désignation de la liaison : " & Me.Spreadsheet2.Cells(Row, 1), "Ajout de liaison")
                            If Trim(LibCode_APP) <> "" Then
                                Me.Spreadsheet2.Cells(Row, 2) = LibCode_APP
                                sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
                                sql = sql & "VALUES ('" & UCase(MyReplace(MyClient)) & "'  , '" & UCase(MyReplace(Me.Spreadsheet2.Cells(Row, 1))) & "', '" & UCase(MyReplace(Me.Spreadsheet2.Cells(Row, 2))) & "' );"
                                Con.Exequte sql
                            End If
                        End If
                        Else
                        
                   sql = "SELECT Ajout_LIAISON.LIAISON "
                   sql = sql & "FROM Ajout_LIAISON "
                   sql = sql & "WHERE Ajout_LIAISON.LIAISON='" & UCase(MyReplace(Me.Spreadsheet2.Cells(Row, 1))) & "' "
                   sql = sql & "AND Ajout_LIAISON.Job= " & NmJob & ";"
                    Set Rs = Con.OpenRecordSet(sql)
                    If Rs.EOF = True Then
                        sql = "INSERT INTO Ajout_LIAISON ( LIAISON, LIB,Job ) "
                        sql = sql & "values ( '" & UCase(MyReplace(Me.Spreadsheet2.Cells(Row, 1))) & "', '" & MyReplace(Me.Spreadsheet2.Cells(Row, 2)) & "'," & NmJob & ");"
                        Con.Exequte sql
                        MyErr = True
                    End If
                    End If
                End If
                Set Rs = Con.CloseRecordSet(Rs)
            
            End If
            
            
            Col3 = 0
        End If
    End If
    NoMacro = False
End If
End Sub

Private Sub Spreadsheet3_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveRow As Long
Dim LibCode_APP As String
Static NoMacro As Boolean
Dim sql As String
Dim Rs As Recordset
Dim BoolOui As Boolean
If SaveRow = 0 Then SaveRow = 1
If NoMacro = True Then GoTo Fin

Row = Me.Spreadsheet3.ActiveCell.Row
Col = Me.Spreadsheet3.ActiveCell.Column
If SaveRow = 0 Then SaveRow = Row
If Row = 1 Then GoTo Fin
NoMacro = True

   If Trim("" & Me.Spreadsheet3.Cells(Row, 1)) <> "" Then Me.Spreadsheet3.Cells(Row, 2) = Row - 1
   NoMacro = False
Fin:
End Sub

Private Sub Spreadsheet3_SelectionChange(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Static SaveRow As Long
Dim LibCode_APP As String
Static NoMacro As Boolean
Dim sql As String
Dim Rs As Recordset
Dim BoolOui As Boolean
If SaveRow = 0 Then SaveRow = 1
If NoMacro = True Then GoTo Fin

Row = Me.Spreadsheet3.ActiveCell.Row
Col = Me.Spreadsheet3.ActiveCell.Column
If SaveRow = 0 Then SaveRow = Row
If Row = 1 Then GoTo Fin
NoMacro = True

   If Trim("" & Me.Spreadsheet3.Cells(Row, 1)) <> "" Then Me.Spreadsheet3.Cells(Row, 2) = Row - 1
    

If Col > 3 Then
    For i = 4 To NbFinOuiNon
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
 For i = 4 To NbFinOuiNon
    If Val(Me.Spreadsheet3.Cells(SaveRow, i)) = 1 Then
        BoolOui = True
        Exit For
    End If
    
    Next i
  If BoolOui = False Then
    MsgBox "Vous devez sélectionner un répertoire.", vbExclamation
    Me.Spreadsheet3.Cells(SaveRow, 4).Select
  End If
   
End If
SaveRow = Row
NoMacro = False
Fin:
End Sub

Private Sub Spreadsheet4_Change(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Dim Col As Long
Dim LibCode_APP As String
Static NoMacro As Boolean
Dim sql As String
Dim Rs As Recordset
Row = Me.Spreadsheet4.ActiveCell.Row
Col = Me.Spreadsheet4.ActiveCell.Column
If Row = 1 Then Exit Sub
If NoMacro = True Then Exit Sub
NoMacro = True
If Col = 1 Then
   If Trim("" & Me.Spreadsheet4.Cells(Row, 1)) <> "" Then Me.Spreadsheet4.Cells(Row, 2) = Row - 1
    
End If
NoMacro = False


End Sub

Private Sub UserForm_Activate()
Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
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
Set MyRange = MyWorkbook.Sheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
MyRange.Copy
Me.Spreadsheet2.ActiveSheet.Range("a1").Paste
Set MyRange = MyWorkbook.Sheets("Connecteurs").Range("a1").CurrentRegion
MyRange.Copy
Me.Spreadsheet1.ActiveSheet.Range("a1").Paste

Set MyRange = MyWorkbook.Sheets("Composants").Range("a1").CurrentRegion
NbFinOuiNon = MyRange.Columns.Count
MyRange.Copy
Me.Spreadsheet3.ActiveSheet.Range("a1").Paste

Set MyRange = MyWorkbook.Sheets("Notas").Range("a1").CurrentRegion
MyRange.Copy
Me.Spreadsheet4.ActiveSheet.Range("a1").Paste

MyEcel.AlertBeforeOverwriting = False

Set MyRange = Nothing
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
Me.Spreadsheet3.ActiveSheet.Range("a1").Select
Me.Spreadsheet2.ActiveSheet.Range("a1").Select
Me.Spreadsheet1.ActiveSheet.Range("a1").Select
Me.Spreadsheet1.Columns(2).NumberFormat = "Yes/No"
For i = 4 To 304
    Me.Spreadsheet3.Columns(i).NumberFormat = "Yes/No"
 Next i
    
DoEvents
End Sub
Public Sub chargement(Fichier As String, Client As String, Optional NouveauF As Boolean)
MyClient = Client
Nouveau = NouveauF
Me.Caption = Fichier
Me.Show vbModal
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NotSortie

End Sub

