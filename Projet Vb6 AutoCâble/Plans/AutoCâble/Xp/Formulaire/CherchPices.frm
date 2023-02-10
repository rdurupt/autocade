VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CherchPices 
   Caption         =   "Chercher Pièces :"
   ClientHeight    =   13800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18825
   OleObjectBlob   =   "CherchPices.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CherchPices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyFormCible As Object
Dim Filtre As String
Dim boolTxts As Boolean
Dim Noquite As Boolean


Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
If boolTxts = True Then
       MyFormCible.Tag = Me.Controls("txt" & CStr(1)).Tag
       GoTo Fin
End If
For i = 1 To 12

    MyFormCible.Controls("txt" & CStr(i)).Caption = Me.Controls("txt" & CStr(i)).Caption
    MyFormCible.Controls("txt" & CStr(i)).Tag = Me.Controls("txt" & CStr(i)).Tag

Next i
Fin:
Noquite = False
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Noquite = False
Me.Hide
End Sub

Private Sub Spreadsheet1_DblClick(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Row = Me.Spreadsheet1.ActiveCell.Row
strStatus = ""
PlanArchive = False
    If Row > 1 Then
        For i = 1 To 12
            Me.Controls("TXT" & CStr(i)).Caption = "" & Me.Spreadsheet1.Cells(Row, i)
             Me.Controls("TXT" & CStr(i)).Tag = "" & Me.Spreadsheet1.Cells(Row, 13)
             Select Case Me.Spreadsheet1.Cells(Row, 1).Interior.Color
                  
                  Case 16777164
                          strStatus = "CRE"
                  Case 10079487
                          strStatus = "MOD"
                
                  Case 13434828
                        strStatus = "VAL"
                   Case &HFFC0FF
                        strStatus = "VAL"
                        PlanArchive = True
        
             End Select
        Next i
    
    Else
        For i = 1 To 13
            Me.Controls("TXT" & CStr(i)).Caption = ""
             Me.Controls("TXT" & CStr(i)).Tag = "0"
        Next i
    End If
End Sub

Public Sub Charge(MyForm As Object, Optional Filtre As String, Optional boolTxt As Boolean, Optional boolArchive As Boolean)
Dim sql As String
Dim Rs As Recordset
Dim IndexRow As Long
Dim IndexCol As Long

boolTxts = boolTxt
IndexRow = 1
IndexCol = 0

sql = "SELECT SelectProjets.* "
sql = sql & "FROM SelectProjets; "
Set Rs = Con.OpenRecordSet(sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
While Rs.EOF = False
IndexRow = IndexRow + 1
For IndexCol = 0 To Rs.Fields.Count - 12
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
    If IndexCol > 3 And IndexCol < 8 Then
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 16))
    Else
        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 11))
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = True

Next IndexCol

Rs.MoveNext
Wend
If boolArchive = True Then
    sql = "SELECT Archive_SelectProjets.* "
sql = sql & "FROM Archive_SelectProjets; "
Set Rs = Con.OpenRecordSet(sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
While Rs.EOF = False
IndexRow = IndexRow + 1
For IndexCol = 0 To Rs.Fields.Count - 12
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
    If IndexCol > 3 And IndexCol < 8 Then
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 16))
    Else
        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = &HFFC0FF
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = True

Next IndexCol

Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)


Me.Show vbModal
End Sub

Function ChoixCouleur(Mode As Long) As Long
   
  
   Select Case Mode
    Case 1
        ChoixCouleur = 16777164
    Case 2
    ChoixCouleur = 10079487
    Case 3
        ChoixCouleur = 13434828
   End Select

End Function

Private Sub UserForm_Activate()
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
