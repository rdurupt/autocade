VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "Codes Liaisons:"
   ClientHeight    =   12390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14070
   OleObjectBlob   =   "UserForm6.dsx":0000
   StartUpPosition =   1  'CenterOwner
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
Dim sql As String
Dim Rs As Recordset
If MsgBox("Voulez vous enregistre les modification.", vbQuestion + vbYesNo) = vbNo Then Exit Sub
Me.CommandButton1.Enabled = False
Me.CommandButton2.Enabled = False
Me.MultiPage1.Enabled = False
sql = "UPDATE LIAISON SET LIAISON.Sup = True;"
Con.Exequte sql
sql = "UPDATE LIAISON_CONNECTEURS SET LIAISON_CONNECTEURS.Sup = True;"
Con.Exequte sql
Set MyRange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion

Rangecount = MyRange.Rows.Count
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion

Rangecount = Rangecount + MyRange.Rows.Count
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = Rangecount
For i = 2 To MyRange.Rows.Count
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
DoEvents
    sql = "select LIAISON_CONNECTEURS.LIAISON from LIAISON_CONNECTEURS WHERE "
    sql = sql & "LIAISON_CONNECTEURS.CLIENT='" & MyReplace("" & MyRange(i, 1)) & "' AND "
     sql = sql & "LIAISON_CONNECTEURS.LIAISON='" & MyReplace("" & MyRange(i, 2)) & "' "
     Set Rs = Con.OpenRecordSet(sql)
     If Rs.EOF = False Then
        sql = "UPDATE LIAISON_CONNECTEURS SET LIAISON_CONNECTEURS.LIB ='" & MyReplace("" & MyRange(i, 3)) & "',  "
        sql = sql & "LIAISON_CONNECTEURS.Sup = False WHERE "
        sql = sql & "LIAISON_CONNECTEURS.CLIENT='" & MyReplace("" & MyRange(i, 1)) & "' AND "
        sql = sql & "LIAISON_CONNECTEURS.LIAISON='" & MyReplace("" & MyRange(i, 2)) & "' "
        Con.Exequte sql
     Else
        sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
        sql = sql & "VALUES( '" & MyReplace("" & MyRange(i, 1)) & "', "
        sql = sql & "'" & MyReplace("" & MyRange(i, 2)) & "' ,"
        sql = sql & "'" & MyReplace("" & MyRange(i, 3)) & "');"
        Con.Exequte sql
     End If
Next i

Set MyRange = Me.Spreadsheet2.ActiveSheet.Range("a1").CurrentRegion
Rangecount = Rangecount + MyRange.Rows.Count
For i = 2 To MyRange.Rows.Count
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
DoEvents
    sql = "select LIAISON.LIAISON from LIAISON WHERE "
    sql = sql & "LIAISON.CLIENT='" & MyReplace("" & MyRange(i, 1)) & "' AND "
     sql = sql & "LIAISON.LIAISON='" & MyReplace("" & MyRange(i, 2)) & "' "
     Set Rs = Con.OpenRecordSet(sql)
     If Rs.EOF = False Then
        sql = "UPDATE LIAISON SET LIAISON.LIB ='" & MyReplace("" & MyRange(i, 3)) & "',  "
        sql = sql & "LIAISON.Sup = False WHERE "
        sql = sql & "LIAISON.CLIENT='" & MyReplace("" & MyRange(i, 1)) & "' AND "
        sql = sql & "LIAISON.LIAISON='" & MyReplace("" & MyRange(i, 2)) & "' "
        Con.Exequte sql
     Else
        sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
        sql = sql & "VALUES( '" & MyReplace("" & MyRange(i, 1)) & "', "
        sql = sql & "'" & MyReplace("" & MyRange(i, 2)) & "' ,"
        sql = sql & "'" & MyReplace("" & MyRange(i, 3)) & "');"
        Con.Exequte sql
     End If
Next i
Con.Exequte "DELETE LIAISON.*, LIAISON.Sup FROM LIAISON WHERE LIAISON.Sup=True;"
Con.Exequte "DELETE LIAISON_CONNECTEURS.*, LIAISON_CONNECTEURS.Sup FROM LIAISON_CONNECTEURS WHERE LIAISON_CONNECTEURS.Sup=True;"

Noquite = False
Me.Hide
End Sub

Public Sub chargement()
Dim Rs As Recordset
Dim Row As Long



Set Rs = Con.OpenRecordSet("SELECT RqLiaisonConnecteur.* FROM RqLiaisonConnecteur;")
FormBarGrah.ProgressBar1Caption.Caption = "Liaisons Connecteur:"
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
FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    Row = Row + 1
    For i = 0 To Rs.Fields.Count - 1
     Spreadsheet1.Cells(Row, i + 1) = "'" & Rs.Fields(i).Value
     Next i
     
Rs.MoveNext
Wend
Set Rs = Con.OpenRecordSet("SELECT RqLiaisonFils.*FROM RqLiaisonFils;")

FormBarGrah.ProgressBar1Caption.Caption = "Liaisons Fils:"
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
FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
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
Me.Show
End Sub

Private Sub UserForm_Activate()

 Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite

End Sub
