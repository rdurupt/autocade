Attribute VB_Name = "ModCompact"
Option Explicit

Sub CompactDb(Db As String)
  Dim fso As FileSystemObject
Dim DbAs As String
DbAs = Replace(Db, ".mdb", "_As.MDB")
Set fso = New FileSystemObject
If fso.FileExists(DbAs) = True Then
    fso.DeleteFile DbAs
End If
DBEngine.CompactDatabase Db, DbAs
MySeconde 1
 fso.DeleteFile Db
 fso.CopyFile DbAs, Db
End Sub

Sub MySeconde(Inter As Integer)
 Dim s As Integer
Dim sSave As Integer
Dim Sm As Integer
s = Second(Time)

If Sm = 0 Then Sm = s: sSave = Inter
While Inter <> 0
    If s <> sSave Then Inter = Val(Inter) - 1
    sSave = s
    s = Second(Time)
    DoEvents
Wend
 End Sub
Sub Main()
If Trim("" & Command) = "" Then
    DbAcces.Show
Else
    Compactage
End If
End Sub
Sub Compactage()
Dim Con As New Ado
Dim Rs As Recordset
Dim Frm As New Scrol
Dim Index As Long
Con.OpenConnetion App.Path & "\Access\CopactDb.mdb"
Set Rs = Con.OpenRecordSet("SELECT CompactDb.DB From CompactDb WHERE CompactDb.Executer=True;")
While Rs.EOF = False
    Index = Index + 1
    Rs.MoveNext
Wend
Rs.Requery
Frm.ProgressBar1.Value = 0
Frm.ProgressBar1.Max = Index
Frm.Visible = True
While Rs.EOF = False

Frm.Label1.Caption = "Compactage de : " & Rs!Db
DoEvents
    CompactDb "" & Rs!Db
Frm.ProgressBar1.Value = Frm.ProgressBar1.Value + 1
    Rs.MoveNext
    
Wend
Set Rs = Con.CloseRecordSet(Rs)
Con.CloseConnection
Unload Frm
End Sub
