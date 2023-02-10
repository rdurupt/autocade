Attribute VB_Name = "Mod_MajBase"
Option Explicit

Public Sub Macro_Demarage_Serveur()
Dim Sql As String
Dim Cmd As New ADODB.Command
Dim ConnString As String
Dim conn As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim Email As String
On Error Resume Next
 ConnString = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "ConnString")

    Set conn = New ADODB.Connection
   
    conn.ConnectionString = UCase(ConnString)
    conn.Open
     Set Cmd.ActiveConnection = conn
Sql = "Macro_Demarage_Serveur"
Cmd.CommandText = Sql
     Set Rs = Cmd.Execute
     While Rs.EOF = False
     DoEvents
     Sql = Rs!Qry
     Debug.Print Sql
     Debug.Print "************************************************"
        conn.Execute Sql
        Rs.MoveNext
     Wend
      Set Rs = Nothing
     Set Cmd = Nothing
      Set conn = Nothing


End Sub
Public Function TransfaireBase()
If Dir(PathBase & "MAJ MCT_DATA.mdb") <> "" Then Kill PathBase & "MAJ MCT_DATA.mdb"
Unzip dirFTP & "MAJ MCT_DATA.zip"
On Error Resume Next
Reprise:

Kill dirFTP & "MAJ MCT_DATA.zip"
If Err Then
    Err.Clear
    GoTo Reprise

End If
On Error GoTo 0
Maj_Data

End Function
Sub Maj_Data()
Dim Sql As String
Dim Cmd As New ADODB.Command
Dim ConnString As String
Dim conn As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim Email As String
On Error Resume Next
 ConnString = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "ConnString")

    Set conn = New ADODB.Connection
   
    conn.ConnectionString = UCase(ConnString)
    conn.Open
     Set Cmd.ActiveConnection = conn
     Sql = "SELECT [Macro Import].* From [Macro Import] " & _
            "ORDER BY [Macro Import].Ordre;"
Cmd.CommandText = Sql
     Set Rs = Cmd.Execute
     While Rs.EOF = False
     DoEvents
     Sql = Rs!Qry
     Debug.Print Sql
     Debug.Print "************************************************"
        conn.Execute Sql
        Rs.MoveNext
     Wend
      Set Rs = Nothing
     Set Cmd = Nothing
      Set conn = Nothing
      If Dir(PathBase & "MAJ MCT_DATA.mdb") <> "" Then
        Kill PathBase & "MAJ MCT_DATA.mdb"
        FileCopy PathBase & "Model_MCT_DATA.mdb", PathBase & "MAJ MCT_DATA.mdb"
     End If
End Sub
Public Function STRATOK() As Boolean
Dim Sql As String
Dim ConnString As String
Dim conn As ADODB.Connection
On Error Resume Next
ConnString = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "ConnString")

    Set conn = New ADODB.Connection
   
    conn.ConnectionString = UCase(ConnString)
    conn.Open
     
     Sql = "UPDATE ServicePop SET ServicePop.[Oui/non] = False " & _
            "WHERE ServicePop.Service='StopAcquit' " & _
            "OR " & _
            "ServicePop.Service='Stop';"
        conn.Execute Sql
     Shell DirServerPop
      Set conn = Nothing
End Function
Public Function STOPSERVEUR() As Boolean
Dim Sql As String
Dim ConnString As String
Dim conn As ADODB.Connection
Dim Cmd As New ADODB.Command
Dim Rs As Recordset

On Error GoTo erreur
ConnString = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "ConnString")

    Set conn = New ADODB.Connection
   
    conn.ConnectionString = UCase(ConnString)
    conn.Open
    Set Cmd.ActiveConnection = conn
     Sql = "UPDATE ServicePop SET ServicePop.[Oui/non] = true " & _
            "WHERE ServicePop.Service='Stop';"
        conn.Execute Sql
        
Sql = "SELECT ServicePop.Service, ServicePop.[Oui/non] From ServicePop " & _
        "WHERE ServicePop.Service='StopAcquit' " & _
        "AND " & _
            "ServicePop.[Oui/non]=True;"
Cmd.CommandText = Sql
     Set Rs = Cmd.Execute
While Rs.EOF = True
    DoEvents
    Rs.Requery
Wend
Sql = "SELECT [Macro Exportaion].* From [Macro Exportaion] ORDER BY [Macro Exportaion].Ordre;"
Cmd.CommandText = Sql
Set Rs = Cmd.Execute
While Rs.EOF = False
Sql = Rs!Qry
DoEvents
'  On Error GoTo 0
   conn.Execute Sql ', , adCmdStoredProc
    Rs.MoveNext
Wend
        Set Rs = Nothing
      Set Cmd = Nothing
      Set conn = Nothing
zip PathBase & "MAJ MCT_DATA.mdb"
erreur:
'MsgBox Err.Description
End Function
