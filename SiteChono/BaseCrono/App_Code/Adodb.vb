Imports Microsoft.VisualBasic

Public Class Adodb
    Public Sql As String
    Private Con As Object
    Public Sub OpenConnection(ByVal Base As String)
        Dim ConecString As String
        Con = CreateObject("ADODB.Connection")
        Con.Mode = 16

        '

        ConecString = "Provider=SQLOLEDB.1;Password=dur1234/*-;Persist Security Info=True;User ID=rd;Initial Catalog=" & Base & ";Data Source=autocable; "
        Con.Open(ConecString)

    End Sub
    Public Function OpenRecordset()

        'Cn.execute(sql)
        'sql = "select test.* from test"
        OpenRecordset = CreateObject("adodb.recordset")
        OpenRecordset.open(Sql, Con)
    End Function
    Public Function Execute(ByVal Sql As String) As Boolean
        On Error Resume Next
        Execute = True
        Con.execute(Sql)
        If Err.Number <> 0 Then
            Err.Clear()
            Execute = False
        End If
    End Function
    Public Sub CloseConnection()
        On Error Resume Next

        Con.close()
        Con = Nothing
    End Sub
End Class
