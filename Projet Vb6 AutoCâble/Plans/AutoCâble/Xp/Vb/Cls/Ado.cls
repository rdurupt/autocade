VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Connexion As ADODB.Connection
Public Function OpenConnetion(db As String) As Boolean
    OpenConnetion = False
    On Error Resume Next
    Dim ConnecString As String
    ConnecString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & db
    Set Connexion = New ADODB.Connection
    Connexion.Open ConnecString
    If Err = 0 Then
        OpenConnetion = True
       Connexion.CommandTimeout = 14400
    End If
    Debug.Print Err.Description
    Err.Clear
    On Error GoTo 0
End Function
Public Function CloseConnection() As Boolean
CloseConnection = False
On Error Resume Next
    Connexion.Close
    Set Connexion = Nothing
     If Err = 0 Then
        CloseConnection = True
    End If
    Err.Clear
    On Error GoTo 0
End Function
Public Function OpenRecordSet(Sql As String)
On Error Resume Next
    Dim Rs As Recordset
    Set OpenRecordSet = New Recordset
    Set Rs = Connexion.Execute(Sql)
    Set OpenRecordSet = Rs
    If Err Then
    MsgBox Err.Description
    End If
End Function
Public Function CloseRecordSet(Rs As Recordset)
    Rs.Close
    Set CloseRecordSet = Nothing
End Function
Public Function Exequte(Sql As String) As Boolean
    Exequte = False
    On Error Resume Next
    Connexion.Execute Sql
    If Err = 0 Then
        Exequte = True
     Else
        MsgBox Err.Description
    End If
    
    Err.Clear


End Function

