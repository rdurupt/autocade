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
Public Event LeftClick()
Private Connexion
Public Function RetournDbName(Extention As String) As String
    RetournDbName = Connexion.DefaultDatabase & "." & Extention
End Function
Public Function OpenConnetion(db As String, Optional DRIVER As String) As Boolean
Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")
If Trim("" & DRIVER) = "" Then DRIVER = "{Microsoft Access Driver (*.mdb)}"
    OpenConnetion = False
    On Error Resume Next
    Dim ConnecString As String
    ConnecString = "DRIVER=" & DRIVER & ";DBQ=" & db ' & ";user=" & PassDb.UserDb & ";Passwors=" & PassDb.PassWordDb
     Dim NbErr As Integer
Reprise:
    Set Connexion = CreateObject("ADODB.Connection")
    Connexion.Open ConnecString  '"DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=C:\webprod\dbsportail\wwwencelade\Encelade_Menu.mdb"
'ConnecString
    
    If Err = 0 Then
        OpenConnetion = True
       Connexion.CommandTimeout = 14400
    Else
    'MsgBox Err.Description
        Err.Clear
    NbErr = NbErr + 1
    If NbErr < 11 Then GoTo Reprise
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
    Dim Rs
Dim NbErr As Integer
Reprise:
Debug.Print Replace(Sql, "%", "*")
    Set Rs = CreateObject("ADODB.Recordset")
'    OpenRecordSet.LockType = adLockOptimistic
'
   ' OpenRecordSet.Open Sql, Connexion, adOpenDynamic
  Set Rs = Connexion.Execute(Sql)
    If Err Then
    MsgBox Err.Description '7701997031
    NbErr = NbErr + 1
        If NbErr < 11 Then
            Set OpenRecordSet = Nothing
            GoTo Reprise
        End If
    
    End If
    Err.Clear
    Set OpenRecordSet = Rs
End Function
Public Function RetournConnection()
Set RetournConnection = Connexion
End Function
Public Function OpenRecordSetParametre(Sql As String, Param)
Dim Commande
Dim Params
Set Commande = CreateObject("ADODB.Command")
Dim MyParameter
Set MyParameter = CreateObject("ADODB.Parameter")
Set Commande.ActiveConnection = Connexion
Commande.CommandText = "select Requête2.* from Requête2;"
 Commande.CommandType = adCmdText

 Set MyParameter = Commande.CreateParameter("[NumJob]", adNumeric)
         MyParameter.Value = 10
Commande.Parameters.Append MyParameter



'aa.Parameters.Append("MyRef") = "243410M660"
Set Rs2 = Commande.Execute

End Function
Public Function CloseRecordSet(Rs)
On Error Resume Next
    Rs.Close
    Set CloseRecordSet = Nothing
End Function
Public Function Execute(Sql As String) As Boolean
    Exequte = False
    On Error Resume Next
    Dim NbErr As Integer
Reprise:
Debug.Print Sql
    Connexion.Execute Sql
    If Err = 0 Then
        Exequte = True
        
        
        
        
'     Else
'    MsgBox Err.Description
'         Err.Clear
'    NbErr = NbErr + 1
'    If NbErr < 11 Then
'
'        GoTo Reprise
'    End If
Else
'    MsgBox Err.Description
    End If
    
    Err.Clear

End Function

