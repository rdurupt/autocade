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
Public TYPEBASE As Integer
Public BASE As String
Public SERVER As String
Public Fichier As String
Public User As String
Public PassWord As String

Private Function GenereCSTRING()
'1 - ODBC
'2 - ORACLE
'3 - ACCESS 2000
'4 - ACCESS 97
'5 - SQL SERVER
'6 - SQL SERVER 2005 Express
If Trim("" & Fichier) = "" Then Fichier = BASE

Select Case TYPEBASE
    Case 1
        GenereCSTRING = "Provider=MSDASQL.1;Password=" & PassWord & ";Persist Security Info=True;User ID=" & User & ";Data Source=" & BASE
    Case 2
        GenereCSTRING = "Provider=OraOLEDB.Oracle.1;Password=" & PassWord & ";Persist Security Info=True;User ID=" & User & ";Data Source=" & BASE
   
    Case 3
        GenereCSTRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Fichier & ";Persist Security Info=False"
    Case 4
        GenereCSTRING = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & Fichier
    Case 5
            GenereCSTRING = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Fichier
            GenereCSTRING = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Fichier & ";user=" & User & ";Passwors=" & PassWord
    Case 6
        GenereCSTRING = "Provider=SQLOLEDB.1;Password=" & PassWord & ";Persist Security Info=True;User ID=" & User & ";Initial Catalog=" & BASE & ";Data Source=" & SERVER
   Case 7
    GenereCSTRING = "Provider=SQLNCLI.1;Password=" & PassWord & ";Persist Security Info=True;User ID=" & User & ";Initial Catalog=" & BASE & ";Data Source=" & SERVER
    

    Case Else
        GenereCSTRING = "PAS ASSEZ DE PARAMETRES RENSEIGNES !!!"

End Select
'Response.Write GenereCSTRING
'Response.End
End Function
Public Function RetournDbName(Extention As String) As String
    RetournDbName = Connexion.DefaultDatabase & "." & Extention
End Function
Public Function OpenConnetion() As Boolean
'Dim Fso As New FileSystemObject
If Trim("" & DRIVER) = "" Then DRIVER = "{Microsoft Access Driver (*.mdb)}"
    OpenConnetion = False
    On Error Resume Next
    Dim ConnecString As String
    ConnecString = "DRIVER=" & DRIVER & ";DBQ=" & BASE & ";user=" & UserDb & ";Passwors=" & PassWord
     Dim NbErr As Integer
Reprise:
    Set Connexion = New ADODB.Connection
    Connexion.Open GenereCSTRING  '"DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=C:\webprod\dbsportail\wwwencelade\Encelade_Menu.mdb"
'ConnecString

    
    If Err = 0 Then
        OpenConnetion = True
       Connexion.CommandTimeout = 14400
    Else
'    MsgBox Err.Description
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
Public Function OpenRecordSet(Connexion, Sql As String, IsRsOk)
On Error Resume Next
    Dim Rs

    Set Rs = CreateObject("ADODB.Recordset")
    'OpenRecordSet.LockType = adLockOptimistic

    Rs.Open Sql, Connexion
  
 
    
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
    Execute = False
    On Error Resume Next
    Dim NbErr As Integer
Reprise:
Debug.Print Sql
    Connexion.Execute Sql
    If Err = 0 Then
        Execute = True
        
        
        
        
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

