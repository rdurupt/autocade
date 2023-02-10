Attribute VB_Name = "mod_Global2"
Public ColecAplication As Collection

Option Explicit
Public AppOff As Boolean

Const sql = "UPDATE ServicePop SET ServicePop.[Oui/non] = True " & _
            "WHERE ServicePop.Service='Stop';"
Public dirFTP As String
Public Interval As String
Public ServiceName As String
Public BdAutocable As String
Type MyLicGene
    Societe As String
    Tous As String
    AficheFrm As String
    DateDeb As String
    DateExecuter As String
    DateFin As String
    Enregistre As String
    NbJeton As String
    NbJetonActif As String
End Type

Type MyLic
    Serial As String
    PassWord As String
    UserName As String
    Enregistre As String
End Type
Type Licence
    Count As Long
    General As MyLicGene
    Record() As MyLic
End Type
Global FiledLicence As Licence

Global CodageX As New CDETXT
Global IsServeur As Boolean
Global Msg As String
Public Function RetournIdApp(Optional Application As String, Optional Retourn As Boolean, Optional Handle As Long) As Boolean
Dim Liste
Dim element
Dim Valid As String
Set Liste = GetObject("winmgmts:").InstancesOf("Win32_Process")

If Retourn = False Then
Set ColecAplication = Nothing
Set ColecAplication = New Collection
               

For Each element In Liste
    Debug.Print element.Name
    
        ColecAplication.Add element.Name, "Handle_" & element.Handle
    
Next element
Else
    On Error Resume Next
   
        Valid = ColecAplication("Handle_" & Handle)
        If Err Then
            Err.Clear
                RetournIdApp = False
        Else
            If UCase(Valid) = UCase(Application) Then
                RetournIdApp = True
            Else
                RetournIdApp = False
            End If
        End If
  
End If

End Function


Public Sub Main()
If App.PrevInstance Then
       End
    End If
Interval = InputDir(App.Path & "\Watch Dog.ini", "Interval")
ServiceName = InputDir(App.Path & "\Watch Dog.ini", "Service")
BdAutocable = InputDir(App.Path & "\Watch Dog.ini", "BdAutocable")
WatchDog.Show
End Sub

Public Function EcrirFile(NameFile As String) As Boolean
Dim FileNumber As Integer

On Error GoTo Fin
FileNumber = FreeFile   ' Get unused file
      ' number.
   Open NameFile For Output As #FileNumber   ' Create file name.
   Write #FileNumber, NameFile ' Output text.
   Close #FileNumber   ' Close file.
    EcrirFile = True
Fin:
End Function

Public Function Tremine() As Boolean
Dim Cmd As New ADODB.Command
Dim conn As ADODB.Connection
Dim ConnString As String
'Dim conn As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim Email As String
Dim strSubject
Dim strMsgId As String
Tremine = False
 ConnString = InputDir(App.Path & "\Watch Dog.ini", "ConnString")

    Set conn = New ADODB.Connection
   
    conn.ConnectionString = ConnString
    conn.Open
' ProgBusy = False
    Set Cmd.ActiveConnection = conn
   Cmd.CommandText = sql
   Cmd.Execute
 Set Cmd = Nothing
  Set conn = Nothing
  dirFTP = InputDir(App.Path & "\Watch Dog.ini", "dirFTP")
  If Right(dirFTP, 1) <> "\" Then dirFTP = dirFTP & "\"

If Dir(dirFTP & "*.*") <> "" Then
    Kill dirFTP & "*.*"
End If
EcrirFile dirFTP & "ServeuroFF.txt"
Unload WatchDog
End
End Function

Public Function InputDir(Fichier As String, Recherche As String) As String
'Permet de rechercher des information dans le fichier Fax.ini
Dim pose As Integer
Dim FileNumber As Integer
Dim MyString As String
FileNumber = FreeFile
On Error Resume Next
If Err = 0 Then
Open Fichier For Input As #FileNumber


Do While Not EOF(FileNumber)   ' Effectue la boucle jusqu'à la fin du fichier.
   Input #FileNumber, MyString  ' Lit les données dans la variables.
   pose = InStr(1, UCase(MyString), UCase(Recherche))
   If pose <> 0 Then
        InputDir = Mid(MyString, Len(Recherche) + 2, Len(MyString) - (Len(Recherche) + 1))
        InputDir = Trim(InputDir)
        
        Close #FileNumber   ' Ferme le fichier.
       
         Exit Function
   End If
Loop
Close #FileNumber   ' Ferme le fichier.

Else
''    boolKill = False
Err.Clear
 End If

End Function
Public Sub KillProcessus()
Dim sql As String
Dim Rs As Recordset
Dim Con As New Ado
On Error Resume Next
sql = "SELECT T_Job.DateDebut, T_Job.FinTraitement, T_Job.IdApp, "
sql = sql & "T_Job.IdAutocad, T_Job.IdExcel, T_Job.IdExcel2, T_Job.IdWord "
sql = sql & "From T_Job "
sql = sql & "WHERE T_Job.DateDebut Is Not Null  "
sql = sql & "AND T_Job.FinTraitement=False;"
Con.BASE = BdAutocable
Con.TYPEBASE = 5
Con.OpenConnetion

Set Rs = Con.OpenRecordSet(sql)
While Rs.EOF = False
    cmddelproc Rs!IdApp, "AutoCable.exe"
    cmddelproc Rs!IdAutocad, "acad.exe"
    cmddelproc Rs!IdExcel, "EXCEL.EXE"
    cmddelproc Rs!IdExcel2, "EXCEL.EXE"
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
sql = "UPDATE T_Job SET T_Job.DateDebut = Null, T_Job.IdApp = 0,  "
sql = sql & "T_Job.IdAutocad = 0, T_Job.IdExcel = 0, T_Job.IdExcel2 = 0,T_Job.IdWord = 0 "
sql = sql & "WHERE T_Job.DateDebut Is Not Null  "
sql = sql & "AND T_Job.FinTraitement=False;"
Con.Execute sql

sql = "SELECT ServerAutocad.IdAutcad From ServerAutocad WHERE ServerAutocad.Id=1;"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
    cmddelproc Rs!IdAutcad, "acad.exe"
    
End If
sql = "UPDATE ServerAutocad SET ServerAutocad.IdAutcad = 0 WHERE ServerAutocad.Id=1;"
Con.Execute sql
Con.CloseConnection
End Sub
Public Sub cmddelproc(Proces As Long, Application As String)
On Error Resume Next
If Proces <> 0 Then
If RetournIdApp(Application, True, Proces) = False Then Exit Sub
Dim ServiceObject As SWbemObject 'Variable de type Objet WMI
Dim Locator As SWbemLocator 'Variable de type Objet de connexion
Dim services As SWbemServices 'Variable de type Objet services
Dim P
Set Locator = New SWbemLocator 'Nouvelle instance d'une connexion


'Connexion au serveur
Set services = Locator.ConnectServer("")

'Récupération du processus selectionné
Set ServiceObject = services.Get("Win32_Process='" & Proces & "'")
        'Destruction du processus
        P = ServiceObject.Terminate
'Le kill a reussi
End If
End Sub

