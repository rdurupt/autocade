Attribute VB_Name = "API"

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Machine  As String
'API pour trouver le hwnd d une fenetre en donnant sa caption
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function CreateProcessWithLogon Lib "Advapi32" Alias "CreateProcessWithLogonW" (ByVal lpUsername As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public AutoApp As Object
Global boolAutoCAD As Boolean
Public CompactOk As Boolean


Function LirMachineName() As String
    'Retourne le nom de l'utilisateur courant de l'ordinateur
   
    Dim stTmp As String, lgTmp As Long
    stTmp = Space$(250)
    lgTmp = 251
    Call GetComputerName(stTmp, lgTmp)
    LirMachineName = Mid$(stTmp, 1, InStr(1, stTmp, Chr$(0)) - 1)
End Function


Public Sub focuswindow(LaWindow As String) 'sub pour amener une fenetre au prenier plan
  Dim rep As Long
  rep = BringWindowToTop(FindWindow(vbNullString, GetTrueTitle(LaWindow)))
End Sub

Public Function GetTrueTitle(Ftitle As String) As String 'fonction ki retourne le vrai nom de la fenetre
'par exemple si vous passer comme argument: "ma fenetre [Disable] [Hidden]" elle
'retourne: "ma fenetre"
  Dim pos1 As Integer, pos2 As Integer
  Dim rep As String
  
  pos1 = InStr(1, Ftitle, "[Hidden]") 'on cherche la position de "[Hidden]"
  pos2 = InStr(1, Ftitle, "[Disabled]") 'on cherche la position de "[Disable]"
  If (pos1 > pos2) And (pos2 > 0) Then 'si on a trouver "[Disable]" & "[Hidden]" alors...
    rep = Mid(Ftitle, 1, Len(Ftitle) - Len(" [Disabled] [Hidden]")) 'on decoupe la chaine
  ElseIf (pos1 = 0) And (pos2 > 0) Then 'si on a trouver que "[Disable]" alors...
    rep = Mid(Ftitle, 1, Len(Ftitle) - Len(" [Disabled]")) 'on decoupe la chaine
  ElseIf (pos1 > 0) And (pos2 = 0) Then 'si on a trouver que "[Hidden]" alors...
    rep = Mid(Ftitle, 1, Len(Ftitle) - Len(" [Hidden]")) 'on decoupe la chaine
  Else 'si il n y a rien alors...
    rep = Ftitle 'pas de modification
  End If
  GetTrueTitle = rep 'on retourne le vrai nom de la fenetre
End Function

Public Function RetournIdAppName(Handle As Long) As String
Dim Liste
Dim element
Dim ColecAplication As New Collection
Dim Valid As String
Set Liste = GetObject("winmgmts:").InstancesOf("Win32_Process")


               

For Each element In Liste
    Debug.Print element.Name&; " : " & element.Handle & " " & Handle
       
            ColecAplication.Add element.Name, "Handle_" & element.Handle
       
Next element
On Error Resume Next
RetournIdAppName = ColecAplication("Handle_" & Handle)
If Err Then
    Err.Clear
    RetournIdAppName = ""
End If
On Error GoTo 0
End Function
Public Function RetournIdApp(Application As String, Optional Retourn As Boolean) As Long
Dim Liste
Dim element
Dim Valid
Dim ColecAplication As New Collection
RetournIdApp = -1
Set Liste = GetObject("winmgmts:").InstancesOf("Win32_Process")

If Retourn = False Then
               

For Each element In Liste
    Debug.Print element.Name
    If UCase(element.Name) = UCase(Application) Then
        ColecAplication.Add element.Handle, element.Handle
    End If
Next element
Else
    On Error Resume Next
    For Each element In Liste
    Debug.Print element.Name & " : " & element.Handle
    If UCase(element.Name) = UCase(Application) Then
        Valid = ColecAplication(element.Handle)
        If Err Then
            Err.Clear
                RetournIdApp = element.Handle
                Exit For
        End If
    End If
Next element
End If

End Function

Public Function MyExecute(Fichier As String, Apli As String) As Long
Dim lapi As Long
On Error Resume Next
RetournIdApp Apli
lapi = ShellExecute(100, "open", Fichier, vbNull, vbNull, 5)
MyExecute = RetournIdApp(Apli, True)
If Err Then MsgBox Err.Description
Err.Clear
On Error GoTo 0
End Function
Public Sub StratProcess(lpApplicationName As String, lpCommandLine As String)
    Dim Sql As String
    Dim rs As Recordset
    
    Dim lpUsername As String, lpDomain As String, lpPassword As String
    Dim lpCurrentDirectory As String
    Dim StartInfo As STARTUPINFO, ProcessInfo As PROCESS_INFORMATION
    Sql = "SELECT AdminAutocable.User, AdminAutocable.PassWord, AdminAutocable.Serveur,AdminAutocable.Service "
Sql = Sql & "FROM AdminAutocable;"
Set rs = Con.OpenRecordSet(Sql)
    lpUsername = "" & rs!User
    lpDomain = ""
    lpPassword = "" & rs!PassWord
    
    Set rs = Con.CloseRecordSet(rs)
    lpApplicationName = Trim("" & lpApplicationName) & " "
    lpCommandLine = " " & Trim("" & lpCommandLine) & " "
'    lpCommandLine = " \\10.30.0.5\production\Cablage-production\RENAULT\PI\662\16-PI\PI_662_05_1445_1\12-PL\PL_662_05_1444_1.dwg "

'    lpCommandLine = vbNullString 'use the same as lpApplicationName
    lpCurrentDirectory = ""  'use standard directory
    StartInfo.cb = LenB(StartInfo) 'initialize structure
    StartInfo.dwFlags = 0&
    CreateProcessWithLogon StrPtr(lpUsername), StrPtr(lpDomain), StrPtr(lpPassword), LOGON_WITH_PROFILE, StrPtr(lpApplicationName), StrPtr(lpCommandLine), CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or CREATE_NEW_PROCESS_GROUP, ByVal 0&, StrPtr(lpCurrentDirectory), StartInfo, ProcessInfo
    CloseHandle ProcessInfo.hThread 'close the handle to the main thread, since we don't use it
    CloseHandle ProcessInfo.hProcess 'close the handle to the process, since we don't use it
    'note that closing the handles of the main thread and the process do not terminate the process
    'unload this application
End Sub
Public Sub LireLicenceAcad()
Dim IdApp As Long
Dim Fso As New FileSystemObject
Dim FileNumber As Long
Dim MyString As String
Dim txt, txtError As String
Dim SpliFile
Dim NuLigne As Long
Dim i As Long
Dim NbLicence As String
Dim Sql As String
Dim rs As Recordset
''If boolAutoCAD = True Then Exit Sub
'If Fso.FileExists("c:\AutocableLicenceAcad\not Map Sur X.txt") = True Then Fso.DeleteFile "c:\AutocableLicenceAcad\not Map Sur X.txt"
'
'If Fso.FileExists("X:\Utilitaires\FLEXLM\lmutil.exe") = False Then
'    On Error Resume Next
'    Fso.CreateTextFile "c:\AutocableLicenceAcad\not Map Sur X.txt"
'   Shell "subst.exe" & " X: ""\\10.30.0.5\donnees d entreprise""", vbHide
' End If
'If Fso.FolderExists("c:\AutocableLicenceAcad") = False Then Fso.CreateFolder "c:\AutocableLicenceAcad"
'If Fso.FileExists("c:\AutocableLicenceAcad\Licence.txt") = True Then Fso.DeleteFile "c:\AutocableLicenceAcad\Licence.txt"
'Set Fso = Nothing
'IdApp = MyExecute(App.path & "\Licence.bat", "cmd.exe")
'  While RetournIdAppName(IdApp) <> ""
'  DoEvents
'  Wend
'
'FileNumber = FreeFile
'MyString = ""
'
'Open "c:\AutocableLicenceAcad\Licence.txt" For Input As #FileNumber
'Do While Not EOF(FileNumber)    ' Effectue la boucle jusqu'à la fin du fichier.
'   Line Input #FileNumber, txt  ' Lit les données dans deux variables.
'    MyString = MyString & txt & vbCrLf
'Loop
'Close #FileNumber    ' Ferme le fichier
'NuLigne = 0
'  SpliFile = Split(MyString, vbCrLf)
'  For i = UBound(SpliFile) To 0 Step -1
'  If Trim(SpliFile(i)) <> "" Then
'    If Trim(SpliFile(i)) = "floating license" Then Exit For
'    If InStr(1, UCase(Trim(SpliFile(i))), UCase("License file(s)")) <> 0 Then NuLigne = NuLigne + 1
'
'
'  End If
'
'  Next
' For i = UBound(SpliFile) To 0 Step -1
'
'    If InStr(SpliFile(i), " (Total of") <> 0 Then Exit For
'
'
'  Next
'  NbLicence = SpliFile(i)
'  NbLicence = Trim(Mid(NbLicence & Space(100), InStr(NbLicence, " (Total of") + Len(" (Total of") + 1, 100))
'   NbLicence = Trim(Replace(NbLicence, "licenses available)", ""))
'   If NuLigne = Val(NbLicence) Then Exit Sub
'loadFichierAppliWin
    RetournIdApp "Acad.exe"
    Err.Clear
    
'     If PathAppliAutocad = "ERR" Then
'                MsgBox "L'exécutable d'Autocad n'a pas été trouvée"
'            Else
'                  StratProcess PathAppliAutocad, ""
'          End If
'          Err.Clear
'          MySeconde 5
'          Set AutoApp = GetObject(, "autocad.application")
'        If Err = 0 Then
'
'            AutoApp.Visible = False
'            Example_AutoAudit
''            AutoApp.Documents(0).Close False
'            DoEvents
'            IsCilent = False
'        Else
'            MsgBox "Plus de licence Autocad disponible", vbInformation, "AutoCâble  licence :"
'            boolAutoCAD = False
'        End If
    
  Set AutoApp = CreateObject("AutoCAD.Application")
'  AutoApp.Visible = True
  If Err = 0 Then
'   sql = "SELECT MachinAplicationChemen.* "
'    sql = sql & "FROM MachinAplicationChemen  "
'    sql = sql & "WHERE MachinAplicationChemen.Machine='" & Machine & "';"
' Set Rs = Con.OpenRecordSet(sql)
' If Rs.EOF = False Then
''    PathAppliExcel = "" & Rs!EXCEL
'    AutoApp.Quit
'    StratProcess "" & Rs!AutoCAD, ""
' End If
''End If
 
  MySeconde 10
' Set AutoApp = GetObject(, "autocad.application")
'   If Err = 0 Then
       
        Sql = "UPDATE ServerAutocad SET ServerAutocad.IdAutcad = " & RetournIdApp("Acad.exe", True) & "  "
        Sql = Sql & "WHERE ServerAutocad.Id=1;"
        Con.Execute Sql
        Con.Execute Sql
        Con.Execute Sql
         boolAutoCAD = True
   
'            AutoApp.Documents(0).Close False
            AutoApp.Visible = False
             DoEvents
        Else
           
'            AutoCableNbTest = AutoCableNbTest + 1
'            If AutoCableNbTest < 4 Then GoTo RepriseTest
            txtError = Err.Description
            boolAutoCAD = False
            Err.Clear
            Debug.Print txtError
        NuLigne = NuLigne + 1
        Trace.Ecrire "Ligne " & str(NuLigne) & " " & txtError
            boolAutoCAD = False
        End If
        DoEvents
End Sub
Public Sub cmddelproc(Id As Long)
On Error Resume Next

Dim ServiceObject As SWbemObject 'Variable de type Objet WMI
Dim Locator As SWbemLocator 'Variable de type Objet de connexion
Dim services As SWbemServices 'Variable de type Objet services
Dim P
Set Locator = New SWbemLocator 'Nouvelle instance d'une connexion


'Connexion au serveur
Set services = Locator.ConnectServer("")

'Récupération du processus selectionné
Set ServiceObject = services.Get("Win32_Process='" & Id & "'")
        'Destruction du processus
        P = ServiceObject.Terminate
'Le kill a reussi

End Sub
