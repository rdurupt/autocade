Attribute VB_Name = "ModApi"
Option Explicit
Private Type CURSOR
    x As Long
    y As Long
End Type
'définitions des variables
Public pos As CURSOR
'var de la position du curseur
Public PosCursor As CURSOR
Public lngTokenHandle, lngLogonType, lngLogonProvider As Long
Public blnResult As Boolean
Public Const LOGON32_LOGON_INTERACTIVE = 2
Public Const LOGON32_PROVIDER_DEFAULT = 0

Public Declare Function LogonUser Lib "advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long


Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long
Declare Function RevertToSelf Lib "advapi32.dll" () As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Public Declare Function LogonUser Lib "advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'pour connaitre la position du curseur
Public Declare Function GetCursorPos Lib "user32" (lpPoint As CURSOR) As Long
'pour definir une position
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const Flags = SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Sub Logoff()
Dim blnResult As Boolean
'MsgBox "Session fermée"
blnResult = RevertToSelf()
End Sub
Public Function SetTopMostWindow(Window As Form, Topmost As Boolean) As Long

    If Topmost = True Then
        SetTopMostWindow = SetWindowPos(Window.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        SetTopMostWindow = SetWindowPos(Window.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If

End Function
Function SecuFill(MPath As String, Lecture As Boolean)
If Lecture = True Then
 SetFileAttributes MPath, FILE_ATTRIBUTE_READONLY

Else
    SetFileAttributes MPath, FILE_ATTRIBUTE_NORMAL
End If
MySeconde 1
End Function
Public Sub SetAttributs(dossier As String, Lecture As Boolean)
    Dim I       As Integer
    Dim nb      As Integer
    Dim Table() As String
    Dim Fichier As String
    Dim LectureEcriture As Long
    If Lecture = True Then
    LectureEcriture = vbReadOnly
    Else
         LectureEcriture = vbNormal
    End If
    nb = 1
    ReDim Preserve Table(1)
    Table(1) = dossier
    
    While I < nb
        I = I + 1
        dossier = Table(I)
        If Right(dossier, 1) <> "\" Then dossier = dossier & "\"
        Fichier = Dir(dossier & "*.*", vbDirectory)
        Do Until Fichier = ""
           If Asc(Fichier) <> 46 Then
              If GetAttr(dossier & Fichier) = vbDirectory Then
                 nb = nb + 1
                 ReDim Preserve Table(nb)
                 Table(nb) = dossier & Fichier
                 Else
                 SetAttr dossier & Fichier, LectureEcriture
                 End If
              End If
           Fichier = Dir()
           Loop
        Wend
       
End Sub


Function LirMachineName() As String
    'Retourne le nom de l'utilisateur courant de l'ordinateur
   
    Dim stTmp As String, lgTmp As Long
    stTmp = Space$(250)
    lgTmp = 251
    Call GetComputerName(stTmp, lgTmp)
    LirMachineName = Mid$(stTmp, 1, InStr(1, stTmp, Chr$(0)) - 1)
End Function
Function LirUserName() As String
    'Retourne le nom de l'utilisateur courant de l'ordinateur
   
    Dim stTmp As String, lgTmp As Long
    stTmp = Space$(250)
    lgTmp = 251
    Call GetUserName(stTmp, lgTmp)
    LirUserName = Mid$(stTmp, 1, InStr(1, stTmp, Chr$(0)) - 1)
End Function
Sub MyExecute(Fichier As String, Optional Param As String = vbNull)

Dim lapi As Long
On Error Resume Next
lapi = ShellExecute(100, "open", Fichier, Param, vbNull, 5)
'If Err Then MsgBox Err.Description
Err.Clear
On Error GoTo 0
End Sub
