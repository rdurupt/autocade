Attribute VB_Name = "Module1"
Option Explicit


Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long
Declare Function RevertToSelf Lib "advapi32.dll" () As Long
'Open the ImpersonateUser class module, and then paste the following code to create the Logon and Logoff methods:
Const LOGON32_LOGON_INTERACTIVE = 2
Const LOGON32_PROVIDER_DEFAULT = 0

Public Declare Function LogonUser Lib "advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long

Dim lngTokenHandle, lngLogonType, lngLogonProvider As Long
Dim blnResult As Boolean
Public AutocadApp As AutoCAD.AcadApplication
Sub Main()
On Error Resume Next
Dim Fso As New FileSystemObject
Dim a
blnResult = RevertToSelf()
a = LogonUser( _
"robert.durupt", _
"INGENICA", _
"dur12345/*-", _
          LOGON32_LOGON_INTERACTIVE, _
         LOGON32_PROVIDER_DEFAULT, _
            lngTokenHandle)


If blnResult = False Then MsgBox "Impossible d'ouvrir LogonUser()"
'MsgBox "Session avec le jeton" & lngTokenHandle & " et " & strAdminUser & ", " & strAdminDomain & ", " & strAdminPassword & " ouverte !"

            
blnResult = ImpersonateLoggedOnUser(lngTokenHandle)
Set AutocadApp = New AutoCAD.AcadApplication
   Logoff
   AutocadApp.Visible = True
End Sub
Public Sub Logoff()
Dim blnResult As Boolean
'MsgBox "Session fermée"
blnResult = RevertToSelf()
End Sub

