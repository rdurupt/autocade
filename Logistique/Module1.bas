Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function RevertToSelf Lib "advapi32.dll" () As Long
Public Declare Function LogonUser Lib "advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long

