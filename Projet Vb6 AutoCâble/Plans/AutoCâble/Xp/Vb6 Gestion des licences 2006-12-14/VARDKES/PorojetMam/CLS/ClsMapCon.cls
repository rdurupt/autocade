VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMapCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ConName As String
Public Colec_Liason As Collection


Private Sub Class_Initialize()
Set Colec_Liason = New Collection
End Sub
Public Sub AjouterLiason(Pin As Integer, Liai As String)
    Colec_Liason.Add Pin, "Lai_" & Liai
End Sub
