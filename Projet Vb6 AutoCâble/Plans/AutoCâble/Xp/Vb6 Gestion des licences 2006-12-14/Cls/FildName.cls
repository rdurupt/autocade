VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FildName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Txt(24, 4) As String
 Sub retournTableur(rr)

Dim I As Long
For I = 0 To 24
            Txt(I, 0) = rr(I, 0)
            Txt(I, 1) = rr(I, 1)
            Txt(I, 2) = rr(I, 2)
            Txt(I, 3) = rr(I, 3)
            Txt(I, 4) = rr(I, 4)
        Next

End Sub

Public Sub Remplie(Value, I)
Txt(I, 4) = Value
End Sub
Public Function RetournName(I)
RetournName = Txt(I, 0)
'RetournName = txt(i, 1)
'RetournName = txt(i, 2)
'RetournName = txt(i, 3)
'RetournName = txt(i, 4)
End Function
Public Function RetournValue(I)
If Val(Txt(I, 1)) = 3 Then
    RetournValue = Val(Txt(I, 4))
Else
    RetournValue = Txt(I, 4)
End If
'RetournValue = txt(i, 1)
'RetournName = txt(i, 2)
'RetournName = txt(i, 3)
'RetournName = txt(i, 4)
End Function

