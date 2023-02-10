VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReyRecordsetMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//         By Renfield
'//     thomas.reynald@caraamail.com
'//     vous pouvez enlever ces commentaires , mais ne dites pas que cette source est de vous........
Public Rs As ADODB.Recordset
Dim Descr As String

Dim MustAddDesc As Boolean
Dim MustAddRecord As Boolean

Dim m_Record As Dictionary
Public Fields As Dictionary

Dim FieldTypeText



'Array        8192
'Big Integer        20
'Binary 128
'Boolean        11
'BSTR 8
'Chapter 135
'Char 129
'Currency        6
'Date 7
'dbDate 133
'dbTime 134
'dbTimeStamp 135
'Decimal        14
'Double        5
'Empty        0
'Error 10
'FileTime 64
'Guid 72
'Dispatch 9
'Integer        3
'Unknown 13
'LongVarBinary 205
'LongVarChar 201
'LongVarWChar 203
'Numeric 131
'PropVariant 138
'Single        4
'SmallInt 2
'TinyInt 16
'UnsignedBigInt 21
'UnsignedInt 19
'UnsignedSmallInt 18
'UnsignedTinyint 17
'Userdefined 132
'VarBinary 204
'VarChar 200
'Variant        12
'VarNumeric 139
'VarWChar 202
'WChar 130
Public Enum FieldTypeEnum
    FT_Char = 0
    FT_Currency = 1
    FT_Date = 2
    FT_Decimal = 3
    FT_Double = 4
    FT_Integer = 5
    FT_Numeric = 6
    FT_Single = 7
    FT_SmallInt = 8
    FT_VarChar = 9
    FT_VarBinary = 10
     FT_Big_Integer = 20
End Enum
Public Function CreatFilds(Tchamps) As Recordset
Dim i As Long
For i = 0 To UBound(Tchamps)
    AddField "" & Tchamps(i, 0), Val(Tchamps(i, 1)), Val(Tchamps(i, 2)), Val(Tchamps(i, 3))
'    AddField "Naiss", FT_VarChar, 10
Next

End Function

Private Sub Class_Initialize()
'    FieldTypeText = Array("adChar", "adCurrency", "adDate", "adDecimal", "adDouble", "adInteger", "adNumeric", "adSingle", "adSmallInt", "adVarChar", "adVarBinary")
'    Set Rs = New ADODB.Recordset
'    Rs.ActiveConnection = "provider=msdatashape;data provider=none;"
'    Descr = " SHAPE APPEND "
'    MustAddDesc = True
'    MustAddRecord = False
'
'    Set m_Record = Nothing
'    Set m_Record = New Dictionary
'
'    Set Fields = Nothing
'    Set Fields = New Dictionary
   
End Sub

Public Sub AddField(f_Name As String, f_Type As Integer, f_Size As Integer, Optional AutoIncrement As Boolean)
On Error GoTo 0
    If Descr <> " SHAPE APPEND " Then Descr = Descr & ","
    If f_Size = 0 Then
        Descr = Descr & "NEW " & FieldTypeText(f_Type) & " As " & f_Name
    Else
        Descr = Descr & "NEW " & FieldTypeText(f_Type) & "(" & f_Size & ") As " & f_Name
    End If
    Fields.Add Fields.Count, f_Name
'    Fields.Add "AutoIncrement", f_Name
'     Fields(f_Name).Properties("AutoIncrement") = AutoIncrement
 '= AutoIncrement
   
End Sub

Public Sub AddRecord()
Dim Names
    If MustAddDesc Then
        Rs.Open Descr, , adOpenStatic, adLockOptimistic
         
        MustAddDesc = False
    End If
    
        Names = m_Record.Keys
        Values = m_Record.Items
        Rs.AddNew
        For i = 0 To UBound(Values)
            Rs(Replace(Replace(Names(i), "[", ""), "]", "")) = Values(i)
        Next
'        Names, Values
        MustAddRecord = False
    
    
    m_Record.RemoveAll
    MustAddRecord = True
End Sub

Public Sub Fill(f_Name As String, f_Value, Optional Num As Boolean)
'    On Error Resume Next                 ' si valeur deja existente .......... ecraser ancienne valeur
    If Trim(f_Value) = "" Then Exit Sub
    If Num = True Then
         m_Record.Add f_Name, Val(Trim("" & Replace(f_Value, ",", ".")))
        m_Record(f_Name) = Val(Trim("" & Replace(f_Value, ",", ".")))
    Else
            m_Record.Add f_Name, f_Value
            m_Record(f_Name) = f_Value
'         m_Record(f_Name).
    End If
End Sub

Public Function Recordset() As Recordset
'Dim Names
'    If MustAddRecord Then
'        Names = m_Record.Keys
'        Values = m_Record.Items
'        Rs.AddNew Names, Values
'    End If
    Set Recordset = Rs
    
    Class_Initialize
End Function

