VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifierUser 
   Caption         =   "Droits utilisateurs :"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   OleObjectBlob   =   "ModifierUser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifierUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim sql As String

For i = 1 To 3
    sql = "UPDATE T_Users SET "
    sql = sql & "T_Users.[PassWord] = '" & MyReplace(Me.Controls(Trim(Me.Controls("User" & CStr(i)).Caption) & "PassWord")) & "' "
    sql = sql & ", T_Users.Admin =  " & Replace(Replace(Me.Controls(Trim(Me.Controls("User" & CStr(i)).Caption) & "Admin"), "Vrai", "True"), "Faux", "False") & " "
    sql = sql & ", T_Users.Verificateur =  " & Replace(Replace(Me.Controls(Trim(Me.Controls("User" & CStr(i)).Caption) & "Verificateur"), "Vrai", "True"), "Faux", "False") & ", "
    sql = sql & "T_Users.Approbateur =  " & Replace(Replace(Me.Controls(Trim(Me.Controls("User" & CStr(i)).Caption) & "Approbateur"), "Vrai", "True"), "Faux", "False") & " "
    sql = sql & "WHERE T_Users.Id=" & Me.Controls(Trim(Me.Controls("User" & CStr(i)).Caption) & "Id") & ";"
    Con.Exequte sql
Next i


Unload Me
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub UserForm_Activate()
Dim sql As String
Dim Rs As Recordset

sql = "SELECT T_Users.Id, T_Users.User, T_Users.PassWord, "
sql = sql & "T_Users.Admin, T_Users.Verificateur, T_Users.Approbateur "
sql = sql & "FROM T_Users;"

Set Rs = Con.OpenRecordSet(sql)
While Rs.EOF = False
    Me.Controls(Rs!User & "Id") = Rs!Id
    For i = 2 To Rs.Fields.Count - 1
        Me.Controls(Rs!User & Rs.Fields(i).Name) = Rs.Fields(i).Value
    Next i
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)


End Sub

