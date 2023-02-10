VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifierUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Droits utilisateurs :"
   ClientHeight    =   3105
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   7680
   Icon            =   "ModifierUser.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "ModifierUser.dsx":08CA
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ModifierUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim Sql As String

For I = 1 To 3
    Sql = "UPDATE T_Users SET "
    Sql = Sql & "T_Users.[PassWord] = '" & MyReplace(Me.Controls(Trim(Me.Controls("User" & CStr(I)).Caption) & "PassWord")) & "' "
    Sql = Sql & ", T_Users.Admin =  " & Replace(Replace(Me.Controls(Trim(Me.Controls("User" & CStr(I)).Caption) & "Admin"), "Vrai", "True"), "Faux", "False") & " "
    Sql = Sql & ", T_Users.Verificateur =  " & Replace(Replace(Me.Controls(Trim(Me.Controls("User" & CStr(I)).Caption) & "Verificateur"), "Vrai", "True"), "Faux", "False") & ", "
    Sql = Sql & "T_Users.Approbateur =  " & Replace(Replace(Me.Controls(Trim(Me.Controls("User" & CStr(I)).Caption) & "Approbateur"), "Vrai", "True"), "Faux", "False") & " "
    Sql = Sql & "WHERE T_Users.Id=" & Me.Controls(Trim(Me.Controls("User" & CStr(I)).Caption) & "Id") & ";"
    Con.Execute Sql
Next I


Unload Me
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub UserForm_Activate()
Dim Sql As String
Dim Rs As Recordset

Sql = "SELECT T_Users.Id, T_Users.User, T_Users.PassWord, "
Sql = Sql & "T_Users.Admin, T_Users.Verificateur, T_Users.Approbateur "
Sql = Sql & "FROM T_Users;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Me.Controls(Rs!User & "Id") = Rs!Id
    For I = 2 To Rs.Fields.Count - 1
        Me.Controls(Rs!User & Rs.Fields(I).Name) = Rs.Fields(I).Value
    Next I
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)


End Sub

