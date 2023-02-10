VERSION 5.00
Begin VB.Form frmMacco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Supp 
      Caption         =   "&Supprimer Macro"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Exec 
      Caption         =   "&Exécuter Macro"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox LstMacro 
      Height          =   3765
      ItemData        =   "frmMacco.frx":0000
      Left            =   120
      List            =   "frmMacco.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMacco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyUserName As String
Dim MyTypeSortie As String
Public NameMacro As String
Public SubMacro As String
Public Sub charger(UserName As String, TypeSortie As String)
 NameMacro = ""
 SubMacro = ""
MyUserName = UserName
MyTypeSortie = TypeSortie
If MyTypeSortie = "EXE" Then
    Exec.Visible = True
    Me.Caption = "Exécuter Macro :"
Else
    Supp.Visible = True
    Me.Caption = "Supprimer Macro :"
End If
LoadMacro
Me.Show vbModal
End Sub

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Exec_Click()
Dim Sql As String
Dim Rs As Recordset
Debug.Print LstMacro.ListIndex
If LstMacro.ListIndex < 0 Then
MsgBox "Vous devez sélectionner une Macro dans la liste.", vbExclamation
Exit Sub
End If
Debug.Print LstMacro.List(LstMacro.ListIndex)

Sql = "SELECT T_Macro.Sub FROM T_Macro "
Sql = Sql & "WHERE T_Macro.Macro='" & Replace(LstMacro.List(LstMacro.ListIndex), "'", "''") & "'  "
Sql = Sql & "AND T_Macro.Formulaire='" & MyUserName & "';"
Set Rs = Con.OpenRecordSet(Sql)
NameMacro = LstMacro.List(LstMacro.ListIndex)
SubMacro = "" & Rs!Sub
Set Rs = Con.CloseRecordSet(Rs)
Me.Hide
End Sub

Private Sub LstMacro_DblClick()
If MyTypeSortie = "EXE" Then
    Exec_Click
Else
    Supp_Click
End If
End Sub

Private Sub Supp_Click()
Dim Sql As String
If LstMacro.ListIndex < 0 Then
    MsgBox "Vous devez sélectionner une Macro dans la liste.", vbExclamation
    Exit Sub
End If
If MsgBox("Voulez-vous vraiment supprimer la macro : " & LstMacro.List(LstMacro.ListIndex), vbQuestion + vbYesNo, "Supprimer Macro") = vbNo Then Exit Sub
Sql = "DELETE T_Macro.* FROM T_Macro "
Sql = Sql & "WHERE T_Macro.Macro='" & Replace(LstMacro.List(LstMacro.ListIndex), "'", "''") & "'  "
Sql = Sql & "AND T_Macro.Formulaire='" & MyUserName & "';"
Con.Execute Sql
LoadMacro
End Sub
Sub LoadMacro()
Dim Sql As String
Dim Rs As Recordset
LstMacro.Clear
Sql = "SELECT T_Macro.Macro FROM T_Macro "
Sql = Sql & "WHERE T_Macro.Formulaire='" & MyUserName & "';"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    LstMacro.AddItem "" & Rs!Macro
    Rs.MoveNext
Wend
End Sub
