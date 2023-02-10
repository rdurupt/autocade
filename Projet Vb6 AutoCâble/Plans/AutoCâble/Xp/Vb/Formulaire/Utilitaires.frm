VERSION 5.00
Begin VB.Form Utilitaires 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Utilitaires :"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   Icon            =   "Utilitaires.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0.264
   ScaleMode       =   0  'User
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exécuter"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      Columns         =   1
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000FFF0A&
      Height          =   2310
      ItemData        =   "Utilitaires.frx":1272
      Left            =   120
      List            =   "Utilitaires.frx":1274
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Utilitaires"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Fso As FileSystemObject
Dim Sql As String
Dim Rs As Recordset
If Me.List1.ListIndex = -1 Then
    MsgBox "devez sélectionner un utilitaire.", vbExclamation
    Exit Sub
End If
Sql = "SELECT Utilitaire.Utilitaire FROM Utilitaire "
Sql = Sql & "WHERE Utilitaire.NameBouton='" & Me.List1.List(Me.List1.ListIndex) & "' "
Sql = Sql & "ORDER BY Utilitaire.NameBouton;"

Set Rs = Con.OpenRecordSet(Sql)
Set Fso = New FileSystemObject
If Rs.EOF = False Then
'   If Fso.FileExists("" & Rs!Utilitaire) = True Then
'   MsgBox ""
'   End If
     'Execute explorer.exe
    MyExecute "" & Rs!Utilitaire
     'Execute la calculette
''    WinExec "Calc.exe", 1
     'autre astuce :Le vrai mode plein ecran est là !!
'    Shell "c:\program files\internet explorer\iexplore -k c:"
'    Shell "" & Rs!Utilitaire, vbMaximizedFocus
End If
Set Rs = Con.CloseRecordSet(Rs)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT Utilitaire.NameBouton, Utilitaire.Utilitaire "
Sql = Sql & "FROM Utilitaire ORDER BY Utilitaire.NameBouton;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
Me.List1.AddItem "" & Rs!NameBouton
   Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
End Sub
   

Private Sub List1_DblClick()
Command1_Click
End Sub
