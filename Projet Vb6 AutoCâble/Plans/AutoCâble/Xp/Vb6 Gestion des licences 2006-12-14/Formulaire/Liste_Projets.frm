VERSION 5.00
Begin VB.Form Liste_Projets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liste Projet Et Base Véhicule:"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   Icon            =   "Liste_Projets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   9720
      Width           =   2295
   End
   Begin VB.TextBox ProjetM1 
      Height          =   615
      Left            =   6120
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Default         =   -1  'True
      Height          =   735
      Left            =   4440
      Picture         =   "Liste_Projets.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Annuler"
      Top             =   300
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   6000
      Picture         =   "Liste_Projets.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Supprimer"
      Top             =   300
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   5220
      Picture         =   "Liste_Projets.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Enregistrer"
      Top             =   300
      Width           =   735
   End
   Begin VB.TextBox Projet 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Liste Des Projets Et Des Bases Véhicules"
      Height          =   8415
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   6375
      Begin VB.ListBox LstProjet 
         Height          =   8055
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Double Click pour Editer"
         Top             =   240
         Width           =   5295
      End
      Begin VB.CommandButton Command3 
         Height          =   735
         Left            =   5520
         Picture         =   "Liste_Projets.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Editer"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Projet :"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Liste_Projets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Sql As String
Dim Rs As Recordset
Dim msg As String
Sql = "SELECT T_Liste_Projet.* FROM T_Liste_Projet "
Sql = Sql & "WHERE T_Liste_Projet.Projet='" & MyReplace(Me.Projet) & "' "
msg = "Ajout"
If Me.Tag <> "" Then
    Sql = Sql & "AND T_Liste_Projet.id<>" & Me.Tag & ";"
    msg = "Modification"
End If
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    MsgBox Me.Projet & " Existe déjà" & vbCrLf & vbCrLf & msg & " non effectuée.", vbExclamation
Else
    If Me.Tag <> "" Then
        Sql = "UPDATE T_Liste_Projet SET T_Liste_Projet.Projet = '" & MyReplace(Me.Projet) & "' "
        Sql = Sql & "WHERE T_Liste_Projet.id=" & Me.Tag & ";"
        Con.Execute Sql
        
        Sql = "UPDATE T_Projet SET T_Projet.Projet = '" & MyReplace(Me.Projet) & "' "
        Sql = Sql & "WHERE T_Projet.Projet='" & MyReplace(Me.ProjetM1) & "';"
        Con.Execute Sql
        
        Sql = "UPDATE Archive_T_Projet SET Archive_T_Projet.Projet = '" & MyReplace(Me.Projet) & "' "
        Sql = Sql & "WHERE Archive_T_Projet.Projet='" & MyReplace(Me.ProjetM1) & "';"
        Con.Execute Sql
    Else
        Sql = "INSERT INTO T_Liste_Projet ( Projet ) VALUES ( '" & MyReplace(Me.Projet) & "');"
        Con.Execute Sql
    End If
End If
Set Rs = Con.CloseRecordSet(Rs)
Me.Projet = ""
Me.ProjetM1 = ""
Me.Tag = ""
Acualise

End Sub

Private Sub Command2_Click()
Dim Sql As String
Dim Rs As Recordset
Dim msg As String
If Me.Tag <> "" Then
    Sql = "SELECT T_Projet.Projet FROM T_Projet "
    Sql = Sql & "WHERE T_Projet.Projet= '" & MyReplace(Me.LstProjet.List(Me.LstProjet.ListIndex)) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
    
        msg = "Le projet : [" & Me.LstProjet.List(Me.LstProjet.ListIndex) & "] pointe sur une ou plusieurs affaires." & vbCrLf & vbCrLf
        msg = msg & "Veuillez modifier toutes les affaires affiliées à ce projet " & vbCrLf
        msg = msg & "avant une prochaine tentative." & vbCrLf & vbCrLf
        msg = msg & "Suppression non effectuée."
        MsgBox msg, vbCritical
        Me.Projet = ""
        Me.ProjetM1 = ""
        Me.Tag = ""
      Set Rs = Con.CloseRecordSet(Rs)
      Exit Sub
    End If
'    sql = "SELECT Archive_T_Projet.Projet FROM Archive_T_Projet "
'    sql = sql & "WHERE Archive_T_Projet.Projet= '" & MyReplace(Me.LstProjet.List(Me.LstProjet.ListIndex)) & "';"
'
'     Set Rs = Con.OpenRecordSet(sql)
'    If Rs.EOF = False Then
'        msg = "Le projet : [" & Me.LstProjet.List(Me.LstProjet.ListIndex) & "] pointe sur une ou plusieurs affaires." & vbCrLf & vbCrLf
'        msg = msg & "Veuillez modifier toutes les affaires affiliées à ce projet " & vbCrLf
'        msg = msg & "avant une prochaine tentative." & vbCrLf & vbCrLf
'        msg = msg & "Suppression non effectuée."
'        MsgBox msg, vbCritical
'
'        Me.Projet = ""
'        Me.ProjetM1 = ""
'        Me.Tag = ""
'      Set Rs = Con.CloseRecordSet(Rs)
'      Exit Sub
'    End If
    Sql = "DELETE T_Liste_Projet.* FROM T_Liste_Projet "
    Sql = Sql & "WHERE T_Liste_Projet.id=" & Me.Tag & ";"
    Con.Execute Sql
     Set Rs = Con.CloseRecordSet(Rs)
End If
Me.Projet = ""
Me.ProjetM1 = ""
Me.Tag = ""
Acualise
End Sub

Private Sub Command3_Click()
Dim Sql As String
Dim Rs As Recordset

If Me.LstProjet.ListIndex = -1 Then
MsgBox "Vous devez sélectionner un projet dans la liste."
Exit Sub
End If
Sql = "SELECT T_Liste_Projet.id FROM T_Liste_Projet "
Sql = Sql & "Where T_Liste_Projet.Projet = '" & MyReplace(Me.LstProjet.List(Me.LstProjet.ListIndex)) & "'"
Set Rs = Con.OpenRecordSet(Sql)
Me.Tag = Rs!Id
Me.Projet = Me.LstProjet.List(Me.LstProjet.ListIndex)
Me.ProjetM1 = Me.Projet
End Sub

Private Sub Command4_Click()
Me.Projet = ""
Me.ProjetM1 = ""
Me.Tag = ""
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Acualise
End Sub
Sub Acualise()
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT T_Liste_Projet.Projet FROM T_Liste_Projet ORDER BY T_Liste_Projet.Projet;"
Set Rs = Con.OpenRecordSet(Sql)
Me.LstProjet.Clear
While Rs.EOF = False
    Me.LstProjet.AddItem "" & Rs!Projet
    Rs.MoveNext
Wend

End Sub

Private Sub LstProjet_DblClick()
Command3_Click
End Sub
