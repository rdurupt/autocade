VERSION 5.00
Begin VB.Form DroitsGroupe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion des Droits aux Groupes :"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17865
   Icon            =   "DroitsGroupe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   17865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   17625
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   4575
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   17295
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2.45745e5
            Left            =   0
            ScaleHeight     =   2.45745e5
            ScaleWidth      =   2.45745e5
            TabIndex        =   11
            Top             =   0
            Width           =   2.45745e5
            Begin VB.Frame Frame5 
               BorderStyle     =   0  'None
               Height          =   5175
               Left            =   2040
               TabIndex        =   12
               Top             =   0
               Width           =   15135
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   615
                  Left            =   0
                  ScaleHeight     =   615
                  ScaleWidth      =   2055
                  TabIndex        =   13
                  Top             =   0
                  Width           =   2055
                  Begin VB.CheckBox Droit 
                     DownPicture     =   "DroitsGroupe.frx":08CA
                     Height          =   615
                     Index           =   0
                     Left            =   0
                     MaskColor       =   &H00E0E0E0&
                     Picture         =   "DroitsGroupe.frx":0D0C
                     Style           =   1  'Graphical
                     TabIndex        =   14
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   2055
                  End
               End
            End
            Begin VB.Label User 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   615
               Index           =   0
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Visible         =   0   'False
               Width           =   2055
            End
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   0
         Width           =   2055
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   370
            Left            =   0
            ScaleHeight     =   345
            ScaleWidth      =   2025
            TabIndex        =   9
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   0
         Width           =   15255
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   2055
            TabIndex        =   6
            Top             =   0
            Width           =   2055
            Begin VB.Label Bouton 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   7
               Top             =   0
               Visible         =   0   'False
               Width           =   2055
            End
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5535
      LargeChange     =   61
      Left            =   17640
      Max             =   3270
      SmallChange     =   6
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   270
      LargeChange     =   205
      Left            =   240
      Max             =   327
      SmallChange     =   20
      TabIndex        =   2
      Top             =   5640
      Width           =   17655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   615
      Left            =   14400
      TabIndex        =   1
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enregistrer"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   6120
      Width           =   1695
   End
End
Attribute VB_Name = "DroitsGroupe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L As Long
Dim C As Long
Dim d As Long
Dim MyGroupe As Long

Private Sub Command1_Click()
Dim Sql As String
Dim Id As Long
Dim IdVal
Sql = "DELETE T_Groupe_Users.* "
Sql = Sql & "FROM T_Groupe INNER JOIN T_Groupe_Users ON T_Groupe.id = T_Groupe_Users.Id_Groupe "
Sql = Sql & "WHERE T_Groupe.Niveaux> " & MyGroupe & ";"

Con.Execute Sql
For Id = 1 To d - 1

If Me.Droit(Id).Value = 1 Then
IdVal = Split(Me.Droit(Id).Tag, ";")
    Sql = "INSERT INTO T_Groupe_Users ( Id_Users, Id_Groupe ) "
    Sql = Sql & "values (" & IdVal(0) & "," & IdVal(1) & ");"
    Con.Execute Sql
End If
Next
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Public Sub chargement(Groupe As Long)
MyGroupe = Groupe
Dim Sql As String
Dim Rs As Recordset
Dim IC As Long
Dim IL As Long
C = 0
L = 0
d = 1

Sql = "SELECT T_Groupe.* FROM T_Groupe WHERE  T_Groupe.Niveaux > " & Groupe & " ORDER BY T_Groupe.Niveaux;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    If C <> 0 Then
        Load Bouton(C)
        Bouton(C).Left = Bouton(C - 1).Left + 2055
    End If
    
    Bouton(C).Caption = Trim("" & Rs!Groupe)
    Bouton(C).Tag = Trim("" & Rs!Id)
      Bouton(C).Visible = True
    C = C + 1
    Me.Picture2.Width = 2055 + (2055 * C)
    Me.Picture4.Width = 2055 + (2055 * C)
    Rs.MoveNext
Wend
'Me.Picture2.Width = 2055 + (2055 * C)

Sql = "SELECT T_Users.* "
Sql = Sql & "FROM T_Users LEFT JOIN (T_Groupe RIGHT JOIN T_Groupe_Users  "
Sql = Sql & "ON T_Groupe.id = T_Groupe_Users.Id_Groupe) ON T_Users.Id = T_Groupe_Users.Id_Users "
Sql = Sql & "Where T_Groupe.Niveaux > " & Groupe & " "
Sql = Sql & "Or T_Groupe.Niveaux Is Null "
Sql = Sql & "ORDER BY T_Users.User;"
Set Rs = Con.OpenRecordSet(Sql)

While Rs.EOF = False
If L <> 0 Then
Load User(L)
    User(L).Top = User(L - 1).Top + User(L - 1).Height
End If
Me.Picture2.Height = Me.Picture2.Height + (User(0).Height)
User(L).Caption = Trim("" & Rs!User)
User(L).Tag = Trim("" & Rs!Id)
User(L).Visible = True
L = L + 1
    Rs.MoveNext
Wend
Sql = "SELECT T_Groupe_Users.* FROM T_Groupe_Users;"
Set Rs = Con.OpenRecordSet(Sql)
For IL = 0 To L - 1
    For IC = 0 To C - 1
'    Rs.Requery
        Rs.Filter = "Id_Users=" & User(IL).Tag & " And Id_Groupe=" & Bouton(IC).Tag
        Load Droit(d)
        Droit(d).Visible = True
        Droit(d).Top = User(IL).Top
        Droit(d).Left = Bouton(IC).Left
        Droit(d).Tag = User(IL).Tag & ";" & Bouton(IC).Tag
        If Rs.EOF = False Then
        Droit(d).Value = 1
       
        End If
        d = d + 1
    Next
Next
Me.Show vbModal
End Sub

Private Sub HScroll1_Change()
 If Me.HScroll1.Value = 0 Then
    Picture2.Left = 0
     Picture4.Left = 0
Else
Picture2.Left = Me.HScroll1.Value * -100
Picture4.Left = Me.HScroll1.Value * -100
End If
End Sub

Private Sub Image1_Click()

End Sub

Private Sub VScroll1_Change()
If VScroll1.Value = 0 Then
    Me.Picture3.Top = 0
Else
     Me.Picture3.Top = VScroll1.Value * -10
End If
End Sub
