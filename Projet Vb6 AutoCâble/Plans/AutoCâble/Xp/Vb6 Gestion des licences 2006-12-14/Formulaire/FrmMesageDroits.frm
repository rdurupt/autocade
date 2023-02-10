VERSION 5.00
Begin VB.Form FrmMesageDroits 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion des destinataires d'Emails."
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "FrmMesageDroits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6705
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5295
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   6375
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
               Width           =   4695
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
                     DownPicture     =   "FrmMesageDroits.frx":030A
                     Height          =   615
                     Index           =   0
                     Left            =   0
                     MaskColor       =   &H00E0E0E0&
                     Picture         =   "FrmMesageDroits.frx":0BD4
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
         Height          =   735
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
      Height          =   5775
      LargeChange     =   61
      Left            =   6720
      Max             =   3270
      SmallChange     =   6
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   270
      LargeChange     =   205
      Left            =   0
      Max             =   327
      SmallChange     =   20
      TabIndex        =   2
      Top             =   5760
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enregistrer"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   6360
      Width           =   1695
   End
End
Attribute VB_Name = "FrmMesageDroits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L As Long
Dim C As Long
Dim d As Long
Dim IninOk As Boolean

Private Sub Bouton_DblClick(Index As Integer)
frmContenuMessage.chargement Val(Bouton(Index).Tag)
Unload frmContenuMessage
End Sub

Private Sub Command1_Click()
Dim Sql As String
Dim Id As Long
Dim IdVal

Sql = "DELETE T_Destinataire.* FROM T_Destinataire;"
Con.Execute Sql
For Id = 1 To d - 1

If Me.Droit(Id).Value = 1 Then
IdVal = Split(Me.Droit(Id).Tag, ";")
    Sql = "INSERT INTO T_Destinataire ( Id_Useur, Id_Message ) "
    Sql = Sql & "values (" & IdVal(0) & "," & IdVal(1) & ");"
    Con.Execute Sql
End If
Next
Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub


Private Sub Form_Activate()
chargement
End Sub

Private Sub Form_Load()
IninOk = False
C = 0
L = 0
d = 1
End Sub
Sub chargement()
Dim Sql As String
Dim Rs As Recordset
Dim IC As Long
Dim IL As Long
If IninOk = True Then Exit Sub
IninOk = True
Sql = "SELECT T_Message_Mail.Id, T_Message_Mail.Routine FROM T_Message_Mail ORDER BY  T_Message_Mail.Routine ;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    If C <> 0 Then
        Load Bouton(C)
        Bouton(C).Left = Bouton(C - 1).Left + 2055
    End If
    
    Bouton(C).Caption = Trim("" & Rs!Routine)
    Bouton(C).Tag = Trim("" & Rs!Id)
      Bouton(C).Visible = True
       Me.Picture2.Width = 2055 + (2055 * C)
    Me.Picture4.Width = 2055 + (2055 * C)
    C = C + 1
    Rs.MoveNext
Wend
Me.Picture2.Width = 2055 + (2055 * C)

Sql = "SELECT T_Users.* FROM T_Users ORDER BY T_Users.User;"
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
Sql = "SELECT T_Destinataire.* FROM T_Destinataire;"
Set Rs = Con.OpenRecordSet(Sql)
For IL = 0 To L - 1
    For IC = 0 To C - 1
'    Rs.Requery
        Rs.Filter = "Id_Useur=" & User(IL).Tag & " And Id_Message=" & Bouton(IC).Tag
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

