VERSION 5.00
Begin VB.Form Param_E_Boutique 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bases de données catalogues:"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   7320
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Valider"
      Height          =   495
      Left            =   480
      TabIndex        =   24
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Height          =   285
      Left            =   8640
      Picture         =   "Param_E_Boutique.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2280
      Width           =   285
   End
   Begin VB.CommandButton Command7 
      Height          =   285
      Left            =   8640
      Picture         =   "Param_E_Boutique.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1986
      Width           =   285
   End
   Begin VB.CommandButton Command6 
      Height          =   285
      Left            =   8640
      Picture         =   "Param_E_Boutique.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1680
      Width           =   285
   End
   Begin VB.CommandButton Command5 
      Height          =   285
      Left            =   8640
      Picture         =   "Param_E_Boutique.frx":09C6
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1404
      Width           =   285
   End
   Begin VB.CommandButton Command4 
      Height          =   285
      Left            =   8640
      Picture         =   "Param_E_Boutique.frx":0D08
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1080
      Width           =   285
   End
   Begin VB.CommandButton Command3 
      Height          =   285
      Left            =   8640
      Picture         =   "Param_E_Boutique.frx":104A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   822
      Width           =   285
   End
   Begin VB.CommandButton Command2 
      Height          =   285
      Left            =   8640
      Picture         =   "Param_E_Boutique.frx":138C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   531
      Width           =   285
   End
   Begin VB.CommandButton Command1 
      Height          =   285
      Left            =   8640
      Picture         =   "Param_E_Boutique.frx":16CE
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   240
      Width           =   285
   End
   Begin VB.TextBox Eb_SUPPORTS 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   15
      Tag             =   "SUPPORTS ET FIXATIONS;SUPPORTS ET FIXATIONS;QRY;TXT;Eb_SUPPORTS"
      Top             =   2280
      Width           =   6375
   End
   Begin VB.TextBox Eb_JOINTS 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Tag             =   "JOINTS ET BOUCHONS;JOINTS ET BOUCHONS;QRY;TXT;Eb_JOINTS"
      Top             =   1986
      Width           =   6375
   End
   Begin VB.TextBox Eb_HABILLAGES 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Tag             =   "HABILLAGES;HABILLAGES;QRY;TXT;Eb_HABILLAGES"
      Top             =   1680
      Width           =   6375
   End
   Begin VB.TextBox Eb_FILS 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Tag             =   "FILS ETCOMPOSANTS;FILS ETCOMPOSANTS;QRY;TXT;Eb_FILS"
      Top             =   1404
      Width           =   6375
   End
   Begin VB.TextBox Eb_CONNECTEURS 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Tag             =   "CONNECTEURS;CONNECTEURS;QRY;TXT;Eb_CONNECTEURS"
      Top             =   1080
      Width           =   6375
   End
   Begin VB.TextBox Eb_CONNECTIQUE 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Tag             =   "ONNECTIQUE;Eb_CONNECTIQUE;QRY;TXT;Eb_CONNECTIQUE"
      Top             =   822
      Width           =   6375
   End
   Begin VB.TextBox Eb_CAPOTS 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Tag             =   "CAPOTS ET VERROUX;CAPOTS ET VERROUX;QRY;TXT;Eb_CAPOTS"
      Top             =   531
      Width           =   6375
   End
   Begin VB.TextBox Eb_BAGUES 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Tag             =   "BAGUES;BAGUES;QRY;TXT;Eb_BAGUES"
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BAGUES"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HABILLAGES"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FILS ETCOMPOSANTS"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1404
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUPPORTS ET FIXATIONS"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JOINTS ET BOUCHONS"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1986
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAPOTS ET VERROUX"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   531
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONNECTIQUE"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   822
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONNECTEURS"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1113
      Width           =   2175
   End
End
Attribute VB_Name = "Param_E_Boutique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Serveur As String

Private Sub Command1_Click()
If Serveur = "" Then
    MsgBox ""
Else
    Eb_BAGUES = Replace(ScanFichier.chargement("MDB", Eb_BAGUES), Serveur, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub Command10_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
If Serveur = "" Then
    MsgBox ""
Else
    Eb_CAPOTS = Replace(ScanFichier.chargement("MDB", Eb_CAPOTS), Serveur, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub Command3_Click()
If Serveur = "" Then
    MsgBox ""
Else
    Eb_CONNECTIQUE = Replace(ScanFichier.chargement("MDB", Eb_CONNECTIQUE), Serveur, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub Command4_Click()
If Serveur = "" Then
    MsgBox ""
Else
    Eb_CONNECTEURS = Replace(ScanFichier.chargement("MDB", Eb_CONNECTEURS), Serveur, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub Command5_Click()
If Serveur = "" Then
    MsgBox ""
Else
    Eb_FILS = Replace(ScanFichier.chargement("MDB", Eb_FILS), Serveur, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub Command6_Click()
If Serveur = "" Then
    MsgBox ""
Else
    Eb_HABILLAGES = Replace(ScanFichier.chargement("MDB", Eb_HABILLAGES), Serveur, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub Command7_Click()
If Serveur = "" Then
    MsgBox ""
Else
    Eb_JOINTS = Replace(ScanFichier.chargement("MDB", Eb_JOINTS), Serveur, "", 1)
    Unload ScanRep

End If
End Sub

Private Sub Command8_Click()
If Serveur = "" Then
    MsgBox ""
Else
    Eb_SUPPORTS = Replace(ScanFichier.chargement("MDB", Eb_SUPPORTS), Serveur, "", 1)
    Unload ScanRep

End If

End Sub

Private Sub Command9_Click()
Dim Sql As String
If MyFormatQRY(Me.Eb_BAGUES) = False Then Exit Sub
If MyFormatQRY(Me.Eb_CAPOTS) = False Then Exit Sub
If MyFormatQRY(Me.Eb_CONNECTIQUE) = False Then Exit Sub
If MyFormatQRY(Me.Eb_CONNECTEURS) = False Then Exit Sub
If MyFormatQRY(Me.Eb_FILS) = False Then Exit Sub
If MyFormatQRY(Me.Eb_HABILLAGES) = False Then Exit Sub
If MyFormatQRY(Me.Eb_JOINTS) = False Then Exit Sub
If MyFormatQRY(Me.Eb_SUPPORTS) = False Then Exit Sub

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Eb_BAGUES) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eb_BAGUES';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Eb_CAPOTS) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eb_CAPOTS';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Eb_CONNECTIQUE) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eb_CONNECTIQUE';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Eb_CONNECTEURS) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eb_CONNECTEURS';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Eb_FILS) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eb_FILS';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Eb_HABILLAGES) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eb_HABILLAGES';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Eb_JOINTS) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eb_JOINTS';"
Con.Execute Sql

Sql = "UPDATE T_Path SET T_Path.PathVar = '" & MyReplace(Eb_SUPPORTS) & "' "
Sql = Sql & "WHERE T_Path.NameVar='Eb_SUPPORTS';"
Con.Execute Sql


Me.Hide
End Sub

Public Sub chargement(MyServeur As String)
Dim Sql As String
Dim Rs As Recordset
Serveur = MyServeur

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eb_BAGUES';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eb_BAGUES = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eb_CAPOTS';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eb_CAPOTS = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eb_CONNECTEURS';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eb_CONNECTEURS = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eb_CONNECTIQUE';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eb_CONNECTIQUE = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eb_FILS';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eb_FILS = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eb_HABILLAGES';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eb_HABILLAGES = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eb_JOINTS';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eb_JOINTS = "" & Rs!PathVar

Sql = "select T_Path.PathVar from T_Path "
Sql = Sql & "WHERE T_Path.NameVar='Eb_SUPPORTS';"
Set Rs = Con.OpenRecordSet(Sql)
Me.Eb_SUPPORTS = "" & Rs!PathVar

Me.Show vbModal
End Sub
