VERSION 5.00
Begin VB.Form frm_Etiquettes_Serie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Création D'étiquettes Câblage :"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   8280
      TabIndex        =   23
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Valider"
      Height          =   495
      Left            =   4740
      TabIndex        =   22
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Rechercher une Pièce"
      Height          =   495
      Left            =   1200
      TabIndex        =   21
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   8640
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "frm_Etiquettes_Serie.frx":0000
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.Label TXT8 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6240
      TabIndex        =   27
      Tag             =   "N° P ;N° PL;QRY;TXT;TXT8"
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label TXT7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6240
      TabIndex        =   26
      Tag             =   "N° P ;N° PL;QRY;TXT;TXT7"
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label TXT6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6240
      TabIndex        =   25
      Tag             =   "N° P ;N° PL;QRY;TXT;TXT6"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label TXT5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6240
      TabIndex        =   24
      Tag             =   "N° P ;N° PL;QRY;TXT;TXT5"
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Projet"
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Vague"
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   1725
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Equipement"
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   2115
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "Ensemble"
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Pièce"
      Height          =   315
      Left            =   5160
      TabIndex        =   16
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "Plan"
      Height          =   315
      Left            =   5160
      TabIndex        =   15
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "Outil"
      Height          =   315
      Left            =   5160
      TabIndex        =   14
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label Label8 
      Caption         =   "Liste"
      Height          =   315
      Left            =   5160
      TabIndex        =   13
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Label Label9 
      Caption         =   "Client"
      Height          =   315
      Left            =   5160
      TabIndex        =   12
      Top             =   2760
      Width           =   1005
   End
   Begin VB.Label Label10 
      Caption         =   "Dessinateur"
      Height          =   315
      Left            =   5160
      TabIndex        =   11
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label Label11 
      Caption         =   " Vérificateur "
      Height          =   315
      Left            =   5160
      TabIndex        =   10
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label Label12 
      Caption         =   " Approbateur:"
      Height          =   315
      Left            =   5160
      TabIndex        =   9
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Label txt1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label txt2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   1725
      Width           =   3135
   End
   Begin VB.Label txt3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   2115
      Width           =   3135
   End
   Begin VB.Label txt4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   1395
      Left            =   1320
      TabIndex        =   5
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label txt9 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6240
      TabIndex        =   4
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label txt10 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label txt11 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label txt12 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6240
      TabIndex        =   1
      Top             =   3840
      Width           =   3135
   End
End
Attribute VB_Name = "frm_Etiquettes_Serie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim sql As String
Dim Rs As Recordset
Dim CherchPicesAnnuler As Boolean
CherchPices.Charge Me, "(VerifieDate= Null   and Archiver=false) OR (IdStatus<4  and Archiver=false)"
CherchPicesAnnuler = CherchPices.Annuler
Unload CherchPices
If Me.txt3.Tag = "" Then CherchPicesAnnuler = True

If CherchPicesAnnuler = True Then Exit Sub

IdFils = 0
sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(sql)
IdFils = 0
Me.Tag = Me.txt3.Tag
If Rs!Pere > 0 Then
IdFils = Me.txt3.Tag
    Me.txt3.Tag = Rs!Pere
    Me.Tag = Me.txt3.Tag
End If
Set Rs = Con.CloseRecordSet(Rs)
'Maj Me.txt3.Tag
End Sub

Private Sub Command2_Click()
Dim Exec As Boolean
Dim I, I2 As Long
If Trim("" & Me.Tag) = "" Then
'    CommandButton1_Click
    Exit Sub
 End If
'If MyFormatQRY(txt13) = False Then Exit Sub
'If MyFormatQRY(txt14) = False Then Exit Sub
'If MyFormatQRY(txt15) = False Then Exit Sub
'If MyFormatQRY(txt16) = False Then Exit Sub
For I = 13 To 16
    For I2 = 13 To 16
        If I <> I2 Then
            If Me.Controls("txt" & CStr(I)) = Me.Controls("txt" & CStr(I2)) Then
                MsgBox "Vous devez saisir des valeurs différentes dans les listes déroulante" & vbCrLf & Me.Controls("txt" & CStr(I)) & " = " & Me.Controls("txt" & CStr(I2)), vbExclamation
                Exit Sub
            End If
        End If
    Next
Next
Dim Fso As New FileSystemObject
If Fso.FileExists(Environ("USERPROFILE") & "\Mes Documents\" & Replace(TXT5.Caption, ":", "_", 1) & ".XLS") = True Then
    Fso.DeleteFile Environ("USERPROFILE") & "\Mes Documents\" & Replace(TXT5.Caption, ":", "_", 1) & ".XLS"
End If

Set FormBarGrah = Me
If ExporteXls(Environ("USERPROFILE") & "\Mes Documents\" & Replace(TXT5.Caption, ":", "_", 1) & ".XLS", CLng(Me.Tag)) = True Then

EnteteClasseurControle = "Contrôle"
bool_MiseEnPage = True
'DossierDeFab Environ("USERPROFILE") & "\Mes Documents\" & Replace(TXT5.Caption, ":", "_", 1) & ".XLS", CLng(Me.Tag), _
'Me.txt1, Me.txt2, Me.txt3, Me.txt4, Me.TXT5, Me.TXT6, Me.TXT7, Me.TXT8, Me.txt9, Me.PieceCLI, Me.txt13, Me.txt14, Me.txt16, True, Me.Affaire, Val(Me.Tag)

EnteteClasseurControle = "Fabrication"

'DossierDeFab Environ("USERPROFILE") & "\Mes Documents\" & Replace(TXT5.Caption, ":", "_", 1) & ".XLS", CLng(Me.Tag), _
'Me.txt1, Me.txt2, Me.txt3, Me.txt4, Me.TXT5, Me.TXT6, Me.TXT7, Me.TXT8, Me.txt9, Me.PieceCLI, Me.txt13, Me.txt15, Me.txt16, False, Me.Affaire, Val(Me.Tag)
bool_MiseEnPage = False
End If
If Fso.FileExists(Environ("USERPROFILE") & "\Mes Documents\" & Replace(TXT5.Caption, ":", "_", 1) & ".XLS") = True Then
    Fso.DeleteFile Environ("USERPROFILE") & "\Mes Documents\" & Replace(TXT5.Caption, ":", "_", 1) & ".XLS"
End If
'Noquite = False
Unload Me
End Sub

Private Sub Command3_Click()
'Noquite = False
Unload Me

End Sub
