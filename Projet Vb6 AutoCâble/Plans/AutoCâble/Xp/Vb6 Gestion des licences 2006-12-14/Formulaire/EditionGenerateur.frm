VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EditeGenerateur 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Générateur d'Etat Mode Edition:"
   ClientHeight    =   6345
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Rechercher une Pièce"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5040
      Width           =   1995
   End
   Begin VB.CommandButton CommandButton3 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   7800
      TabIndex        =   23
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Exécuter 
      Caption         =   "Pas de modèle sélectionné"
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   5040
      Width           =   5055
   End
   Begin VB.ComboBox lstLi 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   3840
      TabIndex        =   21
      Tag             =   "Liste ;Liste;QRY;TXT;TXT8"
      Top             =   4320
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1680
      TabIndex        =   28
      Top             =   5880
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label ProgressBar1Caption 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label txt8 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   27
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label txt7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   26
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label txt6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   25
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label txt5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   24
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   7560
      Picture         =   "EditionGenerateur.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label23 
      Caption         =   "Liste"
      Height          =   315
      Left            =   3240
      TabIndex        =   20
      Top             =   4320
      Width           =   525
   End
   Begin VB.Label txt12 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   19
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label txt11 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label txt10 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   17
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label txt9 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   16
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label txt4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   1395
      Left            =   1440
      TabIndex        =   15
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label txt3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Top             =   1995
      Width           =   3135
   End
   Begin VB.Label txt2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   13
      Top             =   1605
      Width           =   3135
   End
   Begin VB.Label txt1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   " Approbateur:"
      Height          =   315
      Left            =   5280
      TabIndex        =   11
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Label Label11 
      Caption         =   " Vérificateur "
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Top             =   3360
      Width           =   1005
   End
   Begin VB.Label Label10 
      Caption         =   "Dessinateur"
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label Label9 
      Caption         =   "Client"
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Label Label8 
      Caption         =   "Liste"
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "Outil"
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "Plan"
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Pièce"
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "Ensemble"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Equipement"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1995
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Vague"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1605
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Projet"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Menu Model 
      Caption         =   "Modèle"
      Visible         =   0   'False
      Begin VB.Menu lstModel 
         Caption         =   "LstModel"
         Index           =   0
      End
   End
End
Attribute VB_Name = "EditeGenerateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CopieStrtxt7 As String
Dim CopieStrtxt8 As String
Dim CopieStrtxt9 As String
Dim CopieStrtxt10 As String
Dim boolExecute As Boolean
Dim Noquite As Boolean
Dim NewControl() As Object
Dim IndexObj As Long
Dim MuComment As Collection
Dim MyFils As Long

'Public Function Charge(Id As Long, Projet As String, Vague As String, Equipement As String, _
'                        Ensemble As String, Client As String, Affaire As String, strtxt7 As String, _
'                        strtxt8 As String, strtxt9 As String, strtxt10 As String, _
'                        Dessin As String, Verif As String, Approuv As String) As Boolean
'Dim sql As String
'Dim Rs As Recordset
'Dim RsBaseNum As Recordset
'Set MuComment = Nothing
'Set MuComment = New Collection
'If ConBaseNum.OpenConnetion(DbNumPlan) = True Then
'Me.Tag = Id
'sql = "SELECT T_indiceProjet.Equipement,[PI] & '_' & Trim([PI_Indice]) AS Piece, T_indiceProjet.id FROM T_indiceProjet "
'sql = sql & "WHERE T_indiceProjet.Pere=" & Id & ";"
'Set Rs = Con.OpenRecordSet(sql)
'While Rs.EOF = False
'    ChargementFille "" & Rs!Equipement, "" & Rs!Piece, Rs!Id
'    Rs.MoveNext
'Wend
'Set Rs = Con.CloseRecordSet(Rs)
'txt1 = Projet
'txt2 = Vague
'txt3 = Equipement
'txt4 = Ensemble
'txt9 = Client
'txt5.Clear
'
'txt6.Clear
'txt7.Clear
'txt8.Clear
'CopieStrtxt7 = strtxt7
' CopieStrtxt8 = strtxt8
' CopieStrtxt9 = strtxt9
' CopieStrtxt10 = strtxt10
' txt10 = Dessin
' txt11 = Verif
' txt12 = Approuv
'If Trim("" & BdDateTable) <> "" Then
'    RqChronoAnne = "[Chrono Requête " & BdDateTable & "]"
'    ChronoAnnee = "[Chrono " & BdDateTable & "]"
'    ChronoAnnee_M1 = "[Chrono" & Val(BdDateTable) - 1 & "]"
'
'Else
'     RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
'     ChronoAnnee = "[Chrono" & Format(Date, "yyyy]")
'     ChronoAnnee_M1 = "[Chrono" & Val(Format(Date, "yyyy")) - 1 & "]"
'End If
'
'
'Me.Caption = Me.Caption & " Affaire = " & Affaire
'
'
'sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'sql = sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
'sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
'sql = sql & "FROM " & ChronoAnnee & " "
'
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'PI' "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    txt5.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
'
''    txt5.List(txt5.ListCount - 1, 1) = "" & "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
''    "_" & RsBaseNum![Clé Ch]
'''    txt5.List(txt5.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
''    txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
''    txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
''    txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
''    txt5.List(txt5.ListCount - 1, 5) = "" & RsBaseNum![Rév]
''     If strtxt7 = txt5.List(txt5.ListCount - 1) Then txt5.ListIndex = txt5.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'
'
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'PI' "
'sql = sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    txt5.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
'
''    txt5.List(txt5.ListCount - 1, 1) = "" & "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
''    txt5.List(txt5.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
''    txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
''    txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
''    txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
''    txt5.List(txt5.ListCount - 1, 5) = "" & RsBaseNum![Rév]
''     If strtxt7 = txt5.List(txt5.ListCount - 1) Then txt5.ListIndex = txt5.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'
'
'sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'sql = sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
'sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
'sql = sql & "FROM " & ChronoAnnee & " "
'
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'PL' "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    txt6.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
'
''    txt6.List(txt6.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
''    "_" & RsBaseNum![Clé Ch]
''    txt6.List(txt6.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
''    txt6.List(txt6.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
''    txt6.List(txt6.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
''    txt6.List(txt6.ListCount - 1, 5) = " " & RsBaseNum![Rév]
'    If strtxt8 = txt6.List(txt6.ListCount - 1) Then txt6.ListIndex = txt6.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'
'
'sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'sql = sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
'sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
'sql = sql & "FROM " & ChronoAnnee & " "
'
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'PL' "
'sql = sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    txt6.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
'
''    txt6.List(txt6.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
''    "_" & RsBaseNum![Clé Ch]
''    txt6.List(txt6.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
''    txt6.List(txt6.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
''    txt6.List(txt6.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
''    txt6.List(txt6.ListCount - 1, 5) = " " & RsBaseNum![Rév]
'    If strtxt8 = txt6.List(txt6.ListCount - 1) Then txt6.ListIndex = txt6.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'OU' "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    txt7.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
''    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
''    "_" & RsBaseNum![Clé Ch]
''     txt7.List(txt7.ListCount - 1, 2) = " " & RsBaseNum![Rév]
'     If strtxt9 = txt7.List(txt7.ListCount - 1) Then txt7.ListIndex = txt7.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'OU' "
'sql = sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    txt7.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
''    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
''    "_" & RsBaseNum![Clé Ch]
''     txt7.List(txt7.ListCount - 1, 2) = " " & RsBaseNum![Rév]
'     If strtxt9 = txt7.List(txt7.ListCount - 1) Then txt7.ListIndex = txt7.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'LI' "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    txt8.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
''
''    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
''    "_" & RsBaseNum![Clé Ch]
''     txt8.List(txt8.ListCount - 1, 2) = " " & RsBaseNum![Rév]
'       If strtxt10 = txt8.List(txt8.ListCount - 1) Then txt8.ListIndex = txt8.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'LI' "
'sql = sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    txt8.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
''    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
''    "_" & RsBaseNum![Clé Ch]
''     txt8.List(txt8.ListCount - 1, 2) = " " & RsBaseNum![Rév]
'       If strtxt10 = txt8.List(txt8.ListCount - 1) Then txt8.ListIndex = txt8.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'
'
'ReffIndice.Clear
'sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'LI' "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'
'
'sql = "SELECT " & ChronoAnnee & ".[Clé ty] & '_' & " & ChronoAnnee & ".[Clé ac] & '_' & " & ChronoAnnee & ".[Année] & "
'sql = sql & "'_' & " & ChronoAnnee & ".[Clé Ch] & '_' & " & ChronoAnnee & ".[Rév] AS AC, " & ChronoAnnee & ".Objet "
'sql = sql & "FROM " & ChronoAnnee & " "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ty]='AC'  "
'sql = sql & "AND " & ChronoAnnee & ".[Clé ac]=" & Affaire & "  "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    ReffIndice.AddItem "" & RsBaseNum![AC]
'
'    ReffIndice.List(ReffIndice.ListCount - 1) = "" & RsBaseNum![AC]
'    On Error Resume Next
'     MuComment.Add "" & RsBaseNum![Objet], "" & RsBaseNum![AC]
'     Err.Clear
'     On Error GoTo 0
'    RsBaseNum.MoveNext
' Wend
'
' sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
'sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
'sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
'sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
'sql = sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
'sql = sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
'sql = sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
'sql = sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
'sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
'sql = sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'LI' "
'sql = sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
'
'
'sql = "SELECT " & ChronoAnnee & ".[Clé ty] & '_' & " & ChronoAnnee & ".[Clé ac] & '_' & " & ChronoAnnee & ".[Année] & "
'sql = sql & "'_' & " & ChronoAnnee & ".[Clé Ch] & '_' & " & ChronoAnnee & ".[Rév] AS AC, " & ChronoAnnee & ".Objet "
'sql = sql & "FROM " & ChronoAnnee & " "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ty]='AC'  "
'sql = sql & "AND " & ChronoAnnee & ".[Clé ac]=" & Affaire & "  "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    ReffIndice.List(ReffIndice.ListCount - 1) = "" & RsBaseNum![AC]
'    On Error Resume Next
'     MuComment.Add "" & RsBaseNum![Objet], "" & RsBaseNum![AC]
'     Err.Clear
'     On Error GoTo 0
'
'    RsBaseNum.MoveNext
' Wend
'
'sql = "SELECT " & ChronoAnnee_M1 & ".[Clé ty] & '_' & " & ChronoAnnee_M1 & ".[Clé ac] & '_' & " & ChronoAnnee_M1 & ".[Année] & "
'sql = sql & "'_' & " & ChronoAnnee_M1 & ".[Clé Ch] & '_' & " & ChronoAnnee_M1 & ".[Rév] AS AC, " & ChronoAnnee_M1 & ".Objet "
'sql = sql & "FROM " & ChronoAnnee_M1 & " "
'sql = sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ty]='AC'  "
'sql = sql & "AND " & ChronoAnnee_M1 & ".[Clé ac]=" & Affaire & "  "
'sql = sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
'
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    ReffIndice.AddItem "" & RsBaseNum![AC]
'
'    ReffIndice.List(ReffIndice.ListCount - 1) = "" & RsBaseNum![Objet]
'
'    RsBaseNum.MoveNext
' Wend
' sql = "SELECT " & ChronoAnnee & ".[Clé ty] & '_' & " & ChronoAnnee & ".[Clé ac] & '_' & " & ChronoAnnee & ".[Année] & "
'sql = sql & "'_' & " & ChronoAnnee & ".[Clé Ch] & '_' & " & ChronoAnnee & ".[Rév] AS AC, " & ChronoAnnee & ".Objet "
'sql = sql & "FROM " & ChronoAnnee & " "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ty]='NC'  "
'sql = sql & "AND " & ChronoAnnee & ".[Clé ac]=" & Affaire & "  "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    lstNc.AddItem "" & RsBaseNum![AC]
'    RsBaseNum.MoveNext
' Wend
'
' sql = "SELECT " & ChronoAnnee_M1 & ".[Clé ty] & '_' & " & ChronoAnnee_M1 & ".[Clé ac] & '_' & " & ChronoAnnee_M1 & ".[Année] & "
'sql = sql & "'_' & " & ChronoAnnee_M1 & ".[Clé Ch] & '_' & " & ChronoAnnee_M1 & ".[Rév] AS AC, " & ChronoAnnee_M1 & ".Objet "
'sql = sql & "FROM " & ChronoAnnee_M1 & " "
'sql = sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ty]='NC'  "
'sql = sql & "AND " & ChronoAnnee_M1 & ".[Clé ac]=" & Affaire & "  "
'sql = sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
'
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    lstNc.AddItem "" & RsBaseNum![AC]
'    RsBaseNum.MoveNext
' Wend
' sql = "SELECT " & ChronoAnnee & ".[Clé ty] & '_' & " & ChronoAnnee & ".[Clé ac] & '_' & " & ChronoAnnee & ".[Année] & "
'sql = sql & "'_' & " & ChronoAnnee & ".[Clé Ch] & '_' & " & ChronoAnnee & ".[Rév] AS AC, " & ChronoAnnee & ".Objet "
'sql = sql & "FROM " & ChronoAnnee & " "
'sql = sql & "WHERE " & ChronoAnnee & ".[Clé ty]='LI'  "
'sql = sql & "AND " & ChronoAnnee & ".[Clé ac]=" & Affaire & "  "
'sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    lstLi.AddItem "" & RsBaseNum![AC]
'    RsBaseNum.MoveNext
' Wend
'
' sql = "SELECT " & ChronoAnnee_M1 & ".[Clé ty] & '_' & " & ChronoAnnee_M1 & ".[Clé ac] & '_' & " & ChronoAnnee_M1 & ".[Année] & "
'sql = sql & "'_' & " & ChronoAnnee_M1 & ".[Clé Ch] & '_' & " & ChronoAnnee_M1 & ".[Rév] AS AC, " & ChronoAnnee_M1 & ".Objet "
'sql = sql & "FROM " & ChronoAnnee_M1 & " "
'sql = sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ty]='LI'  "
'sql = sql & "AND " & ChronoAnnee_M1 & ".[Clé ac]=" & Affaire & "  "
'sql = sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
'
'Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)
'
'While RsBaseNum.EOF = False
'    lstLi.AddItem "" & RsBaseNum![AC]
'    RsBaseNum.MoveNext
' Wend
'ConBaseNum.CloseConnection
'
'Me.Show vbmodal
'Charge = boolExecute
'Else
' MsgBox "Impossible de se connecter à la base de données : " & vbCrLf & DbNumPlan & vbCrLf & vbCrLf & "Vérifiez qu'elle n'est pas en cours d'utilisation ?" & vbCrLf & "Ou contactez votre Administrateur Réseaux.", vbCritical
' Me.Hide
'End If
'End Function

Private Sub Command1_Click()
Dim Sql  As String
Dim Rs As Recordset
Dim RsBaseNum As Recordset


CherchPices.Charge Me, "(VerifieDate= Null   and Archiver=false) OR (IdStatus=3  and Archiver=false)"
CherchPicesAnnuler = CherchPices.Annuler
Unload CherchPices
If Trim("" & Me.txt1.Tag) = "" Then Exit Sub
If CherchPicesAnnuler = True Then Exit Sub
IdFils = 0
Sql = "SELECT T_indiceProjet.Pere, T_indiceProjet.CleAc FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=  " & Me.txt1.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    If Rs!Pere <> 0 Then
        Me.Tag = Rs!Pere
        MyFils = Me.txt1.Tag
    Else
         Me.Tag = Me.txt1.Tag
         MyFils = 0
    End If
    
If Trim("" & BdDateTable) <> "" Then
    RqChronoAnne = "[Chrono Requête " & BdDateTable & "]"
    ChronoAnnee = "[Chrono " & BdDateTable & "]"
    ChronoAnnee_M1 = "[Chrono" & Val(BdDateTable) - 1 & "]"
     
Else
     RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
     ChronoAnnee = "[Chrono" & Format(Date, "yyyy]")
     ChronoAnnee_M1 = "[Chrono" & Val(Format(Date, "yyyy")) - 1 & "]"
End If

ConBaseNum.TYPEBASE = ADO_TYPEBASE
ConBaseNum.SERVER = ADO_SERVER
ConBaseNum.User = ADO_User
ConBaseNum.PassWord = ADO_PassWord
ConBaseNum.BASE = DbNumPlan


    If ConBaseNum.OpenConnetion() = True Then
    
Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Rs!CleAc & " and " & ChronoAnnee & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
lstLi.Clear
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    Me.lstLi.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
'    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
'     txt8.List(txt8.ListCount - 1, 2) = " " & RsBaseNum![Rév]
'       If strtxt10 = txt8.List(txt8.ListCount - 1) Then txt8.ListIndex = txt8.ListCount - 1
    RsBaseNum.MoveNext
Wend



Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Rs!CleAc & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    lstLi.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
   
'    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
'     txt8.List(txt8.ListCount - 1, 2) = " " & RsBaseNum![Rév]
'       If strtxt10 = txt8.List(txt8.ListCount - 1) Then txt8.ListIndex = txt8.ListCount - 1
    RsBaseNum.MoveNext
Wend
Set RsBaseNum = ConBaseNum.CloseRecordSet(RsBaseNum)
ConBaseNum.CloseConnection
End If
    
End If

End Sub

'Private Sub CommandButton2_Click()
'Dim boolCahnge As Boolean
'Dim I As Long
'Dim sql As String
'Dim Fso As New FileSystemObject
'Dim IndicePi As Long
'Dim Crhono, CrhonoPi As String
'Dim DecomposeChrono
'Dim RacourciCible As String
'Dim RacourciSource As String
'Dim RsFille As Recordset
'Dim PatheFils As String
'Dim PatheFils2 As String
'Dim QuitFor As Boolean
'boolCahnge = True
'
''Set FormBarGrah = Me
'If CopieStrtxt7 = txt5 Or txt5 = "" Then boolCahnge = False
'If CopieStrtxt8 = txt6 Or txt6 = "" Then boolCahnge = False
'If CopieStrtxt9 = txt7 Or txt7 = "" Then boolCahnge = False
'If CopieStrtxt35 = txt8 Or txt8 = "" Then boolCahnge = False
'If MyFormatQRY(ReffIndice) = False Then Exit Sub
'If MyFormatQRY(Me.DescIndice) = False Then Exit Sub
'If MyFormatQRY(Me.lstNc) = False Then Exit Sub
'If MyFormatQRY(Me.lstLi) = False Then Exit Sub
'If IndexObj > 0 Then
'     For I = 1 To IndexObj
'        If Trim(UCase(PiceName(I))) = Trim(UCase(Pièce(I))) Then boolCahnge = False
'        If MyFormatQRY(Pièce(I)) = False Then
'            VScroll1.Value = (35 * (I - 1))
'            Exit Sub
'        End If
'
'    Next
'    For I = 1 To IndexObj
'        If Trim(UCase(txt5)) = Trim(UCase(Pièce(I))) Then
'            MsgBox "Vous ne pouvez pas sélectionner deux foies le même N° de pièce"
'                VScroll1.Value = (35 * (I - 1))
'            Exit Sub
'        End If
'        For I2 = 1 To IndexObj
'            If I <> I2 Then
'                If Trim(UCase(Pièce(I))) = Trim(UCase(Pièce(I2))) Then
'                    MsgBox "Vous ne pouvez pas sélectionner deux foies le même N° de pièce"
'                VScroll1.Value = (35 * (I2 - 1))
'                QuitFor = True
'                Exit Sub
'                End If
'            End If
'        Next
'
'    Next
'
'
'End If
'If boolCahnge = False Then
'    MsgBox "Vous devez changer au moins un N° chrono dans une des liste", vbOKOnly + vbExclamation, "Erreur sur l'indice"
'    Exit Sub
'End If
'
'        sql = "SELECT T_indiceProjet.PlAutoCadSave, T_indiceProjet.OuAutoCadSave,  "
'        sql = sql & "T_indiceProjet.LiAutoCadSave, T_indiceProjet.PI, T_indiceProjet.PI_Indice,  "
'        sql = sql & "T_indiceProjet.Li, T_indiceProjet.PI, T_indiceProjet.PL,  "
'        sql = sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU], T_indiceProjet.Li,T_indiceProjet.Client ,T_indiceProjet.CleAc,T_indiceProjet.Version,T_indiceProjet.Ou_Indice,T_indiceProjet.LI_Indice "
'        sql = sql & "FROM T_indiceProjet "
'        sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'            Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = False Then
'            PathDessin = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![PlAutoCadSave])
'
'                If Fso.FileExists(PathDessin & ".dwg") = True Then
'                    SecuFill PathDessin & ".dwg", False
'                     DecomposeChrono = Split(txt6, "_")
'                      Affaire = Val(DecomposeChrono(1))
'                      Indice = Val(DecomposeChrono(UBound(DecomposeChrono)))
'                      Crhono = ""
'                    For I = 0 To UBound(DecomposeChrono) - 1
'                        Debug.Print DecomposeChrono(I) & "_"
'                        Crhono = Crhono & DecomposeChrono(I) & "_"
'                    Next
'
'                    Crhono = Left(Crhono, Len(Crhono) - 1)
'                    CrhonoPi = ""
'                    DecomposeChrono = Split(txt5, "_")
'                     For I = 0 To UBound(DecomposeChrono) - 1
'                        Debug.Print DecomposeChrono(I) & "_"
'                        CrhonoPi = CrhonoPi & DecomposeChrono(I) & "_"
'                    Next
'                     CrhonoPi = Left(CrhonoPi, Len(CrhonoPi) - 1)
'                     IndicePi = Val(DecomposeChrono(UBound(DecomposeChrono)))
'                     fileCopie = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Affaire, "" & CrhonoPi, "PL", Crhono, Val(Me.Tag), Val(IndicePi), Val(Indice), 1) & ".dwg"
'
'                     If Fso.FileExists(fileCopie) = True Then
'                     SecuFill "" & fileCopie, False
'                        Fso.DeleteFile fileCopie
'                     End If
'                    Fso.CopyFile PathDessin & ".dwg", fileCopie
'                     DecomposeChrono = Split(Crhono, "_")
'                      Affaire = Val(DecomposeChrono(1))
'                      Indice = Val(DecomposeChrono(UBound(DecomposeChrono)))
'                      Crhono = ""
'                    For I = 0 To UBound(DecomposeChrono) - 1
'                        Debug.Print DecomposeChrono(I) & "_"
'                        Crhono = Crhono & DecomposeChrono(I) & "_"
'                    Next
'                    Crhono = Left(Crhono, Len(Crhono) - 1)
'                    sql = "UPDATE T_indiceProjet  SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1,T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "',  T_indiceProjet.PL = '" & Crhono & "', T_indiceProjet.PL_Indice = '" & Indice & "', T_indiceProjet.ApprouveDate = Date() "
'                    sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'                    Con.Execute sql
'                End If
'
'                PathDessin = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![OUAutoCadSave])
'                If Fso.FileExists(PathDessin & ".dwg") = True Then
'                SecuFill PathDessin & ".dwg", False
'                     DecomposeChrono = Split(txt7, "_")
'                      Affaire = Val(DecomposeChrono(1))
'                      Indice = Val(DecomposeChrono(UBound(DecomposeChrono)))
'                      Crhono = ""
'                    For I = 0 To UBound(DecomposeChrono) - 1
'                        Debug.Print DecomposeChrono(I) & "_"
'                        Crhono = Crhono & DecomposeChrono(I) & "_"
'                    Next
'                    Crhono = Left(Crhono, Len(Crhono) - 1)
'                    fileCopie = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Affaire, "" & CrhonoPi, "OU", Crhono, Val(Me.Tag), Val(IndicePi), Val(Indice), 1) & ".dwg"
'
'                     If Fso.FileExists(fileCopie) = True Then
'                     SecuFill "" & fileCopie, False
'                        Fso.DeleteFile fileCopie
'                     End If
'                    Fso.CopyFile PathDessin & ".dwg", fileCopie
''                    Fso.CopyFile PathDessin & ".dwg", PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & Affaire, "" & CrhonoPi, "OU", Crhono, Val(Me.Tag), Val(IndicePi), Val(Indice), 1) & ".dwg"
'                    DecomposeChrono = Split(Crhono, "_")
'                      Affaire = Val(DecomposeChrono(1))
'                      Indice = Val(DecomposeChrono(UBound(DecomposeChrono)))
'                      Crhono = ""
'                    For I = 0 To UBound(DecomposeChrono) - 1
'                        Debug.Print DecomposeChrono(I) & "_"
'                        Crhono = Crhono & DecomposeChrono(I) & "_"
'                    Next
'                    Crhono = Left(Crhono, Len(Crhono) - 1)
'                    sql = "UPDATE T_indiceProjet SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1,T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "', T_indiceProjet.OU = '" & Crhono & "', T_indiceProjet.OU_Indice = '" & Indice & "', T_indiceProjet.ApprouveDate = Date() "
'                    sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'                    Con.Execute sql
'                End If
'                PathDessin = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![LiAutoCadSave])
'                If Fso.FileExists(PathDessin & ".XLS") = True Then
'                SecuFill PathDessin & ".XLS", False
'                    DecomposeChrono = Split(txt8, "_")
'                      Affaire = Val(DecomposeChrono(1))
'                      Indice = Val(DecomposeChrono(UBound(DecomposeChrono)))
'                      Crhono = ""
'                    For I = 0 To UBound(DecomposeChrono) - 1
'                        Debug.Print DecomposeChrono(I) & "_"
'                        Crhono = Crhono & DecomposeChrono(I) & "_"
'                    Next
'                    Crhono = Left(Crhono, Len(Crhono) - 1)
'                    fileCopie = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Affaire, "" & CrhonoPi, "LI", Crhono, Val(Me.Tag), Val(IndicePi), Val(Indice), 1) & ".XLS"
'
'                     If Fso.FileExists(fileCopie) = True Then
'                     SecuFill "" & fileCopie, False
'                        Fso.DeleteFile fileCopie
'                     End If
'                    Fso.CopyFile PathDessin & ".XLS", fileCopie
'                    KilVersionXX "" & PathDessin, "" & fileCopie, True
''                    Fso.CopyFile PathDessin & ".XLS", PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & Affaire, "" & CrhonoPi, "LI", Crhono, Val(Me.Tag), Val(IndicePi), Val(Indice), 1) & ".XLS"
'                    DecomposeChrono = Split(Crhono, "_")
'                      Affaire = Val(DecomposeChrono(1))
'                      Indice = Val(DecomposeChrono(UBound(DecomposeChrono)))
'                      Crhono = ""
'                    For I = 0 To UBound(DecomposeChrono) - 1
'                        Debug.Print DecomposeChrono(I) & "_"
'                        Crhono = Crhono & DecomposeChrono(I) & "_"
'                    Next
'                    Crhono = Left(Crhono, Len(Crhono) - 1)
'                    sql = "UPDATE T_indiceProjet SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1,T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "',T_indiceProjet.Li = '" & Crhono & "', T_indiceProjet.Li_Indice = '" & Indice & "', T_indiceProjet.ApprouveDate = Date() "
'                    sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'                    Con.Execute sql
''                    subExporteXls Val(Me.Tag)
'                End If
'
'                Rs.Requery
'                PathPl = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![LiAutoCadSave])
'                'fileCopie
'                Set Fso = Nothing
'                PathDessin2 = PathDessin
'                PathPl2 = PathPl
'
'             Set Fso = New FileSystemObject
'                 If IndexObj > 0 Then
'                    For I2 = 1 To IndexObj
'                       CrhonoPi = ""
'                    DecomposeChrono = Split(Me.PiceName(I2), "_")
'                     For I = 0 To UBound(DecomposeChrono) - 1
'                        Debug.Print DecomposeChrono(I) & "_"
'                        CrhonoPi = CrhonoPi & DecomposeChrono(I) & "_"
'                    Next
'                     CrhonoPi = Left(CrhonoPi, Len(CrhonoPi) - 1)
'                     IndicePi = Val(DecomposeChrono(UBound(DecomposeChrono)))
'
'                        PathPl = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![PlAutoCadSave])
'                        Racourci PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & CrhonoPi, "PL", Rs!PL, Val(Me.PiceName(I2).Tag), Val(IndicePi), Rs!PL_Indice, 1), "" & PathPl, "DWG"
'
'                        PathPl = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![OUAutoCadSave])
'                        Racourci PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & CrhonoPi, "OU", Rs!OU, Val(Me.PiceName(I2).Tag), Val(IndicePi), Rs!OU_Indice, 1), "" & PathPl, "DWG"
'                        sql = "SELECT T_indiceProjet.LiAutoCadSave FROM T_indiceProjet WHERE T_indiceProjet.Id=" & PiceName(I2).Tag & ";"
'                        Set RsFille = Con.OpenRecordSet(sql)
'
'                        PatheFils = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & RsFille![LiAutoCadSave])
'
'                        Racourci PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & CrhonoPi, "LI", Rs!LI, Val(Me.PiceName(I2).Tag), Val(IndicePi), Rs!LI_Indice, 1), "" & PathPl, "XLS"
'                        RsFille.Requery
'                        PatheFils2 = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & RsFille![LiAutoCadSave])
'                       Set Fso = Nothing
'                       DoEvents
'                        KilVersionXX "" & PatheFils, "" & PathPl, True
'
'                        Set RsFille = Con.CloseRecordSet(RsFille)
'                            sql = "UPDATE T_indiceProjet SET  T_indiceProjet.Version = 1, "
'                            sql = sql & "T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "', "
'                            sql = sql & "T_indiceProjet.PI = '" & CrhonoPi & "',  "
'                            sql = sql & "T_indiceProjet.PI_Indice = '" & IndicePi & "',  "
'                            sql = sql & "T_indiceProjet.PL = '" & Rs!PL & "',  "
'                            sql = sql & "T_indiceProjet.PL_Indice = '" & Rs!PL_Indice & "',  "
'                            sql = sql & "T_indiceProjet.[OU] = '" & Rs!OU & "',  "
'                            sql = sql & "T_indiceProjet.OU_Indice = '" & Rs!OU_Indice & "',  "
'                            sql = sql & "T_indiceProjet.Li = '" & Rs!LI & "',  "
'                            sql = sql & "T_indiceProjet.LI_Indice = '" & Rs!LI_Indice & "',  "
'                            sql = sql & "T_indiceProjet.ApprouveDate =  Date() "
'                            sql = sql & "WHERE T_indiceProjet.Id=" & Me.PiceName(I2).Tag & ";"
'                            Con.Execute sql
'                    Next
'                End If
'            End If
'
' Set Fso = Nothing
' DoEvents
'
'boolExecute = True
'Noquite = False
'Noquite = False
'Me.Hide
'End Sub

Private Sub CommandButton3_Click()
Noquite = False
'frmAutocâble.DesEnabledMenu
Unload Me
'Me.Hide
End Sub

Private Sub LstFils_Click()

End Sub



Private Sub Exécuter_Click()
Dim MenuName
Dim IdIndiceProjet As Long
Dim Sql As String
Dim Rs As Recordset
If MyFormatQRY(Me.lstLi) = False Then Exit Sub
If Exécuter.Caption = "Pas de modèle sélectionné" Then
    MsgBox "Vous devez sélectionner un modèle de document."
    Exit Sub
End If
MenuName = Split(Exécuter.Caption, ":")
Set LstColecDoc = Nothing
Set LstColecDoc = New Collection
IndexTableauDocGen = 0
Dim ColecDoc As New Collection
Dim IndexTableauDoc As Long
Dim TableDoc As GenerateurDoc
Dim Fso As New FileSystemObject
Dim I As Long
'ReDim TableDoc(0)
Set TableauPath = funPath
Set FormBarGrah = Me
' TableDoc(0).Menu = "Menu2"
Set TableDoc = New GenerateurDoc
'TableDoc.indexdoc=TableDoc.c
ColecDoc.Add TableDoc, Replace(Trim(MenuName(1)), " ", "_")
Set TableDoc = Nothing
ColecDoc(Replace(Trim(MenuName(1)), " ", "_")).Menu = Trim(MenuName(1))
Sql = "SELECT T_indiceProjet.Client, T_indiceProjet.CleAc, T_indiceProjet.PI, T_indiceProjet.Version, T_indiceProjet.PI_Indice "
Sql = Sql & "FROM T_indiceProjet WHERE T_indiceProjet.Id="
If MyFils <> 0 Then
  IdIndiceProjet = MyFils
Else
    IdIndiceProjet = Val(Me.Tag)
End If
Sql = Sql & IdIndiceProjet
Set Rs = Con.OpenRecordSet(Sql)
ColecDoc(Replace(Trim(MenuName(1)), " ", "_")).SaveAs = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!PI, "PDF", lstLi, IdIndiceProjet, Rs.Fields("PI_Indice"), "", Rs!Version, True)
ColecDoc(Replace(Trim(MenuName(1)), " ", "_")).LoadColecDoc ColecDoc(Replace(Trim(MenuName(1)), " ", "_")).IndexTableauDoc, ColecDoc, TableDoc, _
    ColecDoc(Replace(Trim(MenuName(1)), " ", "_")).SaveAs
  
    Sql = "SELECT T_ETATS.IsPDF, T_ETATS.id FROM T_ETATS "
    Sql = Sql & "WHERE T_ETATS.Menu='" & MyReplace(Trim(MenuName(1))) & "';"
Set Rs = Con.OpenRecordSet(Sql)


ColecDoc(Replace(Trim(MenuName(1)), " ", "_")).IsPDF = Rs!IsPDF
Set Rs = Con.CloseRecordSet(Rs)
Debug.Print ColecDoc(Replace(Trim(MenuName(1)), " ", "_")).SaveAs
'For I = 1 To ColecDoc.Count
'For I = 1 To ColecDoc.Count
For I = ColecDoc.Count To 1 Step -1
    
    ColecDoc(I).SelectEtat I, ColecDoc, ColecDoc(I), Val(Me.Tag), MyFils, ColecDoc(I).SaveAs, I - 1
Next

For I = 2 To ColecDoc.Count
Debug.Print ColecDoc(I).SaveAs & ".Xls"

    If Fso.FileExists(ColecDoc(I).SaveAs & ".Xls") = True Then
    Second 1
        Fso.DeleteFile ColecDoc(I).SaveAs & ".Xls"
    End If
Next
'frmAutocâble.DesEnabledMenu
Unload Me
End Sub

Private Sub Form_Initialize()
IndexObj = 0
End Sub

Private Sub Form_Load()
Dim Sql As String
Dim Rs As Recordset
Dim MenuCaption As String
Sql = "SELECT T_ETATS.Menu FROM T_ETATS "
Sql = Sql & "Where T_ETATS.Menu Is Not Null And T_ETATS.Visible = True ORDER BY T_ETATS.Menu;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
MenuCaption = "" & Rs!Menu
    If Model.Visible = False Then
        lstModel(0).Caption = MenuCaption
        lstModel(0).Visible = True
        Model.Visible = True
    Else
        Load lstModel(lstModel.Count)
        lstModel(lstModel.Count - 1).Caption = MenuCaption
        lstModel(lstModel.Count - 1).Visible = True
    End If
Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = Noquite
 IndexObj = 0
End Sub

Private Sub Form_Terminate()
IndexObj = 0
End Sub

'Private Sub ReffIndice_Click()
'Me.DescIndice = MuComment(Me.ReffIndice.List(Me.ReffIndice.ListIndex))
'End Sub

Private Sub lstModel_Click(Index As Integer)
Exécuter.Caption = "Exécuter: " & lstModel(Index).Caption
End Sub

Private Sub txt5_Click()
'If Me.txt5.ListIndex <> -1 Then
'    Me.txt10 = txt5.List(Me.txt5.ListIndex, 2)
'     Me.txt11 = txt5.List(Me.txt5.ListIndex, 3)
'      Me.txt12 = txt5.List(Me.txt5.ListIndex, 4)
'Else
'     Me.txt10 = ""
'     Me.txt11 = ""
'     Me.txt12 = ""
'End If
End Sub




'Private Sub ChargementFille(Equipement As String, PI As String, Id As Long)
'IndexObj = IndexObj + 1
'
'Load PiceEquipement(IndexObj)
'PiceEquipement(IndexObj).Top = PiceEquipement(0).Top + (315 * (IndexObj - 1))
'PiceEquipement(IndexObj).Visible = True
'PiceEquipement(IndexObj) = Equipement
'
'Load Pièce(IndexObj)
'Pièce(IndexObj).Top = Pièce(0).Top + (315 * (IndexObj - 1))
'Pièce(IndexObj).Visible = True
'
'Load PiceName(IndexObj)
'PiceName(IndexObj) = PI
'PiceName(IndexObj).Tag = CStr(Id)
'
'PiceName(IndexObj).Top = PiceName(0).Top + (315 * (IndexObj - 1))
'PiceName(IndexObj).Visible = True
'Me.EnvelopePice.Height = 495 + (315 * (IndexObj - 1))
'If IndexObj = 7 Then VScroll1.Visible = True
'DoEvents
'End Sub





'Private Sub VScroll1_Change()
'If Me.VScroll1.Value = 0 Then
'    EnvelopePice.Top = 120
'Else
'EnvelopePice.Top = 120 + (Me.VScroll1.Value * (-1 * IndexObj))
'End If
'End Sub
