VERSION 5.00
Begin VB.Form FrmIndice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indice Description"
   ClientHeight    =   11025
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11025
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandButton3 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   6360
      TabIndex        =   37
      Top             =   10440
      Width           =   1455
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "&Valider"
      Height          =   375
      Left            =   1680
      TabIndex        =   36
      Top             =   10440
      Width           =   1455
   End
   Begin VB.TextBox DescIndice 
      BackColor       =   &H00FFFF80&
      Height          =   1335
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   35
      Tag             =   ".DESIGNATION.LIGNE.1;Description Indice;QRY;TXT;DescIndice"
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1080
      TabIndex        =   30
      Top             =   5520
      Width           =   8415
      Begin VB.VScrollBar VScroll1 
         Height          =   2175
         LargeChange     =   350
         Left            =   8160
         Max             =   9999
         SmallChange     =   35
         TabIndex        =   39
         Top             =   0
         Value           =   120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox EnvelopePice 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   8175
         TabIndex        =   32
         Top             =   120
         Width           =   8175
         Begin VB.ComboBox Pièce 
            BackColor       =   &H00FFFF80&
            Height          =   315
            Index           =   0
            Left            =   5640
            TabIndex        =   33
            Tag             =   "Liste;Liste;QRY;TXT;TXT10"
            Top             =   0
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label PiceEquipement 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label PiceName 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   3120
            TabIndex        =   38
            Top             =   0
            Visible         =   0   'False
            Width           =   2535
         End
      End
   End
   Begin VB.ComboBox lstLi 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   3240
      TabIndex        =   29
      Tag             =   "Liste ;Liste;QRY;TXT;TXT8"
      Top             =   5160
      Width           =   3135
   End
   Begin VB.ComboBox lstNc 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   3240
      TabIndex        =   28
      Tag             =   "Non Conformité;Non Conformité;QRY;TXT;TXT8"
      Top             =   4800
      Width           =   3135
   End
   Begin VB.ComboBox ReffIndice 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   3240
      TabIndex        =   27
      Tag             =   "Action Corrective  ;Action Corrective ;QRY;TXT;TXT8"
      Top             =   4320
      Width           =   3135
   End
   Begin VB.ComboBox txt8 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   6360
      TabIndex        =   23
      Top             =   2160
      Width           =   3135
   End
   Begin VB.ComboBox txt6 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   6360
      TabIndex        =   22
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ComboBox txt7 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   6360
      TabIndex        =   21
      Top             =   1800
      Width           =   3135
   End
   Begin VB.ComboBox txt5 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   6360
      TabIndex        =   20
      Tag             =   "N° P ;N° PL;QRY;TXT;TXT8"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   7560
      Picture         =   "FrmIndice.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "Description Indice :"
      Height          =   315
      Left            =   3120
      TabIndex        =   34
      Top             =   8280
      Width           =   3135
   End
   Begin VB.Label Label24 
      Caption         =   "Pièces Filles"
      Height          =   195
      Left            =   0
      TabIndex        =   31
      Top             =   5640
      Width           =   1005
   End
   Begin VB.Label Label23 
      Caption         =   "Liste"
      Height          =   315
      Left            =   1440
      TabIndex        =   26
      Top             =   5115
      Width           =   1725
   End
   Begin VB.Label Label22 
      Caption         =   "Non Conformité"
      Height          =   315
      Left            =   1440
      TabIndex        =   25
      Top             =   4710
      Width           =   1725
   End
   Begin VB.Label Label21 
      Caption         =   "Action Corrective"
      Height          =   315
      Left            =   1440
      TabIndex        =   24
      Top             =   4320
      Width           =   1725
   End
   Begin VB.Label txt12 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   19
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label txt11 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   18
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label txt10 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   17
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label txt9 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   16
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label txt4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   1395
      Left            =   1440
      TabIndex        =   15
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label txt3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Top             =   1875
      Width           =   3135
   End
   Begin VB.Label txt2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   13
      Top             =   1485
      Width           =   3135
   End
   Begin VB.Label txt1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   " Approbateur:"
      Height          =   315
      Left            =   5280
      TabIndex        =   11
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Label Label11 
      Caption         =   " Vérificateur "
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label Label10 
      Caption         =   "Dessinateur"
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      Top             =   2880
      Width           =   1005
   End
   Begin VB.Label Label9 
      Caption         =   "Client"
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label8 
      Caption         =   "Liste"
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label Label7 
      Caption         =   "Outil"
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "Plan"
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Pièce"
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "Ensemble"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Equipement"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1875
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Vague"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1485
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Projet"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1005
   End
End
Attribute VB_Name = "FrmIndice"
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

Public Function Charge(Id As Long, Projet As String, Vague As String, Equipement As String, _
                        Ensemble As String, Client As String, Affaire As String, strtxt7 As String, _
                        strtxt8 As String, strtxt9 As String, strtxt10 As String, _
                        Dessin As String, Verif As String, Approuv As String) As Boolean
Dim Sql As String
Dim Rs As Recordset
Dim RsBaseNum As Recordset
Set MuComment = Nothing
Set MuComment = New Collection
ConBaseNum.TYPEBASE = ADO_TYPEBASE
ConBaseNum.SERVER = ADO_SERVER
ConBaseNum.User = ADO_User
ConBaseNum.PassWord = ADO_PassWord
ConBaseNum.BASE = DbNumPlan



If ConBaseNum.OpenConnetion = True Then
Me.Tag = Id
Sql = "SELECT T_indiceProjet.Equipement,[PI] & '_' & Trim([PI_Indice]) AS Piece, T_indiceProjet.id "
Sql = Sql & ", T_indiceProjet.Id, T_Status.Status "
Sql = Sql & "FROM T_Status INNER JOIN T_indiceProjet ON T_Status.Id = T_indiceProjet.IdStatus  "
Sql = Sql & "WHERE T_indiceProjet.Pere=" & Id & ";"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    ChargementFille "" & Rs!Equipement, "" & Rs!Piece, Rs!Id, "" & Rs!Status
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
txt1 = Projet
txt2 = Vague
txt3 = Equipement
TXT4 = Ensemble
txt9 = Client
txt5.Clear

txt6.Clear
txt7.Clear
txt8.Clear
CopieStrtxt7 = strtxt7
 CopieStrtxt8 = strtxt8
 CopieStrtxt9 = strtxt9
 CopieStrtxt10 = strtxt10
 txt10 = Dessin
 txt11 = Verif
 txt12 = Approuv
If Trim("" & BdDateTable) <> "" Then
    RqChronoAnne = "[Chrono Requête " & BdDateTable & "]"
    ChronoAnnee = "[Chrono " & BdDateTable & "]"
    ChronoAnnee_M1 = "[Chrono" & Val(BdDateTable) - 1 & "]"
     
Else
     RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
     ChronoAnnee = "[Chrono" & Format(Date, "yyyy]")
     ChronoAnnee_M1 = "[Chrono" & Val(Format(Date, "yyyy")) - 1 & "]"
End If


Me.Caption = Me.Caption & " Affaire = " & Affaire


Sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
Sql = Sql & "FROM " & ChronoAnnee & " "

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
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'PI' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt5.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
    
    
'    txt5.List(txt5.ListCount - 1, 1) = "" & "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
''    txt5.List(txt5.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
'    txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt5.List(txt5.ListCount - 1, 5) = "" & RsBaseNum![Rév]
     If strtxt7 = txt5.List(txt5.ListCount - 1) Then txt5.ListIndex = txt5.ListCount - 1
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
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'PI' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt5.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
    
    
'    txt5.List(txt5.ListCount - 1, 1) = "" & "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch]
'    txt5.List(txt5.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
'    txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt5.List(txt5.ListCount - 1, 5) = "" & RsBaseNum![Rév]
     If strtxt7 = txt5.List(txt5.ListCount - 1) Then txt5.ListIndex = txt5.ListCount - 1
    RsBaseNum.MoveNext
Wend


Sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
Sql = Sql & "FROM " & ChronoAnnee & " "

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
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'PL' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt6.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
    
    
'    txt6.List(txt6.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
'    txt6.List(txt6.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt6.List(txt6.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt6.List(txt6.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt6.List(txt6.ListCount - 1, 5) = " " & RsBaseNum![Rév]
    If strtxt8 = txt6.List(txt6.ListCount - 1) Then txt6.ListIndex = txt6.ListCount - 1
    RsBaseNum.MoveNext
Wend


Sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
Sql = Sql & "FROM " & ChronoAnnee & " "

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
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'PL' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt6.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
    
    
'    txt6.List(txt6.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
'    txt6.List(txt6.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt6.List(txt6.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt6.List(txt6.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt6.List(txt6.ListCount - 1, 5) = " " & RsBaseNum![Rév]
    If strtxt8 = txt6.List(txt6.ListCount - 1) Then txt6.ListIndex = txt6.ListCount - 1
    RsBaseNum.MoveNext
Wend

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
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'OU' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt7.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
     
'    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
'     txt7.List(txt7.ListCount - 1, 2) = " " & RsBaseNum![Rév]
     If strtxt9 = txt7.List(txt7.ListCount - 1) Then txt7.ListIndex = txt7.ListCount - 1
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
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'OU' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt7.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
     
'    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
'     txt7.List(txt7.ListCount - 1, 2) = " " & RsBaseNum![Rév]
     If strtxt9 = txt7.List(txt7.ListCount - 1) Then txt7.ListIndex = txt7.ListCount - 1
    RsBaseNum.MoveNext
Wend
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
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt8.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
'
'    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
'     txt8.List(txt8.ListCount - 1, 2) = " " & RsBaseNum![Rév]
       If strtxt10 = txt8.List(txt8.ListCount - 1) Then txt8.ListIndex = txt8.ListCount - 1
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
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt8.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
   
'    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
'    "_" & RsBaseNum![Clé Ch]
'     txt8.List(txt8.ListCount - 1, 2) = " " & RsBaseNum![Rév]
       If strtxt10 = txt8.List(txt8.ListCount - 1) Then txt8.ListIndex = txt8.ListCount - 1
    RsBaseNum.MoveNext
Wend


ReffIndice.Clear
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
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"


Sql = "SELECT " & ChronoAnnee & ".[Clé ty] & '_' & " & ChronoAnnee & ".[Clé ac] & '_' & " & ChronoAnnee & ".[Année] & "
Sql = Sql & "'_' & " & ChronoAnnee & ".[Clé Ch] & '_' & " & ChronoAnnee & ".[Rév] AS AC, " & ChronoAnnee & ".Objet "
Sql = Sql & "FROM " & ChronoAnnee & " "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ty]='AC'  "
Sql = Sql & "AND " & ChronoAnnee & ".[Clé ac]=" & Affaire & "  "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    ReffIndice.AddItem "" & RsBaseNum![AC]
   
    ReffIndice.List(ReffIndice.ListCount - 1) = "" & RsBaseNum![AC]
    On Error Resume Next
     MuComment.Add "" & RsBaseNum![Objet], "" & RsBaseNum![AC]
     Err.Clear
     On Error GoTo 0
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
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"


Sql = "SELECT " & ChronoAnnee & ".[Clé ty] & '_' & " & ChronoAnnee & ".[Clé ac] & '_' & " & ChronoAnnee & ".[Année] & "
Sql = Sql & "'_' & " & ChronoAnnee & ".[Clé Ch] & '_' & " & ChronoAnnee & ".[Rév] AS AC, " & ChronoAnnee & ".Objet "
Sql = Sql & "FROM " & ChronoAnnee & " "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ty]='AC'  "
Sql = Sql & "AND " & ChronoAnnee & ".[Clé ac]=" & Affaire & "  "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    ReffIndice.List(ReffIndice.ListCount - 1) = "" & RsBaseNum![AC]
    On Error Resume Next
     MuComment.Add "" & RsBaseNum![Objet], "" & RsBaseNum![AC]
     Err.Clear
     On Error GoTo 0
     
    RsBaseNum.MoveNext
 Wend
 
Sql = "SELECT " & ChronoAnnee_M1 & ".[Clé ty] & '_' & " & ChronoAnnee_M1 & ".[Clé ac] & '_' & " & ChronoAnnee_M1 & ".[Année] & "
Sql = Sql & "'_' & " & ChronoAnnee_M1 & ".[Clé Ch] & '_' & " & ChronoAnnee_M1 & ".[Rév] AS AC, " & ChronoAnnee_M1 & ".Objet "
Sql = Sql & "FROM " & ChronoAnnee_M1 & " "
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ty]='AC'  "
Sql = Sql & "AND " & ChronoAnnee_M1 & ".[Clé ac]=" & Affaire & "  "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    ReffIndice.AddItem "" & RsBaseNum![AC]
   
    ReffIndice.List(ReffIndice.ListCount - 1) = "" & RsBaseNum![Objet]
     
    RsBaseNum.MoveNext
 Wend
 Sql = "SELECT " & ChronoAnnee & ".[Clé ty] & '_' & " & ChronoAnnee & ".[Clé ac] & '_' & " & ChronoAnnee & ".[Année] & "
Sql = Sql & "'_' & " & ChronoAnnee & ".[Clé Ch] & '_' & " & ChronoAnnee & ".[Rév] AS AC, " & ChronoAnnee & ".Objet "
Sql = Sql & "FROM " & ChronoAnnee & " "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ty]='NC'  "
Sql = Sql & "AND " & ChronoAnnee & ".[Clé ac]=" & Affaire & "  "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    lstNc.AddItem "" & RsBaseNum![AC]
    RsBaseNum.MoveNext
 Wend
 
 Sql = "SELECT " & ChronoAnnee_M1 & ".[Clé ty] & '_' & " & ChronoAnnee_M1 & ".[Clé ac] & '_' & " & ChronoAnnee_M1 & ".[Année] & "
Sql = Sql & "'_' & " & ChronoAnnee_M1 & ".[Clé Ch] & '_' & " & ChronoAnnee_M1 & ".[Rév] AS AC, " & ChronoAnnee_M1 & ".Objet "
Sql = Sql & "FROM " & ChronoAnnee_M1 & " "
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ty]='NC'  "
Sql = Sql & "AND " & ChronoAnnee_M1 & ".[Clé ac]=" & Affaire & "  "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    lstNc.AddItem "" & RsBaseNum![AC]
    RsBaseNum.MoveNext
 Wend
 Sql = "SELECT " & ChronoAnnee & ".[Clé ty] & '_' & " & ChronoAnnee & ".[Clé ac] & '_' & " & ChronoAnnee & ".[Année] & "
Sql = Sql & "'_' & " & ChronoAnnee & ".[Clé Ch] & '_' & " & ChronoAnnee & ".[Rév] AS AC, " & ChronoAnnee & ".Objet "
Sql = Sql & "FROM " & ChronoAnnee & " "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ty]='LI'  "
Sql = Sql & "AND " & ChronoAnnee & ".[Clé ac]=" & Affaire & "  "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    lstLi.AddItem "" & RsBaseNum![AC]
    RsBaseNum.MoveNext
 Wend
 
 Sql = "SELECT " & ChronoAnnee_M1 & ".[Clé ty] & '_' & " & ChronoAnnee_M1 & ".[Clé ac] & '_' & " & ChronoAnnee_M1 & ".[Année] & "
Sql = Sql & "'_' & " & ChronoAnnee_M1 & ".[Clé Ch] & '_' & " & ChronoAnnee_M1 & ".[Rév] AS AC, " & ChronoAnnee_M1 & ".Objet "
Sql = Sql & "FROM " & ChronoAnnee_M1 & " "
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ty]='LI'  "
Sql = Sql & "AND " & ChronoAnnee_M1 & ".[Clé ac]=" & Affaire & "  "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    lstLi.AddItem "" & RsBaseNum![AC]
    RsBaseNum.MoveNext
 Wend
ConBaseNum.CloseConnection

Me.Show vbModal
Charge = boolExecute
Else
 MsgBox "Impossible de se connecter à la base de données : " & vbCrLf & DbNumPlan & vbCrLf & vbCrLf & "Vérifiez qu'elle n'est pas en cours d'utilisation ?" & vbCrLf & "Ou contactez votre Administrateur Réseaux.", vbCritical
 Me.Hide
End If
End Function

Private Sub CommandButton2_Click()
Dim boolCahnge As Boolean
Dim I As Long
Dim Sql As String
Dim Fso As New FileSystemObject
Dim IndicePi As Long
Dim Crhono, CrhonoPi As String
Dim DecomposeChrono
Dim RacourciCible As String
Dim RacourciSource As String
Dim RsFille As Recordset
Dim PatheFils As String
Dim PatheFils2 As String
Dim QuitFor As Boolean
boolCahnge = False
Dim CronoPl As String
Dim ChronoPl As String
Dim IndicePl As Integer
Dim ChronoLi As String
Dim ChronoOu As String
Dim IndiceLi As Integer
Dim IndiceOu As Integer
'Set FormBarGrah = Me
If CopieStrtxt7 <> txt5 And txt5 <> "" Then boolCahnge = True
If CopieStrtxt8 <> txt6 And txt6 = "" Then boolCahnge = True
If CopieStrtxt9 <> txt7 And txt7 = "" Then boolCahnge = True
If CopieStrtxt35 <> txt8 And txt8 = "" Then boolCahnge = True
If MyFormatQRY(ReffIndice) = False Then Exit Sub
If MyFormatQRY(Me.DescIndice) = False Then Exit Sub
If MyFormatQRY(Me.lstNc) = False Then Exit Sub
If MyFormatQRY(Me.lstLi) = False Then Exit Sub
If IndexObj > 0 Then
     For I = 1 To IndexObj
        If Trim(UCase(PiceName(I))) = Trim(UCase(Pièce(I))) Then boolCahnge = False
        If MyFormatQRY(Pièce(I)) = False Then
            VScroll1.Value = (35 * (I - 1))
            Exit Sub
        End If
        
    Next
    For I = 1 To IndexObj
        If Trim(UCase(txt5)) = Trim(UCase(Pièce(I))) Then
            MsgBox "Vous ne pouvez pas sélectionner deux foies le même N° de pièce"
                VScroll1.Value = (35 * (I - 1))
            Exit Sub
        End If
        For I2 = 1 To IndexObj
            If I <> I2 Then
                If Trim(UCase(Pièce(I))) = Trim(UCase(Pièce(I2))) Then
                    MsgBox "Vous ne pouvez pas sélectionner deux foies le même N° de pièce"
                VScroll1.Value = (35 * (I2 - 1))
                QuitFor = True
                Exit Sub
                End If
            End If
        Next
      
    Next
    
   
End If
If boolCahnge = False Then
    MsgBox "Vous devez changer au moins un N° chrono dans une des liste", vbOKOnly + vbExclamation, "Erreur sur l'indice"
    Exit Sub
End If

        Sql = "UPDATE T_indiceProjet SET T_indiceProjet.ReffIndice = '" & MyReplace(Me.ReffIndice) & " ', T_indiceProjet.DNC = '" & MyReplace(Me.lstNc) & " ', T_indiceProjet.LIEC = '" & MyReplace(Me.lstLi) & "',Descripton='" & MyReplace(Me.DescIndice) & "' "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
        Con.Execute Sql
        
        
        Sql = "SELECT T_indiceProjet.PlAutoCadSave, T_indiceProjet.OuAutoCadSave,  "
        Sql = Sql & "T_indiceProjet.LiAutoCadSave, T_indiceProjet.PI, T_indiceProjet.PI_Indice,  "
        Sql = Sql & "T_indiceProjet.Li, T_indiceProjet.PI, T_indiceProjet.PL,  "
        Sql = Sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU], T_indiceProjet.Li,T_indiceProjet.Client ,T_indiceProjet.CleAc,T_indiceProjet.Version,T_indiceProjet.Ou_Indice,T_indiceProjet.LI_Indice "
        Sql = Sql & "FROM T_indiceProjet "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = False Then
            End If
'            PathDessin = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![PlAutoCadSave])
            
'                If Fso.FileExists(PathDessin & ".dwg") = True Then
'                    SecuFill PathDessin & ".dwg", False
                     DecomposeChrono = Split(txt6, "_")
                      Affaire = Val(DecomposeChrono(1))
                      IndicePl = Val(DecomposeChrono(UBound(DecomposeChrono)))
                      ChronoPl = ""
                    For I = 0 To UBound(DecomposeChrono) - 1
                        Debug.Print DecomposeChrono(I) & "_"
                        ChronoPl = ChronoPl & DecomposeChrono(I) & "_"
                    Next
                   
                    ChronoPl = Left(ChronoPl, Len(ChronoPl) - 1)
                  
                    DecomposeChrono = Split(txt5, "_")
                    CrhonoPi = ""
                    IndicePi = Val(DecomposeChrono(UBound(DecomposeChrono)))
                     For I = 0 To UBound(DecomposeChrono) - 1
                        Debug.Print DecomposeChrono(I) & "_"
                        CrhonoPi = CrhonoPi & DecomposeChrono(I) & "_"
                    Next
                     CrhonoPi = Left(CrhonoPi, Len(CrhonoPi) - 1)
                     
                       DecomposeChrono = Split(txt7, "_")
                    IndiceOu = Val(DecomposeChrono(UBound(DecomposeChrono)))
                    ChronoOu = ""
                     For I = 0 To UBound(DecomposeChrono) - 1
                        Debug.Print DecomposeChrono(I) & "_"
                        ChronoOu = ChronoOu & DecomposeChrono(I) & "_"
                    Next
                     ChronoOu = Left(ChronoOu, Len(ChronoOu) - 1)
                     
                      DecomposeChrono = Split(txt8, "_")
                    IndiceLi = Val(DecomposeChrono(UBound(DecomposeChrono)))
                    ChronoLi = ""
                     For I = 0 To UBound(DecomposeChrono) - 1
                        Debug.Print DecomposeChrono(I) & "_"
                        ChronoLi = ChronoLi & DecomposeChrono(I) & "_"
                    Next
                     ChronoLi = Left(ChronoLi, Len(ChronoLi) - 1)
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
                    Sql = "UPDATE T_indiceProjet  SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1, "
                    Sql = Sql & "T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "',   "
                    Sql = Sql & "T_indiceProjet.PL = '" & ChronoPl & "', T_indiceProjet.PL_Indice = '" & IndicePl & "',  "
                    Sql = Sql & "T_indiceProjet.OU = '" & ChronoOu & "', T_indiceProjet.OU_Indice = '" & IndiceOu & "',  "
                    Sql = Sql & "T_indiceProjet.LI = '" & ChronoLi & "', T_indiceProjet.Li_Indice = '" & IndicePl & "',  "
                    Sql = Sql & "T_indiceProjet.ApprouveDate =" & MyReplaceDate(Date) & " "
                    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
                    Con.Execute Sql
'                End If
                
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
'                    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1,T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "', T_indiceProjet.OU = '" & Crhono & "', T_indiceProjet.OU_Indice = '" & Indice & "', T_indiceProjet.ApprouveDate = " & MyReplaceDate(Date) & " "
'                    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'                    Con.Execute Sql
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
'                    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1,T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "',T_indiceProjet.Li = '" & Crhono & "', T_indiceProjet.Li_Indice = '" & Indice & "', T_indiceProjet.ApprouveDate =" & MyReplaceDate(Date) & " "
'                    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'                    Con.Execute Sql
''                    subExporteXls Val(Me.Tag)
'                End If
'
'                Rs.Requery
'                PathPl = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![LiAutoCadSave])
'                'fileCopie
'                Set Fso = Nothing
'                PathDessin2 = PathDessin
'                PathPl2 = PathPl
'               Rs.Requery
''             Set Fso = New FileSystemObject
                 If IndexObj > 0 Then
                    For I2 = 1 To IndexObj
''                       CrhonoPi = ""
''
'
'                        PathPl = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![PlAutoCadSave])
'                        Racourci PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & CrhonoPi, "PL", Rs!PL, Val(Me.PiceName(I2).Tag), Val(IndicePi), Rs!PL_Indice, 1), "" & PathPl, "DWG"
'
'                        PathPl = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![OUAutoCadSave])
'                        Racourci PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & CrhonoPi, "OU", Rs!OU, Val(Me.PiceName(I2).Tag), Val(IndicePi), Rs!OU_Indice, 1), "" & PathPl, "DWG"
                        
'                     CrhonoPi = Left(CrhonoPi, Len(CrhonoPi) - 1)
'                     IndicePi = Val(DecomposeChrono(UBound(DecomposeChrono)))
'                        PatheFils = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & RsFille![LiAutoCadSave])
'
'                        Racourci PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & CrhonoPi, "LI", Rs!LI, Val(Me.PiceName(I2).Tag), Val(IndicePi), Rs!LI_Indice, 1), "" & PathPl, "XLS"
'                        RsFille.Requery
'                        PatheFils2 = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & RsFille![LiAutoCadSave])
'                       Set Fso = Nothing
'                       DoEvents
'                       If PiceEquipement(I2).Tag <> "CRE" Then
'                         KilVersionXX "" & PatheFils, "" & PathPl, True
'                       End If
                       
'                    Sql = "SELECT T_indiceProjet.LiAutoCadSave FROM T_indiceProjet WHERE T_indiceProjet.Id=" & PiceName(I2).Tag & ";"
'                        Set RsFille = Con.OpenRecordSet(Sql)
'                    For I = 0 To UBound(DecomposeChrono) - 1
'                                           Debug.Print DecomposeChrono(I) & "_"
'                                           CrhonoPi = CrhonoPi & DecomposeChrono(I) & "_"
'                                       Next
'                        Set RsFille = Con.CloseRecordSet(RsFille)
                         DecomposeChrono = Split(Me.Pièce(I2), "_")
                         CrhonoPi = ""
                         IndicePi = Val(DecomposeChrono(UBound(DecomposeChrono)))
                        For I = 0 To UBound(DecomposeChrono) - 1
                           Debug.Print DecomposeChrono(I) & "_"
                           CrhonoPi = CrhonoPi & DecomposeChrono(I) & "_"
                        Next
                        CrhonoPi = Left(CrhonoPi, Len(CrhonoPi) - 1)
                        IndicePi = Val(DecomposeChrono(UBound(DecomposeChrono)))
                        Rs.Requery
                            Sql = "UPDATE T_indiceProjet SET  T_indiceProjet.Version = 1, "
                            Sql = Sql & "T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "', "
                            Sql = Sql & "T_indiceProjet.PI = '" & CrhonoPi & "',  "
                            Sql = Sql & "T_indiceProjet.PI_Indice = '" & IndicePi & "',  "
                            Sql = Sql & "T_indiceProjet.PL = '" & Rs!PL & "',  "
                            Sql = Sql & "T_indiceProjet.PL_Indice = '" & Rs!PL_Indice & "',  "
                            Sql = Sql & "T_indiceProjet.[OU] = '" & Rs!OU & "',  "
                            Sql = Sql & "T_indiceProjet.OU_Indice = '" & Rs!OU_Indice & "',  "
                            Sql = Sql & "T_indiceProjet.Li = '" & Rs!LI & "',  "
                            Sql = Sql & "T_indiceProjet.LI_Indice = '" & Rs!LI_Indice & "',  "
                            Sql = Sql & "T_indiceProjet.ApprouveDate =  " & MyReplaceDate(Date) & " "
                            Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.PiceName(I2).Tag & ";"
                            Con.Execute Sql
                    Next
                End If
'            End If
          
 Set Fso = Nothing
 DoEvents
  
boolExecute = True
Noquite = False
Noquite = False
Me.Hide
'Dim boolCahnge As Boolean
'Dim I As Long
'Dim Sql As String
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
'boolCahnge = False
'
''Set FormBarGrah = Me
'If "" & CopieStrtxt7 <> txt5 And txt5 <> "" Then boolCahnge = True
'If "" & CopieStrtxt8 <> txt6 And txt6 <> "" Then boolCahnge = True
'If "" & CopieStrtxt9 <> txt7 And txt7 <> "" Then boolCahnge = True
'If "" & CopieStrtxt35 <> txt8 And txt8 <> "" Then boolCahnge = True
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
'        Sql = "UPDATE T_indiceProjet SET T_indiceProjet.ReffIndice = '" & MyReplace(Me.ReffIndice) & " ', T_indiceProjet.DNC = '" & MyReplace(Me.lstNc) & " ', T_indiceProjet.LIEC = '" & MyReplace(Me.lstLi) & "',Descripton='" & MyReplace(Me.DescIndice) & "' "
'        Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'        Con.Execute Sql
'
'
'        Sql = "SELECT T_indiceProjet.PlAutoCadSave, T_indiceProjet.OuAutoCadSave,  "
'        Sql = Sql & "T_indiceProjet.LiAutoCadSave, T_indiceProjet.PI, T_indiceProjet.PI_Indice,  "
'        Sql = Sql & "T_indiceProjet.Li, T_indiceProjet.PI, T_indiceProjet.PL,  "
'        Sql = Sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU], T_indiceProjet.Li,T_indiceProjet.Client ,T_indiceProjet.CleAc,T_indiceProjet.Version,T_indiceProjet.Ou_Indice,T_indiceProjet.LI_Indice "
'        Sql = Sql & "FROM T_indiceProjet "
'        Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'            Set Rs = Con.OpenRecordSet(Sql)
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
'                    Sql = "UPDATE T_indiceProjet  SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1,T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "',  T_indiceProjet.PL = '" & Crhono & "', T_indiceProjet.PL_Indice = '" & Indice & "', T_indiceProjet.ApprouveDate = Date() "
'                    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'                    Con.Execute Sql
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
'                    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1,T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "', T_indiceProjet.OU = '" & Crhono & "', T_indiceProjet.OU_Indice = '" & Indice & "', T_indiceProjet.ApprouveDate = Date() "
'                    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'                    Con.Execute Sql
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
'                    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "',T_indiceProjet.Version = 1,T_indiceProjet.PI = '" & CrhonoPi & "', T_indiceProjet.PI_Indice = '" & IndicePi & "',T_indiceProjet.Li = '" & Crhono & "', T_indiceProjet.Li_Indice = '" & Indice & "', T_indiceProjet.ApprouveDate = Date() "
'                    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
'                    Con.Execute Sql
''                    subExporteXls Val(Me.Tag)
'                End If
'
'                Rs.Requery
'                PathPl = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & Rs![LiAutoCadSave])
'                'fileCopie
'                Set Fso = Nothing
'                PathDessin2 = PathDessin
'                PathPl2 = PathPl
'               Rs.Requery
'             Set Fso = New FileSystemObject
'                 If IndexObj > 0 Then
'                    For I2 = 1 To IndexObj
'                       CrhonoPi = ""
'                    DecomposeChrono = Split(Me.Pièce(I2), "_")
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
'                        Sql = "SELECT T_indiceProjet.LiAutoCadSave FROM T_indiceProjet WHERE T_indiceProjet.Id=" & PiceName(I2).Tag & ";"
'                        Set RsFille = Con.OpenRecordSet(Sql)
'
'                        PatheFils = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & RsFille![LiAutoCadSave])
'
'                        Racourci PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & CrhonoPi, "LI", Rs!LI, Val(Me.PiceName(I2).Tag), Val(IndicePi), Rs!LI_Indice, 1), "" & PathPl, "XLS"
'                        RsFille.Requery
'                        PatheFils2 = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & RsFille![LiAutoCadSave])
'                       Set Fso = Nothing
'                       DoEvents
''                       If PiceEquipement(I2).Tag <> "CRE" Then
'                         KilVersionXX "" & PatheFils, "" & PathPl, True
''                       End If
'
'
'                        Set RsFille = Con.CloseRecordSet(RsFille)
'                            Sql = "UPDATE T_indiceProjet SET  T_indiceProjet.Version = 1, "
'                            Sql = Sql & "T_indiceProjet.Description = '" & MyReplace(Me.DescIndice) & "', "
'                            Sql = Sql & "T_indiceProjet.PI = '" & CrhonoPi & "',  "
'                            Sql = Sql & "T_indiceProjet.PI_Indice = '" & IndicePi & "',  "
'                            Sql = Sql & "T_indiceProjet.PL = '" & Rs!PL & "',  "
'                            Sql = Sql & "T_indiceProjet.PL_Indice = '" & Rs!PL_Indice & "',  "
'                            Sql = Sql & "T_indiceProjet.[OU] = '" & Rs!OU & "',  "
'                            Sql = Sql & "T_indiceProjet.OU_Indice = '" & Rs!OU_Indice & "',  "
'                            Sql = Sql & "T_indiceProjet.Li = '" & Rs!LI & "',  "
'                            Sql = Sql & "T_indiceProjet.LI_Indice = '" & Rs!LI_Indice & "',  "
'                            Sql = Sql & "T_indiceProjet.ApprouveDate =  Date() "
'                            Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.PiceName(I2).Tag & ";"
'                            Con.Execute Sql
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
End Sub

Private Sub CommandButton3_Click()
Noquite = False
Unload Me
'Me.Hide
End Sub

Private Sub LstFils_Click()

End Sub




Private Sub Form_Initialize()
IndexObj = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = Noquite
 IndexObj = 0
End Sub

Private Sub Form_Terminate()
IndexObj = 0
End Sub

Private Sub ReffIndice_Click()
Me.DescIndice = MuComment(Me.ReffIndice.List(Me.ReffIndice.ListIndex))
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


Private Sub Form_Activate()
Dim I As Long
Dim I2 As Long
Noquite = True
boolExecute = False
 For I = 1 To IndexObj
     Pièce(I).Clear
 Next
If IndexObj > 0 Then
    For I2 = 0 To txt5.ListCount - 1
    
        For I = 1 To IndexObj
        DoEvents
            Pièce(I).AddItem txt5.List(I2)
            If PiceName(I) = Pièce(I).List(Pièce(I).ListCount - 1) Then Pièce(I).ListIndex = Pièce(I).ListCount - 1
        Next
    Next
    End If
End Sub

Private Sub ChargementFille(Equipement As String, PI As String, Id As Long, MyTag As String)
IndexObj = IndexObj + 1

Load PiceEquipement(IndexObj)
PiceEquipement(IndexObj).Top = PiceEquipement(0).Top + (315 * (IndexObj - 1))
PiceEquipement(IndexObj).Tag = MyTag
PiceEquipement(IndexObj).Visible = True
PiceEquipement(IndexObj) = Equipement

Load Pièce(IndexObj)
Pièce(IndexObj).Top = Pièce(0).Top + (315 * (IndexObj - 1))
Pièce(IndexObj).Visible = True

Load PiceName(IndexObj)
PiceName(IndexObj) = PI
PiceName(IndexObj).Tag = CStr(Id)

PiceName(IndexObj).Top = PiceName(0).Top + (315 * (IndexObj - 1))
PiceName(IndexObj).Visible = True
Me.EnvelopePice.Height = 495 + (315 * (IndexObj - 1))
If IndexObj = 7 Then VScroll1.Visible = True
DoEvents
End Sub





Private Sub VScroll1_Change()
If Me.VScroll1.Value = 0 Then
    EnvelopePice.Top = 120
Else
EnvelopePice.Top = 120 + (Me.VScroll1.Value * (-1 * IndexObj))
End If
End Sub
