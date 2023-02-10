VERSION 5.00
Object = "{79647E82-6BF1-4435-B9A3-02ADECF7452D}#1.0#0"; "Autocable_R_Ocx.ocx"
Begin VB.Form ConeverEtudeCsv 
   Caption         =   "VARDKES Automation & Co:"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   ControlBox      =   0   'False
   Icon            =   "ConeverEtudeCsv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Quitter"
      Height          =   615
      Left            =   7440
      TabIndex        =   32
      Top             =   6120
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   31
      Text            =   "Combo1"
      Top             =   240
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exécuter"
      Height          =   615
      Left            =   4260
      TabIndex        =   29
      Top             =   6120
      Width           =   2055
   End
   Begin AutocableOcx.RecherAutocable RecherAutocable1 
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   6120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Database        =   "\\autocable\Autocable Access\Autocable.mdb"
      Filtre          =   ""
   End
   Begin VB.Label Label151 
      Caption         =   "Bloc d'empreinte"
      Height          =   315
      Left            =   240
      TabIndex        =   30
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Val8 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   28
      Top             =   3240
      Width           =   6975
   End
   Begin VB.Label Label8 
      Height          =   315
      Left            =   240
      TabIndex        =   27
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label12 
      Height          =   315
      Left            =   240
      TabIndex        =   26
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label11 
      Height          =   315
      Left            =   240
      TabIndex        =   25
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label10 
      Height          =   315
      Left            =   240
      TabIndex        =   24
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label14 
      Height          =   315
      Left            =   240
      TabIndex        =   23
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label9 
      Height          =   315
      Left            =   240
      TabIndex        =   22
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label13 
      Height          =   315
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label7 
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label6 
      Height          =   315
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label5 
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label4 
      Height          =   315
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Val14 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   13
      Top             =   5400
      Width           =   6975
   End
   Begin VB.Label Val13 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Top             =   5040
      Width           =   6975
   End
   Begin VB.Label Val12 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   11
      Top             =   4680
      Width           =   6975
   End
   Begin VB.Label Val11 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   10
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Label Val10 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   3960
      Width           =   6975
   End
   Begin VB.Label Val9 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   3600
      Width           =   6975
   End
   Begin VB.Label Val7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   2880
      Width           =   6975
   End
   Begin VB.Label Val6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label Val5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Label Val4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Label Val3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label Val2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Val1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
End
Attribute VB_Name = "ConeverEtudeCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Dim TXT As String
Dim txt2 As String
Dim ColecLiason As New Collection
Dim ICol As Long
Dim SqlEpissur As String
Dim RS As Recordset
Dim Rsliai As Recordset
Dim SplitEquipement
Dim Sql As String
Dim CloseWere As String
Dim CloseWere2 As String
Dim I As Long
Set TableauOnglet = Nothing
Set TableauOnglet = New Collection
Dim PathPl As String
If Me.Combo1.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner un Bloc d'empreinte.", vbOKOnly, "VARDKES Automation & Co"
    Exit Sub
End If
If Trim("" & Me.Tag) = "" Then
    MsgBox "Vous devez sélectionner une Pièce.", vbOKOnly, "VARDKES Automation & Co"
    Exit Sub
End If

For I = 1 To 11
    TXT = TXT & Me.Controls("Label" & CStr(I)).Caption & ";"
    TXT = TXT & Replace(Me.Controls("Val" & CStr(I)).Caption, ";", " ") & ";"
Next
TXT = TXT & UCase("BLOCEmpreinte;") & Me.Combo1.List(Me.Combo1.ListIndex) & vbCrLf
Sql = "SELECT T_indiceProjet.Pere From T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set RS = Con.OpenRecordSet(Sql)
If RS!Pere <> 0 Then
 Me.Tag = RS!Pere
End If
SplitEquipement = Split(Me.Val3.Caption & ";", ";")
Sql = "SELECT T_Dossier_Contrôle.App,T_Dossier_Contrôle.App2 "
Sql = Sql & "From T_Dossier_Contrôle "

 

Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & Me.Tag & "  AND ( "
CloseWere = "t_Dossier_Contrôle.OPTION Like'Tous%' OR t_Dossier_Contrôle.OPTION Like'ALL%' OR "
For I = 0 To UBound(SplitEquipement)
If Trim("" & SplitEquipement(I)) <> "" Then
    CloseWere = CloseWere & "t_Dossier_Contrôle.OPTION Like'" & SplitEquipement(I) & "%' OR "
End If
Next
CloseWere = Trim(CloseWere)
CloseWere = Left(CloseWere, Len(CloseWere) - 2)
Sql = Sql & CloseWere & ") "
Sql = Sql & "GROUP BY T_Dossier_Contrôle.App,T_Dossier_Contrôle.App2 "

'sql = sql & "GROUP BY T_Dossier_Contrôle.LIAI, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT, T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI, T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[POS-OUT2], T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2], T_Dossier_Contrôle.OPTION, T_Dossier_Contrôle.Id_IndiceProjet, Connecteurs.[O/N], Connecteurs_1.[O/N]"

Sql = "SELECT T_Dossier_Contrôle.Onglet, Count(T_Dossier_Contrôle.Onglet) AS CompteDeOnglet "
Sql = Sql & "From T_Dossier_Contrôle "

 

Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & Me.Tag & "  AND ( "
CloseWere = "t_Dossier_Contrôle.OPTION Like'Tous%' OR t_Dossier_Contrôle.OPTION Like'ALL%' OR "
For I = 0 To UBound(SplitEquipement)
If Trim("" & SplitEquipement(I)) <> "" Then
    CloseWere = CloseWere & "t_Dossier_Contrôle.OPTION Like'" & SplitEquipement(I) & "%' OR "
End If
Next
CloseWere = Trim(CloseWere)
CloseWere = Left(CloseWere, Len(CloseWere) - 2)
Sql = Sql & CloseWere & ") "
Sql = Sql & "GROUP BY T_Dossier_Contrôle.Onglet "
Sql = Sql & "ORDER BY Count(T_Dossier_Contrôle.Onglet) DESC;"

Set RS = Con.OpenRecordSet(Sql)
While RS.EOF = False
On Error Resume Next
'
' Public CloseWere As String
'Dim Sql As String
'Public Id_IndiceProjet As Long

TableauOnglet("" & RS!Onglet).Onglet = "" & RS("Onglet")
If Err Then
Err.Clear
 AjouterClassOnglet TableauOnglet, "" & RS("Onglet")
End If
'
'ColecLiason("" & Rs!APP).CloseWere = CloseWere
'ColecLiason("" & Rs!APP).Id_IndiceProjet = Me.Tag
'
'ColecLiason("" & Rs!App2).APP = "" & Rs("app2")
'If Err Then
'Err.Clear
' AjouterClass ColecLiason, "" & Rs("app2")
'End If
'
'ColecLiason("" & Rs!App2).CloseWere = CloseWere
'ColecLiason("" & Rs!App2).Id_IndiceProjet = Me.Tag
     RS.MoveNext
Wend
Sql = "SELECT T_Dossier_Contrôle.TEINT, T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.LIAI, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.VOI, T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[POS-OUT2], T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2] "
Sql = Sql & "From T_Dossier_Contrôle "

 

Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & Me.Tag & "  AND ( "
CloseWere = "t_Dossier_Contrôle.OPTION Like'Tous%' OR t_Dossier_Contrôle.OPTION Like'ALL%' OR "
For I = 0 To UBound(SplitEquipement)
If Trim("" & SplitEquipement(I)) <> "" Then
    CloseWere = CloseWere & "t_Dossier_Contrôle.OPTION Like'" & SplitEquipement(I) & "%' OR "
End If
Next
CloseWere = Trim(CloseWere)
CloseWere = Left(CloseWere, Len(CloseWere) - 2)
Sql = Sql & CloseWere & ") "
Sql = Sql & "GROUP BY T_Dossier_Contrôle.TEINT, T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.LIAI, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.VOI, T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[POS-OUT2], T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2] "
Sql = Sql & "ORDER BY T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.VOI,  "
Sql = Sql & "T_Dossier_Contrôle.[POS-OUT2], T_Dossier_Contrôle.VOI2; "
Set RS = Con.OpenRecordSet(Sql)
While RS.EOF = False
    TableauOnglet(RS("Onglet")).AjouterClass RS
    RS.MoveNext
Wend

SqlEpissur = "SELECT T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.LIAI, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
SqlEpissur = SqlEpissur & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI, T_Dossier_Contrôle.[REF CONNECTEUR],  "
SqlEpissur = SqlEpissur & "T_Dossier_Contrôle.[POS-OUT2], T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2], T_Dossier_Contrôle.FIL "
SqlEpissur = SqlEpissur & "From T_Dossier_Contrôle "
SqlEpissur = SqlEpissur & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & Me.Tag & " AND (" & CloseWere & ")"
SqlEpissur = SqlEpissur & "Group By T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.LIAI, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
SqlEpissur = SqlEpissur & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI, T_Dossier_Contrôle.[REF CONNECTEUR],  "
SqlEpissur = SqlEpissur & "T_Dossier_Contrôle.[POS-OUT2], T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2], T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.Id "
SqlEpissur = SqlEpissur & "ORDER BY T_Dossier_Contrôle.Id;"



SqlEpissur = "SELECT Connecteurs.CODE_APP, Connecteurs.[O/N] "
SqlEpissur = SqlEpissur & "From Connecteurs "
SqlEpissur = SqlEpissur & "Where Connecteurs.Id_IndiceProjet = " & Me.Tag & "  and  "
SqlEpissur = SqlEpissur & "Connecteurs.[O/N]=True "
SqlEpissur = SqlEpissur & "GROUP BY Connecteurs.CODE_APP, Connecteurs.[O/N];"


Set RS = Con.OpenRecordSet(SqlEpissur)
While RS.EOF = False
 TableauOnglet(RS("CODE_APP")).epissure = True
'ColecLiason("" & RS!Onglet).AddEpisure RS
'ColecLiason("" & Rs!App2).AddEpisure Rs
 RS.MoveNext
Wend
For I = 1 To TableauOnglet.Count
    If TableauOnglet(I).epissure = True Then
       TableauOnglet(I).ReplaceEpissure TableauOnglet
    End If
Next

For I = 1 To TableauOnglet.Count
TableauOnglet(I).RetourneTableau TXT

Next
Set TableauPath = funPath
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set RS = Con.OpenRecordSet(Sql)
If RS.EOF = False Then
     PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & RS!Client, "" & RS!CleAc, "" & RS!Pieces, "Li", RS.Fields("Li"), Val(Me.Tag), RS.Fields("PI_Indice"), RS.Fields("LI_Indice"), RS!Version)
     PathPl = PathPl & ".csv"
     Debug.Print PathPl
End If
I = FreeFile
Open PathPl For Output As #I
    Print #I, TXT  ' Écrit le texte.
    Close #I   ' Ferme le fichier.
Set ColecLiason = Nothing
MsgBox "Traitement terminé.", vbOKOnly, "VARDKES Automation & Co"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Fichier As String
Me.RecherAutocable1.Database = Con.Fichier
Fichier = Dir(APP.Path & "\Map\*.map")
Me.Combo1.Clear
While Fichier <> ""
Me.Combo1.AddItem Replace(Fichier, ".map", "")
Fichier = Dir
Wend
End Sub

Private Sub RecherAutocable1_Action(Tableau_Valeur As Variant, Annuler As Variant)

Me.Tag = Tableau_Valeur(15, 1)
Dim I As Long
For I = 1 To 14
Me.Controls("Label" & CStr(I)).Caption = Tableau_Valeur(I, 0)
Me.Controls("Val" & CStr(I)).Caption = Tableau_Valeur(I, 1)
Next

End Sub

