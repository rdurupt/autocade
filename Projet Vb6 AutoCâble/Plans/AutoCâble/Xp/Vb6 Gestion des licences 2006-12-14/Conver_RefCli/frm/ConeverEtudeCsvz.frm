VERSION 5.00
Object = "{79647E82-6BF1-4435-B9A3-02ADECF7452D}#1.0#0"; "Autocable_R_Ocx.ocx"
Begin VB.Form ConeverEtudeCsv 
   Caption         =   "Convertisseur de références Client :"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   ControlBox      =   0   'False
   Icon            =   "ConeverEtudeCsvz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   1080
      Width           =   6975
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   720
      Width           =   6975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00FF00FF&
      Height          =   315
      ItemData        =   "ConeverEtudeCsvz.frx":030A
      Left            =   2520
      List            =   "ConeverEtudeCsvz.frx":030C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   360
      Width           =   6975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Quitter"
      Height          =   615
      Left            =   7320
      TabIndex        =   30
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exécuter"
      Height          =   615
      Left            =   4140
      TabIndex        =   29
      Top             =   7200
      Width           =   2055
   End
   Begin AutocableOcx.RecherAutocable RecherAutocable1 
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   7200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Filtre          =   ""
   End
   Begin VB.Label Label17 
      Caption         =   "Référence par  défaut"
      Height          =   315
      Left            =   240
      TabIndex        =   33
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label16 
      Caption         =   "Nouvelle référence"
      Height          =   315
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Référence actuelle"
      Height          =   315
      Left            =   240
      TabIndex        =   31
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Val8 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   28
      Top             =   4080
      Width           =   6975
   End
   Begin VB.Label Label8 
      Height          =   315
      Left            =   240
      TabIndex        =   27
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label14 
      Height          =   315
      Left            =   240
      TabIndex        =   26
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label13 
      Height          =   315
      Left            =   240
      TabIndex        =   25
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label12 
      Height          =   315
      Left            =   240
      TabIndex        =   24
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label11 
      Height          =   315
      Left            =   240
      TabIndex        =   23
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label10 
      Height          =   315
      Left            =   240
      TabIndex        =   22
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label9 
      Height          =   315
      Left            =   240
      TabIndex        =   21
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label7 
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label6 
      Height          =   315
      Left            =   240
      TabIndex        =   19
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label4 
      Height          =   315
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Val14 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   13
      Top             =   6240
      Width           =   6975
   End
   Begin VB.Label Val13 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Top             =   5880
      Width           =   6975
   End
   Begin VB.Label Val12 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   11
      Top             =   5520
      Width           =   6975
   End
   Begin VB.Label Val11 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   10
      Top             =   5160
      Width           =   6975
   End
   Begin VB.Label Val10 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   4800
      Width           =   6975
   End
   Begin VB.Label Val9 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   4440
      Width           =   6975
   End
   Begin VB.Label Val7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   3720
      Width           =   6975
   End
   Begin VB.Label Val6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   3360
      Width           =   6975
   End
   Begin VB.Label Val5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Label Val4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   2640
      Width           =   6975
   End
   Begin VB.Label Val3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   2280
      Width           =   6975
   End
   Begin VB.Label Val2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   6975
   End
   Begin VB.Label Val1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
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
Dim Txt As String
Dim txt2 As String
Dim ColecLiason As New Collection
Dim ICol As Long
Dim SqlEpissur As String
Dim Rs As Recordset
Dim Rsliai As Recordset
Dim SplitConnecteur
Dim Sql As String
Dim CloseWere As String
Dim CloseWere2 As String
Dim I As Long
Dim PathPl As String
Dim RefActuelle As String
Dim NouvelleRef As String
Dim RefDefault As String
Dim RsConnecteur As Recordset
Dim IdFils As Long
If Me.Combo1.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner la Référence actuelle", vbOKOnly, "Conversion CLI"
    Exit Sub
End If
If Me.Combo2.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner la Nouvelle référence", vbOKOnly, "Conversion CLI"
    Exit Sub
End If
If Me.Combo3.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner la Référence actuelle", vbOKOnly, "Conversion CLI"
    Exit Sub
End If
If Trim("" & Me.Tag) = "" Then
    MsgBox "Vous devez sélectionner une Pièce.", vbOKOnly, "Conversion CLI"
    Exit Sub
End If

For I = 1 To 11
    Txt = Txt & Me.Controls("Label" & CStr(I)).Caption & ";"
    Txt = Txt & Replace(Me.Controls("Val" & CStr(I)).Caption, ";", " ") & ";"
Next

Sql = "SELECT T_indiceProjet.Pere From T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs!Pere <> 0 Then
    IdFils = Me.Tag
 Me.Tag = Rs!Pere
End If
Sql = "SELECT con_FieldDefs.FieldName "
Sql = Sql & "FROM con_FieldDefs IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "' "
Sql = Sql & "WHERE con_FieldDefs.FieldAlias='" & MyReplace("" & Me.Combo1.List(Me.Combo1.ListIndex)) & "' "
Sql = Sql & "AND con_FieldDefs.FieldOrder<>0;"
Set Rs = Con.OpenRecordSet(Sql)

If Rs.EOF = False Then
    RefActuelle = "" & Rs(0)
End If

Sql = "SELECT con_FieldDefs.FieldName "
Sql = Sql & "FROM con_FieldDefs IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "' "
Sql = Sql & "WHERE con_FieldDefs.FieldAlias='" & MyReplace("" & Me.Combo2.List(Me.Combo2.ListIndex)) & "' "
Sql = Sql & "AND con_FieldDefs.FieldOrder<>0;"
Set Rs = Con.OpenRecordSet(Sql)

If Rs.EOF = False Then
    NouvelleRef = "" & Rs(0)
End If

Sql = "SELECT con_FieldDefs.FieldName "
Sql = Sql & "FROM con_FieldDefs IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "' "
Sql = Sql & "WHERE con_FieldDefs.FieldAlias='" & MyReplace("" & Me.Combo3.List(Me.Combo3.ListIndex)) & "' "
Sql = Sql & "AND con_FieldDefs.FieldOrder<>0;"
Set Rs = Con.OpenRecordSet(Sql)

If Rs.EOF = False Then
    RefDefault = "" & Rs(0)
End If

'***********************************************************************************************************************
'*                                          Supprime le contenu des tables Ecart.                                     *
Sql = "DELETE T_Critères_Ecart.*  FROM T_Critères_Ecart "
Sql = Sql & "where T_Critères_Ecart.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql


Sql = "DELETE Ligne_Tableau_fils_Ecart.*  FROM Ligne_Tableau_fils_Ecart "
Sql = Sql & "where Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "DELETE Connecteurs_Ecart.* FROM Connecteurs_Ecart "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "DELETE Nota_Ecart.* FROM Nota_Ecart "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "DELETE Composants_Ecart.* FROM Composants_Ecart "
Sql = Sql & "WHERE Composants_Ecart.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "DELETE T_Noeuds_Ecart.* FROM T_Noeuds_Ecart "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                                        Sauvegarde les anciennes valeurs                                            *

Sql = "INSERT INTO T_Critères_Ecart SELECT T_Critères.* FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO Connecteurs_Ecart SELECT Connecteurs.* FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO Nota_Ecart SELECT Nota.* FROM Nota "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO Composants_Ecart SELECT Composants.* FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO Ligne_Tableau_fils_Ecart SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO T_Noeuds_Ecart SELECT T_Noeuds.* FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql


'MAJ Connecteur
'******************************************************************************************************************
Set Rs = Con.CloseRecordSet(Rs)
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a, con_contacts." & NouvelleRef & " as b, con_contacts." & RefDefault & " as c "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm INNER JOIN Connecteurs ON MyForm.a = Connecteurs.CONNECTEUR  "
Sql = Sql & "SET Connecteurs.CONNECTEUR = [b] "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Me.Tag & "  "
Sql = Sql & "AND MyForm.b<>''  "
Sql = Sql & "AND Connecteurs.ACTIVER=True;"
'
'
Con.Execute Sql

Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a, con_contacts." & NouvelleRef & " as b, con_contacts." & RefDefault & " as c "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm INNER JOIN Connecteurs ON MyForm.a = Connecteurs.CONNECTEUR  "
Sql = Sql & "SET Connecteurs.CONNECTEUR = [c] "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Me.Tag & "  "
Sql = Sql & "AND MyForm.c<>''  "
Sql = Sql & "AND Connecteurs.ACTIVER=True;"

Con.Execute Sql
'sql = "SELECT { fn Spilt§([Connecteurs].[CONNECTEUR]) as toto "
'sql = sql & "From Connecteurs "
'sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & Me.Tag & "  "
'sql = sql & "AND Connecteurs.ACTIVER=True "
'sql = sql & " AND Connecteurs.[O/N]=False;"
'Set Rs = Con.OpenRecordSet(sql)
'While Rs.EOF = False
'SplitConnecteur = Split("" & Rs!CONNECTEUR & "§", "§")
'
'    sql = "SELECT con_contacts." & RefActuelle & ", con_contacts." & NouvelleRef & ", con_contacts." & RefDefault & " "
'sql = sql & "FROM con_contacts IN '"
'sql = sql & TableauPath("Eb_CONNECTEURS")
'
'sql = sql & "' where con_contacts." & RefActuelle & "='" & SplitConnecteur(0) & "';"
'Set RsConnecteur = Con.OpenRecordSet(sql)
'If RsConnecteur.EOF = False Then
'    If Trim("" & RsConnecteur(NouvelleRef)) <> "" Then
'        Rs!CONNECTEUR = Replace("" & Rs!CONNECTEUR, SplitConnecteur(0), "" & RsConnecteur(NouvelleRef))
'    Else
'        If Trim("" & RsConnecteur(RefDefault)) <> "" Then
'            Rs!CONNECTEUR = Replace("" & Rs!CONNECTEUR, SplitConnecteur(0), "" & RsConnecteur(RefDefault))
'        End If
'    End If
'Else
'     sql = "SELECT con_contacts." & RefActuelle & ", con_contacts." & NouvelleRef & ", con_contacts." & RefDefault & " "
'    sql = sql & "FROM con_contacts IN '"
'    sql = sql & TableauPath("Eb_CONNECTEURS")
'    SplitConnecteur = Split("" & Rs!CONNECTEUR & "§", "§")
'    sql = sql & "' where con_contacts." & RefDefault & "='" & SplitConnecteur(0) & "';"
'    Set RsConnecteur = Con.OpenRecordSet(sql)
'    If RsConnecteur.EOF = False Then
'        If Trim("" & RsConnecteur(NouvelleRef)) <> "" Then
'            Rs!CONNECTEUR = Replace("" & Rs!CONNECTEUR, SplitConnecteur(0), "" & RsConnecteur(NouvelleRef))
'        Else
'            If Trim("" & RsConnecteur(RefDefault)) <> "" Then
'                Rs!CONNECTEUR = Replace("" & Rs!CONNECTEUR, SplitConnecteur(0), "" & RsConnecteur(RefDefault))
'            End If
'        End If
'    End If
'End If
'Rs.Update
'    Rs.MoveNext
'Wend

''************************************************************************************************************************
''MAJ Dossier de contrôle
''**********************************************************************************************************************
''REF CONNECTEUR
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm   "
Sql = Sql & "INNER JOIN T_Dossier_Contrôle ON MyForm.a = T_Dossier_Contrôle.[REF CONNECTEUR]   "
Sql = Sql & "SET T_Dossier_Contrôle.[REF CONNECTEUR] = [b]  "
Sql = Sql & "WHERE MyForm.b<>''  "
Sql = Sql & "AND T_Dossier_Contrôle.ACTIVER=True   "
Sql = Sql & "AND T_Dossier_Contrôle.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql



'
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm   "
Sql = Sql & "INNER JOIN T_Dossier_Contrôle ON MyForm.a = T_Dossier_Contrôle.[REF CONNECTEUR]   "
Sql = Sql & "SET T_Dossier_Contrôle.[REF CONNECTEUR] = [c]  "
Sql = Sql & "WHERE MyForm.c<>''  "
Sql = Sql & "AND T_Dossier_Contrôle.ACTIVER=True   "
Sql = Sql & "AND T_Dossier_Contrôle.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql
'
''********************************************************************************
''Maj Dossier de Fab
'
''********************************************************************************
''REF CONNECTEUR
'
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm INNER JOIN T_Dossier_Fabrication ON MyForm.a =  "
Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR]  "
Sql = Sql & "SET T_Dossier_Fabrication.[REF CONNECTEUR] = [b] "
Sql = Sql & "WHERE MyForm.b<>''  "
Sql = Sql & "AND T_Dossier_Fabrication.ACTIVER=True  "
Sql = Sql & "AND T_Dossier_Fabrication.Id_IndiceProjet=" & Me.Tag & ";"

Con.Execute Sql
'
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm INNER JOIN T_Dossier_Fabrication ON MyForm.a =  "
Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR]  "
Sql = Sql & "SET T_Dossier_Fabrication.[REF CONNECTEUR] = [c] "
Sql = Sql & "WHERE MyForm.c<>''  "
Sql = Sql & "AND T_Dossier_Fabrication.ACTIVER=True  "
Sql = Sql & "AND T_Dossier_Fabrication.Id_IndiceProjet=" & Me.Tag & ";"

Con.Execute Sql
'
''************************************************************************************************************************
''MAJ Dossier de contrôle
''**********************************************************************************************************************
'REF CONNECTEUR2
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm   "
Sql = Sql & "INNER JOIN T_Dossier_Contrôle ON MyForm.a = T_Dossier_Contrôle.[REF CONNECTEUR2]   "
Sql = Sql & "SET T_Dossier_Contrôle.[REF CONNECTEUR2] = [b]  "
Sql = Sql & "WHERE MyForm.b<>''   "
Sql = Sql & "AND T_Dossier_Contrôle.ACTIVER=True   "
Sql = Sql & "AND T_Dossier_Contrôle.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql
'
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm   "
Sql = Sql & "INNER JOIN T_Dossier_Contrôle ON MyForm.a = T_Dossier_Contrôle.[REF CONNECTEUR2]   "
Sql = Sql & "SET T_Dossier_Contrôle.[REF CONNECTEUR2] = [c]  "
Sql = Sql & "WHERE MyForm.c<>''  "
Sql = Sql & "AND T_Dossier_Contrôle.ACTIVER=True   "
Sql = Sql & "AND T_Dossier_Contrôle.Id_IndiceProjet=" & Me.Tag & ";"

Con.Execute Sql
'
''********************************************************************************
''Maj Dossier de Fab
'
''********************************************************************************
''REF CONNECTEUR2

Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm INNER JOIN T_Dossier_Fabrication ON MyForm.a =  "
Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR2]  "
Sql = Sql & "SET T_Dossier_Fabrication.[REF CONNECTEUR2] = [b] "
Sql = Sql & "WHERE MyForm.b<>''  "
Sql = Sql & "AND T_Dossier_Fabrication.ACTIVER=True  "
Sql = Sql & "AND T_Dossier_Fabrication.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql
'
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm INNER JOIN T_Dossier_Fabrication ON MyForm.a =  "
Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR2]  "
Sql = Sql & "SET T_Dossier_Fabrication.[REF CONNECTEUR2] = [c] "
Sql = Sql & "WHERE MyForm.c<>''  "
Sql = Sql & "AND T_Dossier_Fabrication.ACTIVER=True  "
Sql = Sql & "AND T_Dossier_Fabrication.Id_IndiceProjet=" & Me.Tag & ";"

Con.Execute Sql

'
'
'
''********************************************************************************
''Maj tableau de fils
'
''********************************************************************************
''Ref Connecteur
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm  "
Sql = Sql & "INNER JOIN Ligne_Tableau_fils  "
Sql = Sql & "ON MyForm.a = Ligne_Tableau_fils.[Ref Connecteur]  "
Sql = Sql & "SET Ligne_Tableau_fils.[Ref Connecteur] = [b] "
Sql = Sql & "WHERE MyForm.b<>''  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True  "
Sql = Sql & "AND Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql
'
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a, con_contacts." & NouvelleRef & " as b, con_contacts." & RefDefault & " as c "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm  "
Sql = Sql & "INNER JOIN Ligne_Tableau_fils  "
Sql = Sql & "ON MyForm.a = Ligne_Tableau_fils.[Ref Connecteur]  "
Sql = Sql & "SET Ligne_Tableau_fils.[Ref Connecteur] = [c] "
Sql = Sql & "WHERE MyForm.c<>''  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True  "
Sql = Sql & "AND Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql
'
'Ref Connecteur2
Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a, con_contacts." & NouvelleRef & " as b, con_contacts." & RefDefault & " as c "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm  "
Sql = Sql & "INNER JOIN Ligne_Tableau_fils  "
Sql = Sql & "ON MyForm.a = Ligne_Tableau_fils.[Ref Connecteur2]  "
Sql = Sql & "SET Ligne_Tableau_fils.[Ref Connecteur2] = [b] "
Sql = Sql & "WHERE MyForm.b<>''  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True  "
Sql = Sql & "AND Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a, con_contacts." & NouvelleRef & " as b, con_contacts." & RefDefault & " as c "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm  "
Sql = Sql & "INNER JOIN Ligne_Tableau_fils  "
Sql = Sql & "ON MyForm.a= Ligne_Tableau_fils.[Ref Connecteur2]  "
Sql = Sql & "SET Ligne_Tableau_fils.[Ref Connecteur2] = [c] "
Sql = Sql & "WHERE MyForm.c<>''  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True  "
Sql = Sql & "AND Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

'
''********************************************************************************
''Maj NomeclatureConnecteurs
'
''********************************************************************************

Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm INNER JOIN NomeclatureConnecteurs ON  "
Sql = Sql & "MyForm.a = NomeclatureConnecteurs.Connecteur  "
Sql = Sql & "SET NomeclatureConnecteurs.Connecteur = [b] "
Sql = Sql & "WHERE MyForm.b<>''  "
Sql = Sql & "AND NomeclatureConnecteurs.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a , con_contacts." & NouvelleRef & " as b , con_contacts." & RefDefault & " as c  "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm INNER JOIN NomeclatureConnecteurs ON  "
Sql = Sql & "MyForm.a = NomeclatureConnecteurs.Connecteur  "
Sql = Sql & "SET NomeclatureConnecteurs.Connecteur = [c] "
Sql = Sql & "WHERE MyForm.c<>''  "
Sql = Sql & "AND NomeclatureConnecteurs.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql
'
''********************************************************************************
''Maj NomenclaturFinal
'
''********************************************************************************
'

Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a, con_contacts." & NouvelleRef & " as b, con_contacts." & RefDefault & " as c "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm  "
Sql = Sql & "INNER JOIN NomenclaturFinal  "
Sql = Sql & "ON MyForm.a = NomenclaturFinal.Ref SET NomenclaturFinal.Ref = [b] "
Sql = Sql & "WHERE MyForm.b<>''  "
Sql = Sql & "AND NomenclaturFinal.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Sql = "UPDATE (SELECT con_contacts." & RefActuelle & " as a, con_contacts." & NouvelleRef & " as b, con_contacts." & RefDefault & " as c "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm  "
Sql = Sql & "INNER JOIN NomenclaturFinal  "
Sql = Sql & "ON MyForm.a = NomenclaturFinal.Ref SET NomenclaturFinal.Ref = [c] "
Sql = Sql & "WHERE MyForm.c<>''  "
Sql = Sql & "AND NomenclaturFinal.Id_IndiceProjet=" & Me.Tag & ";"
Con.Execute Sql

Dim MyExcel As New Excel.Application
MajEcart Me.Tag, IdFils, MyExcel


MsgBox "Traitement terminé.", vbOKOnly, "Conversion CLI"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String
Dim Rs As Recordset
NmJob = LaodJob
Me.RecherAutocable1.Database = ADO_Fichier
 Set TableauPath = funPath
Sql = "SELECT con_FieldDefs.FieldName, con_FieldDefs.FieldAlias "
Sql = Sql & "FROM con_FieldDefs IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "' "
Sql = Sql & "WHERE con_FieldDefs.FieldOrder<>0  "
Sql = Sql & "AND con_FieldDefs.FieldAttribut Like 'txt%';"
Me.Combo1.Clear
Me.Combo2.Clear
Me.Combo3.Clear
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Me.Combo1.AddItem Rs!FieldAlias
    Me.Combo2.AddItem Rs!FieldAlias
    Me.Combo3.AddItem Rs!FieldAlias
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
End Sub

Private Sub RecherAutocable1_Action(Tableau_Valeur As Variant, Annuler As Variant)

Me.Tag = Tableau_Valeur(15, 1)
Dim I As Long
For I = 1 To 14
Me.Controls("Label" & CStr(I)).Caption = Tableau_Valeur(I, 0)
Me.Controls("Val" & CStr(I)).Caption = Tableau_Valeur(I, 1)
Next

End Sub


