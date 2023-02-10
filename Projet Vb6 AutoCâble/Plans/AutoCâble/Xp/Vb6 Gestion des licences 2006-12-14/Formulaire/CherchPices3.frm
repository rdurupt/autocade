VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form UserForm5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importer Archives :"
   ClientHeight    =   12645
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   16890
   Icon            =   "CherchPices3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12645
   ScaleWidth      =   16890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   10110
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   16890
      HTMLURL         =   ""
      HTMLData        =   $"CherchPices3.frx":08CA
      DataType        =   "HTMLDATA"
      AutoFit         =   -1  'True
      DisplayColHeaders=   0   'False
      DisplayGridlines=   -1  'True
      DisplayHorizontalScrollBar=   -1  'True
      DisplayRowHeaders=   0   'False
      DisplayTitleBar =   0   'False
      DisplayToolbar  =   0   'False
      DisplayVerticalScrollBar=   -1  'True
      EnableAutoCalculate=   -1  'True
      EnableEvents    =   -1  'True
      MoveAfterReturn =   0   'False
      MoveAfterReturnDirection=   0
      RightToLeft     =   0   'False
      ViewableRange   =   "1:65536"
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   14640
      TabIndex        =   7
      Top             =   0
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "CherchPices3.frx":18C3
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   11640
      Width           =   3015
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "&Valider"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   11760
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   12360
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   195
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Archive"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label28 
      Caption         =   "Légendes :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Noquite As Boolean
Private Sub CommandButton1_Click()
Dim msg As String
Dim MyRange
Dim IndiceProjet As Long
Dim Id_Pieces As Long
Dim IdProjet As Long
Dim Sql As String
Dim LibPice As String
Dim Rs As Recordset
Dim IdPIndice As Long
If MsgBox("vous réimporter les enregistrements Archivés", vbYesNo + vbQuestion, "Importer Archives") = vbNo Then Exit Sub
RazFiltreEditExcel Me.Spreadsheet1
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
ProgressBar1.Value = 0
ProgressBar1.Max = MyRange.Rows.Count
For I = 2 To MyRange.Rows.Count
ProgressBar1.Value = I
    If UCase(MyRange(I, 1)) <> 0 Then
         IndiceProjet = CInt(Me.Spreadsheet1.Cells(I, 16))
            Sql = "UPDATE T_indiceProjet SET T_indiceProjet.IdStatus = 3 "
            Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & " "
            Sql = Sql & "OR T_indiceProjet.Pere=" & IndiceProjet & " ;"
            Con.Execute Sql
'         sql = "SELECT   Archive_T_indiceProjet.Id "
'        sql = sql & "FROM  Archive_T_indiceProjet "
'        sql = sql & "WHERE [PI] & '_' & [PI_Indice]='" & Replace(Me.Spreadsheet1.Cells(I, 6), " ", "") & "'"
'        Set Rs = Con.OpenRecordSet(sql)
'         If Rs.EOF = False Then
'            IdPIndice = Rs!Id
'         Else
'         IdPIndice = 0
'         End If
'
'        sql = "SELECT  Archive_T_indiceProjet.Id_Pieces FROM Archive_T_indiceProjet "
'        sql = sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
'        Set Rs = Con.OpenRecordSet(sql)
'        If Rs.EOF = False Then
'            Id_Pieces = Rs!Id_Pieces
'
'             sql = "SELECT  Archive_T_Pieces.IdProjet,Archive_T_Pieces.Description  FROM Archive_T_Pieces "
'            sql = sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
'             Set Rs = Con.OpenRecordSet(sql)
'              If Rs.EOF = False Then
'                    IdProjet = Rs!IdProjet
'                    LibPice = "" & Rs!Description
'           End If
'
'
'
'            sql = "SELECT T_Projet.id FROM T_Projet "
'            sql = sql & "WHERE T_Projet.id=" & IdProjet & ";"
'            Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = True Then
'                sql = "INSERT INTO T_Projet ( id, Projet, Description, CleAc ) "
'            sql = sql & "SELECT Archive_T_Projet.id, Archive_T_Projet.Projet, "
'            sql = sql & "Archive_T_Projet.Description, Archive_T_Projet.CleAc "
'            sql = sql & "FROM Archive_T_Projet "
'            sql = sql & "WHERE Archive_T_Projet.id=" & IdProjet & ";"
'            Con.Execute sql
'            End If
'
'
'
'            sql = "SELECT T_Pieces.Id FROM T_Pieces "
'            sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
'             Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = True Then
'                sql = "INSERT INTO T_Pieces ( id, IdProjet, Description ) "
'                sql = sql & "SELECT Archive_T_Pieces.Id, Archive_T_Pieces.IdProjet, Archive_T_Pieces.Description "
'                sql = sql & "FROM Archive_T_Pieces "
'                sql = sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
'                Con.Execute sql
'
'            End If
'
'
'
'            sql = "SELECT T_indiceProjet.Id FROM T_indiceProjet "
'            sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
'            Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = True Then
'                sql = "INSERT INTO T_indiceProjet ( Id, Id_Pieces, Description, PL, PL_Indice,  "
'                sql = sql & "OU, OU_Indice, Li, LI_Indice, PI, IdStatus, PI_Indice, IdStatusSave,  "
'                sql = sql & "IdApprobateur, PlAutoCadSaveAs, PlAutoCadSave, NbErr, Client, Destinataire,  "
'                sql = sql & "Service, DessineDate, DessineNOM, VerifieDate, VerifieNom, ApprouveDate,  "
'                sql = sql & "ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble, CleAc, RefP,  "
'                sql = sql & "Masse, LiAutoCadSaveAs, LiAutoCadSave, OuAutoCadSaveAs, OuAutoCadSave,  "
'                sql = sql & "Archiver, Cartouche, Version, Pere,NbCartouche ) "
'                sql = sql & "SELECT Archive_T_indiceProjet.Id,Archive_T_indiceProjet.Id_Pieces,  "
'                sql = sql & "Archive_T_indiceProjet.Description, Archive_T_indiceProjet.PL,  "
'                sql = sql & "Archive_T_indiceProjet.PL_Indice, Archive_T_indiceProjet.[OU],  "
'                sql = sql & "Archive_T_indiceProjet.OU_Indice, Archive_T_indiceProjet.Li,  "
'                sql = sql & "Archive_T_indiceProjet.LI_Indice, Archive_T_indiceProjet.PI,  "
'                sql = sql & "Archive_T_indiceProjet.IdStatus, Archive_T_indiceProjet.PI_Indice,  "
'                sql = sql & "Archive_T_indiceProjet.IdStatusSave, Archive_T_indiceProjet.IdApprobateur,  "
'                sql = sql & "Archive_T_indiceProjet.PlAutoCadSaveAs, Archive_T_indiceProjet.PlAutoCadSave,  "
'                sql = sql & "Archive_T_indiceProjet.NbErr, Archive_T_indiceProjet.Client,  "
'                sql = sql & "Archive_T_indiceProjet.Destinataire, Archive_T_indiceProjet.Service,  "
'                sql = sql & "Archive_T_indiceProjet.DessineDate, Archive_T_indiceProjet.DessineNOM,  "
'                sql = sql & "Archive_T_indiceProjet.VerifieDate, Archive_T_indiceProjet.VerifieNom,  "
'                sql = sql & "Archive_T_indiceProjet.ApprouveDate, Archive_T_indiceProjet.ApprouveNom,  "
'                sql = sql & "Archive_T_indiceProjet.Responsable, Archive_T_indiceProjet.Vague,  "
'                sql = sql & "Archive_T_indiceProjet.Equipement, Archive_T_indiceProjet.RefPF,  "
'                sql = sql & "Archive_T_indiceProjet.Ensemble , Archive_T_indiceProjet.CleAc,  "
'                sql = sql & "Archive_T_indiceProjet.RefP, Archive_T_indiceProjet.Masse,  "
'                sql = sql & "Archive_T_indiceProjet.LiAutoCadSaveAs, Archive_T_indiceProjet.LiAutoCadSave,  "
'                sql = sql & "Archive_T_indiceProjet.OuAutoCadSaveAs, Archive_T_indiceProjet.OuAutoCadSave,  "
'                sql = sql & "Archive_T_indiceProjet.Archiver, Archive_T_indiceProjet.Cartouche,  "
'                sql = sql & "Archive_T_indiceProjet.Version, Archive_T_indiceProjet.Pere,Archive_T_indiceProjet.NbCartouche "
'                sql = sql & "FROM Archive_T_indiceProjet "
'                sql = sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
'                Con.Execute sql
'                If IdPIndice = 0 Then
'                    sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = False "
'                    sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
'                    Con.Execute sql
'                End If
'            End If
'
'              sql = "SELECT T_Critères.Id FROM T_Critères "
'            sql = sql & "WHERE T_Critères.Id=" & Id_Pieces & ";"
'             Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = True Then
'            sql = "INSERT INTO T_Critères SELECT Archive_T_Critères.* FROM Archive_T_Critères "
'                sql = sql & "WHERE Archive_T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
'
''                Sql = "INSERT INTO T_Critères "
''                Sql = Sql & "SELECT Archive_T_Critères.* "
''                Sql = Sql & "FROM Archive_T_Critères "
''                Sql = Sql & "WHERE Archive_T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
'                Con.Execute sql
'
'            End If
'            sql = "SELECT Connecteurs.Id_IndiceProjet FROM Connecteurs "
'            sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
'             Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = True Then
'
'            sql = "INSERT INTO Connecteurs SELECT Archive_Connecteurs.* FROM Archive_Connecteurs "
'                sql = sql & "WHERE Archive_Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
'
''                Sql = "INSERT INTO Connecteurs ( Numéro, Id_IndiceProjet, CONNECTEUR,  "
''                Sql = Sql & "[O/N], DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2, [100%] ) "
''                Sql = Sql & "SELECT Archive_Connecteurs.Numéro, Archive_Connecteurs.Id_IndiceProjet,  "
''                Sql = Sql & "Archive_Connecteurs.CONNECTEUR,Archive_Connecteurs.[O/N],  "
''                Sql = Sql & "Archive_Connecteurs.DESIGNATION, Archive_Connecteurs.CODE_APP,  "
''                Sql = Sql & "Archive_Connecteurs.N°, Archive_Connecteurs.POS, Archive_Connecteurs.[POS-OUT],  "
''                Sql = Sql & "Archive_Connecteurs.PRECO1, Archive_Connecteurs.PRECO2, Archive_Connecteurs.[100%] "
''                Sql = Sql & "FROM Archive_Connecteurs "
''                Sql = Sql & "WHERE Archive_Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
'                Con.Execute sql
'            End If
'            sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet FROM Ligne_Tableau_fils "
'            sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
'            Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = True Then
'
'              sql = "INSERT INTO Ligne_Tableau_fils SELECT Archive_Ligne_Tableau_fils.* FROM Archive_Ligne_Tableau_fils "
'                sql = sql & "WHERE Archive_Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
'
''                Sql = "INSERT INTO Ligne_Tableau_fils ( Numéro, Id_IndiceProjet, LIAI, DESIGNATION, FIL,  "
''                Sql = Sql & "SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP,  "
''                Sql = Sql & "VOI, POS2, [POS-OUT2], FA2, APP2, VOI2, PRECO, [OPTION] ) "
''                Sql = Sql & "SELECT Archive_Ligne_Tableau_fils.Numéro,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.Id_IndiceProjet,Archive_Ligne_Tableau_fils.LIAI,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.DESIGNATION, Archive_Ligne_Tableau_fils.FIL, "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.SECT, Archive_Ligne_Tableau_fils.TEINT,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.TEINT2, Archive_Ligne_Tableau_fils.ISO,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.LONG, Archive_Ligne_Tableau_fils.[LONG CP],  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.COUPE, Archive_Ligne_Tableau_fils.POS,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.[POS-OUT], Archive_Ligne_Tableau_fils.FA,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.APP, Archive_Ligne_Tableau_fils.VOI,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.POS2, Archive_Ligne_Tableau_fils.[POS-OUT2],  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.FA2, Archive_Ligne_Tableau_fils.APP2,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.VOI2, Archive_Ligne_Tableau_fils.PRECO,  "
''                Sql = Sql & "Archive_Ligne_Tableau_fils.OPTION "
''                Sql = Sql & "FROM Archive_Ligne_Tableau_fils "
''                Sql = Sql & "WHERE Archive_Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
'            Con.Execute sql
'            End If
'
'             sql = "SELECT Composants.Id_IndiceProjet FROM Composants "
'             sql = sql & "WHERE Composants.Id_IndiceProjet=" & IndiceProjet & ";"
'             Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = True Then
'
'            sql = "INSERT INTO Composants SELECT Archive_Composants.* FROM Archive_Composants "
'                sql = sql & "WHERE Archive_Composants.Id_IndiceProjet=" & IndiceProjet & ";"
''
''                Sql = "INSERT INTO Composants ( Id, Id_IndiceProjet, DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
''                Sql = Sql & "SELECT Archive_Composants.Id, Archive_Composants.Id_IndiceProjet,  "
''                Sql = Sql & "Archive_Composants.DESIGNCOMP, Archive_Composants.NUMCOMP,  "
''                Sql = Sql & "Archive_Composants.REFCOMP, Archive_Composants.Path "
''                Sql = Sql & "FROM Archive_Composants "
''                Sql = Sql & "WHERE Archive_Composants.Id_IndiceProjet=" & IndiceProjet & ";"
'            Con.Execute sql
'
'            End If
'
'            sql = "SELECT Nota.Id_IndiceProjet FROM Nota "
'            sql = sql & "WHERE Nota.Id_IndiceProjet=" & IndiceProjet & ";"
'             Set Rs = Con.OpenRecordSet(sql)
'            If Rs.EOF = True Then
'
'             sql = "INSERT INTO Nota SELECT Archive_Nota.* FROM Archive_Nota "
'                sql = sql & "WHERE Archive_Nota.Id_IndiceProjet=" & IndiceProjet & ";"
'
''                Sql = "INSERT INTO Nota ( Id, Id_IndiceProjet, NOTA, NUMNOTA )  "
''                Sql = Sql & "SELECT Archive_Nota.Id, Archive_Nota.Id_IndiceProjet, Archive_Nota.NOTA,   "
''                Sql = Sql & "Archive_Nota.NUMNOTA  "
''                Sql = Sql & "FROM Archive_Nota  "
''                Sql = Sql & "WHERE Archive_Nota.Id_IndiceProjet=" & IndiceProjet & ";"
'                Con.Execute sql
'
'            End If
'            sql = "SELECT T_Noeuds.Id_IndiceProjet FROM T_Noeuds  "
'                sql = sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IndiceProjet & ";"
'              Set Rs = Con.OpenRecordSet(sql)
'        If Rs.EOF = True Then
'                sql = "INSERT INTO T_Noeuds SELECT Archive_T_Noeuds.*  "
'                sql = sql & "FROM Archive_T_Noeuds "
'                sql = sql & "WHERE Archive_T_Noeuds.Id_IndiceProjet=" & IndiceProjet & ";"
'
'                Con.Execute sql
'
'            End If
'
'             sql = "DELETE Archive_T_Pieces.*"
'                sql = sql & "FROM Archive_T_Pieces INNER JOIN Archive_T_indiceProjet  "
'                sql = sql & "ON Archive_T_Pieces.Id = Archive_T_indiceProjet.Id_Pieces "
'                sql = sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
'                Con.Execute sql
'
'        End If
    End If
Next I
Sql = "DELETE Archive_T_Projet.*, Archive_T_Pieces.Id "
            Sql = Sql & "FROM Archive_T_Projet LEFT JOIN Archive_T_Pieces  "
            Sql = Sql & "ON Archive_T_Projet.id = Archive_T_Pieces.IdProjet "
            Sql = Sql & "WHERE Archive_T_Pieces.Id Is Null;"
            Con.Execute Sql
Sql = "DELETE Archive_T_Pieces.* "
Sql = Sql & "FROM Archive_T_Pieces LEFT JOIN Archive_T_indiceProjet ON Archive_T_Pieces.Id = Archive_T_indiceProjet.Id_Pieces  "
Sql = Sql & "WHERE Archive_T_indiceProjet.Id Is Null;"
Con.Execute Sql

 Noquite = False
Me.Hide
End Sub

Public Sub Charge(MyForm As Object, Optional Filtre As String, Optional boolTxt As Boolean, Optional boolArchive As Boolean)
'Dim sql As String
'Dim Rs As Recordset
'Dim IndexRow As Long
'Dim IndexCol As Long
'
'boolTxts = boolTxt
'IndexRow = 1
'IndexCol = 0
'Me.Spreadsheet1.Columns(1).Locked = False
'
' Me.Spreadsheet1.Columns(1).NumberFormat = "Yes/No"
' Me.Spreadsheet1.Columns(1).Locked = True
'
'sql = "SELECT 0 AS Importer ,Archive_SelectProjets.* "
'sql = sql & "FROM Archive_SelectProjets; "
'Set Rs = Con.OpenRecordSet(sql)
'Rs.Filter = Filtre
'Set MyFormCible = MyForm
'While Rs.EOF = False
'IndexRow = IndexRow + 1
'For IndexCol = 0 To Rs.Fields.Count - 11
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
'    If IndexCol > 2 And IndexCol < 7 Then
'         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("'" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 17))
'    Else
'        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("'" & Rs.Fields(IndexCol))
'    End If
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = &HFFC0FF
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = MyLocked(Me.Spreadsheet1.Cells(1, IndexCol + 1), Rs.Fields(Rs.Fields.Count - 11))
'
'Next IndexCol
'
'Rs.MoveNext
'Wend
'
'Set Rs = Con.CloseRecordSet(Rs)
'
'Dim Myrange
'Set Myrange = Me.Spreadsheet1.Range("A1").CurrentRegion
'Myrange.AutoFitColumns
Dim Sql As String
Dim Rs As Recordset
Dim IndexRow As Long
Dim IndexCol As Long
Dim OfsetCol As Long
boolTxts = boolTxt
IndexRow = 1
IndexCol = 0
OfsetCol = 1
Dim I As Long
Sql = "SELECT SelectProjets.* "
Sql = Sql & "FROM SelectProjets; "
Set Rs = Con.OpenRecordSet(Sql)
Rs.Filter = Filtre
Sql = "SELECT  0 AS Suprimers  , 0 AS Archivers,SelectProjets.* "
Sql = Sql & "FROM SelectProjets; "



Sql = "SELECT   0 AS Archivers,SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement, SelectProjets.Ensemble,  "
Sql = Sql & "SelectProjets.CleAc, 0 AS chrono, [SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice]  "
Sql = Sql & "AS Expr1,  [SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice]  "
Sql = Sql & "AS Expr2, [SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice] AS Expr3,  "
Sql = Sql & "[SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice] AS Expr4, SelectProjets.Client,  "
Sql = Sql & "SelectProjets.DessineNOM, SelectProjets.VerifieNom, SelectProjets.ApprouveNom,  "
Sql = Sql & "SelectProjets.Id, SelectProjets.IdStatus, SelectProjets.NbErr, SelectProjets.LiAutoCadSave,  "
Sql = Sql & "SelectProjets.VerifieDate, SelectProjets.Archiver, SelectProjets.PI_Indice,  "
Sql = Sql & "SelectProjets.PL_Indice, SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere,  "
Sql = Sql & "SelectProjets.PlOk, SelectProjets.OuOk "
Sql = Sql & "FROM ((SelectProjets LEFT JOIN T_Job ON SelectProjets.Id = T_Job.Id_Piece)   "
Sql = Sql & "LEFT JOIN T_Job AS T_Job_1 ON SelectProjets.Pere = T_Job_1.Id_Piece) LEFT JOIN T_indiceProjet   "
Sql = Sql & "ON SelectProjets.Id = T_indiceProjet.Id  "
Sql = Sql & "WHERE (T_Job.Id_Piece Is Null   "
Sql = Sql & "AND T_Job_1.Id_Piece Is Null   "
Sql = Sql & "AND (T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "'   "
Sql = Sql & "Or T_indiceProjet.UserName Is Null))   "
Sql = Sql & "OR ((T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "'   "
Sql = Sql & "Or T_indiceProjet.UserName Is Null)   "
Sql = Sql & "AND (T_Job.FinTraitement=True or T_Job_1.FinTraitement=True))   "
'Sql = Sql & "AND ((T_Job_1.FinTraitement)=True))  "
Sql = Sql & "ORDER BY SelectProjets.CleAc DESC , SelectProjets.PI DESC;"
Set Rs = Con.OpenRecordSet(Sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
'Rs.Filter = "IdStatus<4"
If Rs.EOF = False Then
    Const sDelimiteur$ = vbTab
    Debug.Print Asc(vbCrLf)
    Dim toto
    toto = Rs.GetString(, , sDelimiteur$, "¤")
    
    toto = Replace(toto, vbCrLf, " ")
    toto = Replace(toto, Chr(13), "")
    toto = Replace(toto, Chr(10), "")
    toto = Replace(toto, "\", "")
    toto = Replace(toto, "¤", vbCrLf)
    Spreadsheet1.ActiveSheet.Protection.Enabled = False
    Spreadsheet1.ActiveSheet.Range("A2").ParseText _
    toto, sDelimiteur$

End If
Me.Spreadsheet1.Columns(1).NumberFormat = "Yes/No"
' Me.Spreadsheet1.Columns(2).NumberFormat = "Yes/No"
 
Dim MyRange
Set MyRange = Me.Spreadsheet1.Range("A1").CurrentRegion
MyRange.AutoFitColumns
Spreadsheet1.ActiveSheet.Cells(1, 15).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 16).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 17).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 18).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 19).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 20).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 21).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 22).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 23).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 24).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 25).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 26).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 27).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 28).ColumnWidth = 0
Spreadsheet1.ActiveSheet.Cells(1, 29).ColumnWidth = 0
For I = 2 To MyRange.Rows.Count
aa = Split(Trim("" & MyRange(I, 8) & "____"), "_")
MyRange(I, 7) = aa(3)
'             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & aa(1))
Me.Spreadsheet1.Rows(I).Interior.Color = ChoixCouleur(Val(MyRange(I, 17)))
Next
Spreadsheet1.ActiveSheet.Protection.Enabled = True




    Spreadsheet1.ActiveSheet.Columns(1).Locked = False

'    Spreadsheet1.ActiveSheet.Cells(I, 2).Locked = False

Set MyRange = Nothing

Me.Show vbModal
End Sub
Function MyLocked(Mytype As String, Statues As Long) As Boolean
MyLocked = True
If Mytype = "Importer O/N" Then MyLocked = False

End Function

Private Sub CommandButton2_Click()
 Noquite = False
Me.Hide
End Sub

Private Sub UserForm_Activate()
 Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub


