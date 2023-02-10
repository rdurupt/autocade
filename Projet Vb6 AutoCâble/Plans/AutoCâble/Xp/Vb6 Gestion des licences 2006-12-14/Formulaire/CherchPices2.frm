VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form UserForm4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supprimer/Archiver:"
   ClientHeight    =   11715
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   18840
   Icon            =   "CherchPices2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11715
   ScaleWidth      =   18840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   18840
      HTMLURL         =   ""
      HTMLData        =   $"CherchPices2.frx":08CA
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
      Left            =   16560
      TabIndex        =   11
      Top             =   0
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "CherchPices2.frx":1E1C
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   10800
      Width           =   3135
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "&Valider"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   10800
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   -240
      TabIndex        =   10
      Top             =   11400
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   795
      TabIndex        =   8
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1485
      TabIndex        =   7
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "CRE"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   330
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "MOD"
      Height          =   195
      Index           =   1
      Left            =   1050
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "VAL"
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   300
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean
Public Sub Charge(MyForm As Object, Optional Filtre As String, Optional boolTxt As Boolean, Optional boolArchive As Boolean)
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



Sql = "SELECT  0 AS Suprimers  , 0 AS Archivers,SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement, SelectProjets.Ensemble,  "
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
 Me.Spreadsheet1.Columns(2).NumberFormat = "Yes/No"
 
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
aa = Split(Trim("" & MyRange(I, 9) & "____"), "_")
MyRange(I, 8) = aa(3)
'             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & aa(1))
Me.Spreadsheet1.Rows(I).Interior.Color = ChoixCouleur(Val(MyRange(I, 18)))
Next
Spreadsheet1.ActiveSheet.Protection.Enabled = True


For I = 2 To MyRange.Rows.Count

If MyRange(I, 18) < 3 Then
    Spreadsheet1.ActiveSheet.Cells(I, 1).Locked = False
Else
    Spreadsheet1.ActiveSheet.Cells(I, 2).Locked = False
End If
Next
While Rs.EOF = False


IndexRow = IndexRow + 1
OfsetCol = 1
For IndexCol = 0 To Rs.Fields.Count - 13

    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = False
    If IndexCol > 6 And IndexCol < 10 Then
    If IndexCol = 7 Then
           aa = Split(Trim("" & Rs.Fields(IndexCol) & "___"), "_")
             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & aa(3))
             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 12))
            Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = True
            OfsetCol = OfsetCol + 1
        End If
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = False
        
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("'" & Rs.Fields(IndexCol)) & "_" & Trim("" & Rs.Fields(IndexCol + 14))
    Else
        
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("'" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 12))
    
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = True
    If IndexCol + OfsetCol = 1 And (Rs.Fields(Rs.Fields.Count - 12) = 1 Or Rs.Fields(Rs.Fields.Count - 12) = 2) Then
        Me.Spreadsheet1.Cells(IndexRow, 1).Locked = False
    End If
If IndexCol + OfsetCol = 2 And Rs.Fields(Rs.Fields.Count - 12) = 3 Then
        Me.Spreadsheet1.Cells(IndexRow, 2).Locked = False
    End If

Next IndexCol

Rs.MoveNext
Wend

'Dim Sql As String
'Dim Rs As Recordset
'Dim IndexRow As Long
'Dim IndexCol As Long
'Dim OfsetCol As Long
'Me.Spreadsheet1.Columns(1).Locked = False
'  Me.Spreadsheet1.Columns(2).Locked = False
 
' Me.Spreadsheet1.Columns(1).Locked = True
'  Me.Spreadsheet1.Columns(2).Locked = True
'
'
'boolTxts = boolTxt
'IndexRow = 1
'IndexCol = 0
'OfsetCol = 1
'Sql = "SELECT  0 AS Suprimers  , 0 AS Archivers,SelectProjets.* "
'Sql = Sql & "FROM SelectProjets; "
'Set Rs = Con.OpenRecordSet(Sql)
'Rs.Filter = Filtre
'Set MyFormCible = MyForm
'While Rs.EOF = False
'IndexRow = IndexRow + 1
'
'For IndexCol = 0 To Rs.Fields.Count - 13
'OfsetCol = 1
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = False
'    If IndexCol > 5 And IndexCol < 10 Then
'    If IndexCol = 6 Then
'           aa = Split(Trim("" & Rs.Fields(IndexCol)), "_")
'             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & aa(UBound(aa)))
'             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 12))
'            Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = True
'            OfsetCol = OfsetCol + 1
'        End If
'         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = False
'
'         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & Rs.Fields(IndexCol)) & "_" & Trim("" & Rs.Fields(IndexCol + 18))
'    Else
'
'         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & Rs.Fields(IndexCol))
'    End If
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 12))
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = True
'
'Next IndexCol
''
''For IndexCol = 0 To Rs.Fields.Count - 13
''    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
''    If IndexCol > 4 And IndexCol < 9 Then
''         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("'" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 18))
''    Else
''        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("'" & Rs.Fields(IndexCol))
''    End If
''    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 12))
''    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = MyLocked(Me.Spreadsheet1.Cells(1, IndexCol + 1), Rs.Fields(Rs.Fields.Count - 11))
''
''Next IndexCol
'
'Rs.MoveNext
'Wend
'
'Set Rs = Con.CloseRecordSet(Rs)
'Dim Myrange
'Set Myrange = Me.Spreadsheet1.Range("A1").CurrentRegion
'Myrange.AutoFitColumns
'Set Myrange = Nothing
Me.Show vbModal
End Sub
'Function ChoixCouleur(Mode As Long) As Long
'
'
'   Select Case Mode
'    Case 1
'        ChoixCouleur = 16777164
'    Case 2
'    ChoixCouleur = 10079487
'    Case 3
'        ChoixCouleur = 13434828
'   End Select
'
'End Function
Function MyLocked(Mytype As String, Statues As Long) As Boolean
MyLocked = True
If Mytype = "Supprimer O/N" And Statues = 1 Then MyLocked = False
If Mytype = "Supprimer O/N" And Statues = 2 Then MyLocked = False
If Mytype = "Supprimer O/N" And Statues = 3 Then MyLocked = False
If Mytype = "Archiver O/N" And Statues = 3 Then MyLocked = False

End Function

Private Sub Command1_Click()

End Sub

Private Sub CommandButton1_Click()
Dim msg As String
Dim MyRange
Dim IndiceProjet As Long
Dim Id_Pieces As Long
Dim IdProjet As Long
Dim Sql As String
Dim LibPice As String
Dim Rs As Recordset
msg = "Attention tous les enregistrements supprimés seront définitivement perdus." & vbCrLf & vbCrLf
msg = msg & "Tous les enregistrements archivés pourront être réinsérés par la suite" & vbCrLf & vbCrLf
msg = msg & "Voulez-vous continuer." & vbCrLf & vbCrLf
If MsgBox(msg, vbYesNo + vbQuestion, "Supprimer/Archiver") = vbNo Then Exit Sub
RazFiltreEditExcel Me.Spreadsheet1
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
ProgressBar1.Value = 0
ProgressBar1.Max = MyRange.Rows.Count



For I = 2 To MyRange.Rows.Count
ProgressBar1.Value = I
    If UCase(MyRange(I, 2)) <> 0 Then
        IndiceProjet = CInt(Me.Spreadsheet1.Cells(I, 17))
'        sql = "SELECT  T_indiceProjet.Id_Pieces FROM T_indiceProjet "
'        sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
'        Set Rs = Con.OpenRecordSet(sql)
'        If Rs.EOF = False Then
'            Id_Pieces = Rs!Id_Pieces
'
'             sql = "SELECT  T_Pieces.IdProjet,T_Pieces.Description  FROM T_Pieces "
'            sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
'             Set Rs = Con.OpenRecordSet(sql)
'              If Rs.EOF = False Then
'                    IdProjet = Rs!IdProjet
'                    LibPice = "" & Rs!Description
'                    sql = "SELECT Archive_T_Projet.id FROM Archive_T_Projet "
'                    sql = sql & "WHERE Archive_T_Projet.id=" & IdProjet & ";"
'                     Set Rs = Con.OpenRecordSet(sql)
'                     If Rs.EOF = True Then
'                        sql = "INSERT INTO Archive_T_Projet ( id, Projet, Description, CleAc ) "
'                        sql = sql & "SELECT T_Projet.id, T_Projet.Projet, T_Projet.Description, T_Projet.CleAc "
'                        sql = sql & "FROM T_Projet "
'                   sql = sql & "WHERE T_Projet.id=" & IdProjet & ";"
'
'                        Con.Execute sql
'
'                     End If
'                     sql = "SELECT Archive_T_Pieces.Id FROM Archive_T_Pieces "
'                     sql = sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
'                        Set Rs = Con.OpenRecordSet(sql)
'                     If Rs.EOF = True Then
'                        sql = "INSERT INTO Archive_T_Pieces ( id, IdProjet, Description ) "
'                        sql = sql & "SELECT T_Pieces.Id, T_Pieces.IdProjet, T_Pieces.Description "
'                        sql = sql & "FROM T_Pieces "
'                        sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
'                        Con.Execute sql
'                     End If
'                     sql = "SELECT Archive_T_indiceProjet.Id FROM Archive_T_indiceProjet "
'                     sql = sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
'                       Set Rs = Con.OpenRecordSet(sql)
'                     If Rs.EOF = True Then
'
'
'                        sql = "INSERT INTO Archive_T_indiceProjet  "
'                        sql = sql & "( Id, Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice,  "
'                        sql = sql & "Li, LI_Indice, PI, IdStatus, PI_Indice, IdStatusSave,  "
'                        sql = sql & "IdApprobateur, PlAutoCadSaveAs, PlAutoCadSave, NbErr, Client,  "
'                        sql = sql & "Destinataire, Service, DessineDate, DessineNOM, VerifieDate,  "
'                        sql = sql & "VerifieNom, ApprouveDate, ApprouveNom, Responsable, Vague,  "
'                        sql = sql & "Equipement, RefPF, Ensemble, CleAc, RefP, Masse, LiAutoCadSaveAs,  "
'                        sql = sql & "LiAutoCadSave, OuAutoCadSaveAs, OuAutoCadSave, Archiver, Cartouche,  "
'                        sql = sql & "Version, Pere,NbCartouche) "
'                        sql = sql & "SELECT T_indiceProjet.Id, T_indiceProjet.Id_Pieces,  "
'                        sql = sql & "T_indiceProjet.Description, T_indiceProjet.PL, "
'                        sql = sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU],  "
'                        sql = sql & "T_indiceProjet.OU_Indice, T_indiceProjet.Li,  "
'                        sql = sql & "T_indiceProjet.LI_Indice, T_indiceProjet.PI,  "
'                        sql = sql & "T_indiceProjet.IdStatus, T_indiceProjet.PI_Indice,  "
'                        sql = sql & "T_indiceProjet.IdStatusSave, T_indiceProjet.IdApprobateur,  "
'                        sql = sql & "T_indiceProjet.PlAutoCadSaveAs, T_indiceProjet.PlAutoCadSave, "
'                        sql = sql & " T_indiceProjet.NbErr, T_indiceProjet.Client,  "
'                        sql = sql & "T_indiceProjet.Destinataire, T_indiceProjet.Service, "
'                        sql = sql & " T_indiceProjet.DessineDate, T_indiceProjet.DessineNOM, "
'                        sql = sql & " T_indiceProjet.VerifieDate, T_indiceProjet.VerifieNom, "
'                        sql = sql & " T_indiceProjet.ApprouveDate, T_indiceProjet.ApprouveNom, "
'                        sql = sql & " T_indiceProjet.Responsable, T_indiceProjet.Vague, "
'                        sql = sql & " T_indiceProjet.Equipement, T_indiceProjet.RefPF, "
'                        sql = sql & " T_indiceProjet.Ensemble, T_indiceProjet.CleAc, "
'                        sql = sql & " T_indiceProjet.RefP, T_indiceProjet.Masse, "
'                        sql = sql & " T_indiceProjet.LiAutoCadSaveAs, T_indiceProjet.LiAutoCadSave, "
'                        sql = sql & " T_indiceProjet.OuAutoCadSaveAs, T_indiceProjet.OuAutoCadSave, "
'                        sql = sql & " T_indiceProjet.Archiver, T_indiceProjet.Cartouche ,  "
'                        sql = sql & "T_indiceProjet.Version, T_indiceProjet.Pere,T_indiceProjet.NbCartouche "
'
'                        sql = sql & "FROM T_indiceProjet "
'                        sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
'                        Con.Execute sql
'
'                     End If
'                     sql = "SELECT Archive_Connecteurs.Id_IndiceProjet FROM Archive_Connecteurs "
'                     sql = sql & "WHERE Archive_Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
'                    Set Rs = Con.OpenRecordSet(sql)
'                    If Rs.EOF = True Then
'
'                     sql = "INSERT INTO Archive_Connecteurs SELECT Connecteurs.* FROM Connecteurs "
'                            sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
'
'
''                        Sql = "INSERT INTO Archive_Connecteurs ( Numéro, Id_IndiceProjet, CONNECTEUR, "
''                        Sql = Sql & "[O/N], DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2, [100%] )"
''                        Sql = Sql & "SELECT Connecteurs.Numéro, Connecteurs.Id_IndiceProjet, "
''                        Sql = Sql & "Connecteurs.CONNECTEUR, Connecteurs.[O/N], "
''                        Sql = Sql & "Connecteurs.DESIGNATION, Connecteurs.CODE_APP, Connecteurs.N°, "
''                        Sql = Sql & "Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, "
''                        Sql = Sql & "Connecteurs.PRECO2, Connecteurs.[100%]"
''                        Sql = Sql & "FROM Connecteurs "
''                        Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
'                        Con.Execute sql
'                    End If
'                        sql = "SELECT Archive_Ligne_Tableau_fils.Id_IndiceProjet FROM Archive_Ligne_Tableau_fils "
'                        sql = sql & "WHERE Archive_Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
'
'                        Set Rs = Con.OpenRecordSet(sql)
'                        If Rs.EOF = True Then
'
'                         sql = "INSERT INTO Archive_Ligne_Tableau_fils SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils "
'                            sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
'
''
''                            Sql = "INSERT INTO Archive_Ligne_Tableau_fils ( Numéro, Id_IndiceProjet, LIAI, "
''                        Sql = Sql & "DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE,  "
''                        Sql = Sql & "POS, [POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2, VOI2,  "
''                        Sql = Sql & "PRECO, [OPTION] ) "
''                        Sql = Sql & "SELECT Ligne_Tableau_fils.Numéro, Ligne_Tableau_fils.Id_IndiceProjet,  "
''                        Sql = Sql & "Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
''                        Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT,  "
''                        Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
''                        Sql = Sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
''                        Sql = Sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE,  "
''                        Sql = Sql & "Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT],  "
''                        Sql = Sql & "Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,  "
''                        Sql = Sql & "Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
''                        Sql = Sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
''                        Sql = Sql & "Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
''                        Sql = Sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
''                        Sql = Sql & "FROM Ligne_Tableau_fils "
''                        Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
'                        Con.Execute sql
'
'                        End If
'
'                         sql = "SELECT Archive_Composants.Id_IndiceProjet FROM Archive_Composants "
'                        sql = sql & "WHERE Archive_Composants.Id_IndiceProjet=" & IndiceProjet & ";"
'
'                        Set Rs = Con.OpenRecordSet(sql)
'                        If Rs.EOF = True Then
'
'                        sql = "INSERT INTO Archive_Composants SELECT Composants.* FROM Composants "
'                            sql = sql & "WHERE Composants.Id_IndiceProjet=" & IndiceProjet & ";"
'
''                            Sql = "INSERT INTO Archive_Composants ( Id, Id_IndiceProjet, DESIGNCOMP,  "
''                            Sql = Sql & "NUMCOMP, REFCOMP, Path ) "
''                            Sql = Sql & "SELECT Composants.Id, Composants.Id_IndiceProjet,  "
''                            Sql = Sql & "Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP,  "
''                            Sql = Sql & "Composants.Path "
''                            Sql = Sql & "FROM Composants "
''                            Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IndiceProjet & ";"
'                            Con.Execute sql
'                        End If
'
'                        sql = "SELECT Archive_Nota.Id_IndiceProjet FROM Archive_Nota "
'                        sql = sql & "WHERE Archive_Nota.Id_IndiceProjet=" & IndiceProjet & ";"
'
'                        Set Rs = Con.OpenRecordSet(sql)
'                        If Rs.EOF = True Then
'
'                        sql = "INSERT INTO Archive_Nota SELECT Nota.* FROM Nota "
'                            sql = sql & "WHERE Nota.Id_IndiceProjet=" & IndiceProjet & ";"
'
''                            Sql = "INSERT INTO Archive_Nota ( Id, Id_IndiceProjet, NOTA, NUMNOTA ) "
''                            Sql = Sql & "SELECT Nota.Id, Nota.Id_IndiceProjet, Nota.NOTA, Nota.NUMNOTA "
''                            Sql = Sql & "FROM Nota "
''                            Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IndiceProjet & ";"
'                            Con.Execute sql
'
'                        End If
'
'                        sql = "SELECT Archive_T_Critères.Id_IndiceProjet FROM Archive_T_Critères "
'                        sql = sql & "WHERE Archive_T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
'
'                        Set Rs = Con.OpenRecordSet(sql)
'                        If Rs.EOF = True Then
'
'                        sql = "INSERT INTO Archive_T_Critères SELECT T_Critères.* FROM T_Critères "
'                            sql = sql & "WHERE T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
'
''                            Sql = "INSERT INTO Archive_T_Critères ( Id, Id_IndiceProjet, CODE_CRITERE, CRITERES ) "
''                            Sql = Sql & "SELECT T_Critères.Id, T_Critères.Id_IndiceProjet, T_Critères.CODE_CRITERE, T_Critères.CRITERES "
''                            Sql = Sql & "FROM T_Critères "
''                            Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
'                            Con.Execute sql
'
'                        End If
'                        sql = "SELECT Archive_T_Noeuds.Id_IndiceProjet FROM Archive_T_Noeuds "
'                            sql = sql & "WHERE Archive_T_Noeuds.Id_IndiceProjet=" & IndiceProjet & ";"
'
'                        Set Rs = Con.OpenRecordSet(sql)
'                        If Rs.EOF = True Then
'                           sql = "INSERT INTO Archive_T_Noeuds SELECT T_Noeuds.* FROM T_Noeuds "
'                            sql = sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IndiceProjet & ";"
'                            Con.Execute sql
'                        End If
'                        sql = "DELETE T_Pieces.*, T_Pieces.Id FROM T_Pieces "
'                        sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
'
'                        sql = "DELETE  T_indiceProjet.*  "
'                        sql = sql & "FROM T_Pieces INNER JOIN T_indiceProjet ON  "
'                        sql = sql & "T_Pieces.Id = T_indiceProjet.Id_Pieces "
'                        sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
'
'                        Con.Execute sql
''                        Sql = "SELECT T_Pieces.Id FROM T_Pieces "
''                        Sql = Sql & "WHERE T_Pieces.IdProjet=" & IdProjet & ";"
''                        Set Rs = Con.OpenRecordSet(Sql)
''                        If Rs.EOF = True Then
''                            Sql = "DELETE T_Projet.* FROM T_Projet "
''                            Sql = Sql & "WHERE T_Projet.id=" & IdProjet & ";"
''                            Con.Execute Sql
''
''                        End If
'
'              End If
''               Sql = "SELECT T_indiceProjet.Id FROM T_indiceProjet "
''            Sql = Sql & "WHERE T_indiceProjet.PI='" & MyReplace(LibPice) & "' "
''            Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
''            Set Rs = Con.OpenRecordSet(Sql)
''            If Rs.EOF = False Then
''                Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = true "
''                Sql = Sql & "WHERE T_indiceProjet.Id=" & Rs!Id & ";"
''                Con.Execute Sql
''
''            End If
'               Set Rs = Con.CloseRecordSet(Rs)
'        End If
'
        Sql = "UPDATE T_indiceProjet SET T_indiceProjet.IdStatus = 4 "
            Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & " "
            Sql = Sql & "OR T_indiceProjet.Pere=" & IndiceProjet & " ;"
            Con.Execute Sql
    End If
    If UCase(MyRange(I, 1)) <> 0 Then
         IndiceProjet = CInt(Me.Spreadsheet1.Cells(I, 17))
         Sql = "DELETE  T_indiceProjet.*  "
                        Sql = Sql & "FROM T_Pieces INNER JOIN T_indiceProjet ON  "
                        Sql = Sql & "T_Pieces.Id = T_indiceProjet.Id_Pieces "
                        Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"

                        Con.Execute Sql
                        zz = Split(Replace(Me.Spreadsheet1.Cells(I, 9), " ", ""), "_")
                        y = 0
                        Dim yy As Long
                        aa = ""
                        For yy = LBound(zz) To UBound(zz) - 1
                        aa = aa & zz(yy) & "_"
                        Next
                        aa = Left(aa, Len(aa) - 1)
                        Sql = "SELECT T_indiceProjet.Id "
                        Sql = Sql & "FROM T_indiceProjet "
                        Sql = Sql & "WHERE [PI]='" & aa & "' "
                        Sql = Sql & "ORDER BY T_indiceProjet.Id DESC; "
                        Set Rs = Con.OpenRecordSet(Sql)
                        If Rs.EOF = False Then
                            Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = False "
                            Sql = Sql & "WHERE T_indiceProjet.Id=" & Rs!Id & ";"
                            Con.Execute Sql
                        End If
       aa = Me.Spreadsheet1.Cells(I, 9)
       
      
    End If
Next I
 Sql = "DELETE T_Projet.*, T_Pieces.Id  "
Sql = Sql & "FROM T_Projet LEFT JOIN T_Pieces ON T_Projet.id = T_Pieces.IdProjet "
Sql = Sql & "WHERE T_Pieces.Id Is Null;"
'
Con.Execute Sql
Sql = "DELETE T_Pieces.*, T_indiceProjet.Id  "
Sql = Sql & "FROM T_Pieces LEFT JOIN T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces  "
Sql = Sql & "WHERE T_indiceProjet.Id Is Null;"
Con.Execute Sql


Noquite = False
Me.Hide
End Sub

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


