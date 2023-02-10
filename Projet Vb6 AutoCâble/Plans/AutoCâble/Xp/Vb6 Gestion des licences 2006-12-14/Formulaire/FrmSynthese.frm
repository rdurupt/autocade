VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form FrmSynthese 
   Caption         =   "Fichier de Synthèse Générale."
   ClientHeight    =   11715
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   18840
   Icon            =   "FrmSynthese.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   11715
   ScaleWidth      =   18840
   StartUpPosition =   2  'CenterScreen
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   10530
      Left            =   0
      TabIndex        =   0
      Top             =   1155
      Width           =   18840
      HTMLURL         =   ""
      HTMLData        =   $"FrmSynthese.frx":08CA
      DataType        =   "HTMLDATA"
      AutoFit         =   -1  'True
      DisplayColHeaders=   0   'False
      DisplayGridlines=   -1  'True
      DisplayHorizontalScrollBar=   -1  'True
      DisplayRowHeaders=   0   'False
      DisplayTitleBar =   0   'False
      DisplayToolbar  =   -1  'True
      DisplayVerticalScrollBar=   -1  'True
      EnableAutoCalculate=   -1  'True
      EnableEvents    =   -1  'True
      MoveAfterReturn =   -1  'True
      MoveAfterReturnDirection=   0
      RightToLeft     =   0   'False
      ViewableRange   =   "1:65536"
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   16800
      TabIndex        =   1
      Top             =   0
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "FrmSynthese.frx":195D
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Archive"
      Height          =   195
      Index           =   3
      Left            =   2760
      TabIndex        =   10
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   9
      Top             =   720
      Width           =   195
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
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "VAL"
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   7
      Top             =   720
      Width           =   300
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "MOD"
      Height          =   195
      Index           =   1
      Left            =   1290
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "CRE"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   330
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1725
      TabIndex        =   4
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1035
      TabIndex        =   3
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   195
   End
End
Attribute VB_Name = "FrmSynthese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

On Error Resume Next
Dim Sql As String
Dim Myrep As String
Dim Rs As Recordset
Dim MyExcel As EXCEL.Application
MyExcel.DisplayAlerts = False
Dim MyWorkbook As EXCEL.Workbook
Dim Spreadsheet1 As Worksheet
Dim Fso As New FileSystemObject
Dim PathAffair As String
Dim Row As Long
Dim NameFichier
Dim Mabarr As Long
Dim Client As String
Dim coll As Long
Dim I As Long
Sql = "SELECT T_Status.Id, T_Status.Status FROM T_Status ORDER BY T_Status.Id;"
Set Rs = Con.OpenRecordSet(Sql)
coll = 1

'
'While Rs.EOF = False
'
'    Me.Spreadsheet1.Cells(1, coll) = "" & Rs!Status
'      Me.Spreadsheet1.Cells(1, coll).Interior.Color = ChoixCouleur(Val(Rs!Id), False)
'       coll = coll + 1
'    Rs.MoveNext
'Wend
'Me.Spreadsheet1.Cells(1, coll) = "ARCHIVE"
'      Me.Spreadsheet1.Cells(1, coll).Interior.Color = ChoixCouleur(4, False)

'aa = ChoixCouleur(1)
Sql = "SELECT Rq_Synthese_Total.* FROM Rq_Synthese_Total;"
Sql = "SELECT T_indiceProjet.CleAc AS Affaire, T_indiceProjet.Client, T_Projet.Projet, T_indiceProjet.Ensemble, T_indiceProjet.Equipement, "
Sql = Sql & "[RefP] & '_' & [Ref_PF] AS [Ref PF], [RefPieceClient] & '_' & [Ref_Piece_CLI] AS [Pièce CLI], [RefP] & '_' & [Ref_Plan_CLI] AS  "
Sql = Sql & "[Plan CLI], [PI] & '_' & [PI_Indice] AS Pièce, [PL] & '_' & [PL_Indice] AS Plan, [Ou] & '_' & [OU_Indice] AS Outil, [LI] & '_' &  "
Sql = Sql & "[LI_Indice] AS Liste, T_indiceProjet.NbErr, T_indiceProjet.DessineNOM, T_indiceProjet.VerifieNom, T_indiceProjet.ApprouveNom,  "
Sql = Sql & "T_Status.Id "
Sql = Sql & "FROM (T_Status INNER JOIN (T_Projet INNER JOIN (T_Pieces INNER  "
Sql = Sql & "JOIN (T_indiceProjet LEFT JOIN T_Clients ON T_indiceProjet.Client = T_Clients.Client)  "
Sql = Sql & "ON T_Pieces.Id = T_indiceProjet.Id_Pieces) ON T_Projet.id = T_Pieces.IdProjet)  "
Sql = Sql & "ON T_Status.Id = T_indiceProjet.IdStatus) LEFT JOIN T_Job ON T_indiceProjet.Id = T_Job.Id_Piece "
Sql = Sql & "Where T_indiceProjet.UserName Is Null  "
Sql = Sql & "And T_Job.Id_Piece Is Null  "
Sql = Sql & "Or T_indiceProjet.UserName Is Null  "
Sql = Sql & "And T_Job.Id_Piece Is Not Null  "
Sql = Sql & "And T_Job.FinTraitement = True "
Sql = Sql & "ORDER BY T_indiceProjet.CleAc, T_indiceProjet.Client, T_Projet.Projet;"


Set Rs = Con.OpenRecordSet(Sql)
For I = 0 To Rs.Fields.Count - 2
    Me.Spreadsheet1.Cells(1, I + 1) = Rs.Fields(I).Name
     Me.Spreadsheet1.Cells(1, I + 1).Interior.Color = ChoixCouleur(0, False)
Next
If Rs.EOF = False Then
    Const sDelimiteur$ = vbTab
    Debug.Print Asc(vbCrLf)
    Dim toto
    toto = Rs.GetString(, , sDelimiteur$, "¤")
    toto = Replace(toto, sDelimiteur$, sDelimiteur$ & "'")
    toto = Replace(toto, vbCrLf, " ")
    toto = Replace(toto, Chr(13), "")
    toto = Replace(toto, Chr(10), "")
'    toto = Replace(toto, "\", "")
    toto = Replace(toto, "¤", vbCrLf)
    Me.Spreadsheet1.ActiveSheet.Protection.Enabled = False
    Me.Spreadsheet1.ActiveSheet.Range("A2").ParseText _
    toto, sDelimiteur$
End If

Set Rs = Con.CloseRecordSet(Rs)
Spreadsheet1.Range("A1").CurrentRegion.Replace Chr(13), ""
For Row = 2 To Me.Spreadsheet1.Range("A3").CurrentRegion.Rows.Count
       Me.Spreadsheet1.Rows(Row).Interior.Color = ChoixCouleur(Me.Spreadsheet1.Cells(Row, Me.Spreadsheet1.Range("A3").CurrentRegion.Columns.Count), False)
       
        Me.Spreadsheet1.Range("j" & Row & ":l" & Row).Font.Underline = 1
        Me.Spreadsheet1.Range("j" & Row & ":l" & Row).Font.Color = 16711680
Next
Me.Spreadsheet1.Columns(Me.Spreadsheet1.Range("A1").CurrentRegion.Columns.Count).Clear
Me.Spreadsheet1.Range("A1").Select
Me.Spreadsheet1.ActiveSheet.AutoFilterMode = True
Me.Spreadsheet1.ActiveSheet.Range("A1").AutoFilter
Me.Spreadsheet1.ActiveSheet.Range("A1").CurrentRegion.ColumnWidth = 130
Me.Spreadsheet1.ActiveSheet.Range("A1").CurrentRegion.AutoFitColumns
Me.Spreadsheet1.Range("A1").AutoFitColumns
Me.Spreadsheet1.ActiveSheet.Protection.Enabled = True
End Sub

Private Sub Spreadsheet1_DblClick(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Rs As Recordset
Dim Sql As String
Dim Splittxt
Dim NumChrono As String
Dim Process As String
Dim I As Integer
Dim Fichier As String
Dim Fso As FileSystemObject
Dim Row As Long
Dim Col As Long
Row = Me.Spreadsheet1.ActiveCell.Row
Col = Me.Spreadsheet1.ActiveCell.Column

If EventInfo.Range.Row < 2 Then Exit Sub
Set TableauPath = funPath
Splittxt = Split(Trim("" & EventInfo.Range) & "____", "_")
NumChrono = ""
For I = 0 To 3
    NumChrono = NumChrono & Splittxt(I) & "_"
Next
NumChrono = Left(NumChrono, Len(NumChrono) - 1)
Select Case EventInfo.Range.Column
        Case 10
            If PathAppliAutocad = "ERR" Then
                MsgBox "L'exécutable d'Autocad n'a pas été trouvée"
            Else
                Sql = "SELECT T_indiceProjet.PlAutoCadSave "
                Sql = Sql & "FROM T_indiceProjet "
                Sql = Sql & "WHERE T_indiceProjet.PL='" & NumChrono & "'  "
                Sql = Sql & "AND T_indiceProjet.PL_Indice='" & Splittxt(4) & "'  "
                Sql = Sql & "AND T_indiceProjet.PlAutoCadSave Is Not Null "
                Sql = Sql & "AND T_indiceProjet.IdStatus=" & Status(Me.Spreadsheet1.Cells(EventInfo.Range.Row, EventInfo.Range.Column).Interior.Color) & ";"

                Set Rs = Con.OpenRecordSet(Sql)
                If Rs.EOF = False Then
                    Fichier = DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("PathArchiveAutocad"))
                    Fichier = DefinirChemienComplet(Fichier, Trim("" & Rs!PlAutoCadSave)) & ".dwg"
                   
                    Process = PathAppliAutocad & " " '"C:\Program Files\AutoCAD 2002 Fra\acad.exe "
                    Set Fso = New FileSystemObject
                    If Fso.FileExists(Fichier) = True Then
                       StratProcess Process, Fichier
                    Else
                       MsgBox "Fichier Introuvable"
                    End If
                 Else
                       MsgBox "Fichier Introuvable"
                End If
                Set Fso = Nothing
                Set Rs = Con.CloseRecordSet(Rs)
             End If
        Case 11
            If PathAppliAutocad = "ERR" Then
                MsgBox "L'exécutable d'Autocad n'a pas été trouvée"
            Else
            Sql = "SELECT T_indiceProjet.OuAutoCadSave "
            Sql = Sql & "FROM T_indiceProjet "
            Sql = Sql & "WHERE T_indiceProjet.Ou='" & NumChrono & "'  "
            Sql = Sql & "AND T_indiceProjet.Ou_Indice='" & Splittxt(4) & "'  "
            Sql = Sql & "AND T_indiceProjet.ouAutoCadSave Is Not Null;"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = False Then
                Fichier = DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("PathArchiveAutocad"))
                Fichier = DefinirChemienComplet(Fichier, Trim("" & Rs!OUAutoCadSave)) & ".dwg"
                Process = PathAppliAutocad & " " '"C:\Program Files\AutoCAD 2002 Fra\acad.exe "
                
               
                Set Fso = New FileSystemObject
                If Fso.FileExists(Fichier) = True Then
                  StratProcess Process, Fichier
                Else
                   MsgBox "Fichier Introuvable"
                End If
             Else
                   MsgBox "Fichier Introuvable"
            End If
            Set Fso = Nothing
            Set Rs = Con.CloseRecordSet(Rs)
            
            End If
        Case 12
            If PathAppliExcel = "ERR" Then
                MsgBox "L'exécutable d'Excel n'a pas été trouvée"
            Else
                Sql = "SELECT T_indiceProjet.LiAutoCadSave  "
                Sql = Sql & "FROM T_indiceProjet "
                Sql = Sql & "WHERE T_indiceProjet.Li='" & NumChrono & "'  "
                Sql = Sql & "AND T_indiceProjet.Li_Indice='" & Splittxt(4) & "'  "
                Sql = Sql & "AND T_indiceProjet.LiAutoCadSave Is Not Null;"
                Set Rs = Con.OpenRecordSet(Sql)
                If Rs.EOF = False Then
                    Fichier = DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("PathArchiveAutocad"))
                    Fichier = DefinirChemienComplet(Fichier, Trim("" & Rs!LiAutoCadSave)) & ".xls"
                    Process = PathAppliExcel & " "
                    Set Fso = New FileSystemObject
                    If Fso.FileExists(Fichier) = True Then
                       Shell Process & Fichier, vbMaximizedFocus
                    Else
                       MsgBox "Fichier Introuvable"
                    End If
                 Else
                       MsgBox "Fichier Introuvable"
                End If
                Set Fso = Nothing
                Set Rs = Con.CloseRecordSet(Rs)
             End If
        
End Select
End Sub
Function Status(Couleur) As Long
 Select Case Couleur
   Case 12632256
        Status = 0
    Case 16777164
        Status = 1
    Case 10079487
        Status = 2
    Case 13434828
        Status = 3
    Case &HFFC0FF
        Status = 4
   End Select
End Function
     
Private Sub Spreadsheet1_MouseOver(ByVal EventInfo As OWC.SpreadsheetEventInfo)
If EventInfo.Range.Row > 9 And EventInfo.Range.Row < 12 Then
'    Spreadsheet1.
Else
End If
End Sub
