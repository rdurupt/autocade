VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form CherchPices 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chercher Pièces :"
   ClientHeight    =   10875
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   18510
   Icon            =   "CherchPices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   18510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   5970
      Left            =   0
      TabIndex        =   33
      Top             =   3915
      Width           =   18510
      HTMLURL         =   ""
      HTMLData        =   $"CherchPices.frx":030A
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
      Left            =   16200
      TabIndex        =   36
      Top             =   120
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "CherchPices.frx":14F5
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   9840
      TabIndex        =   35
      Top             =   10200
      Width           =   3135
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "&Valider"
      Height          =   375
      Left            =   3480
      TabIndex        =   34
      Top             =   10200
      Width           =   3135
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
      TabIndex        =   32
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Archive"
      Height          =   195
      Index           =   3
      Left            =   2520
      TabIndex        =   31
      Top             =   3600
      Width           =   540
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "VAL"
      Height          =   195
      Index           =   2
      Left            =   1920
      TabIndex        =   30
      Top             =   3600
      Width           =   300
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "MOD"
      Height          =   195
      Index           =   1
      Left            =   1170
      TabIndex        =   29
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "CRE"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   28
      Top             =   3600
      Width           =   330
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   27
      Top             =   3600
      Width           =   195
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1605
      TabIndex        =   26
      Top             =   3600
      Width           =   195
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   915
      TabIndex        =   25
      Top             =   3600
      Width           =   195
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   3600
      Width           =   195
   End
   Begin VB.Label txt4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   1215
      Left            =   1680
      TabIndex        =   23
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Ensemble"
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label txt12 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   10560
      TabIndex        =   21
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label txt8 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label12 
      Caption         =   " Approbateur"
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Liste"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label txt11 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   10560
      TabIndex        =   17
      Top             =   1600
      Width           =   2775
   End
   Begin VB.Label txt10 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   10560
      TabIndex        =   16
      Top             =   920
      Width           =   2775
   End
   Begin VB.Label txt9 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   10560
      TabIndex        =   15
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label txt7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   1600
      Width           =   2775
   End
   Begin VB.Label txt6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   920
      Width           =   2775
   End
   Begin VB.Label txt5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label txt3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   1040
      Width           =   2775
   End
   Begin VB.Label txt2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   640
      Width           =   2775
   End
   Begin VB.Label txt1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label11 
      Caption         =   " Vérificateur "
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   1600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Dessinateur"
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   920
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Client"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Outil"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Plan"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   920
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Pièce"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Equipement"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Vague"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Projet"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "CherchPices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyFormCible As Object
Dim Filtre As String
Dim boolTxts As Boolean
Dim Noquite As Boolean
Public Annuler As Boolean
Dim MyBooolBloque As Boolean
Dim MyAvecForm As Boolean



Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
IdFils = 0
Dim Sql As String
Dim Rs As Recordset
Dim Pere As Long
Annuler = False
If boolTxts = True Then
       MyFormCible.Tag = Me.Controls("txt" & CStr(1)).Tag
       GoTo Fin
End If
 If Trim("" & Me.Controls("txt" & CStr(1)).Tag) = "" Then
       CommandButton2_Click
       Exit Sub
End If
Sql = "UPDATE T_indiceProjet SET T_indiceProjet.UserName = Null "
Sql = Sql & "WHERE T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "';"
Con.Execute Sql

Sql = "SELECT T_indiceProjet.Id, T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Controls("txt1").Tag & ";"

Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Pere = Rs!Pere
End If
Set Rs = Con.CloseRecordSet(Rs)

Sql = "UPDATE T_indiceProjet SET T_indiceProjet.UserName = '" & Replace(Machine, "'", "''") & "' "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Controls("txt1").Tag & " OR T_indiceProjet.Pere=" & Me.Controls("txt1").Tag & " "
If Pere <> 0 Then
Sql = Sql & "or T_indiceProjet.Id=" & Pere
End If
Sql = Sql & ";"
Con.Execute Sql

If Trim("" & Me.txt1) <> "" Then
For I = 1 To 12

    MyFormCible.Controls("txt" & CStr(I)).Caption = Me.Controls("txt" & CStr(I)).Caption
    MyFormCible.Controls("txt" & CStr(I)).Tag = Me.Controls("txt" & CStr(I)).Tag

Next I
Else
CommandButton2_Click
Exit Sub
End If
Fin:
Noquite = False

If MyBooolBloque = True Then MyFormCible.BooolBloque = MyBooolBloque
'If  = False Then MyFormCible. =
If Trim("" & txt3.Tag) <> "" Then
If MyAvecForm = True Then
    MyFormCible.chargement Me, MyBooolBloque
 Else
    Me.Hide
 End If
Else
    If MyAvecForm = True Then
    Unload Me
    Else
        Me.Hide
    End If
End If

End Sub

Private Sub CommandButton2_Click()

IdFils = 0
Noquite = False
If MyAvecForm = True Then
    
    Unload MyFormCible
    Unload Me
Else
    Me.Hide
End If
End Sub

Private Sub Form_Load()
Annuler = True
'Me.CommandButton1.Anchor = "Bottom, Right"

End Sub

Private Sub Spreadsheet1_DblClick(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Row = Me.Spreadsheet1.ActiveCell.Row
Dim Ofset As Long
strStatus = ""
Ofset = 0
PlanArchive = False
    If Row > 1 Then
        For I = 1 To 12
        If I = 5 Then Ofset = Ofset + 2
            Me.Controls("TXT" & CStr(I)).Caption = "" & Me.Spreadsheet1.Cells(Row, I + Ofset)
             Me.Controls("TXT" & CStr(I)).Tag = "" & Me.Spreadsheet1.Cells(Row, 15)
             Select Case Me.Spreadsheet1.Cells(Row, 1).Interior.Color
                  
                  Case 16777164
                          strStatus = "CRE"
                  Case 10079487
                          strStatus = "MOD"
                
                  Case 13434828
                        strStatus = "VAL"
                   Case &HFFC0FF
                        strStatus = "VAL"
                        PlanArchive = True
        
             End Select
        Next I
    
    Else
    
        For I = 1 To 12
            Me.Controls("TXT" & CStr(I)).Caption = ""
             Me.Controls("TXT" & CStr(I)).Tag = "0"
        Next I
    End If
End Sub

Public Sub Charge(MyForm As Object, Optional Filtre As String, Optional boolTxt As Boolean, Optional boolArchive As Boolean, Optional BooolBloque As Boolean, Optional AvecForm As Boolean)
Dim Sql As String
Dim Rs As Recordset
Dim IndexRow As Long
Dim IndexCol As Long
Dim OfsetCol As Long

MyAvecForm = AvecForm
boolTxts = boolTxt
IndexRow = 1
IndexCol = 0
OfsetCol = 1
MyBooolBloque = BooolBloque
Sql = "SELECT SelectProjets.* "
Sql = Sql & "FROM (SelectProjets LEFT JOIN T_Job ON SelectProjets.Id = T_Job.Id_Piece) "
Sql = Sql & "LEFT JOIN T_Job AS T_Job_1 ON SelectProjets.Pere = T_Job_1.Id_Piece "
Sql = Sql & "Where T_Job.Id_Piece Is Null "
'Sql = Sql & "And T_Job_1.Id_Piece Is Null "
Sql = Sql & "Or T_Job.FinTraitement = True "
Sql = Sql & "ORDER BY SelectProjets.CleAc DESC , SelectProjets.PI DESC;"

'
Sql = "SELECT distinct SelectProjets.* "
Sql = Sql & "FROM ((SelectProjets LEFT JOIN T_Job ON SelectProjets.Id = T_Job.Id_Piece)  "
Sql = Sql & "LEFT JOIN T_Job AS T_Job_1 ON SelectProjets.Pere = T_Job_1.Id_Piece) LEFT JOIN T_indiceProjet  "
Sql = Sql & "ON SelectProjets.Id = T_indiceProjet.Id "
Sql = Sql & "WHERE (T_Job.Id_Piece Is Null or T_Job.FinTraitement=True) "
Sql = Sql & "AND (T_Job_1.Id_Piece Is Null or T_Job_1.FinTraitement=True) "
Sql = Sql & "AND (T_indiceProjet.UserName='robert.durupt' Or T_indiceProjet.UserName Is Null ) "
'Sql = Sql & "OR (((T_Job.FinTraitement)=True)  "
'Sql = Sql & "AND ((T_indiceProjet.UserName)='robert.durupt'  "
'Sql = Sql & "Or (T_indiceProjet.UserName) Is Null)) "
Sql = Sql & "ORDER BY SelectProjets.CleAc DESC , SelectProjets.PI DESC;"


Sql = "SELECT SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement, SelectProjets.Ensemble,  "
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
'Sql = Sql & "WHERE (T_Job.Id_Piece Is Null   "
'Sql = Sql & "AND T_Job_1.Id_Piece Is Null   "
'Sql = Sql & "AND (T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "'   "
'Sql = Sql & "Or T_indiceProjet.UserName Is Null))   "
'Sql = Sql & "OR ((T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "'   "
'Sql = Sql & "Or T_indiceProjet.UserName Is Null)   "
'Sql = Sql & "AND (T_Job.FinTraitement=True or T_Job_1.FinTraitement=True))   "
''Sql = Sql & "AND ((T_Job_1.FinTraitement)=True))  "
'Sql = Sql & "ORDER BY SelectProjets.CleAc DESC , SelectProjets.PI DESC;"

 Sql = "SELECT SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement, SelectProjets.Ensemble, SelectProjets.CleAc,  "
Sql = Sql & "0 AS chrono, [SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice] AS Expr1,  "
Sql = Sql & "[SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice] AS Expr2, [SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice] AS Expr3,  "
Sql = Sql & "[SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice] AS Expr4, SelectProjets.Client, SelectProjets.DessineNOM,  "
Sql = Sql & "SelectProjets.VerifieNom, SelectProjets.ApprouveNom, SelectProjets.Id, SelectProjets.IdStatus, SelectProjets.NbErr,  "
Sql = Sql & "SelectProjets.LiAutoCadSave, SelectProjets.VerifieDate, SelectProjets.Archiver, SelectProjets.PI_Indice,  "
Sql = Sql & "SelectProjets.PL_Indice, SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere, SelectProjets.PlOk,  "
Sql = Sql & "SelectProjets.OuOk "
Sql = Sql & "FROM (SelectProjets LEFT JOIN T_Job ON SelectProjets.Id = T_Job.Id_Piece)  "
Sql = Sql & "LEFT JOIN T_indiceProjet ON SelectProjets.Id = T_indiceProjet.Id "
Sql = Sql & "WHERE (T_Job.Id_Piece Is Null  "
Sql = Sql & "Or [FinTraitement]=True)  "
Sql = Sql & "AND (T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "'   "
Sql = Sql & "Or T_indiceProjet.UserName Is Null) "
Sql = Sql & "GROUP BY SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement,  "
Sql = Sql & "SelectProjets.Ensemble, SelectProjets.CleAc, 0, [SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice],  "
Sql = Sql & "[SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice], [SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice],  "
Sql = Sql & "[SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice], SelectProjets.Client, SelectProjets.DessineNOM,  "
Sql = Sql & "SelectProjets.VerifieNom, SelectProjets.ApprouveNom, SelectProjets.Id, SelectProjets.IdStatus, SelectProjets.NbErr,  "
Sql = Sql & "SelectProjets.LiAutoCadSave, SelectProjets.VerifieDate, SelectProjets.Archiver, SelectProjets.PI_Indice, SelectProjets.PL_Indice,  "
Sql = Sql & "SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere, SelectProjets.PlOk, SelectProjets.OuOk, SelectProjets.PI "
Sql = Sql & "ORDER BY SelectProjets.CleAc DESC , SelectProjets.PI DESC;"
'Machine = "PIP-IACOB-01"
Sql = "SELECT SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement,  "
Sql = Sql & "SelectProjets.Ensemble, SelectProjets.CleAc, 0 AS chrono,  "
Sql = Sql & "[SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice] AS Expr1,  "
Sql = Sql & "[SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice] AS Expr2,  "
Sql = Sql & "[SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice] AS Expr3,  "
Sql = Sql & "[SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice] AS Expr4,  "
Sql = Sql & "SelectProjets.Client, SelectProjets.DessineNOM, SelectProjets.VerifieNom,  "
Sql = Sql & "SelectProjets.ApprouveNom, SelectProjets.Id, SelectProjets.IdStatus,  "
Sql = Sql & "SelectProjets.NbErr, SelectProjets.LiAutoCadSave, SelectProjets.VerifieDate,  "
Sql = Sql & "SelectProjets.Archiver, SelectProjets.PI_Indice, SelectProjets.PL_Indice,  "
Sql = Sql & "SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere,  "
Sql = Sql & "SelectProjets.PlOk, SelectProjets.OuOk "
Sql = Sql & "FROM ((SelectProjets LEFT JOIN T_Job ON SelectProjets.Id = T_Job.Id_Piece)  "
Sql = Sql & "LEFT JOIN T_indiceProjet ON SelectProjets.Id = T_indiceProjet.Id)  "
Sql = Sql & "LEFT JOIN T_Job AS T_Job_1 ON SelectProjets.Id = T_Job_1.Id_Fils "
Sql = Sql & "WHERE (T_Job.Id_Piece Is Null Or [T_Job].[FinTraitement]=True)  "
Sql = Sql & "AND (T_Job_1.Id_Fils Is Null Or [T_Job_1].[FinTraitement]=True)  "
Sql = Sql & "AND (T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "' Or T_indiceProjet.UserName Is Null) "
Sql = Sql & "GROUP BY SelectProjets.Projet, SelectProjets.Vague,  "
Sql = Sql & "SelectProjets.Equipement, SelectProjets.Ensemble, SelectProjets.CleAc,  "
Sql = Sql & "0, [SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice],  "
Sql = Sql & "[SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice],  "
Sql = Sql & "[SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice],  "
Sql = Sql & "[SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice],  "
Sql = Sql & "SelectProjets.Client, SelectProjets.DessineNOM, SelectProjets.VerifieNom,  "
Sql = Sql & "SelectProjets.ApprouveNom, SelectProjets.Id, SelectProjets.IdStatus,  "
Sql = Sql & "SelectProjets.NbErr, SelectProjets.LiAutoCadSave, SelectProjets.VerifieDate,  "
Sql = Sql & "SelectProjets.Archiver, SelectProjets.PI_Indice, SelectProjets.PL_Indice,  "
Sql = Sql & "SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere,  "
Sql = Sql & "SelectProjets.PlOk, SelectProjets.OuOk, SelectProjets.PI "
Sql = Sql & "ORDER BY SelectProjets.CleAc DESC , SelectProjets.PI DESC;"

Sql = "SELECT SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement, SelectProjets.Ensemble, SelectProjets.CleAc, 0 AS "
Sql = Sql & "chrono, [SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice] AS Expr1, "
Sql = Sql & "[SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice] AS Expr2, [SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice] AS Expr3, "
Sql = Sql & "[SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice] AS Expr4, SelectProjets.Client, SelectProjets.DessineNOM, "
Sql = Sql & "SelectProjets.VerifieNom, SelectProjets.ApprouveNom, SelectProjets.Id, SelectProjets.IdStatus, SelectProjets.NbErr, "
Sql = Sql & "SelectProjets.LiAutoCadSave, SelectProjets.VerifieDate, SelectProjets.Archiver, SelectProjets.PI_Indice, SelectProjets.PL_Indice, "
Sql = Sql & "SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere, SelectProjets.PlOk, SelectProjets.OuOk, SelectProjets.PI, "
Sql = Sql & "T_indiceProjet.UserName, FrmJob.Id_Piece, FrmJob2.Id_Piece "
Sql = Sql & "FROM ((SelectProjets LEFT JOIN T_indiceProjet ON SelectProjets.Id = T_indiceProjet.Id) LEFT JOIN "

Sql = Sql & "(SELECT T_Job.Id_Piece, T_Job.Id_Fils, T_Job.FinTraitement "
Sql = Sql & "FROM T_Job "
Sql = Sql & "WHERE T_Job.FinTraitement=False) AS FrmJob "

Sql = Sql & "ON SelectProjets.Id = FrmJob.Id_Piece) LEFT JOIN "

Sql = Sql & "(SELECT T_Job.Id_Piece, T_Job.Id_Fils, T_Job.FinTraitement "
Sql = Sql & "FROM T_Job"
Sql = Sql & "WHERE T_Job.FinTraitement=False) AS FrmJob2 "

Sql = Sql & "ON SelectProjets.Id = FrmJob2.Id_Piece "
Sql = Sql & "WHERE (T_indiceProjet.UserName='AUTOCAD' Or T_indiceProjet.UserName Is Null) "
Sql = Sql & "AND FrmJob.Id_Piece Is Null AND FrmJob2.Id_Piece Is Null;"



Sql = "SELECT SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement, SelectProjets.Ensemble, SelectProjets.CleAc, 0 AS chrono,  "
Sql = Sql & "[SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice] AS Expr1, [SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice] AS Expr2,  "
Sql = Sql & "[SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice] AS Expr3, [SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice] AS Expr4,  "
Sql = Sql & "SelectProjets.Client, SelectProjets.DessineNOM, SelectProjets.VerifieNom, SelectProjets.ApprouveNom, SelectProjets.Id,  "
Sql = Sql & "SelectProjets.IdStatus, SelectProjets.NbErr, SelectProjets.LiAutoCadSave, SelectProjets.VerifieDate, SelectProjets.Archiver,  "
Sql = Sql & "SelectProjets.PI_Indice, SelectProjets.PL_Indice, SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere,  "
Sql = Sql & "SelectProjets.PlOk, SelectProjets.OuOk, SelectProjets.PI, T_indiceProjet.UserName, FrmJob.Id_Piece, FrmJob2.Id_Piece "
Sql = Sql & "FROM ((SelectProjets LEFT JOIN T_indiceProjet ON SelectProjets.Id = T_indiceProjet.Id) LEFT JOIN "

Sql = Sql & "(SELECT T_Job.Id_Piece, T_Job.Id_Fils, T_Job.FinTraitement FROM T_Job WHERE T_Job.FinTraitement=False ) AS FrmJob  "

Sql = Sql & "ON SelectProjets.Id = FrmJob.Id_Piece) LEFT JOIN "

Sql = Sql & "(SELECT T_Job.Id_Piece, T_Job.Id_Fils, T_Job.FinTraitement FROM T_Job WHERE T_Job.FinTraitement=False) AS FrmJob2 "

Sql = Sql & "ON SelectProjets.Pere= FrmJob2.Id_Piece "
Sql = Sql & "WHERE (T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "' Or T_indiceProjet.UserName Is Null)  "
Sql = Sql & "AND FrmJob.Id_Piece Is Null AND FrmJob2.Id_Piece Is Null "
Sql = Sql & "ORDER BY SelectProjets.CleAc DESC , SelectProjets.PI DESC;"


'
'Sql = "SELECT SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement, SelectProjets.Ensemble, SelectProjets.CleAc,  "
'Sql = Sql & "0 AS chrono, [SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice] AS Expr1,  "
'Sql = Sql & "[SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice] AS Expr2, [SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice] AS "
'Sql = Sql & "Expr3, [SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice] AS Expr4, SelectProjets.Client, SelectProjets.DessineNOM,  "
'Sql = Sql & "SelectProjets.VerifieNom, SelectProjets.ApprouveNom, SelectProjets.Id, SelectProjets.IdStatus, SelectProjets.NbErr,  "
'Sql = Sql & "SelectProjets.LiAutoCadSave, SelectProjets.VerifieDate, SelectProjets.Archiver, SelectProjets.PI_Indice,  "
'Sql = Sql & "SelectProjets.PL_Indice, SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere, SelectProjets.PlOk,  "
'Sql = Sql & "SelectProjets.OuOk, SelectProjets.PI, T_indiceProjet.UserName, FrmJob.Id_Piece "
'Sql = Sql & "FROM (SelectProjets LEFT JOIN T_indiceProjet ON SelectProjets.Id = T_indiceProjet.Id)  "
'Sql = Sql & "LEFT JOIN  "
'Sql = Sql & "(SELECT T_Job.Id_Piece, T_Job.Id_Fils, T_Job.FinTraitement FROM T_Job WHERE T_Job.FinTraitement=False ) AS FrmJob  "
'
'Sql = Sql & "ON  SelectProjets.Id = FrmJob.Id_Piece "
'Sql = Sql & "WHERE (SelectProjets.Pere<>Val('' & [FrmJob].[Id_Piece]) Or SelectProjets.Pere=0)  "
'Sql = Sql & "AND (T_indiceProjet.UserName='AUTOCAD' Or T_indiceProjet.UserName Is Null)  "
'Sql = Sql & "AND FrmJob.Id_Piece Is Null;"
'



'Sql = "SELECT SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement, SelectProjets.Ensemble, SelectProjets.CleAc,  "
'Sql = Sql & "0 AS chrono, [SelectProjets].[PI] & '_' & [SelectProjets].[PI_Indice] AS Expr1,  "
'Sql = Sql & "[SelectProjets].[PL] & '_' & [SelectProjets].[PL_Indice] AS Expr2,  "
'Sql = Sql & "[SelectProjets].[OU] & '_' & [SelectProjets].[OU_Indice] AS Expr3,  "
'Sql = Sql & "[SelectProjets].[Li] & '_' & [SelectProjets].[LI_Indice] AS Expr4, SelectProjets.Client, SelectProjets.DessineNOM,  "
'Sql = Sql & "SelectProjets.VerifieNom, SelectProjets.ApprouveNom, SelectProjets.Id, SelectProjets.IdStatus, SelectProjets.NbErr,  "
'Sql = Sql & "SelectProjets.LiAutoCadSave, SelectProjets.VerifieDate, SelectProjets.Archiver, SelectProjets.PI_Indice,  "
'Sql = Sql & "SelectProjets.PL_Indice, SelectProjets.OU_Indice, SelectProjets.LI_Indice, SelectProjets.Pere, SelectProjets.PlOk,  "
'Sql = Sql & "SelectProjets.OuOk, SelectProjets.PI "
'Sql = Sql & "FROM (((((SelectProjets LEFT JOIN T_indiceProjet ON SelectProjets.Id = T_indiceProjet.Id) LEFT JOIN  "
'
'Sql = Sql & "(SELECT T_Job.Id_Piece, T_Job.Id_Fils, T_Job.FinTraitement FROM T_Job WHERE T_Job.FinTraitement=False) AS FrmJob  "
'
'Sql = Sql & "ON SelectProjets.Id = FrmJob.Id_Piece) LEFT JOIN  "
'
'Sql = Sql & "(SELECT T_Job.Id_Piece, T_Job.Id_Fils, T_Job.FinTraitement FROM T_Job WHERE T_Job.FinTraitement=False) AS FrmJob2  "
'
'Sql = Sql & "ON SelectProjets.Id = FrmJob2.Id_Piece) LEFT JOIN a_Piece_Ouverte_Par ON T_indiceProjet.Id = a_Piece_Ouverte_Par.Id)  "
'Sql = Sql & "LEFT JOIN T_indiceProjet AS T_indiceProjet_1 ON SelectProjets.Id = T_indiceProjet_1.Pere)  "
'Sql = Sql & "LEFT JOIN a_Piece_Ouverte_Par AS a_Piece_Ouverte_Par_1 ON T_indiceProjet_1.Id = a_Piece_Ouverte_Par_1.Id "
'Sql = Sql & "WHERE FrmJob.Id_Piece Is Null "
'Sql = Sql & "AND FrmJob2.Id_Piece Is Null "
'Sql = Sql & "AND (a_Piece_Ouverte_Par.Machine='autocad' Or a_Piece_Ouverte_Par.Machine Is Null)  "
'Sql = Sql & "AND (a_Piece_Ouverte_Par_1.Machine='autocad' Or a_Piece_Ouverte_Par_1.Machine Is Null) "
'Sql = Sql & "ORDER BY SelectProjets.CleAc DESC , SelectProjets.PI DESC;"




Set Rs = Con.OpenRecordSet(Sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
'While Rs.EOF = False
'IndexRow = IndexRow + 1
'OfsetCol = 1
'For IndexCol = 0 To Rs.Fields.Count - 13
'DoEvents
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = False
'    If IndexCol > 4 And IndexCol < 9 Then
'    If IndexCol = 5 Then
'           aa = Split(Trim("" & Rs.Fields(IndexCol) & "____"), "_")
'             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & aa(1))
'             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 12))
'            Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = True
'            OfsetCol = OfsetCol + 1
'        End If
'         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = False
'
'         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & Rs.Fields(IndexCol)) & "_" & Trim("" & Rs.Fields(IndexCol + 14))
'    Else
'
'         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & Rs.Fields(IndexCol))
'    End If
'
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 12))
'    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = True
'
'Next IndexCol
'
'Rs.MoveNext
'Wend
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

Set Rs = Con.CloseRecordSet(Rs)

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
For I = 2 To MyRange.Rows.Count
aa = Split(Trim("" & MyRange(I, 7) & "____"), "_")
MyRange(I, 6) = aa(3)
'             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & aa(1))
Me.Spreadsheet1.Rows(I).Interior.Color = ChoixCouleur(Val(MyRange(I, 16)))
Next
Spreadsheet1.ActiveSheet.Protection.Enabled = True
Set MyRange = Nothing
Me.Show vbModal
End Sub



Private Sub UserForm_Activate()
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub

