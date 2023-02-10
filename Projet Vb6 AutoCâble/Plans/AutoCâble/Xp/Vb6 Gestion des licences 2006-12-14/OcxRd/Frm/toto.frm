VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form Form1 
   Caption         =   "Rechercher pièce AutoCâble :"
   ClientHeight    =   11415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16980
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   16980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   11295
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   16815
      Begin VB.CommandButton CommandButton2 
         Caption         =   "&Annuler"
         Height          =   375
         Left            =   12600
         TabIndex        =   2
         Top             =   10320
         Width           =   3135
      End
      Begin VB.CommandButton CommandButton1 
         Caption         =   "&Valider"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   10560
         Width           =   3135
      End
      Begin OWC.Spreadsheet Spreadsheet1 
         Height          =   6210
         Left            =   0
         TabIndex        =   3
         Top             =   3675
         Width           =   16815
         HTMLURL         =   ""
         HTMLData        =   $"toto.frx":0000
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
      Begin VB.Image Image1 
         Height          =   855
         Left            =   14640
         Picture         =   "toto.frx":1242
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
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
         TabIndex        =   36
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Archive"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   35
         Top             =   3360
         Width           =   540
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "VAL"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   34
         Top             =   3360
         Width           =   300
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "MOD"
         Height          =   195
         Index           =   1
         Left            =   1170
         TabIndex        =   33
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "CRE"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   32
         Top             =   3360
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
         TabIndex        =   31
         Top             =   3360
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
         TabIndex        =   30
         Top             =   3360
         Width           =   195
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   915
         TabIndex        =   29
         Top             =   3360
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   3360
         Width           =   195
      End
      Begin VB.Label txt4 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   1215
         Left            =   1560
         TabIndex        =   27
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Ensemble"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label txt12 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   10440
         TabIndex        =   25
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label txt8 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   6000
         TabIndex        =   24
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   " Approbateur"
         Height          =   375
         Left            =   9000
         TabIndex        =   23
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Liste"
         Height          =   375
         Left            =   4440
         TabIndex        =   22
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label txt11 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   10440
         TabIndex        =   21
         Top             =   1605
         Width           =   2775
      End
      Begin VB.Label txt10 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   10440
         TabIndex        =   20
         Top             =   915
         Width           =   2775
      End
      Begin VB.Label txt9 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   10440
         TabIndex        =   19
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label txt7 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   6000
         TabIndex        =   18
         Top             =   1605
         Width           =   2775
      End
      Begin VB.Label txt6 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   6000
         TabIndex        =   17
         Top             =   915
         Width           =   2775
      End
      Begin VB.Label txt5 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   6000
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label txt3 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   1035
         Width           =   2775
      End
      Begin VB.Label txt2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   645
         Width           =   2775
      End
      Begin VB.Label txt1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label11 
         Caption         =   " Vérificateur "
         Height          =   375
         Left            =   9000
         TabIndex        =   12
         Top             =   1605
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Dessinateur"
         Height          =   375
         Left            =   9000
         TabIndex        =   11
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Client"
         Height          =   375
         Left            =   9000
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Outil"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   1605
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Plan"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Pièce"
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Equipement"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Vague"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Projet"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New Ado
Private TableauAttribus(12) As String
Public Function MyRetourTableau()
MyRetourTableau = TableauAttribus
End Function
Private Sub CommandButton1_Click()
IdFils = 0
Dim Sql As String
Dim Rs As Object
Dim Pere As Long

MyAnnuler = False




If Trim("" & txt1) <> "" Then
For i = 1 To 12

   TableauAttribus(i) = Controls("txt" & CStr(i)).Caption
    

Next i
 TableauAttribus(0) = Controls("txt" & CStr(1)).Tag

End If
Fin:
Noquite = False

Me.Hide


   
End Sub

Private Sub CommandButton2_Click()
IdFils = 0
Dim Sql As String
Dim Rs As Object
Dim Pere As Long

MyAnnuler = True




For i = 1 To 12

   TableauAttribus(i) = ""
    

Next i
 TableauAttribus(0) = ""

Fin:
Noquite = False

Me.Hide


End Sub


Private Sub Form_Load()
LoadParam
End Sub

Private Sub Form_Unload(Cancel As Integer)
Con.CloseConnection
End Sub

Private Sub Spreadsheet1_DblClick(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Row = Spreadsheet1.ActiveCell.Row
Dim Ofset As Long
strStatus = ""
Ofset = 0
PlanArchive = False
    If Row > 1 Then
        For i = 1 To 12
        If i = 5 Then Ofset = Ofset + 2
            Controls("TXT" & CStr(i)).Caption = "" & Spreadsheet1.Cells(Row, i + Ofset)
             Controls("TXT" & CStr(i)).Tag = "" & Spreadsheet1.Cells(Row, 15)
             Debug.Print i & " " & Spreadsheet1.Cells(1, i + Ofset)
             Select Case Spreadsheet1.Cells(Row, 1).Interior.Color
                  
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
        Next i
        For i = 1 To 20
            ValeurTableau(i, 0) = "" & Spreadsheet1.Cells(1, i)
            ValeurTableau(i, 1) = "" & Spreadsheet1.Cells(Row, i)
        Next
    Else
    
        For i = 1 To 12
            Controls("TXT" & CStr(i)).Caption = ""
             Controls("TXT" & CStr(i)).Tag = "0"
        Next i
        For i = 1 To 20
            ValeurTableau(i, 0) = ""
            ValeurTableau(i, 1) = ""
        Next
    End If
MyTableau = ValeurTableau
End Sub

Private Function LoadParam()
On Error Resume Next
Dim Sql As String
Dim Rs
Dim IndexRow As Long
Dim IndexCol As Long
Dim OfsetCol As Long
If Con.OpenConnetion(Mydb) = False Then GoTo Error
boolTxts = boolTxt
IndexRow = 1
IndexCol = 0
OfsetCol = 1






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






Set Rs = Con.OpenRecordSet(Sql)

DoEvents
Rs.Filter = MyFiltre
Spreadsheet1.ActiveSheet.Protection.Enabled = False

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
    
    Spreadsheet1.ActiveSheet.Range("A2").ParseText _
    toto, sDelimiteur$

End If

Set Rs = Con.CloseRecordSet(Rs)

Dim myrange
Set myrange = Spreadsheet1.ActiveSheet.Range("A2").CurrentRegion
myrange.AutoFitColumns
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
For i = 2 To myrange.Rows.Count
aa = Split(Trim("" & myrange(i, 7) & "____"), "_")
myrange(i, 6) = aa(3)
'             Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & aa(1))
Spreadsheet1.ActiveSheet.Rows(i).Interior.Color = ChoixCouleur(Val(myrange(i, 16)))
Next
Spreadsheet1.ActiveSheet.Protection.Enabled = True
Set myrange = Nothing
GoTo Fin
Error:
Me.Hide
Fin:
End Function
Function ChoixCouleur(Mode As Long, Optional BoolExcel As Boolean)
   
  If BoolExcel = False Then
   Select Case Mode
   Case 0
        ChoixCouleur = 12632256
    Case 1
        ChoixCouleur = 16777164
    Case 2
    ChoixCouleur = 10079487
    Case 3
        ChoixCouleur = 13434828
    Case 4
        ChoixCouleur = &HFFC0FF
   End Select

Else
    Select Case Mode
    Case 0
        ChoixCouleur = 15
    Case 1
        ChoixCouleur = 34
    Case 2
    ChoixCouleur = 40
    Case 3
        ChoixCouleur = 35
    Case 4
        ChoixCouleur = 38
   End Select
End If
End Function

Private Sub UserForm_Click()

'**************************************************************************************
'            permet de rechercher un proget dans la base de données AutoCâble
'......................................................................................
'                               appel du formulaire

'Private Sub Cherche_Click()
'Dim MyTableau
'MyTableau = UserForm6.Charge("ApprouveDate<> null and Archiver=false and IdStatus<4 ", "\\Autocable\Autocable Access\AutoCable.mdb")
'Unload UserForm6
'End Sub
'......................................................................................
' Les parametre d'entrés
'......................................................................................
'   Filtre permet d'éfectuer une restriction sur la recherche
'   "ApprouveDate<> null and Archiver=false and IdStatus<4 "
'   dans cet exp. on filtre sur les proget qui on été approuvés
'   sur l'indice en cours
'   qui ne sont pas archivés
'......................................................................................
'   L'emplacement de la base de données ACESS
'......................................................................................
'   "\\Autocable\Autocable Access\AutoCable.mdb"
'......................................................................................
' Sortie:
'......................................................................................
' Les données de sortie sont sous forme de tableau MyTableau dans notre Exp.
' Valeur du tableau index de 1 à 12 :
'1 Projet
'2 Vagues
'3 Equipements
'4 Ensembles
'5 Pièces
'6 Plans
'7 Outils
'8 Listes
'9 Clients
'10 Déssinateurs
'11 Vérificateur
'12  Approbateur
'**************************************************************************************


End Sub






