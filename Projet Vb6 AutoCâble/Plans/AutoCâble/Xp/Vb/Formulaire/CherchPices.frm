VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form CherchPices 
   Caption         =   "Chercher Pièces :"
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18510
   Icon            =   "CherchPices.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   18510
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
      MoveAfterReturn =   0   'False
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
         Picture         =   "CherchPices.frx":154C
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   9840
      TabIndex        =   35
      Top             =   10200
      Width           =   3135
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "Valider"
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



Private Sub ListBox1_Click()

End Sub

Private Sub CommandButton1_Click()
Annuler = False

If boolTxts = True Then
       MyFormCible.Tag = Me.Controls("txt" & CStr(1)).Tag
       GoTo Fin
End If
For i = 1 To 12

    MyFormCible.Controls("txt" & CStr(i)).Caption = Me.Controls("txt" & CStr(i)).Caption
    MyFormCible.Controls("txt" & CStr(i)).Tag = Me.Controls("txt" & CStr(i)).Tag

Next i
Fin:
Noquite = False
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Noquite = False
Me.Hide
End Sub

Private Sub Form_Load()
Annuler = True
End Sub

Private Sub Spreadsheet1_DblClick(ByVal EventInfo As OWC.SpreadsheetEventInfo)
Dim Row As Long
Row = Me.Spreadsheet1.ActiveCell.Row
Dim Ofset As Long
strStatus = ""
Ofset = 0
PlanArchive = False
    If Row > 1 Then
        For i = 1 To 12
        If i = 5 Then Ofset = Ofset + 1
            Me.Controls("TXT" & CStr(i)).Caption = "" & Me.Spreadsheet1.Cells(Row, i + Ofset)
             Me.Controls("TXT" & CStr(i)).Tag = "" & Me.Spreadsheet1.Cells(Row, 15)
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
        Next i
    
    Else
        For i = 1 To 13
            Me.Controls("TXT" & CStr(i)).Caption = ""
             Me.Controls("TXT" & CStr(i)).Tag = "0"
        Next i
    End If
End Sub

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
Sql = "SELECT SelectProjets.* "
Sql = Sql & "FROM SelectProjets; "
Set Rs = Con.OpenRecordSet(Sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
While Rs.EOF = False
IndexRow = IndexRow + 1
OfsetCol = 1
For IndexCol = 0 To Rs.Fields.Count - 11

    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = False
    If IndexCol > 4 And IndexCol < 9 Then
    If IndexCol = 5 Then
           aa = Split(Trim("" & Rs.Fields(IndexCol)), "_")
             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & aa(UBound(aa)))
             Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 10))
            Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = True
            OfsetCol = OfsetCol + 1
        End If
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = False
        
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 15))
    Else
        
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol) = Trim("" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 10))
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + OfsetCol).Locked = True

Next IndexCol

Rs.MoveNext
Wend
If boolArchive = True Then
    Sql = "SELECT Archive_SelectProjets.* "
Sql = Sql & "FROM Archive_SelectProjets; "
Set Rs = Con.OpenRecordSet(Sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
While Rs.EOF = False
IndexRow = IndexRow + 1
For IndexCol = 0 To Rs.Fields.Count - 12
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
    If IndexCol > 3 And IndexCol < 8 Then
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 16))
    Else
        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = &HFFC0FF
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = True

Next IndexCol

Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)

Dim Myrange
Set Myrange = Me.Spreadsheet1.Range("A1").CurrentRegion
Myrange.AutoFitColumns
Set Myrange = Nothing
Me.Show vbModal
End Sub

Function ChoixCouleur(Mode As Long) As Long
   
  
   Select Case Mode
    Case 1
        ChoixCouleur = 16777164
    Case 2
    ChoixCouleur = 10079487
    Case 3
        ChoixCouleur = 13434828
   End Select

End Function

Private Sub UserForm_Activate()
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub

