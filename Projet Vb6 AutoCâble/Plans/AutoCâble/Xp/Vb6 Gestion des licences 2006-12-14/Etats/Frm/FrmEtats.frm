VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FrmEtats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Générateur d'états (Menu):"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   ControlBox      =   0   'False
   FillStyle       =   3  'Vertical Line
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   4320
      Picture         =   "FrmEtats.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format de sauvegarde"
      Height          =   975
      Left            =   3120
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton IsPDF 
         Height          =   375
         Index           =   1
         Left            =   2280
         Picture         =   "FrmEtats.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.OptionButton IsPDF 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "FrmEtats.frx":0894
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   3480
      Picture         =   "FrmEtats.frx":0945
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Enregistrer"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtMenu 
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ComboBox LstEtat 
      Height          =   315
      Left            =   480
      TabIndex        =   9
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox LstEtatAs 
      Height          =   315
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Height          =   315
      Left            =   3480
      Picture         =   "FrmEtats.frx":1157
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   3960
      Picture         =   "FrmEtats.frx":11FC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Enregistrer"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "FrmEtats.frx":1A0E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Annuler"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quitter"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CheckBox MenuVisible 
      Alignment       =   1  'Right Justify
      Caption         =   "Visible"
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid DetailEtatat 
      Height          =   5415
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9551
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   100
      BackColor       =   -2147483628
      ForeColorSel    =   -2147483635
      BackColorBkg    =   14079702
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      MergeCells      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Menu 
      Alignment       =   2  'Center
      Caption         =   "Menu"
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Document"
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "FrmEtats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CreateDate As Date

Private Sub Load(Optional Value As String)
Dim Sql As String
Dim Rs As Recordset
Me.LstEtat.Clear

Sql = "SELECT T_ETATS.* From T_ETATS ORDER BY T_ETATS.EtatName;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Me.LstEtat.AddItem "" & Rs!EtatName
    If UCase(Value) = UCase("" & Rs!EtatName) Then
        Me.LstEtat.ListIndex = Me.LstEtat.ListCount - 1
    End If
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
End Sub

Private Sub Command1_Click()
'Set LstColecDoc = Nothing
'Set LstColecDoc = New Collection
'IndexTableauDocGen = 0
'Dim ColecDoc As New Collection
'Dim IndexTableauDoc As Long
'Dim TableDoc As GenerateurDoc
'Dim Fso As New FileSystemObject
'Dim I As Long
''ReDim TableDoc(0)
'Set FormBarGrah = Me
'' TableDoc(0).Menu = "Menu2"
'Set TableDoc = New GenerateurDoc
'ColecDoc.Add TableDoc, "Menu2"
'Set TableDoc = Nothing
'ColecDoc("Menu2").Menu = "Menu2"
'ColecDoc("Menu2").LoadColecDoc IndexTableauDoc, ColecDoc, TableDoc, "c:\durupt\testFinal"
'ColecDoc("Menu2").IsPDF = True
'For I = ColecDoc.Count To 1 Step -1
'    ColecDoc(I).SelectEtat I, ColecDoc, ColecDoc(I), 572, 573, ColecDoc(I).SaveAs
'Next
'For I = 2 To ColecDoc.Count
'    If Fso.FileExists(ColecDoc(I).SaveAs & ".Xls") = True Then Fso.DeleteFile ColecDoc(I).SaveAs & ".Xls"
'Next
Unload Me
End Sub

Private Sub Command2_Click()
Dim Sql As String
If Me.LstEtatAs.Tag <> "" Then
    If MsgBox("Voulez vous varaiment supprimer : " & Me.LstEtatAs, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Sql = "DELETE T_ETATS.* FROM T_ETATS WHERE T_ETATS.ID=" & Me.LstEtatAs.Tag & ";"
    Con.Execute Sql
    Load
 ChargeDetail
Command4_Click
End If
End Sub

Private Sub Command3_Click()
Dim Rs As Recordset
Dim Sql As String
If Trim("" & LstEtatAs) = "" Then Exit Sub

Sql = "SELECT T_ETATS.* From T_ETATS where T_ETATS.EtatName='" & MyReplace(Me.LstEtatAs) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    If Me.LstEtatAs.Tag <> "" Then
        If Rs!Id <> Val(Me.LstEtatAs.Tag) Then
            MsgBox "le nenu: " & LstEtatAs & " Existe déjà Mise à jour non effectuée."
        Else
            Sql = "UPDATE T_ETATS SET T_ETATS.EtatName = '" & MyReplace(Me.LstEtatAs) & "' WHERE T_ETATS.ID=" & Me.LstEtatAs.Tag & " ;"
            Con.Execute Sql
        End If
    Else
    End If
Else
     If Me.LstEtatAs.Tag <> "" Then
         Sql = "UPDATE T_ETATS SET T_ETATS.EtatName = '" & MyReplace(Me.LstEtatAs) & "' WHERE T_ETATS.ID=" & Me.LstEtatAs.Tag & " ;"
     Else
        Sql = "INSERT INTO T_ETATS ( EtatName ) values ( '" & MyReplace(Me.LstEtatAs) & "');"
     End If
    Con.Execute Sql
    
End If
Rs.Requery
 Me.Tag = Rs!Id
 Me.txtMenu = "" & Rs!Menu
 If Rs!Visible = True Then
    MenuVisible.Value = 1
 Else
    MenuVisible.Value = 0
 End If
Load LstEtatAs
 ChargeDetail
Command4_Click
End Sub

Private Sub Command4_Click()
Me.LstEtatAs = ""
Me.LstEtatAs.Tag = ""

End Sub

Private Sub Command5_Click()
Dim Rs As Recordset
Dim Sql As String
If Trim("" & LstEtat) = "" Then Exit Sub
Sql = "SELECT T_ETATS.* From T_ETATS where T_ETATS.EtatName='" & MyReplace(Me.LstEtat) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
Me.LstEtatAs = Me.LstEtat
Me.LstEtatAs.Tag = Rs!Id
End If
End Sub

Private Sub Command6_Click()
Dim Sql As String
Dim Rs As Recordset


    If Trim("" & Me.Tag) = "" Then MsgBox "Vous devez saisir ou sélectionner un nom de document": LstEtat.SetFocus: Exit Sub
     Sql = "Select T_ETATS.*  from T_ETATS "
    Sql = Sql & "WHERE T_ETATS.ID<>" & Me.Tag & " and T_ETATS.Menu = '" & MyReplace(txtMenu) & "';"
 Set Rs = Con.OpenRecordSet(Sql)
 If Rs.EOF = False Then
    MsgBox "Le Menu: " & txtMenu & " Existe déjà, Mise a jour non effectuée"
    Set Rs = Con.CloseRecordSet(Rs)
    LstEtat_Click
    Exit Sub
 End If
  Set Rs = Con.CloseRecordSet(Rs)
    Sql = "UPDATE T_ETATS SET T_ETATS.Menu = '" & MyReplace(txtMenu) & "'"
    Sql = Sql & "WHERE T_ETATS.ID=" & Me.Tag & ";"
Con.Execute Sql
End Sub

Private Sub Command7_Click()
Dim Sql As String
If Me.LstEtatAs.Tag <> "" Then
    If MsgBox("Voulez vous varaiment supprimer : " & Me.LstEtatAs, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Sql = "DELETE T_ETATS.* FROM T_ETATS WHERE T_ETATS.ID=" & Me.LstEtatAs.Tag & ";"
    Con.Execute Sql
    Load
 ChargeDetail
Command4_Click
txtMenu = ""
End If
End Sub

Private Sub DetailEtatat_DblClick()
If Trim("" & Me.Tag) = "" Then MsgBox "Vous devez saisir ou sélectionner un nom de document": LstEtat.SetFocus: Exit Sub
If Trim("" & txtMenu) = "" Then MsgBox "Vous devez saisir un nom de Menu": txtMenu.SetFocus: Exit Sub
Me.DetailEtatat.Col = 2
frmDetailOnglet.chargement Val(Trim("" & Me.DetailEtatat)), Me.Tag, CreateDate
Load LstEtat
End Sub

Private Sub DetailEtatat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Sql As String
If Button = 2 Then
Me.DetailEtatat.Col = 2
 If Trim("" & Me.DetailEtatat) = "" Then Exit Sub
Me.DetailEtatat.Col = 0
 If MsgBox("Voulez-vou supprimer la macro: " & Me.DetailEtatat, vbYesNo) = vbNo Then Exit Sub
 Me.DetailEtatat.Col = 2
 Sql = "DELETE T_Etats_Onglet.* From T_Etats_Onglet "
Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Me.DetailEtatat & ";"
Con.Execute Sql
LstEtat_Click
End If
End Sub

Private Sub Form_Load()
Dim aa
Me.DetailEtatat.ColWidth(0) = 1462
Me.DetailEtatat.ColWidth(0) = 1462
Me.DetailEtatat.Row = 0
Me.DetailEtatat.Col = 0
    Me.DetailEtatat = "Macro"
Me.DetailEtatat.Row = 0
Me.DetailEtatat.Col = 1
    Me.DetailEtatat = "Onglet"

Load
End Sub

Private Sub IsPDF_Click(Index As Integer)
If Trim("" & Me.Tag) = "" Then Exit Sub
Dim Sql As String
If IsPDF(0).Value = True Then
    Sql = "UPDATE T_ETATS SET T_ETATS.IsPDF = True "
Else
     Sql = "UPDATE T_ETATS SET T_ETATS.IsPDF = False "
End If
Sql = Sql & "WHERE T_ETATS.ID=" & Me.Tag & ";"
Con.Execute Sql
End Sub

Private Sub LstEtat_Click()
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT T_ETATS.* From T_ETATS where T_ETATS.EtatName='" & MyReplace(Me.LstEtat) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
   
    Me.Tag = Rs!Id
 Me.txtMenu = "" & Rs!Menu
 If Rs!Visible = True Then
    MenuVisible.Value = 1
 Else
    MenuVisible.Value = 0
 End If
If Rs!IsPDF = True Then
    IsPDF(0).Value = 1
     IsPDF(1).Value = 0
 Else
    IsPDF(0).Value = 0
     IsPDF(1).Value = 1
 End If
 ChargeDetail
CreateDate = Rs!CreateDate
End If

 
End Sub

Private Sub LstEtat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT T_ETATS.* From T_ETATS where T_ETATS.EtatName='" & MyReplace(Me.LstEtat) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
   
    Sql = "INSERT INTO T_ETATS ( EtatName ) values ( '" & MyReplace(Me.LstEtat) & "');"
    Con.Execute Sql

End If
Rs.Requery
 Me.Tag = Rs!Id
Load MyReplace(Me.LstEtat)
ChargeDetail
End If
End Sub

Sub ChargeDetail()
Dim Sql As String
Dim Rs As Recordset
Dim nb As Long
Dim I As Long

    Me.DetailEtatat.Clear
  
 For I = Me.DetailEtatat.Rows - 2 To 1 Step -1
    Me.DetailEtatat.RemoveItem I
Next
Me.DetailEtatat.Row = 0
 Me.DetailEtatat.Col = 0
    Me.DetailEtatat = "Macro"
Me.DetailEtatat.Col = 1
    Me.DetailEtatat = "Onglet"
   ' Me.DetailEtatat.Height = 30
Sql = "SELECT T_Etats_Onglet.* From T_Etats_Onglet WHERE T_Etats_Onglet.Id_Etat=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
nb = nb + 1
    Me.DetailEtatat.AddItem ""
    Rs.MoveNext
Wend
'If Nb > 0 Then _
'  Me.DetailEtatat.RemoveItem Nb + 1
  Rs.Requery
  nb = 0
While Rs.EOF = False
nb = nb + 1
    Me.DetailEtatat.Row = nb
    Me.DetailEtatat.Col = 0

    Me.DetailEtatat = "" & Rs!Macro
    Me.DetailEtatat.Col = 1
    Me.DetailEtatat = "" & Rs!Onglet
    Me.DetailEtatat.Col = 2
     Me.DetailEtatat = "" & Rs!Id
    Rs.MoveNext
Wend
End Sub

Private Sub MenuVisible_Click()
Dim Sql As String
If MenuVisible.Value = 1 Then
    Sql = "UPDATE T_ETATS SET T_ETATS.Visible = True "
Else
     Sql = "UPDATE T_ETATS SET T_ETATS.Visible = False "
End If
Sql = Sql & "WHERE T_ETATS.ID=" & Me.Tag & ";"
Con.Execute Sql

End Sub

