VERSION 5.00
Begin VB.Form FrmSelectCriteres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selection Critères:"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   Icon            =   "FrmSelectCriteres.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Annuler"
      Height          =   735
      Left            =   5280
      TabIndex        =   8
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Valider"
      Height          =   735
      Left            =   840
      TabIndex        =   7
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options:"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   3600
         Picture         =   "FrmSelectCriteres.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3600
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   3600
         Picture         =   "FrmSelectCriteres.frx":114C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   495
      End
      Begin VB.ListBox lstSupp 
         Height          =   5130
         Left            =   4560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.ListBox lstGarder 
         Height          =   5130
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Ecarter"
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Garder"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
   End
End
Attribute VB_Name = "FrmSelectCriteres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sorite As Boolean
Dim Annuler As Boolean

Private Sub Command1_Click()
If Me.lstSupp.ListIndex = -1 Then
    MsgBox "Vous devez selectionner un élément dans la liste"
Else
    Me.lstGarder.AddItem Me.lstSupp.List(Me.lstSupp.ListIndex)
    Me.lstSupp.RemoveItem Me.lstSupp.ListIndex
   End If
End Sub

Private Sub List2_Click()

End Sub

Private Sub Command2_Click()
If Me.lstGarder.ListIndex = -1 Then
    MsgBox "Vous devez selectionner un élément dans la liste"
Else
    Me.lstSupp.AddItem Me.lstGarder.List(Me.lstGarder.ListIndex)
    Me.lstGarder.RemoveItem Me.lstGarder.ListIndex
   End If
End Sub




Private Sub Command3_Click()
Sorite = False
Me.Hide
End Sub

Private Sub Command4_Click()
Sorite = False
Annuler = True
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = Sorite
End Sub

Private Sub lstGarder_DblClick()
    Command2_Click
End Sub

Private Sub lstSupp_DblClick()
Command1_Click
End Sub
Public Function Chargement(OngletCritaire, Txt)
Dim Myrange
Dim SlpitTxt
Dim I As Long
Set Myrange = OngletCritaire.ActiveSheet.Range("a1").CurrentRegion
If InStr(1, Txt, "TOUS") = 0 Then
        Me.lstSupp.AddItem "TOUS"
    End If
For I = 2 To Myrange.Rows.Count
    If InStr(1, Txt, "" & Myrange(I, 2)) = 0 Then
        Me.lstSupp.AddItem "" & Myrange(I, 2)
    End If

Next

'sql = "SELECT T_Critères.CODE_CRITERE FROM T_Critères "
'sql = sql & "WHERE T_Critères.Id_IndiceProjet=" & IdProjet & ";"
'Set Rs = Con.OpenRecordSet(sql)
'While Rs.EOF = False
'If InStr(1, txt, "" & Rs!CODE_CRITERE) = 0 Then
'        Me.lstSupp.AddItem "" & Rs!CODE_CRITERE
'    End If
'    Rs.MoveNext
'Wend
'Set Rs = Con.CloseRecordSet(Rs)
SlpitTxt = Split(Txt, ";")
For I = 0 To UBound(SlpitTxt) - 1
Me.lstGarder.AddItem SlpitTxt(I)
Next
Chargement = Txt
Sorite = True
Me.Show vbModal
If Annuler = False Then
    Chargement = ""
    For I = 0 To Me.lstGarder.ListCount - 1
      Chargement = Chargement & Me.lstGarder.List(I) & ";"
    Next
End If
End Function

