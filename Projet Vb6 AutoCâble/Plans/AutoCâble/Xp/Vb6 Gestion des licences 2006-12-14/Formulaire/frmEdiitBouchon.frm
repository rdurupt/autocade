VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditBouchon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Valider"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   7920
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   7455
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   13150
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Bouchons"
      TabPicture(0)   =   "frmEdiitBouchon.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Capots"
      TabPicture(1)   =   "frmEdiitBouchon.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Verrous"
      TabPicture(2)   =   "frmEdiitBouchon.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   9720
         Begin VB.VScrollBar VScroll3 
            Height          =   6975
            LargeChange     =   10
            Left            =   9480
            Max             =   80
            TabIndex        =   27
            Top             =   120
            Width           =   255
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   6375
            Left            =   0
            TabIndex        =   23
            Top             =   490
            Width           =   9495
            Begin VB.Frame framDetailVerous 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   99885
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Width           =   9495
               Begin VB.TextBox VerrousRefFour 
                  Height          =   375
                  Index           =   0
                  Left            =   4920
                  TabIndex        =   34
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   4605
               End
               Begin VB.TextBox VerrousRef 
                  Height          =   375
                  Index           =   0
                  Left            =   360
                  TabIndex        =   26
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   4605
               End
               Begin VB.CommandButton VerrousKill 
                  Height          =   375
                  Index           =   0
                  Left            =   0
                  Picture         =   "frmEdiitBouchon.frx":0054
                  Style           =   1  'Graphical
                  TabIndex        =   25
                  ToolTipText     =   "Supprimer"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   375
               End
            End
         End
         Begin VB.CommandButton VerrousAjouter 
            Height          =   375
            Left            =   0
            Picture         =   "frmEdiitBouchon.frx":00D5
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Ajouter"
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Référence"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            TabIndex        =   33
            Top             =   120
            Width           =   4605
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Référence"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   120
            Width           =   4605
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7095
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   9735
         Begin VB.CommandButton BouchonAjouter 
            Height          =   375
            Left            =   0
            Picture         =   "frmEdiitBouchon.frx":015B
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Ajouter"
            Top             =   120
            Width           =   375
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   6375
            Left            =   0
            TabIndex        =   13
            Top             =   490
            Width           =   9495
            Begin VB.Frame framDetailBouchon 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   99885
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   9495
               Begin VB.TextBox BouchonRefFour 
                  Height          =   375
                  Index           =   0
                  Left            =   3960
                  TabIndex        =   30
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   3590
               End
               Begin VB.CommandButton BouchonKill 
                  Height          =   375
                  Index           =   0
                  Left            =   0
                  Picture         =   "frmEdiitBouchon.frx":01E1
                  Style           =   1  'Graphical
                  TabIndex        =   17
                  ToolTipText     =   "Supprimer"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin VB.TextBox BouchonQts 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   7560
                  TabIndex        =   16
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.TextBox BouchonRef 
                  Height          =   375
                  Index           =   0
                  Left            =   360
                  TabIndex        =   15
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   3590
               End
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   6975
            LargeChange     =   10
            Left            =   9480
            Max             =   80
            TabIndex        =   12
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Référence Fournisseur"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   29
            Top             =   120
            Width           =   3590
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Quantités"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7560
            TabIndex        =   20
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Référence"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   120
            Width           =   3590
         End
      End
      Begin VB.Frame Frame3 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   9735
         Begin VB.VScrollBar VScroll2 
            Height          =   6975
            LargeChange     =   10
            Left            =   9480
            Max             =   80
            TabIndex        =   9
            Top             =   120
            Width           =   255
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   6375
            Left            =   0
            TabIndex        =   5
            Top             =   490
            Width           =   9495
            Begin VB.Frame framDetailCapot 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   99885
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   9495
               Begin VB.TextBox CapotRefFour 
                  Height          =   375
                  Index           =   0
                  Left            =   4920
                  TabIndex        =   32
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   4550
               End
               Begin VB.TextBox CapotRef 
                  Height          =   375
                  Index           =   0
                  Left            =   360
                  TabIndex        =   8
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   4550
               End
               Begin VB.CommandButton CapotKill 
                  Height          =   375
                  Index           =   0
                  Left            =   0
                  Picture         =   "frmEdiitBouchon.frx":0262
                  Style           =   1  'Graphical
                  TabIndex        =   7
                  ToolTipText     =   "Supprimer"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   375
               End
            End
         End
         Begin VB.CommandButton CapotAjout 
            Height          =   375
            Left            =   0
            Picture         =   "frmEdiitBouchon.frx":02E3
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Ajouter"
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Référence"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            TabIndex        =   31
            Top             =   120
            Width           =   4550
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Référence"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   120
            Width           =   4550
         End
      End
   End
End
Attribute VB_Name = "frmEditBouchon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mygrid
Dim MyNumColonne As Collection


Private Sub TabStrip1_Click()

End Sub


Private Sub BouchonAjouter_Click()
LoadControle "Bouchon"

End Sub

Private Sub BouchonKill_Click(Index As Integer)
    UnloadControle Index, "Bouchon"
End Sub

Private Sub CapotAjout_Click()
LoadControle "Capot"
End Sub

Private Sub CapotKill_Click(Index As Integer)
UnloadControle Index, "Capot"
End Sub

Private Sub Command1_Click()
Dim I As Integer
Dim Txt As String
Dim Txt_P As String
On Error Resume Next
'REFBOUCHONFOUR
For I = 1 To Me.BouchonKill.Count - 1
    Txt = Txt & Me.BouchonRef(I) & "(" & Me.BouchonQts(I) & ")©"
    Txt_P = Txt_P & Me.BouchonRefFour(I) & "(" & Me.BouchonQts(I) & ")©"
Next
If Len(Txt) > 0 Then Txt = Left(Txt, Len(Txt) - 1)
If Len(Txt_P) > 0 Then Txt_P = Left(Txt_P, Len(Txt_P) - 1)

Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("ConREFBOUCHON")) = Txt
Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("ConREFBOUCHONFOUR")) = Txt_P

Txt = ""
Txt_P = ""
'REFCAPOTFOUR
For I = 1 To Me.CapotKill.Count - 1
    Txt = Txt & Me.CapotRef(I) & "©"
    Txt_P = Txt_P & Me.CapotRefFour(I) & "©"
Next

If Len(Txt) > 0 Then Txt = Left(Txt, Len(Txt) - 1)
If Len(Txt_P) > 0 Then Txt_P = Left(Txt_P, Len(Txt_P) - 1)
Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("ConREFCAPOT")) = Txt
Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("ConREFCAPOTFOUR")) = Txt_P
'   REFVERROU
Txt = ""
Txt_P = ""
For I = 1 To Me.VerrousKill.Count - 1
    Txt = Txt & Me.VerrousRef(I) & "©"
    Txt_P = Txt_P & Me.VerrousRefFour(I) & "©"
Next
If Len(Txt) > 0 Then Txt = Left(Txt, Len(Txt) - 1)
If Len(Txt_P) > 0 Then Txt_P = Left(Txt_P, Len(Txt_P) - 1)
Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("ConREFVERROU")) = Txt
Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("ConREFVERROUFOUR")) = Txt_P
Command2_Click
On Error GoTo 0
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub SSTab2_DblClick()
LoadControle "Verrous"
End Sub

Private Sub VerrousAjouter_Click()
LoadControle "Verrous"
End Sub

Private Sub VerrousKill_Click(Index As Integer)
UnloadControle Index, "Verrous"
End Sub

Private Sub VScroll1_Change()
Me.framDetailBouchon.Top = Me.VScroll1.Value * -375
End Sub
Sub LoadControle(TypeControle As String)
Dim Index As Integer
Dim MyControl As Control
'Bouchon
Select Case TypeControle
        Case "Bouchon"
                Index = Me.BouchonKill.Count
                
                Load Me.BouchonKill(Index)
                Load Me.BouchonRef(Index)
                Load Me.BouchonRefFour(Index)
                Load Me.BouchonQts(Index)
                
                Me.BouchonKill(Index).Top = 370 * (Index - 1)
                Me.BouchonRef(Index).Top = 370 * (Index - 1)
                Me.BouchonRefFour(Index).Top = 370 * (Index - 1)
                Me.BouchonQts(Index).Top = 370 * (Index - 1)
                
                Me.BouchonKill(Index).Visible = True
                Me.BouchonRef(Index).Visible = True
                Me.BouchonRefFour(Index).Visible = True
                Me.BouchonQts(Index).Visible = True
        
        Case "Capot"
                Index = Me.CapotKill.Count
                Load Me.CapotKill(Index)
                Load Me.CapotRef(Index)
                Load Me.CapotRefFour(Index)
                
                CapotKill(Index).Top = 370 * (Index - 1)
                CapotRef(Index).Top = 370 * (Index - 1)
                CapotRefFour(Index).Top = 370 * (Index - 1)
                
                CapotKill(Index).Visible = True
                CapotRef(Index).Visible = True
                CapotRefFour(Index).Visible = True
                
        Case "Verrous"
                Index = Me.VerrousKill.Count
                Load Me.VerrousKill(Index)
                Load Me.VerrousRef(Index)
                Load Me.VerrousRefFour(Index)
                
                Me.VerrousKill(Index).Top = 370 * (Index - 1)
                Me.VerrousRef(Index).Top = 370 * (Index - 1)
                Me.VerrousRefFour(Index).Top = 370 * (Index - 1)
                
                Me.VerrousKill(Index).Visible = True
                Me.VerrousRef(Index).Visible = True
                Me.VerrousRefFour(Index).Visible = True
            
        
End Select
    
    
'    Load Me.BouchonRef(Index)
'    Load BouchonQts(Index)
'
'    Me.BouchonKill(Index).Top = 370 * (Index - 1)
'    Me.BouchonRef(Index).Top = 370 * (Index - 1)
'    Me.BouchonQts(Index).Top = 370 * (Index - 1)
'
'    Me.BouchonKill(Index).Visible = True
'    Me.BouchonRef(Index).Visible = True
'    Me.BouchonQts(Index).Visible = True
End Sub
Sub UnloadControle(Index As Integer, TypeControle As String)
Dim IndexCount As Integer
'Bouchon
Dim I As Integer
Select Case TypeControle
        Case "Bouchon"
            IndexCount = Me.BouchonKill.Count - 2
    
            For I = Index To IndexCount
                Me.BouchonRef(I).Text = Me.BouchonRef(I + 1).Text
                Me.BouchonQts(I).Text = Me.BouchonQts(I + 1).Text
                      
            Next
            Unload Me.BouchonKill(I)
            Unload BouchonRef(I)
            Unload BouchonQts(I)
            
          Case "Capot"
             IndexCount = Me.CapotKill.Count - 2
    
            For I = Index To IndexCount
                Me.CapotRef(I).Text = Me.CapotRef(I + 1).Text
            Next
            Unload Me.CapotKill(I)
            Unload CapotRef(I)
         Case "Verrous"
             IndexCount = Me.VerrousKill.Count - 2
    
            For I = Index To IndexCount
                Me.VerrousRef(I).Text = Me.VerrousRef(I + 1).Text
            Next
            Unload Me.VerrousKill(I)
            Unload VerrousRef(I)
           
         
End Select
End Sub
Public Sub charger(grid, NumColonne)
Dim Bouchon
Dim Bouchon_p
Dim Capot
Dim Capot_p
Dim Verrou
Dim Verrou_p
Dim Qts
Dim Qts_p
Dim I As Integer
Set Mygrid = grid
Set MyNumColonne = NumColonne
On Error Resume Next
'REFBOUCHONFOUR
Bouchon = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("ConREFBOUCHON")) & "©", "©")
Bouchon_p = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("ConREFBOUCHONFOUR")) & "©", "©")

For I = 0 To UBound(Bouchon)
    If Trim("" & Bouchon(I)) <> "" Then
        LoadControle "Bouchon"
        Qts = Split(Trim("" & Bouchon(I)) & "(", "(")
        Qts_p = Split(Trim("" & Bouchon_p(I)) & "(", "(")
        Me.BouchonRef(Me.BouchonRef.Count - 1).Text = Trim("" & Qts(0))
        Me.BouchonRefFour(Me.BouchonRefFour.Count - 1).Text = Trim("" & Qts_p(0))
        If Trim("" & Qts(1)) <> "" Then
            Me.BouchonQts(Me.BouchonQts.Count - 1).Text = Left(Trim("" & Qts(1)), Len(Trim("" & Qts(1))) - 1)
        End If
    End If
Next

Capot = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("ConREFCAPOT")) & "©", "©")
Capot_p = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("ConREFCAPOTFOUR")) & "©", "©")
'
For I = 0 To UBound(Capot)
    If Trim("" & Capot(I)) <> "" Then
        LoadControle "Capot"
                Me.CapotRef(Me.CapotRef.Count - 1).Text = Trim("" & Capot(I))
                Me.CapotRefFour(Me.CapotRefFour.Count - 1).Text = Trim("" & Capot_p(I))
        
    End If
Next
'REFVERROUFOUR
Verrou = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("ConREFVERROU")) & "©", "©")
Verrou_p = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("ConREFVERROUFOUR")) & "©", "©")

For I = 0 To UBound(Verrou)
    If Trim("" & Verrou(I)) <> "" Then
        LoadControle "Verrous"
                Me.VerrousRef(Me.VerrousRef.Count - 1).Text = Trim("" & Verrou(I))
                Me.VerrousRefFour(Me.VerrousRefFour.Count - 1).Text = Trim("" & Verrou_p(I))
        
    End If
Next
' NumColonne ("ConREFVERROU")
Me.Show vbModal
On Error GoTo 0
End Sub

Private Sub VScroll2_Change()
Me.framDetailCapot.Top = Me.VScroll2.Value * -375
End Sub

Private Sub VScroll3_Change()
Me.framDetailVerous.Top = Me.VScroll3.Value * -375
End Sub
