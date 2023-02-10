VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDediitBouchon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Anuller"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   8640
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Valider"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   8520
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8070
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   14235
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmDitCilp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmDitCilp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin TabDlg.SSTab SSTab2 
         Height          =   7455
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   13150
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Bouchons"
         TabPicture(0)   =   "frmDitCilp.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Joints"
         TabPicture(1)   =   "frmDitCilp.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Clips"
         TabPicture(2)   =   "frmDitCilp.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.Frame Frame1 
            Height          =   7095
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   9735
            Begin VB.CommandButton Ajouter 
               Height          =   375
               Left            =   0
               Picture         =   "frmDitCilp.frx":008C
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Ajouter"
               Top             =   120
               Width           =   375
            End
            Begin VB.Frame Frame2 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   6375
               Left            =   0
               TabIndex        =   4
               Top             =   490
               Width           =   9495
               Begin VB.Frame framDetail 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   99885
                  Left            =   0
                  TabIndex        =   5
                  Top             =   0
                  Width           =   9495
                  Begin VB.CommandButton Kill 
                     Height          =   375
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmDitCilp.frx":0112
                     Style           =   1  'Graphical
                     TabIndex        =   10
                     ToolTipText     =   "Supprimer"
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   375
                  End
                  Begin VB.TextBox BouchoQts 
                     Height          =   375
                     Index           =   0
                     Left            =   7440
                     TabIndex        =   9
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   2055
                  End
                  Begin VB.TextBox BouchonRef 
                     Height          =   375
                     Index           =   0
                     Left            =   360
                     TabIndex        =   8
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   7095
                  End
               End
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   6975
               LargeChange     =   10
               Left            =   9480
               Max             =   80
               TabIndex        =   3
               Top             =   120
               Width           =   255
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
               Left            =   7440
               TabIndex        =   7
               Top             =   120
               Width           =   2055
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
               TabIndex        =   6
               Top             =   120
               Width           =   7095
            End
         End
      End
   End
End
Attribute VB_Name = "frmDediitBouchon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mygrid
Dim MyNumColonne As Collection


Private Sub TabStrip1_Click()

End Sub

Private Sub Ajouter_Click()
LoadControle
End Sub

Private Sub Command1_Click()
Dim I As Integer
Dim Txt As String
For I = 1 To Me.Kill.Count - 1
    Txt = Txt & Me.BouchonRef(I) & "(" & Me.BouchoQts(I) & ")©"
Next
If Len(Txt) > 0 Then Txt = Left(Txt, Len(Txt) - 1)
Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("ConREFBOUCHON")) = Txt
 
   

Command2_Click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Kill_Click(Index As Integer)
UnloadControle Index
End Sub

Private Sub VScroll1_Change()
Me.framDetail.Top = Me.VScroll1.Value * -375
End Sub
Sub LoadControle()
Dim Index As Integer
    Index = Me.Kill.Count
    Load Me.Kill(Index)
    Load BouchonRef(Index)
    Load BouchoQts(Index)
    
    Me.Kill(Index).Top = 370 * (Index - 1)
    Me.BouchonRef(Index).Top = 370 * (Index - 1)
    Me.BouchoQts(Index).Top = 370 * (Index - 1)
    
    Me.Kill(Index).Visible = True
    Me.BouchonRef(Index).Visible = True
    Me.BouchoQts(Index).Visible = True
End Sub
Sub UnloadControle(Index As Integer)
Dim IndexCount As Integer
Dim I As Integer
IndexCount = Me.Kill.Count - 2
    
    For I = Index To IndexCount
        Me.BouchonRef(I).Text = Me.BouchonRef(I + 1).Text
        Me.BouchoQts(I).Text = Me.BouchoQts(I + 1).Text
              
    Next
    Unload Me.Kill(I)
    Unload BouchonRef(I)
    Unload BouchoQts(I)
End Sub
Public Sub Charger(grid, NumColonne)
Dim Bouchon
Dim Qts
Dim I As Integer
Set Mygrid = grid
Set MyNumColonne = NumColonne

Bouchon = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("ConREFBOUCHON")) & "©", "©")
For I = 0 To UBound(Bouchon)
    If Trim("" & Bouchon(I)) <> "" Then
        LoadControle
        Qts = Split(Trim("" & Bouchon(I)) & "(", "(")
        Me.BouchonRef(Me.BouchonRef.Count - 1).Text = Trim("" & Qts(0))
        If Trim("" & Qts(1)) <> "" Then
            Me.BouchoQts(Me.BouchoQts.Count - 1).Text = Left(Trim("" & Qts(1)), Len(Trim("" & Qts(1))) - 1)
        End If
    End If
Next
Me.Show vbModal
End Sub
