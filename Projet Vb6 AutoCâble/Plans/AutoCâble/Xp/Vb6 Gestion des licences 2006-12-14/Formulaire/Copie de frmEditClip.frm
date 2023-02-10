VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditClip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Valider"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   8760
      Width           =   1695
   End
   Begin TabDlg.SSTab OngletConnecteur 
      Height          =   8295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmEditClip.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmEditClip.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   7695
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   13573
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Clips"
         TabPicture(0)   =   "frmEditClip.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Joints"
         TabPicture(1)   =   "frmEditClip.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   6975
            Left            =   -74760
            TabIndex        =   37
            Top             =   480
            Width           =   9015
            Begin VB.VScrollBar VScroll1 
               Height          =   6855
               LargeChange     =   10
               Left            =   8760
               Max             =   80
               TabIndex        =   38
               Top             =   120
               Width           =   255
            End
            Begin VB.CommandButton JointAjouter1 
               Height          =   375
               Left            =   0
               Picture         =   "frmEditClip.frx":0070
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Ajouter"
               Top             =   120
               Width           =   375
            End
            Begin VB.Frame Frame2 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   6375
               Left            =   0
               TabIndex        =   39
               Top             =   480
               Width           =   9495
               Begin VB.Frame detaitJointRef1 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   99885
                  Left            =   0
                  TabIndex        =   40
                  Top             =   120
                  Width           =   9495
                  Begin VB.TextBox JointRef1 
                     Height          =   375
                     Index           =   0
                     Left            =   360
                     TabIndex        =   41
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   8415
                  End
                  Begin VB.CommandButton JointKill1 
                     Height          =   375
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmEditClip.frx":00F6
                     Style           =   1  'Graphical
                     TabIndex        =   42
                     ToolTipText     =   "Supprimer"
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   375
                  End
               End
            End
            Begin VB.Label Label10 
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
               TabIndex        =   44
               Top             =   120
               Width           =   8415
            End
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   6975
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   9015
            Begin VB.VScrollBar VScroll2 
               Height          =   6855
               LargeChange     =   10
               Left            =   8760
               Max             =   80
               TabIndex        =   11
               Top             =   120
               Width           =   255
            End
            Begin VB.Frame Frame4 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   6375
               Left            =   0
               TabIndex        =   6
               Top             =   490
               Width           =   9495
               Begin VB.Frame DetailClip1 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   99885
                  Left            =   0
                  TabIndex        =   7
                  Top             =   0
                  Width           =   9495
                  Begin VB.TextBox ClipRefFour1 
                     Height          =   375
                     Index           =   0
                     Left            =   5640
                     TabIndex        =   15
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   3135
                  End
                  Begin VB.TextBox ClipFamille1 
                     Height          =   375
                     Index           =   0
                     Left            =   360
                     TabIndex        =   10
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   2175
                  End
                  Begin VB.TextBox ClipRef1 
                     Height          =   375
                     Index           =   0
                     Left            =   2520
                     TabIndex        =   9
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   3135
                  End
                  Begin VB.CommandButton ClipKill1 
                     Height          =   375
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmEditClip.frx":0177
                     Style           =   1  'Graphical
                     TabIndex        =   8
                     ToolTipText     =   "Supprimer"
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   375
                  End
               End
            End
            Begin VB.CommandButton ClpAjouter1 
               Height          =   375
               Left            =   0
               Picture         =   "frmEditClip.frx":01F8
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Ajouter"
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Réf Fourniseur"
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
               Left            =   5640
               TabIndex        =   14
               Top             =   120
               Width           =   3135
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Famille"
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
               TabIndex        =   13
               Top             =   120
               Width           =   2175
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
               Left            =   2520
               TabIndex        =   12
               Top             =   120
               Width           =   3135
            End
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7695
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   13573
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Clips"
         TabPicture(0)   =   "frmEditClip.frx":027E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Joints"
         TabPicture(1)   =   "frmEditClip.frx":029A
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   6975
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   9015
            Begin VB.VScrollBar VScroll5 
               Height          =   6855
               LargeChange     =   10
               Left            =   8760
               Max             =   80
               TabIndex        =   46
               Top             =   120
               Width           =   255
            End
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   6375
               Left            =   -360
               TabIndex        =   48
               Top             =   480
               Width           =   9495
               Begin VB.Frame detaitJointRef2 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   99885
                  Left            =   360
                  TabIndex        =   49
                  Top             =   0
                  Width           =   9495
                  Begin VB.TextBox JointRef2 
                     Height          =   375
                     Index           =   0
                     Left            =   360
                     TabIndex        =   51
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   8415
                  End
                  Begin VB.CommandButton JointKill2 
                     Height          =   375
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmEditClip.frx":02B6
                     Style           =   1  'Graphical
                     TabIndex        =   50
                     ToolTipText     =   "Supprimer"
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   375
                  End
               End
            End
            Begin VB.CommandButton JointAjouter2 
               Height          =   375
               Left            =   0
               Picture         =   "frmEditClip.frx":0337
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Ajouter"
               Top             =   120
               Width           =   375
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
               TabIndex        =   52
               Top             =   120
               Width           =   8415
            End
         End
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Height          =   7095
            Left            =   -74880
            TabIndex        =   29
            Top             =   480
            Width           =   9735
            Begin VB.CommandButton Command4 
               Height          =   375
               Left            =   0
               Picture         =   "frmEditClip.frx":03BD
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Ajouter"
               Top             =   120
               Width           =   375
            End
            Begin VB.Frame Frame10 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   6375
               Left            =   0
               TabIndex        =   31
               Top             =   490
               Width           =   9495
               Begin VB.Frame Frame11 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   99885
                  Left            =   0
                  TabIndex        =   32
                  Top             =   0
                  Width           =   9495
                  Begin VB.CommandButton JointKill1 
                     Height          =   375
                     Index           =   1
                     Left            =   0
                     Style           =   1  'Graphical
                     TabIndex        =   34
                     ToolTipText     =   "Supprimer"
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   375
                  End
                  Begin VB.TextBox JointRef1 
                     Height          =   375
                     Index           =   1
                     Left            =   360
                     TabIndex        =   33
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   9135
                  End
               End
            End
            Begin VB.VScrollBar VScroll4 
               Height          =   6975
               LargeChange     =   10
               Left            =   9480
               Max             =   80
               TabIndex        =   30
               Top             =   120
               Width           =   255
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
               Left            =   360
               TabIndex        =   36
               Top             =   120
               Width           =   9135
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   6975
            Left            =   -74880
            TabIndex        =   17
            Top             =   480
            Width           =   9015
            Begin VB.VScrollBar VScroll3 
               Height          =   6855
               LargeChange     =   10
               Left            =   8760
               Max             =   80
               TabIndex        =   18
               Top             =   120
               Width           =   255
            End
            Begin VB.CommandButton ClpAjouter2 
               Height          =   375
               Left            =   0
               Picture         =   "frmEditClip.frx":0443
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "Ajouter"
               Top             =   120
               Width           =   375
            End
            Begin VB.Frame Frame7 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   6375
               Left            =   0
               TabIndex        =   19
               Top             =   490
               Width           =   9495
               Begin VB.Frame DetailClip2 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   99885
                  Left            =   0
                  TabIndex        =   20
                  Top             =   0
                  Width           =   9495
                  Begin VB.CommandButton ClipKill2 
                     Height          =   375
                     Index           =   0
                     Left            =   0
                     Picture         =   "frmEditClip.frx":04C9
                     Style           =   1  'Graphical
                     TabIndex        =   24
                     ToolTipText     =   "Supprimer"
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   375
                  End
                  Begin VB.TextBox ClipRef2 
                     Height          =   375
                     Index           =   0
                     Left            =   2520
                     TabIndex        =   23
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   3135
                  End
                  Begin VB.TextBox ClipFamille2 
                     Height          =   375
                     Index           =   0
                     Left            =   360
                     TabIndex        =   22
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   2175
                  End
                  Begin VB.TextBox ClipRefFour2 
                     Height          =   375
                     Index           =   0
                     Left            =   5640
                     TabIndex        =   21
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   3135
                  End
               End
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
               Left            =   2520
               TabIndex        =   28
               Top             =   120
               Width           =   3135
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Famille"
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
               TabIndex        =   27
               Top             =   120
               Width           =   2175
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Réf Fourniseur"
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
               Left            =   5640
               TabIndex        =   26
               Top             =   120
               Width           =   3135
            End
         End
      End
   End
End
Attribute VB_Name = "frmEditClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mygrid As Spreadsheet
Dim MyNumColonne As Collection


Private Sub TabStrip1_Click()

End Sub



Private Sub ClipillKil2_Click(Index As Integer)
UnloadControle Index, "Clip2"
End Sub

Private Sub ClipillKill1_Click(Index As Integer)
    UnloadControle Index, "Clip1"
End Sub

Private Sub ClipKill1_Click(Index As Integer)
UnloadControle Index, "Clip1"
End Sub

Private Sub ClipKill2_Click(Index As Integer)
UnloadControle Index, "Clip2"
End Sub

Private Sub ClpAjouter_Click()

End Sub

Private Sub ClpAjouter1_Click()
LoadControle "Clip1"
End Sub

Private Sub ClpAjouter2_Click()
LoadControle "Clip2"
End Sub
Function RechercheSpreadsheet(Myrange, MyCellule, strRecherche) As Long
'Permet de rechercher une valeur dans un tableau Excel.
'MyxlWhole = MyxlWhole + 1
On Error Resume Next
'Recherche = Myrange.Find(What:=strRecherche, After:=Myrange.Cells(MyCellule, 1), _
'            LookIn:=xlFormulas, LookAt _
'         :=MyxlWhole, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
'        False).Row
        
RechercheSpreadsheet = Myrange.Find(What:=MyCellule, After:=Myrange.Cells(MyCellule.Row + 1, MyCellule.Column), findlookin:=ssFormulas, findlookat:=ssPart, SearchOrder:=ssByRows, SearchDirection:=ssNext, MatchCase:=False).Row
        
 
    If Err Then
        Err.Clear
        RechercheSpreadsheet = 0
    End If
End Function
Private Sub Command1_Click()
Dim I As Integer
Dim Txt As String
Dim Txt_P As String
Dim txt1
Dim txt1_p
Dim RetourRows As Long
Dim SherchTrouve As Boolean
'Set Myrange = VOI
'Mygrid.ActiveSheet.Range("A1").CurrentRegion.Find( "toto",
RetourRows = Mygrid.ActiveCell.Row
 SherchTrouve = False
While SherchTrouve = False
    RetourRows = RechercheSpreadsheet(Mygrid.ActiveSheet.Range("A1").CurrentRegion, Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi")), "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi")))
    If RetourRows = 0 Then
        SherchTrouve = True
    Else
        If RetourRows = Mygrid.ActiveCell.Row Then SherchTrouve = True
        If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("Filsapp")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("Filsapp")) Then SherchTrouve = True
        If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("Filsapp")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("Filsapp2")) Then SherchTrouve = True
    End If
Wend
'Set Myrange = Mygrid.ActiveSheet.Range("A1").CurrentRegion
'Myrange.
For I = 1 To Me.ClipKill1.Count - 1
    Txt = Txt & Me.ClipFamille1(I) & ":" & Me.ClipRef1(I) & "©"
    Txt_P = Txt_P & Me.ClipFamille1(I) & ":" & Me.ClipRefFour1(I) & "©"
Next
If Right(Txt, 1) = "©" Then Txt = Left(Txt, Len(Txt) - 1)
If Right(Txt_P, 1) = "©" Then Txt_P = Left(Txt_P, Len(Txt_P) - 1)

'If UBound(Split(Txt, "©")) = 0 Then
'    txt1 = Split(Txt & ":", ":")
'    Txt = Trim("" & txt1(1))
'End If
   Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP")) = Txt
   Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP FOUR")) = Txt_P
   
'
'   If RetourRows <> 0 Then
'        If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsApp")) ="" &  Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsApp")) Then
'            If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi")) ="" &  Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi")) Then
'                Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP")) ="" &  Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP"))
'            Else
'                If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi")) ="" &  Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi2")) Then
'                    Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP2")) ="" &  Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP"))
'                End If
'            End If
'        Else
'            If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsApp")) ="" &  Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsApp2")) Then
'                If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi")) ="" &  Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi2")) Then
'                    Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP2")) ="" &  Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP"))
'                End If
'            End If
'        End If
'
'   End If
'

   
   
   
   
 Txt = ""
 Txt_P = ""
For I = 1 To Me.JointKill1.Count - 1
    Txt = Txt & Me.JointRef1(I) & "©"
Next
If Right(Txt, 1) = "©" Then Txt = Left(Txt, Len(Txt) - 1)
Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT")) = Txt

If RetourRows <> 0 Then
        If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsApp")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsApp")) Then
            If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi2")) Then
                Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP"))
                Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT"))
            Else
                 If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi2")) Then
                        Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP2")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP"))
                        Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT2")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT"))
                End If
            End If
        Else
            If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsApp2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsApp")) Then
                 If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi")) Then
                    Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT2"))
                    Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP2"))
                End If
            End If
'            Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT")) ="" &  Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT2"))
        End If
    
   End If
   
   
   
RetourRows = Mygrid.ActiveCell.Row
 SherchTrouve = False
While SherchTrouve = False
    RetourRows = RechercheSpreadsheet(Mygrid.ActiveSheet.Range("A1").CurrentRegion, Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi2")), "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi2")))
    If RetourRows = 0 Then
        SherchTrouve = True
    Else
        If RetourRows = Mygrid.ActiveCell.Row Then SherchTrouve = True
        If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("Filsapp2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("Filsapp2")) Then SherchTrouve = True
        If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("Filsapp2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("Filsapp")) Then SherchTrouve = True
    End If
Wend
   
   
 Txt = ""
For I = 1 To Me.ClipKill2.Count - 1
    Txt = Txt & Me.ClipFamille2(I) & ":" & Me.ClipRef2(I) & "©"
    Txt_P = Txt_P & Me.ClipFamille2(I) & ":" & Me.ClipRefFour2(I) & "©"
Next
If Right(Txt, 1) = "©" Then Txt = Left(Txt, Len(Txt) - 1)
If Right(Txt_P, 1) = "©" Then Txt_P = Left(Txt_P, Len(Txt_P) - 1)
'If UBound(Split(Txt, "©")) = 0 Then
'    txt1 = Split(Txt & ":", ":")
'    Txt = Trim("" & txt1(1))
'End If
   Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP2")) = Txt
    Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP FOUR2")) = Txt_P
      


   
 Txt = ""
For I = 1 To Me.JointKill2.Count - 1
    Txt = Txt & Me.JointRef2(I) & "©"
Next
If Right(Txt, 1) = "©" Then Txt = Left(Txt, Len(Txt) - 1)
   Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT2")) = Txt
   
If RetourRows <> 0 Then
        If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsApp2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsApp2")) Then
            If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi2")) Then
                Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP2")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP2"))
                Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT2")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT2"))
                Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP FOUR2")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP FOUR2"))
            Else
                 If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi")) Then
                        Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP2"))
                        Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT2"))
                        Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP FOUR")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP FOUR2"))
                End If
            End If
        Else
             If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsApp")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsApp")) Then
                If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi")) Then
                    Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP"))
                    Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT"))
                    Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP FOUR")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP FOUR"))
                 End If
            Else
                If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsApp2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsApp")) Then
                     If Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsVoi2")) = "" & Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsVoi")) Then
                        Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT2"))
                        Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP2"))
                        Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF CLIP FOUR")) = "" & Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF CLIP FOUR2"))
                    End If
                 End If
            End If
'            Mygrid.ActiveSheet.Cells(RetourRows, MyNumColonne("FilsREF JOINT")) ="" &  Mygrid.ActiveSheet.Cells(Mygrid.ActiveCell.Row, MyNumColonne("FilsREF JOINT2"))
        End If
    
   End If
Command2_Click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Load()

End Sub

Private Sub JointAjouter1_Click()
LoadControle "Joint1"
End Sub

Private Sub JointAjouter2_Click()
LoadControle "Joint2"
End Sub

Private Sub Command3_Click()

End Sub

Private Sub JointKill1_Click(Index As Integer)
UnloadControle Index, "Joint1"
End Sub

Private Sub JointKill2_Click(Index As Integer)
UnloadControle Index, "Joint2"
End Sub

'Private Sub VScroll1_Change()
'Me.framDetail.Top = Me.VScroll1.Value * -375
'End Sub
Sub LoadControle(TypeControl As String)
Dim Index As Integer
Dim I As Integer
Select Case TypeControl
        Case "Clip1"
            Index = Me.ClipKill1.Count
            Load Me.ClipKill1(Index)
            Load Me.ClipFamille1(Index)
            Load Me.ClipRef1(Index)
            Load Me.ClipRefFour1(Index)
            
            Me.ClipKill1(Index).Top = 370 * (Index - 1)
            Me.ClipFamille1(Index).Top = 370 * (Index - 1)
            Me.ClipRef1(Index).Top = 370 * (Index - 1)
            Me.ClipRefFour1(Index).Top = 370 * (Index - 1)
            
            Me.ClipKill1(Index).Visible = True
            Me.ClipFamille1(Index).Visible = True
            Me.ClipRef1(Index).Visible = True
            Me.ClipRefFour1(Index).Visible = True
        
        Case "Clip2"
            Index = Me.ClipKill2.Count
            Load Me.ClipKill2(Index)
            Load Me.ClipFamille2(Index)
            Load Me.ClipRef2(Index)
            Load Me.ClipRefFour2(Index)
            
            Me.ClipKill2(Index).Top = 370 * (Index - 1)
            Me.ClipFamille2(Index).Top = 370 * (Index - 1)
            Me.ClipRef2(Index).Top = 370 * (Index - 1)
             Me.ClipRefFour2(Index).Top = 370 * (Index - 1)
            
            Me.ClipKill2(Index).Visible = True
            Me.ClipFamille2(Index).Visible = True
            Me.ClipRef2(Index).Visible = True
              Me.ClipRefFour2(Index).Visible = True
            
            Case "Joint1"
                Index = Me.JointKill1.Count
                Load Me.JointKill1(Index)
                Load Me.JointRef1(Index)
                
                Me.JointKill1(Index).Top = 370 * (Index - 1)
                Me.JointRef1(Index).Top = 370 * (Index - 1)
                
                Me.JointKill1(Index).Visible = True
                Me.JointRef1(Index).Visible = True
                
           Case "Joint2"
                Index = Me.JointKill2.Count
                Load Me.JointKill2(Index)
                Load Me.JointRef2(Index)
                
                Me.JointKill2(Index).Top = 370 * (Index - 1)
                Me.JointRef2(Index).Top = 370 * (Index - 1)
                
                Me.JointKill2(Index).Visible = True
                Me.JointRef2(Index).Visible = True
        
End Select
End Sub
Sub UnloadControle(Index As Integer, TypeControl As String)
Dim IndexCount As Integer
Dim I As Integer
Select Case TypeControl
        Case "Clip1"
            IndexCount = Me.ClipKill1.Count - 2
        
            For I = Index To IndexCount
                Me.ClipFamille1(I).Text = Me.ClipFamille1(I + 1).Text
                Me.ClipRef1(I).Text = Me.ClipRef1(I + 1).Text
                 Me.ClipRefFour1(I).Text = Me.ClipRefFour1(I + 1).Text
                      
            Next
            Unload Me.ClipFamille1(I)
            Unload Me.ClipRef1(I)
            Unload Me.ClipKill1(I)
            Unload Me.ClipRefFour1(I)
        Case "Clip2"
            IndexCount = Me.ClipKill2.Count - 2
        
            For I = Index To IndexCount
                Me.ClipFamille2(I).Text = Me.ClipFamille2(I + 1).Text
                Me.ClipRef2(I).Text = Me.ClipRef2(I + 1).Text
                 Me.ClipRefFour2(I).Text = Me.ClipRefFour2(I + 1).Text
                      
            Next
            Unload Me.ClipFamille2(I)
            Unload Me.ClipRef2(I)
            Unload Me.ClipKill2(I)
            Unload Me.ClipRefFour2(I)
            
        Case "Joint1"
            IndexCount = Me.JointKill1.Count - 2
        
            For I = Index To IndexCount
                Me.JointRef1(I).Text = Me.JointRef1(I + 1).Text
                                     
            Next
            Unload Me.JointKill1(I)
            Unload Me.JointRef1(I)
            
        Case "Joint2"
            IndexCount = Me.JointKill2.Count - 2
        
            For I = Index To IndexCount
                Me.JointRef2(I).Text = Me.JointRef2(I + 1).Text
                                     
            Next
            Unload Me.JointKill2(I)
            Unload Me.JointRef2(I)
          
End Select
End Sub
Public Sub Charger(grid As Spreadsheet, NumColonne)
Dim Bouchon
Dim Clip
Dim ClipFour
Dim CliFamilleFour
Dim CliFamille
Dim Qts
Dim Joint
Dim I As Integer
Set Mygrid = grid
'APP
Set MyNumColonne = NumColonne
Me.Caption = "Désignation Liaison: " & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsDESIGNATION"))
OngletConnecteur.TabCaption(0) = "" & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsAPP")) & " : " & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsREF CONNECTEUR"))
OngletConnecteur.TabCaption(1) = "" & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsAPP2")) & " : " & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsREF CONNECTEUR2"))

Clip = Split("©:" & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsREF CLIP")) & "©", "©")
ClipFour = Split("©:" & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsREF CLIP FOUR")) & "©", "©")

For I = 0 To UBound(Clip)

    If Len(Trim("" & Clip(I))) > 1 Then
    
       LoadControle "Clip1"
        CliFamille = Split(Trim("" & Clip(I)) & ":", ":")
         CliFamilleFour = Split(Trim("" & ClipFour(I)) & ":", ":")
         If UBound(CliFamille) = 3 Then
            Me.ClipFamille1(Me.ClipFamille1.Count - 1).Text = Trim("" & CliFamille(1))
            Me.ClipRef1(Me.ClipRef1.Count - 1).Text = Trim("" & CliFamille(2))
            Me.ClipRefFour1(Me.ClipRefFour1.Count - 1).Text = Trim("" & CliFamilleFour(2))
         
         Else
            Me.ClipFamille1(Me.ClipFamille1.Count - 1).Text = Trim("" & CliFamille(0))
            Me.ClipRef1(Me.ClipRef1.Count - 1).Text = Trim("" & CliFamille(1))
            Me.ClipRefFour1(Me.ClipRefFour1.Count - 1).Text = Trim("" & CliFamilleFour(1))
        End If
    End If
Next

Clip = Split("©:" & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsREF CLIP2")) & "©", "©")
ClipFour = Split("©:" & grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsREF CLIP FOUR2")) & "©", "©")
For I = 0 To UBound(Clip)
    If Len(Trim("" & Clip(I))) > 1 Then
    
        LoadControle "Clip2"
        CliFamille = Split(Trim("" & Clip(I)) & ":", ":")
        CliFamilleFour = Split(Trim("" & ClipFour(I)) & ":", ":")
        If UBound(CliFamille) = 3 Then
            Me.ClipFamille2(Me.ClipFamille2.Count - 1).Text = Trim("" & CliFamille(1))
            Me.ClipRef2(Me.ClipRef2.Count - 1).Text = Trim("" & CliFamille(2))
             Me.ClipRefFour2(Me.ClipRefFour2.Count - 1).Text = Trim("" & CliFamilleFour(2))
        Else
            Me.ClipFamille2(Me.ClipFamille2.Count - 1).Text = Trim("" & CliFamille(0))
            Me.ClipRef2(Me.ClipRef2.Count - 1).Text = Trim("" & CliFamille(1))
             Me.ClipRefFour2(Me.ClipRefFour2.Count - 1).Text = Trim("" & CliFamilleFour(1))
         End If
    End If
Next

Joint = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsREF Joint")) & "©", "©")
For I = 0 To UBound(Joint)
    If Trim("" & Joint(I)) <> "" Then
    
        LoadControle "Joint1"
        
        Me.JointRef1(Me.JointRef1.Count - 1).Text = Trim("" & Joint(I))
        
    End If
Next

Joint = Split(grid.ActiveSheet.Cells(grid.ActiveCell.Row, NumColonne("FilsREF Joint2")) & "©", "©")
For I = 0 To UBound(Joint)
    If Trim("" & Joint(I)) <> "" Then
    
        LoadControle "Joint2"
        
        Me.JointRef2(Me.JointRef2.Count - 1).Text = Trim("" & Joint(I))
        
    End If
Next
'REF Joint
Me.Show vbModal
End Sub

Private Sub VScroll1_Change()
 detaitJointRef1.Top = VScroll1.Value * -1 * 370

End Sub

Private Sub VScroll2_Change()
  DetailClip1.Top = VScroll2.Value * -1 * 370
End Sub

Private Sub VScroll3_Change()
DetailClip2.Top = VScroll3.Value * -1 * 370

End Sub

Private Sub VScroll5_Change()
 detaitJointRef2.Top = VScroll5.Value * -1 * 370
End Sub
