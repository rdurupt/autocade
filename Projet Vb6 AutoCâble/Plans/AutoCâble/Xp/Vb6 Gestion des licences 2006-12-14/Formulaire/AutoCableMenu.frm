VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form AutoCableMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11535
   ScaleWidth      =   3105
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   3836
      _Version        =   327682
      HideSelection   =   0   'False
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      MousePointer    =   1
      OLEDragMode     =   1
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AutoCableMenu.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "AutoCableMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim nodX As Node    ' Crée une arborescence.
    Set nodX = TreeView1.Nodes.Add(, , , "Parent1", 1)
    Set nodX = TreeView1.Nodes.Add(, , , "Parent2")
    Set nodX = TreeView1.Nodes.Add(1, tvwChild, , "Fils 1", 1)
    Set nodX = TreeView1.Nodes.Add(1, tvwChild, , "Fils 2", 1)
    Set nodX = TreeView1.Nodes.Add(2, tvwChild, , "Fils 3", 1)
    Set nodX = TreeView1.Nodes.Add(2, tvwChild, , "Fils 4", 1)
    Set nodX = TreeView1.Nodes.Add(3, tvwChild, , "Fils 5", 1)
End Sub

Private Sub Form_Resize()
TreeView1.Width = Me.Width
Me.Height = frmAutocâble.Height
TreeView1.Height = Me.Height
End Sub

