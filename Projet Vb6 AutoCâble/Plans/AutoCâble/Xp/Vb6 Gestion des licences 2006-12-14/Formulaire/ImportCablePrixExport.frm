VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ImportCablePrixExport 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12645
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label ProgressBar1Caption 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "ImportCablePrixExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ImporOk As Boolean

Private Sub Form_Activate()
Set FormBarGrah = Me
If ImporOk = True Then
    ImportCablePrix
Else
    ExportCablePrix
End If
Me.Hide
End Sub

