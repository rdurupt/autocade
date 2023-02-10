VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2430
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10470
   ControlBox      =   0   'False
   DrawWidth       =   10
   FillColor       =   &H80000012&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   5
      TabIndex        =   0
      Top             =   10
      Width           =   10455
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   1508
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Phase d'Initialisation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   10215
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   8160
         Picture         =   "frmInit.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

