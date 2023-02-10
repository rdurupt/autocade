VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Scrol 
   Caption         =   "Compactage"
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   Icon            =   "Scrol.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "Scrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

