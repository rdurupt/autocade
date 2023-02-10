VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{25B2612A-C0EB-452C-BF1F-1F43AC892C8D}#2.0#0"; "RdSmtp.ocx"
Begin VB.Form frmPo3Start 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmPo3Start.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin RdSmtp.Email Email1 
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   582
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   327681
      FullWidth       =   169
      FullHeight      =   33
   End
   Begin VB.Label LblReponse 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   4455
   End
End
Attribute VB_Name = "frmPo3Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Email1_Reponse(Evenement As String)
LblReponse = Evenement
End Sub

Private Sub Form_Activate()
'Me.Animation1.Open App.Path & "\Image\animaux_178.gif"
Me.Email1.SeveurSmtp = "10.11.1.148"
Me.Email1.Envoie "TOTO", "robert.durupt@encelade.fr", "robert.durupt@encelade.fr", _
"Robert Durupt", "Test", "Voici Mon Test", "C:\Rd\POP3\Mal2\Ocx\RdSmtp.ocx"
End Sub

