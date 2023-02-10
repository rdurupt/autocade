VERSION 5.00
Object = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "MSHTML.TLB"
Begin VB.Form Poisson 
   ClientHeight    =   1680
   ClientLeft      =   9045
   ClientTop       =   7260
   ClientWidth     =   1575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   1575
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin MSHTMLCtl.Scriptlet Scriptlet1 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
      Scrollbar       =   0   'False
      URL             =   "file://192.168.1.194/Autocable%20Access/AutoCable%20Client/Images/poisson.gif"
   End
End
Attribute VB_Name = "Poisson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

