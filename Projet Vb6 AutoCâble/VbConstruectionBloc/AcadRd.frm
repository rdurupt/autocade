VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   12750
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Execute"
      Height          =   615
      Left            =   2520
      TabIndex        =   25
      Top             =   10080
      Width           =   3375
   End
   Begin VB.CheckBox Cercle 
      Caption         =   "Cercle ?"
      Height          =   495
      Left            =   840
      TabIndex        =   24
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox NBC 
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Nbl 
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   1080
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "10Z,10Y,10X..."
      Height          =   375
      Index           =   20
      Left            =   2280
      TabIndex        =   19
      Tag             =   "10Z"
      Top             =   9120
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1Z,1Y,1Z..."
      Height          =   375
      Index           =   19
      Left            =   2280
      TabIndex        =   18
      Tag             =   "1Z"
      Top             =   8760
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   " Z10,Z9,Z8..."
      Height          =   375
      Index           =   18
      Left            =   2280
      TabIndex        =   17
      Tag             =   "Z10"
      Top             =   8400
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Z1,Z2,Z3..."
      Height          =   375
      Index           =   17
      Left            =   2280
      TabIndex        =   16
      Tag             =   "Z1"
      Top             =   8040
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "10A,10B,10C..."
      Height          =   375
      Index           =   16
      Left            =   2280
      TabIndex        =   15
      Tag             =   "1A"
      Top             =   7680
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1A,1B,1C..."
      Height          =   375
      Index           =   15
      Left            =   2280
      TabIndex        =   14
      Tag             =   "1A"
      Top             =   7320
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Z-Z,Z-Y,Z-X..."
      Height          =   375
      Index           =   14
      Left            =   2280
      TabIndex        =   13
      Tag             =   "ZZ"
      Top             =   6960
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Z-A,Z-B,Z-C.."
      Height          =   375
      Index           =   13
      Left            =   2280
      TabIndex        =   12
      Tag             =   "ZA"
      Top             =   6600
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "A-Z,A-Y,A-X..."
      Height          =   375
      Index           =   12
      Left            =   2280
      TabIndex        =   11
      Tag             =   "AZ"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "A-A,A-B,A-C..."
      Height          =   375
      Index           =   11
      Left            =   2280
      TabIndex        =   10
      Tag             =   "AA"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   " Z,Y,X..."
      Height          =   375
      Index           =   10
      Left            =   2280
      TabIndex        =   9
      Tag             =   "Z"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   " A10,A9,A8..."
      Height          =   375
      Index           =   9
      Left            =   2280
      TabIndex        =   8
      Tag             =   "A10"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   " A1,A2,A3.."
      Height          =   375
      Index           =   8
      Left            =   2280
      TabIndex        =   7
      Tag             =   "A1"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   " A,B,C..."
      Height          =   375
      Index           =   7
      Left            =   2280
      TabIndex        =   6
      Tag             =   "A"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   " 10-10,10-9,10-8..."
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   5
      Tag             =   "1010"
      Top             =   4080
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1-10,1-9,1-8..."
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   4
      Tag             =   "110"
      Top             =   3720
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "10-1,10-2,10-3..."
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   3
      Tag             =   "101"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1-1,1-2,1-3..."
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   2
      Tag             =   "11"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "10,9,8.."
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Tag             =   "10"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   " 1,2,3..."
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Tag             =   "1"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "NB Lignes"
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nb Colonnes :"
      Height          =   375
      Left            =   840
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyType As String

Private Sub Command1_Click()
Set MyDocDefault = MyAutocad.Documents.Add
MyAutocad.Documents(0).Activate
Demarage MyType, Cercle.Value, Nbl, NBC
MyAutocad.Documents(0).Application.ZoomAll
MsgBox "Fin:"
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim aa As New FunCreateObjet
Dim Atrib
Dim i As Long
Dim Entity
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double
 Dim BlocRef  As AcadBlockReference
Set MyAutocad = New AcadApplication
If Err Then
    MsgBox "Pas de Licence"
    End
End If
MyAutocad.Visible = True
XY1(0) = 0
XY1(1) = 0
XY2(0) = 10
XY2(1) = 10
For i = 0 To MyAutocad.Documents(0).ModelSpace.Count - 1
Set Entity = MyAutocad.Documents(0).ModelSpace.Item(i)
If Entity.ObjectName = "AcDbBlockReference" Then
    Set BlocRef = Entity
    Atrib = BlocRef.GetAttributes
    XY2(0) = BlocRef.InsertionPoint(0) + 10
XY2(1) = BlocRef.InsertionPoint(1) + 10
    aa.CreateLigne BlocRef.Document, BlocRef.InsertionPoint, XY2, 256
'    BlocRef.Delete
 End If
Next
End Sub

Private Sub Option1_Click(Index As Integer)
MyType = Me.Option1(Index).Tag
End Sub
