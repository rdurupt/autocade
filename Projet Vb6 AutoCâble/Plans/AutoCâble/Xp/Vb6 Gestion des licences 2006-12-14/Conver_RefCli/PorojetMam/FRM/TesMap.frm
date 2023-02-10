VERSION 5.00
Object = "{A32A88B3-817C-11D1-A762-00AA0044064C}#1.0#0"; "mscecomdlg.dll"
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form TesMap 
   Caption         =   "Form1"
   ClientHeight    =   12180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17625
   LinkTopic       =   "Form1"
   ScaleHeight     =   12180
   ScaleWidth      =   17625
   StartUpPosition =   3  'Windows Default
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   17415
      HTMLURL         =   ""
      HTMLData        =   $"TesMap.frx":0000
      DataType        =   "HTMLDATA"
      AutoFit         =   0   'False
      DisplayColHeaders=   -1  'True
      DisplayGridlines=   -1  'True
      DisplayHorizontalScrollBar=   -1  'True
      DisplayRowHeaders=   -1  'True
      DisplayTitleBar =   -1  'True
      DisplayToolbar  =   -1  'True
      DisplayVerticalScrollBar=   -1  'True
      EnableAutoCalculate=   -1  'True
      EnableEvents    =   -1  'True
      MoveAfterReturn =   -1  'True
      MoveAfterReturnDirection=   0
      RightToLeft     =   0   'False
      ViewableRange   =   "1:65536"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lencer le &Test"
      Height          =   495
      Left            =   1080
      TabIndex        =   26
      Top             =   11520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Height          =   1215
      Left            =   11520
      Picture         =   "TesMap.frx":0DED
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin CEComDlgCtl.CommonDialog CommonDialog1 
      Left            =   11520
      Top             =   4080
      _cx             =   847
      _cy             =   847
      CancelError     =   0   'False
      Color           =   0
      DefaultExt      =   ""
      DialogTitle     =   ""
      FileName        =   "*.csv"
      Filter          =   ""
      FilterIndex     =   0
      Flags           =   0
      HelpCommand     =   0
      HelpContext     =   ""
      HelpFile        =   ""
      InitDir         =   ""
      MaxFileSize     =   256
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   ""
      FontSize        =   10
      FontUnderline   =   0   'False
      Max             =   0
      Min             =   0
      FontStrikethru  =   0   'False
   End
   Begin VB.Label Combo1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   25
      Top             =   240
      Width           =   6975
   End
   Begin VB.Label Val1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   24
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label Val2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   23
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Label Val3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   22
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label Val4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   21
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Label Val5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   20
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Label Val6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   19
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label Val7 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   18
      Top             =   2880
      Width           =   6975
   End
   Begin VB.Label Val9 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   17
      Top             =   3600
      Width           =   6975
   End
   Begin VB.Label Val10 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   16
      Top             =   3960
      Width           =   6975
   End
   Begin VB.Label Val11 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   15
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Label Label1 
      Height          =   315
      Left            =   1560
      TabIndex        =   14
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label3 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label4 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label5 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label6 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label7 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label9 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label10 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label11 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label8 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Val8 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   6975
   End
   Begin VB.Label Label15 
      Caption         =   "Bloc d'empreinte"
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "TesMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MesCon As Collection


Private Sub Command1_Click()
Set MesCon = Nothing
Set MesCon = New Collection
Dim MonCon As ClsMapCon
On Error Resume Next
Dim NumFil As Long
Dim NumFil2 As Long
Dim txt
Dim I As Long
Dim L As Long
Dim C As Long
Dim MyRange
Dim IndexCo As Long
CommonDialog1.CancelError = True
CommonDialog1.ShowOpen
If Trim("" & CommonDialog1.FileName) <> "*.csv" Then
NumFil = FreeFile
IndexCo = 1
Open CommonDialog1.FileName For Input As #NumFil
If Not EOF(NumFil) Then


Set MyRange = Me.Spreadsheet1.Range("A1").CurrentRegion
For I = MyRange.Rows.Count To 2 Step -1
MyRange(I, 1).DeleteRows
Next
MyRange(1, 1).Select
    Line Input #NumFil, txt
    txt = Split(txt, ";")
    Combo1.Caption = "" & txt(UBound(txt))
    For I = 0 To (UBound(txt) - 2) Step 2
        Controls("Label" & CStr(IndexCo)).Caption = "" & txt(I)
         Controls("Val" & CStr(IndexCo)).Caption = "" & txt(I + 1)
        IndexCo = IndexCo + 1
    Next
    Set MyRange = Nothing
End If

NumFil2 = FreeFile
Open APP.Path & "\Map\" & Combo1.Caption & ".map" For Input As #NumFil2
If Not EOF(NumFil) Then
 Line Input #NumFil2, txt
    txt = Split(txt, ";")
    For I = 0 To UBound(txt)
        Set MonCon = New ClsMapCon
        MonCon.ConName = "" & txt(I)
        MesCon.Add MonCon, "Con_" & txt(I)
        Set MonCon = Nothing
    Next
    
End If
While Not EOF(NumFil2)
     Line Input #NumFil2, txt
    txt = Split(txt, ";")
    MesCon("Con_" & txt(1)).AjouterLiason Val("" & txt(0)), "" & txt(2)
Wend
Close #NumFil2
L = 2
While Not EOF(NumFil)
    Line Input #NumFil, txt
     txt = Split(txt, ";")
    
    For C = 0 To UBound(txt)
    DoEvents
     Me.Spreadsheet1.Cells(L, 1).Select
    Me.Spreadsheet1.Cells(L, C + 1) = "'" & txt(C)
    Next
    L = L + 1
Wend
Close #NumFil
End If
Me.Spreadsheet1.Range("a1").CurrentRegion.AutoFitColumns
Me.Spreadsheet1.Range("a1").CurrentRegion.AutoFitRows
 Me.Spreadsheet1.Range("A2").Select
End Sub

Private Sub Command2_Click()
Dim MyRange
Dim I As Long
Dim NuError As Integer
Dim NumPin As Integer
Set MyRange = Me.Spreadsheet1.Range("A1").CurrentRegion
Visue.Liai = ""
Visue.APP = ""
Visue.VOIE = ""
Visue.APP2 = ""
Visue.Voie2 = ""
Visue.Erreur = ""
Visue.OUT = ""
Visue.OUT2 = ""
Visue.Visible = True
For I = 2 To MyRange.Rows.Count
NuError = 0
NumPin = 257
Visue.Erreur = ""
MyRange(I, 1).Select
Visue.Liai = MyRange(I, 1)
Visue.OUT = MyRange(I, 4)
Visue.APP = MyRange(I, 5)
Visue.VOIE = MyRange(I, 6)
Visue.OUT2 = MyRange(I, 8)
Visue.APP2 = MyRange(I, 9)
Visue.Voie2 = MyRange(I, 10)

Reprise:
Visue.Erreur = MyRange(I, 12)
MyRange(I, 1).EntireRow.Interior.Color = 39423
InputBox "entrez la valeur"
If (I Mod 5) = 0 Then
MyRange(I, 1).EntireRow.Interior.Color = 255
MyRange(I, 12) = "ERR"

NuError = NuError + 1

If NuError < 4 Then GoTo Reprise
Else
   MyRange(I, 12) = ""
MyRange(I, 1).EntireRow.Interior.Color = 65280 '65280
End If
Debug.Print Me.Spreadsheet1.Range("b2").Interior.Color
Next
Unload Visue
MsgBox "Fin du traitement"
End Sub
