VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2025
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2025
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   2025
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   1560
   End
   Begin MSRDC.MSRDC MSRDC1 
      Height          =   270
      Left            =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   476
      _Version        =   393216
      Options         =   0
      CursorDriver    =   0
      BOFAction       =   0
      EOFAction       =   1
      RecordsetType   =   1
      LockType        =   4
      QueryType       =   0
      Prompt          =   3
      Appearance      =   0
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   0   'False
      DataSourceName  =   "RSA"
      RecordSource    =   $"Image.frx":0000
      UserName        =   ""
      Password        =   ""
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "MSRDC1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DataField       =   "Libellé"
      DataSource      =   "MSRDC1"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      DataField       =   "Référence"
      DataSource      =   "MSRDC1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OLE OLE1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DataField       =   "Dessin"
      DataSource      =   "MSRDC1"
      Height          =   2055
      Left            =   0
      SizeMode        =   1  'Stretch
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim Conn As New Ado
 Dim Rs As ADODB.Recordset
Dim Execute As Boolean
Dim Sql As String
Dim Action As Integer


Dim Reserved As Integer
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const Flags = SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long





Private Sub Form_Activate()
SetTopMostWindow Me, True
End Sub

Private Sub Form_Load()
Dim aa
'Me.MSRDC1.Connection '"\\10.30.0.5\donnees d entreprise\Utilitaires\cablage\Catalogues connectique\RENAULT\Catalogue.mdb"
aa = Me.MSRDC1.Connection.Connect
aa = Split(aa, "=")
aa = Split(aa(2), ";")
Sql = "SELECT Objet.Référence, Symbol.Dessin "
Sql = Sql & "FROM Symbol RIGHT JOIN Objet ON Symbol.ID = Objet.IDSymbol "
Sql = Sql & "WHERE Symbol.Dessin Is Not Null;"
Conn.ConnectionString = Me.MSRDC1.Connection.Connect
Conn.OpenConnetion Trim("" & aa(0))
Set Rs = Conn.OpenRecordSet(Sql)
Me.Timer1.Interval = 50
Me.Timer1.Enabled = True
End Sub

Private Sub MSRDC1_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
Err.Clear
End Sub

Private Sub MSRDC1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
On Error Resume Next
Err.Clear
End Sub

Private Sub MSRDC1_Error(ByVal Number As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Err.Clear
Exit Sub
End Sub

Private Sub MSRDC1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Err.Clear
End Sub

Private Sub MSRDC1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Err.Clear
End Sub

Private Sub MSRDC1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Err.Clear
End Sub

Private Sub MSRDC1_QueryCompleted()
On Error Resume Next
Err.Clear
End Sub

Private Sub MSRDC1_Reposition()
On Error Resume Next
Err.Clear
End Sub

Private Sub MSRDC1_Validate(Action As Integer, Reserved As Integer)
On Error Resume Next
Err.Clear
End Sub

Private Sub OLE1_Click()
Call keybd_event(&H2C, 1, 0, 0)
Execute = True
End Sub

Private Sub OLE1_ObjectMove(Left As Single, Top As Single, Width As Single, Height As Single)
On Error Resume Next
Err.Clear
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Rs.EOF = True Then End
Select Case Action
        Case 0
            Action = 1
                 Sql = "SELECT Objet.Référence, Catégorie.Libellé, Symbol.Dessin "
                Sql = Sql & "FROM Catégorie RIGHT JOIN (Symbol RIGHT JOIN Objet ON Symbol.ID = Objet.IDSymbol) ON Catégorie.ID = Objet.IDCatégorie "
                Sql = Sql & "WHERE Objet.Référence='" & MyReplace("" & Rs!Référence) & "' AND Symbol.Dessin Is Not Null;"
                Me.MSRDC1.Sql = Sql
                Me.MSRDC1.Refresh
        Case 1
            Action = 2
            OLE1_Click
        Case 2
            Action = 0
           If Execute = True Then
           Dim Fso As New FileSystemObject
           If Trim("" & Label2) = "" Then
                Label2 = "Autres"
            End If
            If Fso.FolderExists(App.Path & "\Images Connectique") = False Then
            Fso.CreateFolder App.Path & "\Images Connectique"
           End If
           
           If Fso.FolderExists(App.Path & "\Images Connectique\" & Label2) = False Then
            Fso.CreateFolder App.Path & "\Images Connectique\" & Label2
           End If
           
           If Fso.FileExists(App.Path & "\Images Connectique\" & Label2 & "\" & Label1 & ".bmp") = True Then
                Fso.DeleteFile App.Path & "\Images Connectique\" & Label2 & "\" & Label1 & ".bmp"
           End If
           Set Fso = Nothing
                Execute = False
                SavePicture Clipboard.GetData(vbCFBitmap), App.Path & "\Images Connectique\" & Label2 & "\" & Label1 & ".bmp"
                DoEvents
                Rs.MoveNext
            End If
End Select

    
   
    
'Else
'If rs.EOF Then End
'Call keybd_event(&H2C, 1, 0, 0)
''Execute = True
'End If
End Sub
Public Function SetTopMostWindow(Window As Form, Topmost As Boolean) As Long

    If Topmost = True Then
        SetTopMostWindow = SetWindowPos(Window.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        SetTopMostWindow = SetWindowPos(Window.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If

End Function
Function MyReplace(strVal As String) As String
strVal = Trim(strVal)
MyReplace = strVal
MyReplace = Replace(MyReplace, "'", "''")
MyReplace = Replace(MyReplace, Chr(34), Chr(34) & Chr(34))
MyReplace = Trim("" & MyReplace)
End Function
