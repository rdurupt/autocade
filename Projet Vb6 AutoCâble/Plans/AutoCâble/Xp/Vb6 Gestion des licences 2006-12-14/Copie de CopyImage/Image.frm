VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   DrawStyle       =   6  'Inside Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   5820
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   5415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   6960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim Conn As New Ado
 Dim rs As ADODB.Recordset
Dim Execute As Boolean
Dim Sql As String
Dim Action As Integer
Dim I As Long

Dim Reserved As Integer
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const Flags = SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long






Private Sub Drive1_Change()
Me.Dir1.Path = Me.Drive1.Drive
End Sub

Private Sub Form_Activate()
SetTopMostWindow Me, True
Me.Drive1.Drive = "c:\"

End Sub

Private Sub Form_Load()
'Dim aa
''Me.MSRDC1.Connection '"\\10.30.0.5\donnees d entreprise\Utilitaires\cablage\Catalogues connectique\RENAULT\Catalogue.mdb"
'aa = Me.MSRDC1.Connection.Connect
'aa = Split(aa, "=")
'aa = Split(aa(2), ";")
'Sql = "SELECT Objet.Référence, Symbol.Dessin "
'Sql = Sql & "FROM Symbol RIGHT JOIN Objet ON Symbol.ID = Objet.IDSymbol "
'Sql = Sql & "WHERE Symbol.Dessin Is Not Null;"
'Conn.ConnectionString = Me.MSRDC1.Connection.Connect
'Conn.OpenConnetion Trim("" & aa(0))
'Set rs = Conn.OpenRecordSet(Sql)
Me.Timer1.Interval = 10
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

Private Sub List1_DblClick()
Dim Image As String
Image = Me.Dir1.Path
If Right(" " & Image, 1) <> "\" Then Image = Image & "\"
MyExecute Image & Me.List1.List(Me.List1.ListIndex)
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
Dim NumBouton As Long
Dim Bouton
Dim a As Long
Dim fso As New FileSystemObject
NumBouton = 0
Dim Image As String

 Bouton = GetAsyncKeyState(&H2C)
 If Bouton <> 0 Then
 I = I + 1
 Image = Me.Dir1.Path
 If Right(" " & Image, 1) <> "\" Then Image = Image & "\"
     If fso.FileExists(Image & "Image" & I & ".bmp") = True Then
                fso.DeleteFile Image & "Image" & I & ".bmp"
           End If
           
'                Execute = False
                SavePicture Clipboard.GetData(vbCFBitmap), Image & "Image" & I & ".bmp"
                Me.List1.AddItem "Image" & I & ".bmp", 0
 End If
Set fso = Nothing
'If rs.EOF = True Then End
'Select Case Action
'        Case 0
'            Action = 1
'                 Sql = "SELECT Objet.Référence, Catégorie.Libellé, Symbol.Dessin "
'                Sql = Sql & "FROM Catégorie RIGHT JOIN (Symbol RIGHT JOIN Objet ON Symbol.ID = Objet.IDSymbol) ON Catégorie.ID = Objet.IDCatégorie "
'                Sql = Sql & "WHERE Objet.Référence='" & MyReplace("" & rs!Référence) & "' AND Symbol.Dessin Is Not Null;"
'                Me.MSRDC1.Sql = Sql
'                Me.MSRDC1.Refresh
'        Case 1
'            Action = 2
'            OLE1_Click
'        Case 2
'            Action = 0
'           If Execute = True Then
'           Dim Fso As New FileSystemObject
'           If Trim("" & Label2) = "" Then
'                Label2 = "Autres"
'            End If
'            If Fso.FolderExists(App.Path & "\Images Connectique") = False Then
'            Fso.CreateFolder App.Path & "\Images Connectique"
'           End If
'
'           If Fso.FolderExists(App.Path & "\Images Connectique\" & Label2) = False Then
'            Fso.CreateFolder App.Path & "\Images Connectique\" & Label2
'           End If
'
'           If Fso.FileExists(App.Path & "\Images Connectique\" & Label2 & "\" & Label1 & ".bmp") = True Then
'                Fso.DeleteFile App.Path & "\Images Connectique\" & Label2 & "\" & Label1 & ".bmp"
'           End If
'           Set Fso = Nothing
'                Execute = False
'                SavePicture Clipboard.GetData(vbCFBitmap), App.Path & "\Images Connectique\" & Label2 & "\" & Label1 & ".bmp"
'                DoEvents
'                rs.MoveNext
'            End If
'End Select
'
'
'
'
''Else
''If rs.EOF Then End
''Call keybd_event(&H2C, 1, 0, 0)
'''Execute = True
''End If
End Sub
Public Function SetTopMostWindow(Window As Form, Topmost As Boolean) As Long

    If Topmost = True Then
        SetTopMostWindow = SetWindowPos(Window.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        SetTopMostWindow = SetWindowPos(Window.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
    End If

End Function
Function MyReplace(strVal As String) As String
strVal = Trim(strVal)
MyReplace = strVal
MyReplace = Replace(MyReplace, "'", "''")
MyReplace = Replace(MyReplace, Chr(34), Chr(34) & Chr(34))
MyReplace = Trim("" & MyReplace)
End Function
Sub MyExecute(Fichier As String, Optional Param As String = vbNull)

Dim lapi As Long
On Error Resume Next
lapi = ShellExecute(100, "open", Fichier, Param, vbNull, 5)
'If Err Then MsgBox Err.Description
Err.Clear
On Error GoTo 0
End Sub

