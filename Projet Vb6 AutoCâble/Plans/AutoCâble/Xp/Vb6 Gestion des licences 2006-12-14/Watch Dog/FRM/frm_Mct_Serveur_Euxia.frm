VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.OCX"
Begin VB.Form frm_Mct_Serveur_Euxia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serveur POP"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "frm_Mct_Serveur_Euxia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin NTService.NTService NTService1 
      Left            =   2520
      Top             =   1320
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "Serveur MCT (EUXIA)"
      ServiceName     =   "Serveur MCT (EUXIA)"
      StartMode       =   2
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1800
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label LblVersion 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Serveur en attente d'ordre."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   45
      Picture         =   "frm_Mct_Serveur_Euxia.frx":1CFA
      Stretch         =   -1  'True
      Top             =   45
      Width           =   4455
   End
End
Attribute VB_Name = "frm_Mct_Serveur_Euxia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AppOff As Boolean

Private Sub Form_Activate()
Me.Visible = True
On Error Resume Next
Kill dirFTP & "*.*"
DoEvents
    EcrirFile dirFTP & "SERVEURON.TXT"
    DoEvents
Macro_Demarage_Serveur
Iconify Me, MainTitle
AppOff = True
Timer1.Enabled = True
End Sub

Private Sub Form_Load()

    'Call DebugBreak

On Error GoTo Err_Load
    Dim strDisplayName As String
    Dim bStarted As Boolean
'    MsgBox Command
    strDisplayName = NTService1.DisplayName
    
'    StatusBar.Panels(1).Text = "Loading"
    If Command = "?" Then
    Form1.Show vbModal
'    MsgBox App.EXEName & ".EXE & : " & Chr(10) & _
'            "-install" & _
'            " ou " & _
'            "-uninstall" & _
'            " ou " & _
'            "-debug"
    End
    End If
    If Command = "-install" Then
        ' enable interaction with desktop
        NTService1.Interactive = True
        
        If NTService1.Install Then
            Call NTService1.SaveSetting("Parameters", "TimerInterval", "1000")
            MsgBox strDisplayName & " installed successfully"
        Else
            MsgBox strDisplayName & " failed to install"
        End If
        End
        End
    ElseIf Command = "-uninstall" Then
        If NTService1.Uninstall Then
            MsgBox strDisplayName & " uninstalled successfully"
        Else
            MsgBox strDisplayName & " failed to uninstall"
        End If
            End
        End
    ElseIf Command = "-debug" Then
        NTService1.Debug = True
        End
    ElseIf Command <> "" Then
        MsgBox "Invalid command option"
        
        End
    End If
    
'    StatusBar.Panels(1).Text = "Loading configuration"
    Dim parmInterval As String
    parmInterval = NTService1.GetSetting("Parameters", "TimerInterval", "2000")
'    Timer.Interval = CInt(parmInterval)
    
    ' enable Pause/Continue. Must be set before StartService
    ' is called or in design mode
'    StatusBar.Panels(1).Text = "Enabling control mode"
    NTService1.ControlsAccepted = svcCtrlPauseContinue
    
    ' connect service to Windows NT services controller
'    StatusBar.Panels(1).Text = "Starting"
    NTService1.StartService
    
Err_Load:
    If NTService1.Interactive Then
        MsgBox "[" & Err.Number & "] " & Err.Description
        End
    Else
        Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    End If

Me.Label1.Caption = Format(Time, "hh:mm")
Me.LblVersion.Caption = Msg
PathBase = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "DIR")
If Right(PathBase, 1) <> "\" Then PathBase = PathBase & "\"
dirFTP = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "dirFTP")
If Dir(dirFTP, vbDirectory) = "" Then MkDir dirFTP
If Right(dirFTP, 1) <> "\" Then dirFTP = dirFTP & "\"
DirServerPop = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "DirServerPop")
DirZip = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "zip")
If Right(DirZip, 1) <> "\" Then DirZip = DirZip & "\"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim hProcess As Long
    If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
            Case WM_MOUSEMOVE
            Case WM_LBUTTONDOWN
            Case WM_LBUTTONUP
            Case WM_LBUTTONDBLCLK
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
                DeIconify Me
'                Me.StartUpPosition = vbStartUpScreen
            Case WM_RBUTTONDOWN
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
'                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
            Case WM_RBUTTONDBLCLK
        End Select
    End If
End Sub

Private Sub Form_Resize()
 If WindowState = vbMinimized Then
        Iconify Me, MainTitle
        DoEvents
        Exit Sub
    End If
End Sub

Private Sub NTService1_Continue(Success As Boolean)
On Error GoTo Err_Continue
    Timer1.Enabled = True
    AppOff = True
     Success = True
    Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, "Service continued")
    
Err_Continue:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub






Private Sub NTService1_Control(ByVal nEvent As Long)
Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)

End Sub

Private Sub NTService1_Pause(Success As Boolean)
On Error GoTo Err_Pause
    AppOff = False
    Call NTService1.LogEvent(svcEventError, svcMessageError, "Service paused")
    Success = True
    
Err_Pause:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)

End Sub

Private Sub NTService1_Start(Success As Boolean)
On Error Resume Next

 
    Success = True
    
Err_Start:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    
   Form_Load

End Sub

Private Sub NTService1_Stop()
Me.Timer1.Enabled = False
AppOff = False
 DoEvents
'     On Error Resume Next

Dim Cmd As New ADODB.Command
Dim conn As ADODB.Connection
Dim ConnString As String
'Dim conn As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Email As String
Dim strSubject
Dim strMsgId As String
'Tremine = False
Timer1.Enabled = False
Timer2.Enabled = False
 ConnString = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "ConnString")
    Set conn = New ADODB.Connection
   
    conn.ConnectionString = ConnString
    conn.Open
' ProgBusy = False
    Set Cmd.ActiveConnection = conn
    Sql = "UPDATE ServicePop SET ServicePop.[Oui/non] = True " & _
            "WHERE ServicePop.Service='Stop';"
            
   Cmd.CommandText = Sql
   Cmd.Execute

 Set Cmd = Nothing
  Set conn = Nothing
  dirFTP = InputDir(App.Path & "\MCT_Serveur_Euxia.ini", "dirFTP")
  If Right(dirFTP, 1) <> "\" Then dirFTP = dirFTP & "\"

If Dir(dirFTP & "*.*") <> "" Then Kill dirFTP & "*.*"

EcrirFile dirFTP & "ServeuroFF.txt"
'MsgBox ""
 DeIconify Me
Err.Clear

Unload Me

 Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)

End Sub


Private Sub Timer1_Timer()
On Error Resume Next
Dim FileTxt As String

If Dir(App.Path & "\Watch Dog.exe") <> "" Then
    Shell App.Path & "\Watch Dog.exe", vbNormalNoFocus
End If
Me.Timer1.Enabled = False
If Dir(dirFTP & "\SERVEURON.TXT") <> "" Then
    Me.Timer1.Enabled = True
    Exit Sub
End If


FileTxt = Dir(dirFTP & "*.txt", vbNormal)
Select Case UCase(FileTxt)
    Case ""
        EcrirFile dirFTP & "ServeurEncours.txt"
     Case "SERVEURSTOP.TXT"
            Me.LblVersion.Caption = "Arrêt du serveur Pop3 demandé.    "
            STOPSERVEUR
            UnloadeMain
            Kill dirFTP & FileTxt
             EcrirFile dirFTP & "ServeurEncours.txt"
            
            
    Case "SERVEUROFF.TXT"
'    MsgBox "rd"
        EcrirFile dirFTP & "SERVEURON.TXT"
         Kill dirFTP & FileTxt
    Case "STOPSERVEUR.TXT"
   
    Case "TRANSFERBASE.TXT"
        Me.LblVersion.Caption = "Prise en compte du transfère de la base."
        TransfaireBase
         Kill dirFTP & FileTxt
        
    Case "STRATOK.TXT"
            Me.LblVersion.Caption = "Démarrage su serveur Pop3."
            STRATOK
          Kill dirFTP & FileTxt
          Me.LblVersion.Caption = Msg
End Select
Me.Timer1.Enabled = AppOff
DoEvents
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Me.Label1.Caption = Format(Time, "hh:mm")
If Dir(App.Path & "\ServeurStop.txt") <> "" Then
    Kill App.Path & "\ServeurStop.txt"
End If
End Sub
