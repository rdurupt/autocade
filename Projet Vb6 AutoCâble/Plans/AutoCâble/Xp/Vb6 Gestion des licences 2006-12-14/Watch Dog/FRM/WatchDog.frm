VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Begin VB.Form WatchDog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Watch Dog"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3120
   Icon            =   "WatchDog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3120
   StartUpPosition =   1  'CenterOwner
   Begin NTService.NTService NTService1 
      Left            =   2160
      Top             =   960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "Autoâble Watch Dog"
      Interactive     =   -1  'True
      ServiceName     =   "WatchDog"
      StartMode       =   2
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   840
   End
   Begin VB.Image Image2 
      Height          =   1455
      Left            =   50
      Picture         =   "WatchDog.frx":08CA
      Stretch         =   -1  'True
      Top             =   50
      Width           =   3015
   End
End
Attribute VB_Name = "WatchDog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
 



Private Sub Form_Activate()
Me.Caption = "Watch Dog"
Iconify Me, "Watch Dog"
End Sub

Private Sub Form_Load()
Dim Fso As FileSystemObject
Set Fso = New FileSystemObject
Fso.CreateTextFile App.Path & "\KillService.txt"
Set Fso = Nothing
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
            MsgBox strDisplayName & " Service installé."
        Else
            MsgBox strDisplayName & "  Erreur Service pas  installé."
        End If
        End
        End
    ElseIf Command = "-uninstall" Then
        If NTService1.Uninstall Then
            MsgBox strDisplayName & " Service Désinstallé."
        Else
            MsgBox strDisplayName & " Erreur Service pas Désinstallé."
        End If
            End
        End
    ElseIf Command = "-debug" Then
        NTService1.Debug = True
        End
    ElseIf Command <> "" Then
        MsgBox "Command option Invalide"
        
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
GoTo Fin
Err_Load:
    If NTService1.Interactive Then
        MsgBox "[" & Err.Number & "] " & Err.Description
        End
    Else
        Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    End If

Fin:
Me.Timer1.Interval = 100
Me.Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim hProcess As Long
    If Not Visible Then
        Select Case x \ Screen.TwipsPerPixelX
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
        Iconify Me, "Watch Dog"
        DoEvents
        Exit Sub
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 DeIconify Me
End Sub
Function MySeconde(NuSeconde As Integer)
Dim a
 a = Second(Time)
    While Abs(a - Second(Time)) < NuSeconde
        DoEvents
    Wend
End Function

Private Sub NTService1_Continue(Success As Boolean)
On Error GoTo Err_Continue
    Timer1.Enabled = True
    AppOff = True
     Success = True
    Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, "Service continued")
    
Err_Continue:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
    
End Sub











Private Sub NTService1_Control(ByVal Event As Long)
Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
DoEvents
End Sub

Private Sub NTService1_Pause(Success As Boolean)
On Error GoTo Err_Pause
    AppOff = False
    Me.Timer1.Enabled = False
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
    Me.Timer1.Enabled = True


End Sub

Private Sub NTService1_Stop()
DeIconify Me
Con.CloseConnection
Unload Me

 Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)

End Sub

Private Sub Timer1_Timer()
Dim Fso As FileSystemObject
Dim ServiceObject As SWbemObject 'Objet WMI ( Windows Management Instrumentation)
Dim Locator As SWbemLocator 'Objet de connexion
Dim services As SWbemServices 'objet service
Dim sql As String
Dim Rs As Recordset
Static BbTour As Long
 
Static MyMinuteSave
Static PremierPassage As Boolean
Dim MyMinute As Date

Interval = InputDir(App.Path & "\Watch Dog.ini", "Interval")
ServiceName = InputDir(App.Path & "\Watch Dog.ini", "Service")
BdAutocable = InputDir(App.Path & "\Watch Dog.ini", "BdAutocable")

sql = "SELECT T_Job.Job, T_Job.DateDebut, T_Job.IdApp, T_Job.IdAutocad, T_Job.IdExcel,  "
sql = sql & "T_Job.IdExcel2, T_Job.IdWord "
sql = sql & "From T_Job "
sql = sql & "WHERE DateDiff('h',[DateDebut],Now())>=1  "
sql = sql & "AND T_Job.FinTraitement=False;"
Con.BASE = BdAutocable
Con.TYPEBASE = 5
Con.OpenConnetion
Set Rs = Con.OpenRecordSet(sql)
While Rs.EOF = False
    cmddelproc Rs!IdApp, "AutoCable.exe"
    cmddelproc Rs!IdAutocad, "acad.exe"
    cmddelproc Rs!IdExcel, "EXCEL.EXE"
    cmddelproc Rs!IdExcel2, "EXCEL.EXE"
    CodageX.DcrJenton
   
    Rs.MoveNext
Wend
sql = "UPDATE T_Job SET T_Job.DateDebut = Null, T_Job.IdApp = 0, T_Job.IdAutocad = 0,   "
sql = sql & "T_Job.IdExcel = 0, T_Job.IdExcel2 = 0, T_Job.IdWord = 0  "
sql = sql & "WHERE DateDiff('h',[DateDebut],Now())>=1   "
sql = sql & "AND T_Job.FinTraitement=False;"
Con.Execute sql
Set Rs = Con.CloseRecordSet(Rs)
Con.CloseConnection
On Error Resume Next
Timer1.Enabled = False
'Me.Label1.Caption = I & ": " & Format(Time, "hh:mm:ss")
MyMinute = Format(Time, "hh:mm:ss")
If "" & MyMinuteSave = "" Then MyMinuteSave = Time
If DateDiff("s", MyMinuteSave, Time) < Val(Interval) Then
    GoTo Fin2
End If
MyMinuteSave = Time
If Dir(App.Path & "\KillService.txt") <> "" Then
    BbTour = BbTour + 1

    If BbTour = 5 Then
    BbTour = 0
           Kill App.Path & "\KillService.txt"
        BbTour = 0
        
        
        Set Locator = New SWbemLocator
        Set services = Locator.ConnectServer("")
        Set ServiceObject = services.Get("Win32_Service='" & Trim(ServiceName) & "'")
        ServiceObject.StopService
        ServiceObject.ResumeService
        MySeconde 3
        RetournIdApp
        KillProcessus
         CodageX.ReInitJenton
        ServiceObject.StartService
        ServiceObject.ResumeService
        Set Locator = Nothing
      Set services = Nothing
       Set ServiceObject = Nothing
     End If
 Else
 BbTour = 0
End If
Set Fso = New FileSystemObject
Fso.CreateTextFile App.Path & "\KillService.txt"
Set Fso = Nothing
DoEvents
GoTo Fin2
Fin:
Fin2:
Timer1.Enabled = True
End Sub
