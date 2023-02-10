VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "ntsvc.ocx"
Object = "{5983D6D0-93AB-11CF-A14F-0080C80B2692}#1.1#0"; "ESPOP32.OCX"
Begin VB.Form frmPricipal 
   Caption         =   "AutoCâble Serveur"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "frmPricipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1500
      Left            =   240
      Top             =   5760
   End
   Begin ESPOPLib.EsPop EsPop1 
      Left            =   1800
      Top             =   4440
      _Version        =   65537
      _ExtentX        =   2778
      _ExtentY        =   1508
      _StockProps     =   0
      POPServer       =   ""
      Username        =   "autocable@encelade.fr"
      TimeOut         =   5000
      Port            =   110
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   337
      Width           =   4455
   End
   Begin NTService.NTService NTService1 
      Left            =   3120
      Top             =   5880
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "AutoCâble Serveur"
      Interactive     =   -1  'True
      ServiceName     =   "AutoCableServeur"
      StartMode       =   3
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   960
      Top             =   5760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   75
      Width           =   4200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5392
      Width           =   4200
   End
End
Attribute VB_Name = "frmPricipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NuLigne As Long
Dim InitLogOk As Boolean


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Iconify Me, MainTitle
IinitLog
    
    
   
    
        
   
End Sub
Sub IinitLog()
If InitLogOk = True Then Exit Sub
InitLogOk = True
Set Trace = Nothing
Set Trace = New ClsLog

NuLigne = 0
Trace.Enable = True         ' Active l'ecriture des log
    Trace.MaxSize = 10 * 1024   ' Taille en Ko
    ' Definition de l'entete des fichiers LOG
    ' %NOM% sont remplacé par des valeurs
    Trace.Entete = "******************************************************************************************************" & vbCrLf & _
                   "Date : %DATE%" & vbTab & "Time : %HEURE%" & vbCrLf & _
                   "Programme de Autocâble serveur : %APP%.exe [%DATEMODIFAPP%] %VERSION% Build %BUILD%" & vbCrLf & _
                   "Répertoire courant : " & App.path & vbCrLf & _
                   "******************************************************************************************************" & vbCrLf
    Trace.LogDate = True                        ' Activation de l'affichage de la date
    Trace.Format_Date = "dd-mm-yy"              ' Format d'affichage de la date
    Trace.LogTime = True                        ' Activation de l'ecriture de l'heure sur chaque ligne
    Trace.Format_Heure = "hh:mm:ss"             ' Format d'affichage de l'heure
    Trace.FileName = App.path & "\Server\AutocâbleServer.Log"    ' Definition du chemin
    
    ' On ecrit qq ligne pour voir :)
    Trace.Ecrire ""
End Sub
Private Sub Form_Initialize()
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
Fin: LoadDb
MyTimer = Format(Now, "hh:mm:ss")
Me.Timer1.Interval = 100
'Me.Timer1.Enabled = True
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
        Iconify Me, MainTitle
        DoEvents
        Exit Sub
    End If
End Sub

Private Sub Form_Terminate()
Dim Sql As String
If boolAutoCAD = True Then
    AutoApp.Quit
    Set AutoApp = Nothing
    Con.TYPEBASE = 5
    Con.SERVER = ""
    Con.User = ""
    Con.PassWord = ""
    Con.BASE = Db

    Con.OpenConnetion
     Sql = "UPDATE ServerAutocad SET ServerAutocad.IdAutcad = 0 "
        Sql = Sql & "WHERE ServerAutocad.Id=1;"
        Con.Execute Sql
        Con.CloseConnection
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

Private Sub SMTP1_ConnectSMTP()

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Timer1.Enabled = False
Me.Text1.Refresh
DoEvents
Dim KillJob As Boolean
Dim IdApp As Long
Dim Sql As String
Dim rs As Recordset
Dim RsJobOk As Recordset
Dim NbJeton As Long
Dim NbJetonActif As Long
Dim IndexJon As Long
Static Heurdeb As Date
Static NbTour As Long

Dim AutocadDoc As Object
Dim TableauJob() As String
Dim TableauJob2() As String
Dim AutoCableNbTest As Long
Dim i As Long
Dim NoKill As Boolean
Dim ServiceObject As SWbemObject 'Objet WMI ( Windows Management Instrumentation)
Dim Locator As SWbemLocator 'Objet de connexion
Dim services As SWbemServices 'objet service
Dim IdJob As Long
Dim txtError As String
Dim CompactApp As String
Dim CompactHeure As String
Dim CompactExecute As String
Dim CompactHeureeStart As String
Dim CompactHeureFourchette As String
Set CodageX = New CDETXT
RepriseJeton:
NbJeton = Val(CodageX.LireJeton)
NbJetonActif = CodageX.Decrypt(FiledLicence.General.NbJetonActif, "")
If NbJetonActif < 0 Then
    NbJetonActif = 0
       CodageX.DcrJenton
GoTo RepriseJeton
End If
'if NbJeton>
CompactApp = CherCheInFihier("CompactApp")
CompactHeure = CherCheInFihier("CompactHeure")
CompactExecute = CherCheInFihier("CompactExecute")
CompactHeureFourchette = CherCheInFihier("CompactHeureFourchette ")
If UCase(CompactExecute) = "TRUE" Then
    If DateDiff("s", CDate(CompactHeure), Format(Hour(Now), "0#") & ":" & Format(Minute(Now), "0#") & ":" & Format(Second(Now), "0#")) <= Val(CompactHeureFourchette) Then
        If CompactOk = False Then
            Set Locator = New SWbemLocator
            Set services = Locator.ConnectServer("localhost")
            Set ServiceObject = services.Get("Win32_Service='" & Trim(ServiceName) & "'")
            ServiceObject.StopService
            ServiceObject.ResumeService
        MySeconde 3
        On Error Resume Next
        IdApp = 0
         IdApp = Shell(CompactApp, vbNormalFocus)
         If Err Then GoTo FinExecute
            While RetournIdAppName(IdApp) <> ""
                DoEvents
            Wend
            CompactOk = True
FinExecute:
            ServiceObject.StartService
            ServiceObject.ResumeService
            MySeconde 3
            
         End If
    Else
        CompactOk = False
    End If
End If

If UCase(Format(Label1, "dddd  dd mmm yyyy")) <> UCase(Format(Now, "dddd  dd mmm yyyy")) Then
    NuLigne = 0
End If

If Heurdeb = "00:00:00" Then Heurdeb = Time
If Hour(Heurdeb) + 1 <= Hour(Time) Then
Heurdeb = Time
Me.Text1 = ""
End If

LoadDb

Label1 = UCase(Format(Now, "dddd  dd mmm yyyy hh:mm:ss"))
Set CodageX = Nothing
Label2.Caption = NbJeton & " Jeton(s) disponible(s)"


Dim NbAutocable As Integer
If Trim("" & MyTimer) = "" Then
   MyTimer = Time
End If

TimerInerval = CherCheInFihier("TimerInerval")
If Abs(DateDiff("s", MyTimer, Time)) < Val(TimerInerval) Then


    GoTo Fin
End If
If Dir(App.path & "\KillService.txt") <> "" Then
 Kill App.path & "\KillService.txt"
 NoKill = False
Else
   NoKill = True
End If
If NoKill = False Then
    NbTour = 0
Else
NbTour = NbTour + 1
    If NbTour = 7 Then
        NbTour = 0
        Set Locator = New SWbemLocator
        Set services = Locator.ConnectServer("localhost")
        Set ServiceObject = services.Get("Win32_Service='" & Trim(ServiceName) & "'")
        ServiceObject.StopService
        ServiceObject.ResumeService
        MySeconde 3
        ServiceObject.StartService
        ServiceObject.ResumeService
    End If
End If
MyTimer = Time
If NbJeton = 0 Then GoTo Fin
Con.TYPEBASE = 5
    Con.SERVER = ""
    Con.User = ""
    Con.PassWord = ""
    Con.BASE = Db
If Con.OpenConnetion = False Then

'Me.Text1 = "Connexion False"
GoTo Fin
End If
'Shell App.Path & "\Watch Dog.exe"
Sql = "DELETE T_Job.*  From T_Job "
Sql = Sql & "WHERE DateDiff('d',[DateDebut],Date())>=2  "
Sql = Sql & "AND T_Job.FinTraitement=True;"
Con.Execute Sql

  
RepriseTest:
DoEvents
    On Error Resume Next
 
 
   
If boolAutoCAD = False Then LireLicenceAcad
   

If boolAutoCAD = False Then GoTo Fin
On Error GoTo Erreur
  Sql = "SELECT T_Job.Job, T_Job.DateDebut, T_Job.BarGraphMaj, T_Job.FinTraitement,  "
   Sql = Sql & "T_Job.IdApp, T_Job.IdExcel, T_Job.IdExcel2, T_Job.IdWord, T_Job.AutocadDoc "
   Sql = Sql & "From T_Job "
   Sql = Sql & "Where (((T_Job.DateDebut) Is Not Null) And ((T_Job.FinTraitement) = false)) "
   Sql = Sql & "ORDER BY T_Job.Job;"

 Set rs = Con.OpenRecordSet(Sql)
 KillJob = False
 While rs.EOF = False
    If Trim("" & rs!BarGraphMaj) = "" Then
        If DateDiff("n", rs!DateDebut, Now) = 30 Then KillJob = True
        
    Else
         If DateDiff("n", rs!BarGraphMaj, Now) > 30 Then KillJob = True
    End If
    If KillJob = True Then
           Set AutocadDoc = AutoApp.Documents(Trim("" & rs!AutocadDoc))
        If Err = 0 Then
            AutocadDoc.Close False
        End If
            cmddelproc rs!IdApp
            cmddelproc rs!IdExcel
            cmddelproc rs!IdExcel
            cmddelproc rs!IdWord
            
            Set AutocadDoc = Nothing
            Sql = "UPDATE T_Job SET T_Job.DateDebut = Null, T_Job.BarGraphMaj = Null,  "
            Sql = Sql & "T_Job.FinTraitement = False "
            Sql = Sql & "WHERE T_Job.Job=" & rs!Job & ";"
            Con.Execute Sql
            Con.Execute Sql
            Con.Execute Sql
            Set CodageX = New CDETXT
            CodageX.DcrJenton
            Set CodageX = Nothing
    End If
    rs.MoveNext
 Wend
 
ReDim TableauJob(NbJeton)
ReDim TableauJob2(NbJeton)
Sql = "SELECT T_Job.*, T_Job.Job, T_Job.DateDebut From T_Job Where (((T_Job.DateDebut) Is Null)) ORDER BY T_Job.Job;"


Set rs = Con.OpenRecordSet(Sql)
   
    Do While rs.EOF = False
    
        IndexJon = IndexJon + 1
     Me.Text1 = "Exécution du Job N°  :" & rs!Job & " " & rs!Action & " " & rs!Name & vbCrLf & Me.Text1
     TableauAtotocable(NbAutocable).Job = rs!Job
   TableauJob(NbAutocable) = rs!Job
   TableauJob2(NbAutocable) = "Exécution du Job N°  :" & rs!Job & " " & rs!Action & " " & rs!Name
     NbAutocable = NbAutocable + 1
     If NbAutocable = 255 Then NbAutocable = 0
     'If IndexJon = NbJeton Then Exit Do
     Exit Do
rs.MoveNext
Loop
For i = 0 To NbAutocable - 1
If Trim("" & TableauJob(i)) <> "" Then
'    If Con.Execute("UPDATE T_Job SET T_Job.DateDebut =#" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "# WHERE T_Job.Job=" & TableauJob(i) & ";") = True Then
'        Set RsJobOk = Con.OpenRecordSet("SELECT T_Job.Job, T_Job.DateDebut From T_Job WHERE T_Job.Job=" & TableauJob(i) & " " & _
'        "AND T_Job.DateDebut Is Not Null;")
'        RsJobOk.Requery
'        If Err = 0 Then
'            If RsJobOk.EOF = False Then
            
                NuLigne = NuLigne + 1
'                Trace.Ecrire "Ligne " & str(NuLigne) & " " & TableauJob2(i)
                
                Con.Execute "UPDATE T_Job SET T_Job.DateDebut =#" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "#,T_Job.IdApp = " & _
                Shell(App.path & "\AutoCable.exe " & TableauJob(i), vbNormalFocus) & " WHERE T_Job.Job=" & TableauJob(i) & ";"
                MySeconde 3
'
'            End If
        End If
            Err.Clear
        RsJobOk.Close
        Set RsJobOk = Nothing
'     End If
'    End If
    Exit For
Next
Con.CloseConnection
Set Con = Nothing

 MyTimer = Time
' AutoApp.Quit
' Set AutoApp = Nothing
Fin:
Timer2.Enabled = True
On Error GoTo 0
DoEvents
Exit Sub
Erreur:
txtError = Err.Description
NuLigne = NuLigne + 1
'Trace.Ecrire "Ligne " & str(NuLigne) & " " & txtError
Resume Next
End Sub

Private Sub Timer2_Timer()
Dim a As Long
Dim i As Long
Dim Sql As String
Dim rs As Recordset
Static Depop As Date
On Error Resume Next
Timer1.Enabled = False
Timer2.Enabled = False

LoadDb
If Depop = "00:00:00" Then
Depop = POpInter
End If

If DateDiff("s", Depop, Format(Now, "hh:nn:ss")) >= (((Hour(POpInter) * 60) + Minute(POpInter)) * 60) + Second(POpInter) Then
Depop = Format(Now, "hh:nn:ss")

Con.TYPEBASE = 5
    Con.SERVER = ""
    Con.User = ""
    Con.PassWord = ""
    Con.BASE = Db
If Con.OpenConnetion = True Then
    Sql = "SELECT T_Serveur_POP3.* FROM T_Serveur_POP3;"
    Set rs = Con.OpenRecordSet(Sql)
    If rs.EOF = False Then
        
        Me.EsPop1.Username = "" & rs!Utilisatuer
        EsPop1.PassWord = "" & rs!PassWord
        EsPop1.POPServer = "" & rs!POP3
        EsPop1.Port = "" & rs!Port
        EsPop1.LogOn
        a = EsPop1.NumOfMsg
        For i = a To 1 Step -1
            EsPop1.DeleteMail (i)
        Next
        EsPop1.LogOff
        
       
    End If
    Set rs = Con.CloseRecordSet(rs)
    Con.CloseConnection
End If
End If
Fin:
 Timer1.Enabled = True
On Error GoTo 0
End Sub
