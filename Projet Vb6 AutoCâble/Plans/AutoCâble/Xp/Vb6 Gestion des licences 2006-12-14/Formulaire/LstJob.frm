VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LstJob 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liste des JOBS:"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   15810
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Stop Serveur"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   5410
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15645
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   99
         TabIndex        =   1
         Top             =   99
         Width           =   15615
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   10350
            Left            =   0
            ScaleHeight     =   10350
            ScaleWidth      =   15255
            TabIndex        =   3
            Top             =   0
            Width           =   15255
            Begin VB.CommandButton KillJob 
               Height          =   300
               Index           =   0
               Left            =   0
               Picture         =   "LstJob.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   0
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton Note 
               Height          =   300
               Index           =   0
               Left            =   11400
               Picture         =   "LstJob.frx":04A2
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   0
               Visible         =   0   'False
               Width           =   375
            End
            Begin MSComctlLib.ProgressBar Barre 
               Height          =   300
               Index           =   0
               Left            =   11760
               TabIndex        =   6
               Top             =   0
               Visible         =   0   'False
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   529
               _Version        =   393216
               Appearance      =   1
               Max             =   1
               Scrolling       =   1
            End
            Begin VB.Label Satatus 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00FF00FF&
               Height          =   300
               Index           =   0
               Left            =   8160
               TabIndex        =   11
               Top             =   0
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.Label PI 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00FF00FF&
               Height          =   300
               Index           =   0
               Left            =   2760
               TabIndex        =   10
               Top             =   0
               Visible         =   0   'False
               Width           =   2775
            End
            Begin VB.Label AC 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00FF00FF&
               Height          =   300
               Index           =   0
               Left            =   5520
               TabIndex        =   9
               Top             =   0
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.Label User 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00FF00FF&
               Height          =   300
               Index           =   0
               Left            =   360
               TabIndex        =   8
               Top             =   0
               Visible         =   0   'False
               Width           =   2415
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   5295
            LargeChange     =   3100
            Left            =   15240
            SmallChange     =   310
            TabIndex        =   2
            Top             =   0
            Value           =   60
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "LstJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As Recordset
Private Sub SqlKillJob(Job)
Dim Sql As String
Sql = "DELETE T_Job.*, T_Job.Job FROM T_Job WHERE T_Job.Job=" & Job & ";"

Con.Execute Sql
End Sub
Private Sub Command1_Click()

Unload Me
End Sub

Private Sub Command2_Click()
Dim Sql As String
Dim Rs As Recordset
Dim P
Dim ServiceObject As SWbemObject  'Objet WMI ( Windows Management Instrumentation)
Dim Locator As SWbemLocator 'Objet de connexion
Dim services As SWbemServices 'objet service
Dim ServiceName As String
On Error Resume Next
Dim Serveur As String, User As String, PassWord As String
Sql = "SELECT AdminAutocable.User, AdminAutocable.PassWord, AdminAutocable.Serveur,AdminAutocable.Service "
Sql = Sql & "FROM AdminAutocable;"
Set Rs = Con.OpenRecordSet(Sql)
Serveur = "" & Rs!Serveur
User = "" & Rs!User
PassWord = "" & Rs!PassWord
Set Locator = New SWbemLocator
Set services = Locator.ConnectServer(Serveur, struser:=User, strpassword:=PassWord)
Set ServiceObject = services.Get("Win32_Service='" & Trim("" & Rs!Service) & "'")
ServiceObject.StopService
ServiceObject.ResumeService
Sql = "SELECT T_Job.* FROM T_Job;"
Set Rs = Con.OpenRecordSet(Sql)

While Rs.EOF = False
    Set ServiceObject = services.Get("Win32_Process='" & Rs!IdApp & "'")
    P = ServiceObject.Terminate
    Set ServiceObject = services.Get("Win32_Process='" & Rs!IdExcel & "'")
    P = ServiceObject.Terminate
    Set ServiceObject = services.Get("Win32_Process='" & Rs!IdExcel2 & "'")
    P = ServiceObject.Terminate
     Set ServiceObject = services.Get("Win32_Process='" & Rs!IdWord & "'")
    P = ServiceObject.Terminate
'    Set ServiceObject = services.Get("Win32_Process='" & Rs!IdAutocad & "'")
'    P = ServiceObject.Terminate
    Rs.MoveNext
Wend

Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT ServerAutocad.IdAutcad FROM ServerAutocad;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
Dim a
a = 1
'While a <> -1
'a = RetournIdApp("acad.exe", True, Serveur, PassWord)
'If a > -1 Then
    Set ServiceObject = services.Get("Win32_Process='" & Rs!IdAutcad & "'")
        'Destruction du processus
        P = ServiceObject.Terminate
End If
'Wend
'End If
Set Locator = Nothing
Set services = Nothing
Set ServiceObject = Nothing


Set Rs = Con.CloseRecordSet(Rs)
MsgBox "Serveur Interrompu"
End Sub

Private Sub Form_Load()
'Dim MeIndex As Long
'Dim Sql As String
'Dim NbJob As Long
'Dim I As Long
''MsgBox "1"
'Sql = "SELECT T_Job.* "
'Sql = Sql & "FROM  T_Job  "
'Sql = Sql & "ORDER BY T_Job.Job;"
'Set Rs = Con.OpenRecordSet(Sql)
''MsgBox "2"
'While Rs.EOF = False
''MsgBox "3"
'    NbJob = NbJob + 1
'
'    'MsgBox "5"
'    Rs.MoveNext
'    'MsgBox "6"
'Wend
''MsgBox "7"
'If NbJob > 0 Then
'    For I = 1 To NbJob
'
'        LoadJob I
'    Next
'End If
'
'Picture1.Height = Me.Barre(0).Height * (NbJob + 1)
'
'Me.VScroll1.Max = Picture1.Height - 60
'Picture1.Height = Picture1.Height * 2
Me.VScroll1.SmallChange = Me.User(0).Height
Me.VScroll1.LargeChange = Me.User(0).Height * 10
MajTimer
Me.Timer1.Interval = 1500
Me.Timer1.Enabled = True

End Sub

Private Sub KillJob_Click(Index As Integer)
    SqlKillJob Me.KillJob(Index).Tag
End Sub

Private Sub Note_Click(Index As Integer)
Shell "notepad.exe " & Note(Index).Tag, vbMaximizedFocus
End Sub

Private Sub Timer1_Timer()
Me.Timer1.Enabled = False
MajTimer
DoEvents
Me.Timer1.Enabled = True
End Sub

Private Sub VScroll1_Change()
Picture1.Top = Me.VScroll1.Value * -1
End Sub
Sub LoadJob(MeIndex As Long)
Load Me.KillJob(MeIndex)
Load Me.PI(MeIndex)
Load Me.User(MeIndex)
Load Me.Satatus(MeIndex)
Load Me.Barre(MeIndex)
Load Me.Note(MeIndex)
Load Me.AC(MeIndex)
'Me.Frame2.Height
Me.User(MeIndex).Visible = True
Me.Note(MeIndex).Visible = True
Me.PI(MeIndex).Visible = True
Me.Satatus(MeIndex).Visible = True
Me.Barre(MeIndex).Visible = True
Me.KillJob(MeIndex).Visible = True
Me.AC(MeIndex).Visible = True
If MeIndex > 1 Then
    Me.Note(MeIndex).Top = Me.Note(MeIndex - 1).Top + (Me.Note(MeIndex - 1).Height)
    Me.KillJob(MeIndex).Top = Me.KillJob(MeIndex - 1).Top + (Me.KillJob(MeIndex - 1).Height)
    Me.User(MeIndex).Top = Me.User(MeIndex - 1).Top + (Me.User(MeIndex - 1).Height)
    Me.PI(MeIndex).Top = Me.PI(MeIndex - 1).Top + (Me.PI(MeIndex - 1).Height)
    Me.Satatus(MeIndex).Top = Me.Satatus(MeIndex - 1).Top + (Me.Satatus(MeIndex - 1).Height)
    Me.Barre(MeIndex).Top = Me.Barre(MeIndex - 1).Top + (Me.Barre(MeIndex - 1).Height)
    Me.AC(MeIndex).Top = Me.AC(MeIndex - 1).Top + (Me.AC(MeIndex - 1).Height)
End If
End Sub
Sub UnLoadJob(MeIndex As Long)
Unload Me.KillJob(MeIndex)
Unload Me.PI(MeIndex)
Unload Me.User(MeIndex)
Unload Me.Satatus(MeIndex)
Unload Me.Barre(MeIndex)
Unload Me.Note(MeIndex)
Unload Me.AC(MeIndex)
'Me.Frame2.Height
End Sub
Sub MajTimer()
Dim NbJob As Long
Dim IndexJob As Long
Dim Sql As String
Dim I As Long

'MsgBox "1"
Sql = "SELECT T_Job.* "
Sql = Sql & "FROM  T_Job  "
Sql = Sql & "ORDER BY T_Job.Job;"
Set Rs = Con.OpenRecordSet(Sql)
'MsgBox "2"
NbJob = 0
While Rs.EOF = False
'MsgBox "3"
    NbJob = NbJob + 1
   
    'MsgBox "5"
    Rs.MoveNext
    'MsgBox "6"
Wend
Rs.Requery
Do While Rs.EOF = False
    IndexJob = IndexJob + 1
    If NbJob = 0 Then Exit Do
    If IndexJob > Me.Barre.Count - 1 Then
        LoadJob IndexJob
    End If
    If Rs!ValBarGraph > Me.Barre(IndexJob).Max Then
        Me.Barre(IndexJob).Value = 0
          End If
    If Rs!MaxBarGraph = 0 Then
    Me.Barre(IndexJob).Max = 1
    Else
        Me.Barre(IndexJob).Max = Val("" & Rs!MaxBarGraph)
    End If
    Me.User(IndexJob) = "" & Rs!Machine
    Me.Barre(IndexJob).Value = Val("" & Rs!ValBarGraph)
    Me.PI(IndexJob) = "" & Rs!Name
    Me.Satatus(IndexJob) = "" & Rs!Status
    Me.KillJob(IndexJob).Tag = Rs!Job
    Me.AC(IndexJob) = "" & Rs!Action

     If Trim("" & Rs!FichierErr <> "") Then
        Me.Note(IndexJob).Enabled = True
        Me.Note(IndexJob).Tag = "" & Rs!FichierErr
     Else
        Me.Note(IndexJob).Enabled = False
     End If
    If Rs!FinTraitement = True And UserName = "" & Rs!Machine Then
        Me.KillJob(IndexJob).Enabled = True
    Else
        Me.KillJob(IndexJob).Enabled = False
    End If
    Rs.MoveNext
Loop


For I = Me.Barre.Count - 1 To NbJob + 1 Step -1
    UnLoadJob I
Next
DoEvents
End Sub
