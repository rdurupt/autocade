VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EditUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion des utilisateurs"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   Icon            =   "EdiUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Droits"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Double click pour Enregistrer."
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Negotiate       =   -1  'True
      Top             =   4320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   2
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "EditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Groupe As Long


Private Sub DBGrid1_Click()

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Droits.chargement Groupe
Unload Droits
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
If MsgBox("Voulez vous vraiment supprimer le User :  " & Me.Adodc1.Recordset.Fields(0), vbYesNo + vbQuestion) = vbNo Then Cancel = 1
End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next

Dim aa

aa = "" & Me.DataGrid1.Columns(1).Value
If Err = 0 Then
    Me.Adodc1.Recordset.Update
    DoEvents
    Me.Adodc1.Recordset.Requery
End If
On Error GoTo 0
End Sub

Private Sub Form_Load()
Dim sql As String
Dim MyB As Boolean
Dim a
Dim Rs As ADODB.Recordset
sql = "SELECT T_Users.Id, T_Groupe.Niveaux "
sql = sql & "FROM T_Users INNER JOIN (T_Groupe INNER JOIN T_Groupe_Users  "
sql = sql & "ON T_Groupe.id = T_Groupe_Users.Id_Groupe) ON T_Users.Id = T_Groupe_Users.Id_Users "
sql = sql & "WHERE T_Users.Id=" & Id_Users & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = True Then GoTo Fin
 Set a = DataGrid1.Columns(0)
' Id_Users
Groupe = Rs!Niveaux
Command2.Enabled = True
sql = "SELECT T_Users.User, T_Users.PassWord, T_Users.Email, T_Users.Cloturer "
sql = sql & "FROM (T_Users LEFT JOIN T_Groupe_Users ON T_Users.Id =  "
sql = sql & "T_Groupe_Users.Id_Users) LEFT JOIN T_Groupe ON T_Groupe_Users.Id_Groupe = T_Groupe.id "
sql = sql & "Where T_Groupe.Niveaux >=  " & Groupe & " "
sql = sql & "Or (((T_Groupe.Niveaux) Is Null)) "
sql = sql & "ORDER BY T_Users.User;"

  Me.Adodc1.ConnectionString = Con.RetournConnection
'  Me.Adodc1.CommandType = adCmdUnknown
'Adodc1.CursorLocation = adUseClient
   Me.Adodc1.RecordSource = sql
  Adodc1.CursorType = adOpenDynamic
  Set DataGrid1.DataSource = Adodc1
 Adodc1.Refresh
 
 
 
' DataGrid1.Refresh
 
 DataGrid1.Columns(3).DataFormat.Format = "Yes/no"
'  DataGrid1.Columns(2).DataFormat.Type = 5
'  DataGrid1.Columns(2).Button = True
 

' DataGrid1.Columns(2).DataFormat = 2
Fin:
End Sub
