VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form EditGroupe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion des Groupes."
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   Icon            =   "EdiGroupe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2280
      Top             =   4800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Droits"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Double click pour Enregistrer."
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   7
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
      Left            =   0
      Top             =   4080
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
      EOFAction       =   0
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
Attribute VB_Name = "EditGroupe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Groupe As Long
Dim MyValues
Dim IdDelet As Boolean
Dim IsError As Boolean
Dim Row As Long
Private Sub DBGrid1_Click()

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
DroitsGroupe.chargement Groupe
Unload DroitsGroupe
End Sub

Private Sub DataGrid1_AfterInsert()
On Error Resume Next

Dim aa
aa = Groupe
aa = Val("" & Me.DataGrid1.Columns(1).Value)
If Err = 0 Then
   If Val("" & aa) <= Groupe Then
        MsgBox "Vous ne pouvez  saisir un niveau de groupe inférieur ou égal au votre." & vbCrLf & "Votre niveau est sécurité est : " & Groupe, vbCritical
      
        Me.DataGrid1.Columns(1).Value = Groupe + 1
End If
End If
On Error GoTo 0

End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
IdDelet = True
If MsgBox("Voulez vous vraiment supprimer le Groupe :  " & Me.Adodc1.Recordset.Fields(0), vbYesNo + vbQuestion) = vbNo Then IdDelet = False: Cancel = 1

End Sub

Private Sub DataGrid1_BeforeUpdate(Cancel As Integer)
On Error Resume Next
Dim aa
If IdDelet = False Then
aa = Groupe
aa = Val("" & Me.DataGrid1.Columns(1).Value)
     If Val("" & aa) <= Groupe Then
        MsgBox "Vous ne pouvez  saisir un niveau de groupe inférieur ou égal au votre." & vbCrLf & "Votre niveau est sécurité est : " & Groupe, vbCritical
      
'        Me.DataGrid1.Columns(1).Value = Groupe + 1
       IsError = True
       Row = DataGrid1.Row
       Cancel = 1
     End If
End If
IdDelet = False
On Error GoTo 0
End Sub

Private Sub DataGrid1_DblClick()
On Error Resume Next
Dim aa
If IdDelet = False Then
aa = Groupe
aa = Val("" & Me.DataGrid1.Columns(1).Value)
     If Val("" & aa) <= Groupe Then
        MsgBox "Vous ne pouvez  saisir un niveau de groupe inférieur ou égal au votre." & vbCrLf & "Votre niveau est sécurité est : " & Groupe, vbCritical
      
'        Me.DataGrid1.Columns(1).Value = Groupe + 1
       IsError = True
       Row = DataGrid1.Row
       
     End If
End If
IdDelet = False
On Error GoTo 0
If IsError = False Then
Me.Adodc1.Recordset.Update
'If Err Then MsgBox Err.Description
Me.Adodc1.Recordset.Requery
IsError = False
End If
End Sub

Private Sub Form_Load()
Dim sql As String
Dim MyB As Boolean
Dim a
Dim Rs As ADODB.Recordset
 Set a = DataGrid1.Columns(0)
sql = "SELECT T_Groupe.Groupe, T_Groupe.Niveaux, T_Users.Id "
sql = sql & "FROM T_Users INNER JOIN (T_Groupe INNER JOIN T_Groupe_Users ON T_Groupe.id =  "
sql = sql & "T_Groupe_Users.Id_Groupe) ON T_Users.Id = T_Groupe_Users.Id_Users "
sql = sql & "Where T_Users.Id = " & Id_Users & " "
sql = sql & "ORDER BY T_Groupe.Niveaux;"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = True Then GoTo Fin
Command2.Enabled = True
Groupe = Rs!Niveaux
Set Rs = Con.CloseRecordSet(Rs)
sql = "SELECT T_Groupe.Groupe, T_Groupe.Niveaux FROM T_Groupe WHERE  T_Groupe.Niveaux > " & Groupe & " ORDER BY T_Groupe.Niveaux;"

  Me.Adodc1.ConnectionString = Con.RetournConnection
  Me.Adodc1.RecordSource = sql
 Adodc1.Refresh
 Set DataGrid1.DataSource = Me.Adodc1
 DataGrid1.Rebind
'  Adodc1.Recordset.Requery
'
' DataGrid1.Columns(3).DataFormat.Format = "Yes/no"
'  DataGrid1.Columns(2).DataFormat.Type = 5
'  DataGrid1.Columns(2).Button = True
 
Me.Timer1.Enabled = True
' DataGrid1.Columns(2).DataFormat = 2
Fin:
End Sub


Private Sub Timer1_Timer()
If IsError = True Then
'        Me.DataGrid1.Columns(1).Value = Groupe + 1
DataGrid1.Row = Row
       DataGrid1.Columns(1).Value = Groupe + 1
       Me.Adodc1.Recordset.Update
       IsError = False
     End If
End Sub
