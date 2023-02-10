VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Private m_file As String
Private m_taille As Long
Private m_date As Boolean
Private m_time As Boolean
Private m_ptrfile As Integer
Private m_DateDebut As String
Private m_Block As Boolean
Private m_Buffer() As String
Private m_pos_log As Byte
Private m_Format_Date As String
Private m_Format_Heure As String
Private m_date_log As String * 8
Private m_entete As String

Public Enable As Boolean

Private Function init() As Boolean
On Error Resume Next
    If (Not Enable) Then
        Exit Function
    End If
    Dim Tmp As String
    Dim pos As Long
    Dim i As Integer
    i = 1
    
    m_Block = True
    If (m_ptrfile <> 0) Then
        Close #m_ptrfile
    End If
    
    m_ptrfile = FreeFile()
    
    pos = InStrRev(m_file, ".", , vbTextCompare)
    
    m_DateDebut = Format(Now(), "yymmdd_hhmm")
    
    If (pos > 0) Then
        Tmp = Mid$(m_file, 1, pos - 1) & m_DateDebut & Mid$(m_file, pos)
    Else
        Tmp = m_file & m_DateDebut
    End If
    
    m_date_log = Format$(Now(), "ddmmyyyy")
    Open Tmp For Append As #m_ptrfile
    If (m_taille <> 0) Then
        If CInt(Size) > m_taille Then
            init
        End If
    End If
    If (Len(m_entete) > 0) Then
        Tmp = m_entete
        Tmp = Replace(Tmp, "%DATE%", Format$(Now(), "dd/mm/yyyy"))
        Tmp = Replace(Tmp, "%HEURE%", Format$(Now(), "hh:mm:ss"))
        Tmp = Replace(Tmp, "%VERSION%", App.Major & "." & App.Minor)
        Tmp = Replace(Tmp, "%BUILD%", App.Revision)
        Tmp = Replace(Tmp, "%NUMFICH%", CStr(i))
        Tmp = Replace(Tmp, "%APP%", App.EXEName)
        If (FileExist(App.path & IIf(Right$(App.path, 1) <> "\", "\", "") & App.EXEName & ".exe")) Then
            Tmp = Replace(Tmp, "%DATEMODIFAPP%", Format$(FileDateTime(App.path & IIf(Right$(App.path, 1) <> "\", "\", "") & App.EXEName & ".exe"), "dd/mm/yyyy hh:mm"))
        Else
            Tmp = Replace(Tmp, "%DATEMODIFAPP%", Format$(Now(), "dd/mm/yyyy hh:mm"))
        End If
        Tmp = Replace(Tmp, "%DIRCUR%", CurDir$())
        Tmp = Replace(Tmp, "%DIRAPP%", App.path)
        Print #m_ptrfile, Tmp
    End If
    For pos = (m_pos_log + 1) To UBound(m_Buffer)
        If (Len(m_Buffer(pos)) > 0) Then Print #m_ptrfile, m_Buffer(pos)
    Next
    For pos = 0 To m_pos_log
        If (Len(m_Buffer(pos)) > 0) Then Print #m_ptrfile, m_Buffer(pos)
    Next
    m_Block = False
    
End Function

Public Property Let FileName(path As String)
On Error Resume Next
    Dim pos As Long
    
    If (Not Enable) Then
        Call Err.Raise(vbObjectError + 1, "ClsLog::FileName", "Log désactivé")
        Exit Property
    End If
    
    If ((Left$(path, 2) <> "\\") And (Mid$(path, 2, 2) <> ":\")) Then
        m_file = CurDir$ & "\" & path
    Else
        m_file = path
    End If
    pos = InStrRev(m_file, "\", , vbTextCompare)
    If (MakeSureDirectoryPathExists(Left$(m_file, pos)) = 0) Then
        Err.Raise 1, "Class Log", "Impossible de créer l'arborescence de Log"
    Else
        init
    End If
End Property

Public Property Get FileName() As String
    FileName = m_file
End Property

Public Property Let MaxSize(taille As Long)
On Error Resume Next
    If (((taille < 100) Or (taille > 5 * 1024)) And (taille <> 0)) Then
        Err.Raise 2, "Class Log", "Taille spécifiée incorrecte !"
    Else
        m_taille = taille
    End If
End Property

Public Property Get MaxSize() As Long
    MaxSize = m_taille
End Property

Public Property Let LogTime(Tmp As Boolean)
    m_time = Tmp
End Property

Public Property Get LogTime() As Boolean
    LogTime = m_time
End Property

Public Property Let LogDate(Tmp As Boolean)
    m_date = Tmp
End Property

Public Property Get LogDate() As Boolean
    LogDate = m_date
End Property

Public Property Get Size() As Double
    Size = CDbl(LOF(m_ptrfile) / 1024)
End Property

Public Sub Ecrire(texte As String)
    Dim Tmp As String
    
    If (Not Enable) Then
        Exit Sub
    End If
    
    If (Len(Trim$(texte)) > 0) Then
        Tmp = IIf(LogDate, Format(Now(), m_Format_Date) & " ", "")
        Tmp = Tmp & IIf(LogTime, Format(Now(), m_Format_Heure) & " ", "")
        Tmp = Tmp & texte
        If (Not m_Block) Then
            If (m_date_log <> Format$(Now(), "ddmmyyyy")) Then
                init
            End If
            Print #m_ptrfile, Tmp
        Else
            m_pos_log = m_pos_log + 1
            If (m_pos_log > 250) Then m_pos_log = 0
            m_Buffer(m_pos_log) = Tmp
        End If
        Debug.Print Tmp
    End If
    If (m_taille <> 0) Then
        If CInt(Size) > m_taille Then
            init
        End If
    End If
End Sub

Private Sub Class_Initialize()
    m_taille = 1024
    m_date = True
    m_time = True
    m_Block = True
    ReDim m_Buffer(0 To 250)
    Enable = False
    m_file = vbNullString
    m_pos_log = 250
    m_Format_Date = "dd-mm-yyyy"
    m_Format_Heure = "hh:mm:ss"
End Sub

Public Property Let Format_Heure(frt As String)
    m_Format_Heure = frt
End Property

Public Property Get Format_Heure() As String
    Format_Heure = m_Format_Heure
End Property

Public Property Let Format_Date(frt As String)
    m_Format_Date = frt
End Property

Public Property Get Format_Date() As String
    Format_Date = m_Format_Date
End Property

Public Property Let Entete(str As String)
    m_entete = str
End Property

Private Function FileExist(ByVal Nom As String) As Boolean
On Error Resume Next
    Dim attrib As Integer
    
    attrib = GetAttr(Nom)
    If (Err <> 0) Then
        FileExist = False
    Else
        If ((attrib And vbDirectory) = vbDirectory) Then
            FileExist = False
        Else
            FileExist = True
        End If
    End If
End Function

