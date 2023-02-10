Attribute VB_Name = "mod_Global"
Option Explicit
Public PathBase As String
Public dirFTP As String
Public DirServerPop As String
Public DirZip As String
Public Const MainTitle = "Server Pop Euxia"
Public Const Msg = "Serveur en attente d'ordre."
Public Sub Main()
If App.PrevInstance Then
       End
    End If

frm_Mct_Serveur_Euxia.Show
End Sub

Public Function EcrirFile(NameFile As String) As Boolean
Dim FileNumber As Integer

On Error GoTo Fin
FileNumber = FreeFile   ' Get unused file
      ' number.
   Open LCase(NameFile) For Output As #FileNumber   ' Create file name.
   Write #FileNumber, NameFile ' Output text.
   Close #FileNumber   ' Close file.
    EcrirFile = True
Fin:
End Function
Public Function InputDir(Fichier As String, Recherche As String) As String
'Permet de rechercher des information dans le fichier Fax.ini
Dim pose As Integer
Dim FileNumber As Integer
Dim MyString As String
FileNumber = FreeFile
On Error Resume Next
If Err = 0 Then
Open Fichier For Input As #FileNumber


Do While Not EOF(FileNumber)   ' Effectue la boucle jusqu'à la fin du fichier.
   Input #FileNumber, MyString  ' Lit les données dans la variables.
   pose = InStr(1, UCase(MyString), UCase(Recherche))
   If pose <> 0 Then
        InputDir = Mid(MyString, Len(Recherche) + 2, Len(MyString) - (Len(Recherche) + 1))
        InputDir = Trim(InputDir)
        
        Close #FileNumber   ' Ferme le fichier.
       
         Exit Function
   End If
Loop
Close #FileNumber   ' Ferme le fichier.

Else
''    boolKill = False
Err.Clear
 End If

End Function
Public Function Unzip(MyFile As String) As String
On Error GoTo reprise_err
Dim CmdLine As String
Dim TempDir As String
Dim MyDir As String
Dim I As Integer

'TempDir = MyFile
'MyDir = my_work_path & "\temp\"

'If Dir(TempDir) <> "" Then Kill TempDir & "*.*"
Dim a
Shell (DirZip & "wzunzip -d " & Chr(34) & MyFile & Chr$(34) & " " & Chr$(34) & PathBase & Chr$(34))
Do While Dir(PathBase & "MAJ MCT_DATA.mdb") = ""
I = I + 1
If I = 32767 Then
    Exit Function
End If
    DoEvents
Loop

'Unzip = Dir(TempDir)
Exit Function
reprise_err:
    
'    LogEvent "Erreur dans la fonction unzip " & Err.Description & " " & Err.Number
MsgBox Err.Description
    Err.Clear
    
    Exit Function
End Function

Public Function zip(MyFile As String) As String
On Error GoTo reprise_err
Dim CmdLine As String
Dim TempDir As String
Dim MyDir As String
Dim I As Integer

'TempDir = MyFile
'MyDir = my_work_path & "\temp\"

'If Dir(TempDir) <> "" Then Kill TempDir & "*.*"
Dim a
Shell (DirZip & "WZZIP.EXE  " & dirFTP & Chr(34) & "MAJ MCT_DATA" & Chr(34) & " " & PathBase & Chr(34) & "MAJ MCT_DATA.mdb" & Chr(34))

For I = 1 To 30
        If File_Exists(dirFTP & "MAJ MCT_DATA.zip") Then
            Exit For
        End If
        AttendreSecondes 1
    Next I

'Unzip = Dir(TempDir)
Exit Function
reprise_err:
    
'    LogEvent "Erreur dans la fonction unzip " & Err.Description & " " & Err.Number
MsgBox Err.Description
    Err.Clear
    
    Exit Function
End Function
Public Sub UnloadeMain()
Dim txtCompact As String

DoEvents




    txtCompact = PathBase & "\MCT_DATA.mdb"
    frm_Mct_Serveur_Euxia.LblVersion.Caption = "Compactage de: MCT_DATA.mdb en cours..."
    DoEvents
    txtCompact = Left(txtCompact, Len(txtCompact) - 4)
    subCompact txtCompact
    
    txtCompact = PathBase & "\MCT_DATA_Commune.mdb"
    frm_Mct_Serveur_Euxia.LblVersion.Caption = "Compactage de: MCT_DATA_Commune.mdb en cours..."
    DoEvents
    txtCompact = Left(txtCompact, Len(txtCompact) - 4)
    subCompact txtCompact
    
'    txtCompact = dbDirectory & "\MCT_IHM_SERVEUR.mdb"
'    frm_Sortie.LblVersion.Caption = "Compactage de: MCT_IHM_SERVEUR.mdb en cours..."
'    DoEvents
'    txtCompact = Left(txtCompact, Len(txtCompact) - 4)
'    subCompact txtCompact

    txtCompact = PathBase & "\MCT_IHM.mdb"
    frm_Mct_Serveur_Euxia.LblVersion.Caption = "Compactage de: MCT_IHM.mdb en cours..."
    DoEvents
    txtCompact = Left(txtCompact, Len(txtCompact) - 4)
    subCompact txtCompact
    
     txtCompact = PathBase & "\MCT_in.mdb"
    frm_Mct_Serveur_Euxia.LblVersion.Caption = "Compactage de: MCT_in.mdb en cours..."
    DoEvents
    txtCompact = Left(txtCompact, Len(txtCompact) - 4)
    subCompact txtCompact
' frm_Sortie.LblVersion.FontSize = Police
frm_Mct_Serveur_Euxia.LblVersion.Caption = Msg
End Sub
Public Sub subCompact(dbName As String)
Dim dbBack As String, dbOld As String, dbRepair As String
Dim Rs As Recordset
Dim Email As String
Dim strSubject
Dim strMsgId As String


On Error GoTo Fin
 

dbBack = dbName & "_Bak.mdb"
'dbOld = dbName & "_Bak.mdb"
dbRepair = dbName & "_Save.mdb"
Screen.MousePointer = vbHourglass
''''doevents
File_Delete dbBack
 File_Delete dbRepair
 FileCopy dbName & ".mdb", dbRepair
DBEngine.CompactDatabase dbName, dbBack
    
    File_Delete dbName & ".mdb"
    FileCopy dbBack, dbName & ".mdb"
File_Delete dbBack
Screen.MousePointer = vbDefault
'Pop.sBar.Caption = "Fichier " & dbName & " sauvegardé"

Fin:


End Sub
Public Sub File_Delete(FileSpec As String, Optional ALL As Boolean)
Dim fso As New FileSystemObject
On Error GoTo Err_DelFile
If ALL Then
    fso.DeleteFile FileSpec
Else
    If fso.FileExists(FileSpec) Then
        fso.DeleteFile FileSpec
    End If
End If
Exit Sub
Err_DelFile:
MsgBox Error(Err), vbCritical, "File_Delete: " & FileSpec
End Sub
Public Sub AttendreSecondes(nSecondes As Integer)
' Attend au moins le nombre de secondes prescrit
    Dim date1 As Date
    date1 = Now
    Do
        DoEvents
    Loop While DateDiff("s", date1, Now) < nSecondes + 1
End Sub

Public Function File_Exists(FileSpec As String) As Boolean
Dim fso As New FileSystemObject
On Error GoTo Err_FExist
File_Exists = fso.FileExists(FileSpec)
Exit Function
Err_FExist:
MsgBox Error(Err), vbCritical, "File_Exists: " & FileSpec
End Function
