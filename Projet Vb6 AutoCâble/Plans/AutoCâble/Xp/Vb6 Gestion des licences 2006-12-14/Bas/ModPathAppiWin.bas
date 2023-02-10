Attribute VB_Name = "ModPathAppliWin"
 Private result As String
 Public Function SerchFile(DefaultRep As String, MyFichier As String, obj) As String
 Dim retourFile As String
Dim lecteur As Drive
Dim dossier As Folder
Dim Fichier As File
Dim sousdossier As Folder
Dim txtPath As String
Dim SpitPath
Dim UboundIndex As Long
Dim I As Long
 Dim Fso As New FileSystemObject
 txtPath = DefaultRep
  SpitPath = Split(txtPath, "\")
    UboundIndex = UBound(SpitPath) + 1
    If InStr(1, UCase(DefaultRep), UCase(MyFichier)) = 0 Then GoTo Reprise
If Fso.FileExists(txtPath) = False Then
   
Reprise:
    UboundIndex = UboundIndex - 1
    txtPath = ""
    For I = 0 To UboundIndex
        txtPath = txtPath & SpitPath(I) & "\"
    Next
    If Fso.FolderExists(txtPath) = False Then GoTo Reprise
    txtPath = Left(txtPath, Len(txtPath) - 1)
'For Each lecteur In fso.Drives
'     If lecteur.IsReady Then

     Set dossier = Fso.GetFolder(txtPath)
      retourFile = scanDosier(dossier, MyFichier, obj)
      DoEvents
      If Trim("" & retourFile) = "" Then
             For Each sousdossier In dossier.SubFolders
             retourFile = scan(CStr(sousdossier), MyFichier, Fso, obj)
             If Trim("" & retourFile) <> "" Then Exit For
             Next sousdossier
     End If
'     End If
'Next lecteur
Else
    retourFile = DefaultRep
End If
If Trim("" & retourFile) = "" And UboundIndex <> 0 Then GoTo Reprise
If Fso.FileExists(Trim("" & retourFile)) = False And UboundIndex <> 0 Then GoTo Reprise

obj.Caption = retourFile
 SerchFile = retourFile
 Set Fso = Nothing
  If InStr(1, UCase(DefaultRep), UCase(MyFichier)) = 0 Then SerchFile = "ERR"
  
End Function
  
Private Function scan(dd As String, MyFichier As String, Fso, obj) As String
Dim dossier As Folder
Dim sousdossier As Folder
Dim Fichier As File
Set dossier = Fso.GetFolder(dd)
On Error Resume Next
     If dossier.SubFolders.Count <> 0 Then
      For Each sousdossier In dossier.SubFolders
        DoEvents
       scan = scan(CStr(sousdossier), MyFichier, Fso, obj)
       If Trim("" & scan) <> "" Then Exit For
      Next sousdossier
     End If
    For Each Fichier In dossier.Files
    obj.Caption = Fichier.Path
    DoEvents
    Debug.Print Fichier.Path
    'à la place de "dbx" mettez l'extension souhaitée ex : ".txt"
        If InStr(1, UCase(Fichier.Path), UCase(MyFichier)) <> 0 Then
        scan = Fichier.Path
       
        DoEvents
        Exit For
        
        End If
    Next Fichier
    DoEvents
  
End Function

Private Function scanDosier(dossier, MyFichier As String, obj) As String
Dim sousdossier As Folder
Dim Fichier As File
On Error Resume Next
     
    For Each Fichier In dossier.Files
    obj.Caption = Fichier.Path
    DoEvents
    Debug.Print Fichier.Path
    'à la place de "dbx" mettez l'extension souhaitée ex : ".txt"
        If InStr(1, UCase(Fichier.Path), UCase(MyFichier)) <> 0 Then
        scanDosier = Fichier.Path
       
        DoEvents
        Exit For
        
        End If
    Next Fichier
    DoEvents
  
End Function
Public Sub loadFichierAppliWin()
 Dim sql As String
 Dim Rs As Recordset
Dim Fso As New FileSystemObject

Dim FRM As Form
    PathAppliAutocad = CherCheInFihier("AutocadExe")
    PathAppliExcel = CherCheInFihier("Excelexe")

sql = "SELECT MachinAplicationChemen.* "
sql = sql & "FROM MachinAplicationChemen  "
sql = sql & "WHERE MachinAplicationChemen.Machine='" & Machine & "';"
 Set Rs = Con.OpenRecordSet(sql)
 If Rs.EOF = True Then
    Rs.AddNew
    Rs!Machine = Machine
    Rs!EXCEL = PathAppliExcel
     Rs!AutoCAD = PathAppliAutocad
     Rs.Update
 End If
Rs.Requery

PathAppliExcel = "" & Rs!EXCEL
PathAppliAutocad = "" & Rs!AutoCAD
If PathAppliExcel <> "ERR" Then
    If InStr(1, UCase(PathAppliExcel), "EXCEL.EXE") = 0 Then
        If Fso.FileExists("" & Rs!EXCEL) = False Then
            If MsgBox("Autocâble doit rechercher l'emplacement d'EXCEL sur votre disque local." & vbCrLf & _
                "Volez-vous effectuer cette opération.", vbYesNo + vbQuestion) = vbNo Then
                    PathAppliExcel = "ERR"
                Else
ReprisExcel:
                Set FRM = New FrmPathAppliWin
                FRM.Show
                FRM.Visible = True
                FRM.Caption = "Recherer le chemin d'Excel"
                PathAppliExcel = SerchFile("" & Rs!EXCEL, "Excel.exe", FRM.Label1)
                FRM.Hide
                Unload FRM
            End If
        End If
    End If
End If


If PathAppliAutocad <> "ERR" Then
    If InStr(1, UCase(PathAppliAutocad), "ACAD.EXE") = 0 Then
        If Fso.FileExists("" & Rs!AutoCAD) = False Then
            If MsgBox("Autocâble doit rechercher l'emplacement d'Autocad  sur votre disque local." & vbCrLf & _
            "Volez-vous effectuer cette opération.", vbYesNo + vbQuestion) = vbNo Then
                PathAppliAutocad = "ERR"
            Else
ReprisAcad:
            Set FRM = New FrmPathAppliWin
            FRM.Show
                FRM.Visible = True
                FRM.Caption = "Recherer le chemin d'Autocad"
                PathAppliAutocad = SerchFile("" & Rs!AutoCAD, "acad.exe", FRM.Label1)
                FRM.Hide
                Unload FRM
            End If
         End If
    End If
End If
Rs!Machine = Machine
Rs!EXCEL = PathAppliExcel
Rs!AutoCAD = PathAppliAutocad
Rs.Update

Set Rs = Con.CloseRecordSet(Rs)
End Sub
