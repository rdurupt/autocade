Attribute VB_Name = "ModulePricipal"
Option Explicit
Public PathAppliAutocad As String
Public PathAppliExcel As String
Type MyLicGene
    Societe As String
    Tous As String
    AficheFrm As String
    DateDeb As String
    DateExecuter As String
    DateFin As String
    Enregistre As String
    NbJeton As String
    NbJetonActif As String
End Type

Type MyLic
    Serial As String
    PassWord As String
    Useur As String
    Enregistre As String
End Type
Type Licence
    Count As Long
    General As MyLicGene
    Record() As MyLic
End Type
Public FiledLicence As Licence
Type T_Job
    AppActivate As Object
    Job As Long
End Type
Public Con As New Ado
Public CodageX As CDETXT
Public IsServeur As Boolean
Public Msg As String
Public Db As String
Public TimerInerval As String
Public KillJob As String
Public TableauAtotocable(255) As T_Job
Public AppOff As Boolean
Public ServiceName As String
Public Const MainTitle = "AutoCâble Serveur"
Public MyTimer As String
Public Trace As ClsLog
Public POpInter As String
Sub Main()
If App.PrevInstance = True Then

        MsgBox "Une instance du programme à déjà été lancer, impossible de lancer une nouvelle instance.", vbOKOnly + vbCritical, "Autocâble"
        ' Ferme le programme
        End
 End If
 Machine = LirMachineName
 frmPricipal.Show
End Sub

Sub LoadDb()

POpInter = CherCheInFihier("POpInter")
Db = CherCheInFihier("BdAutocable")
TimerInerval = CherCheInFihier("TimerInerval")
KillJob = CherCheInFihier("KillJob")
ServiceName = CherCheInFihier("Service")
End Sub
Function CherCheInFihier(Cherher As String) As String
Dim FileNumber As Long
Dim MyString As String
FileNumber = FreeFile

  
Open App.path & "\Autocable.ini" For Input As #FileNumber
Do While Not EOF(FileNumber)    ' Effectue la boucle jusqu'à la fin du fichier.
    Input #FileNumber, MyString ' Lit les données dans deux variables.
    If InStr(1, MyString, Cherher) <> 0 Then
       CherCheInFihier = Right(MyString, Len(MyString) - (Len(Cherher) + InStr(1, MyString, Cherher)))
       CherCheInFihier = Trim(Replace(CherCheInFihier, "=", ""))
       Exit Do
    End If
Loop
Close #FileNumber    ' Ferme le fichier
CherCheInFihier = Replace(CherCheInFihier, "§remplaceDate§", Format(Date, "yyyy"))
CherCheInFihier = Trim(CherCheInFihier)
End Function

Public Function MySeconde(NuSeconde As Integer)
Dim a
 a = Second(Time)
    While Abs(a - Second(Time)) < NuSeconde
    DoEvents
    Wend
End Function
Public Sub loadFichierAppliWin()
 Dim Sql As String
 Dim rs As Recordset
Dim Fso As New FileSystemObject

Dim FRM As Form

    PathAppliAutocad = CherCheInFihier("AutocadExe")
    PathAppliExcel = CherCheInFihier("Excelexe")

Sql = "SELECT MachinAplicationChemen.* "
Sql = Sql & "FROM MachinAplicationChemen  "
Sql = Sql & "WHERE MachinAplicationChemen.Machine='" & Machine & "';"
 Set rs = Con.OpenRecordSet(Sql)
 If rs.EOF = True Then
    rs.AddNew
    rs!Machine = Machine
    rs!EXCEL = PathAppliExcel
     rs!AutoCAD = PathAppliAutocad
     rs.Update
 End If
rs.Requery

PathAppliExcel = "" & rs!EXCEL
PathAppliAutocad = "" & rs!AutoCAD
If PathAppliExcel <> "ERR" Then
    If InStr(1, UCase(PathAppliExcel), "EXCEL.EXE") = 0 Then
        If Fso.FileExists("" & rs!EXCEL) = False Then
            If MsgBox("Autocâble doit rechercher l'emplacement d'EXCEL sur votre disque local." & vbCrLf & _
                "Volez-vous effectuer cette opération.", vbYesNo + vbQuestion) = vbNo Then
                    PathAppliExcel = "ERR"
                Else
ReprisExcel:
                Set FRM = New FrmPathAppliWin
                FRM.Show
                FRM.Visible = True
                FRM.Caption = "Recherer le chemin d'Excel"
                PathAppliExcel = SerchFile("" & rs("EXCEL"), "Excel.exe", FRM.Label1)
                FRM.Hide
                Unload FRM
            End If
        End If
    End If
End If


If PathAppliAutocad <> "ERR" Then
    If InStr(1, UCase(PathAppliAutocad), "ACAD.EXE") = 0 Then
        If Fso.FileExists("" & rs!AutoCAD) = False Then
            If MsgBox("Autocâble doit rechercher l'emplacement d'Autocad  sur votre disque local." & vbCrLf & _
            "Volez-vous effectuer cette opération.", vbYesNo + vbQuestion) = vbNo Then
                PathAppliAutocad = "ERR"
            Else
ReprisAcad:
            Set FRM = New FrmPathAppliWin
            FRM.Show
                FRM.Visible = True
                FRM.Caption = "Recherer le chemin d'Autocad"
                PathAppliAutocad = SerchFile("" & rs!AutoCAD, "acad.exe", FRM.Label1)
                FRM.Hide
                Unload FRM
            End If
         End If
    End If
End If
rs!Machine = Machine
rs!EXCEL = PathAppliExcel
rs!AutoCAD = PathAppliAutocad
rs.Update

Set rs = Con.CloseRecordSet(rs)
End Sub
Public Function SerchFile(DefaultRep As String, MyFichier As String, obj) As String
 Dim retourFile As String
Dim lecteur As Drive
Dim dossier As Folder
Dim Fichier As File
Dim sousdossier As Folder
Dim txtPath As String
Dim SpitPath
Dim UboundIndex As Long
Dim i As Long
 Dim Fso As New FileSystemObject
 txtPath = DefaultRep
  SpitPath = Split(txtPath, "\")
    UboundIndex = UBound(SpitPath) + 1
    If InStr(1, UCase(DefaultRep), UCase(MyFichier)) = 0 Then GoTo Reprise
If Fso.FileExists(txtPath) = False Then
   
Reprise:
    UboundIndex = UboundIndex - 1
    txtPath = ""
    For i = 0 To UboundIndex
        txtPath = txtPath & SpitPath(i) & "\"
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
Private Function scanDosier(dossier, MyFichier As String, obj) As String
Dim sousdossier As Folder
Dim Fichier As File
On Error Resume Next
     
    For Each Fichier In dossier.Files
    obj.Caption = Fichier.path
    DoEvents
    Debug.Print Fichier.path
    'à la place de "dbx" mettez l'extension souhaitée ex : ".txt"
        If InStr(1, UCase(Fichier.path), UCase(MyFichier)) <> 0 Then
        scanDosier = Fichier.path
       
        DoEvents
        Exit For
        
        End If
    Next Fichier
    DoEvents
  
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
    obj.Caption = Fichier.path
    DoEvents
    Debug.Print Fichier.path
    'à la place de "dbx" mettez l'extension souhaitée ex : ".txt"
        If InStr(1, UCase(Fichier.path), UCase(MyFichier)) <> 0 Then
        scan = Fichier.path
       
        DoEvents
        Exit For
        
        End If
    Next Fichier
    DoEvents
  
End Function

Public Sub Example_AutoAudit()
    ' This example returns the current setting of
    ' AutoAudit. It then changes the value, and finally
    ' it resets the value back to the original setting.
    
    Dim preferences As Object
    Dim currAutoAudit As Boolean
    Dim newAutoAudit As Boolean
    
    Set preferences = CreateObject("AutoApp.preferences")
    
    ' Retrieve the current AutoAudit value
'    currAutoAudit = preferences.OpenSave.AutoAudit
'    MsgBox "The current value for AutoAudit is " & currAutoAudit, vbInformation, "AutoAudit Example"
'
    ' Toggle the value for AutoAudit
    newAutoAudit = Not (currAutoAudit)
    preferences.OpenSave.AutoAudit = newAutoAudit
'    MsgBox "The new value for AutoAudit is " & newAutoAudit, vbInformation, "AutoAudit Example"
'
'    ' Reset AutoAudit to its original value
'    preferences.OpenSave.AutoAudit = newAutoAudit
'    MsgBox "The AutoAudit value is reset to " & currAutoAudit, vbInformation, "AutoAudit Example"
End Sub
