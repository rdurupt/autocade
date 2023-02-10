Attribute VB_Name = "FonctionsGlobales"
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Dans votre code, ici un click sur un label (les labels peuvent être transparents et donc se placer sur un BMP) :

Public Sub MyExecute(Fichier As String)
Dim lapi As Long
On Error Resume Next
lapi = ShellExecute(100, "open", Fichier, vbNull, vbNull, 5)
If Err Then MsgBox Err.Description
Err.Clear
On Error GoTo 0
End Sub


Public Function LoadCalque()
   AutoApp.Documents(0).ActiveLayer = AutoApp.Documents(0).Layers(0)
   AutoApp.Documents(0).ActiveLayer.ViewportDefault = True
  AutoApp.Documents(0).PurgeAll
End Function

Public Function ValideChampsTexte(Formulaire, NbChamps As Long) As Boolean
   
    ValideChampsTexte = False
    For i = 1 To NbChamps
        If MyFormatQRY(Formulaire.Controls("txt" & CStr(i))) = False Then Exit Function
        DoEvents
       
    Next i
    ValideChampsTexte = True
    End Function
 Public Function MyFormatQRY(txt As Object) As Boolean
 Dim MyTag
 MyFormatQRY = False
 MyTag = Split(txt.Tag, ";")
    
        If Trim("" & txt) = "" Then
            If UCase(Trim(MyTag(2))) = "QRY" Then
                MsgBox "Valeur de : " & MyTag(1) & " obligatoire", vbExclamation
                txt.SetFocus
                Exit Function
            End If
        Else
            If MyFormat("" & MyTag(3), txt, "" & MyTag(1)) = False Then
                txt.SetFocus
                Exit Function
            End If
           
    
        End If
        MyFormatQRY = True
 End Function
 
 
 
 Public Function MyFormat(MyType As String, MyText As Object, MyLib As String) As Boolean
 MyFormat = True
 If MyText = "" Then Exit Function
 
  Select Case UCase(MyType)
                    Case "DATE"
                        If Not IsDate(MyText) Then
                            MsgBox "Vous devez saisir une date pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            
                            Exit Function
                        Else
                            MyText = Format(MyText, "dd/mm/yyyy")
                        End If
                    Case "ENT"
                        If Not IsNumeric(MyText) Then
                            MsgBox "Vous devez saisir un nombre entier pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            Exit Function
                        Else
                            If (InStr(1, (MyText), ",") <> 0) Or (InStr(1, (MyText), ".") <> 0) Then
                                MsgBox "Vous devez saisir un nombre entier pour : " & MyLib, vbExclamation
                                MyText = ""
                                MyFormat = False
                                Exit Function
                            End If
                        End If
                    Case "DBL"
                        If Not IsNumeric(MyText) Then
                            MyText = Replace(MyText, ".", ",")
                        End If
                        If Not IsNumeric(MyText) Then
                            MsgBox "Vous devez saisir un nombre à virgule pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            Exit Function
                        End If
            End Select
 End Function
 
Function DecalInsertPointLigneTableau_fils_Bas(Y, Ofset)
    DecalInsertPointLigneTableau_fils_Bas = Y - Ofset
End Function
Function DecalInsertPointLigneTableau_fils_Gauche(X, Ofset)
    DecalInsertPointLigneTableau_fils_Gauche = X + Ofset
End Function
Public Function funAttributesLigne_Tableau_fils(MyName As String, Attributes, Tableau, Nb, Optional RangeTitre As Recordset, Optional BoolTirte As Boolean, Optional MyColection As Collection, Optional vignette As Boolean, Optional EPISSURE As Boolean)
Dim DESIGNATION As String
Dim MyNb As Long
Dim MyNbStart As Long
Dim msgAttib As String
If vignette = True Then
    If EPISSURE = False Then
        DESIGNATION = ".HAUT"
        MyNb = 4
        MyNbStart = 2
    Else
        MyNbStart = 3
         DESIGNATION = ""
      MyNb = 3
    End If
Else
    DESIGNATION = ""
      MyNb = Nb
      If BoolTirte = True Then
        MyNbStart = 2
      Else
      MyNbStart = 0
      End If
End If
On Error GoTo MsgError

For i = MyNbStart To MyNb
DoEvents
    If BoolTirte = False Then
        Attributes(i).TextString = "" & Tableau(i)
    Else
        msgAttib = Tableau(i)
       If EPISSURE = False Then
            Attributes(MyColection.Item(Replace(RangeTitre.Fields(i).Name, "PRECO", "PRECO.") & DESIGNATION)).TextString = Trim("" & Tableau(i))
        Else
            Attributes(MyColection.Item("EPISSURE")).TextString = Trim("" & Tableau(i))

        End If
            DESIGNATION = ""
    End If
Next i
Exit Function
MsgError:
FunError 2, RangeTitre.Fields(i).Name, MyName
Resume Next
End Function

Public Function FunInsBlock(PathName, InsertPoint, Name, Optional Rotation As Double, Optional XScaleFactor As Double, Optional YScaleFactor As Double, Optional ZScaleFactor As Double) As AcadBlockReference
On Error GoTo GesERR
Dim DDDD As AutoCAD.AcadApplication

ErrInsert = False
  
    Set FunInsBlock = AutoApp.Documents(0).ModelSpace.InsertBlock(InsertPoint, PathName, 1#, 1#, 1#, 0#)
   Dim layerObj As AcadLayer
   a = FunInsBlock.Name
  e = FunInsBlock.GetAttributes
  FunInsBlock.Rotation = Rotation
  If XScaleFactor = 0 Then XScaleFactor = 1
   If YScaleFactor = 0 Then YScaleFactor = 1
   If ZScaleFactor = 0 Then ZScaleFactor = 1
  FunInsBlock.XScaleFactor = XScaleFactor
   FunInsBlock.YScaleFactor = YScaleFactor
   FunInsBlock.ZScaleFactor = ZScaleFactor
    AutoApp.Documents(0).Application.ZoomAll
'    AutoApp
    Exit Function
GesERR:
    FunError 100, "*", PathName & vbCrLf & Err.Description
    ErrInsert = True
   Resume Next
End Function

Public Function FunInsBlock2(PathName, InsertPoint, Name, Optional Rotation As Double, Optional XScaleFactor As Double, Optional YScaleFactor As Double, Optional ZScaleFactor As Double) As AcadBlockReference
On Error GoTo GesERR
ErrInsert = False
    Set FunInsBlock2 = AutoApp.Documents(0).ModelSpace.InsertBlock(InsertPoint, PathName, 1#, 1#, 1#, 0#)
   Dim layerObj As AcadLayer
   a = FunInsBlock2.Name
  e = FunInsBlock2.GetAttributes
  FunInsBlock2.Rotation = Rotation
  If XScaleFactor = 0 Then XScaleFactor = 1
   If YScaleFactor = 0 Then YScaleFactor = 1
   If ZScaleFactor = 0 Then ZScaleFactor = 1
  FunInsBlock2.XScaleFactor = XScaleFactor
   FunInsBlock2.YScaleFactor = YScaleFactor
   FunInsBlock2.ZScaleFactor = ZScaleFactor

   ' AutoApp.Documents(0).ZoomAll
    
    Exit Function
GesERR:
' Msg = Msg & "***************************************************" & vbCrLf
'           Msg = Msg & PathName & vbCrLf
'            Msg = Msg & Err.Description & vbCrLf
'            Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
ErrInsert = True
End Function
Function DirNUMEROFIL(Path As String, Index As Long) As String
    Dim MyDir As String
    Dim IndexDir As Long
    Dim Fso As New FileSystemObject
    DirNUMEROFIL = ""
    MyDir = ""
     
    For IndexDir = 0 To 3
        If Fso.FileExists(Path & "NUMEROFIL" & CStr(IndexDir + Index) & ".dwg") = True Then
            MyDir = Path & "NUMEROFIL" & CStr(IndexDir + Index) & ".dwg"
            Exit For
        End If
               
    Next IndexDir
    If Trim(MyDir) = "" Then Exit Function
    DirNUMEROFIL = MyDir
    Set Fso = Nothing
End Function
Public Function FunError(NumErr As Long, Lib1 As String, Msg As String, Optional Lib2 As String)
Dim Sql As String
If Trim("" & Lib1) = "" Then Exit Function
If JobError = 0 Then JobError = AtrbNumError
Msg = MsgErreur(NumErr, Lib1, Lib2, Msg)
Sql = "INSERT INTO T_Error ( JobError, ValError ) "
Sql = Sql & "values(" & JobError & ",'" & Msg & "' );"
Con.Exequte Sql

End Function
Public Function AfficheErreur(Path As String, Entete)

    Dim NuFichier As Long
    Dim Text
    Dim MyTxtErr
    Dim Sql As String
    Dim RsErreur As Recordset
    Dim Fichier As String
    NuFichier = FreeFile
    
    Text = ""
    Sql = "SELECT T_Error.ValError FROM T_Error "
    Sql = Sql & "WHERE T_Error.JobError=" & JobError & ";"
   
    Set RsErreur = Con.OpenRecordSet(Sql)
    While RsErreur.EOF = False
        Text = Text & RsErreur!ValError & vbCrLf
        RsErreur.MoveNext
    Wend
    Set RsErreur = Con.CloseRecordSet(RsErreur)
    If Trim("" & Text) = "" Then Exit Function
    
    Dim Fso As New FileSystemObject

    MyTxtErr = Entete & Text
    pathUser = Environ("USERPROFILE")
    pathUser = pathUser + "\Mes Documents"
    
    If Fso.FolderExists(Path & "RepErrorLog") = False Then
        Fso.CreateFolder Path & "RepErrorLog"
    End If
    Fichier = Path & "RepErrorLog\Error_" & Format(Now, "yyyy-mm-dd_hh_mm_ss") & ".log"
    
    While Fso.FileExists(Fichier) = True
        Fichier = Path & "RepErrorLog\Error_" & Format(Now, "yyyy-mm-dd_hh_mm_ss") & ".log"
    
    Wend
    Open Fichier For Output As #NuFichier
    Print #NuFichier, MyTxtErr
    Close #NuFichier
    Sql = "DELETE T_Error.* FROM T_Error "
    Sql = Sql & "WHERE T_Error.JobError=" & JobError & ";"
    Con.Exequte Sql
    Set Fso = Nothing
    Shell "notepad.exe " & Fichier, vbMaximizedFocus
End Function
 
Public Function MyReplace(strVal As String) As String
MyReplace = strVal
MyReplace = Replace(MyReplace, "'", "''")
MyReplace = Replace(MyReplace, Chr(34), Chr(34) & Chr(34))
End Function
Public Function MyReplaceDate(strVal As String) As String
    If Trim(strVal) = "" Then
        MyReplaceDate = "NULL"
    Else
        MyReplaceDate = "#" & strVal & "#"
    End If
    
End Function

Public Function OpenFichier(Fichier)

On Error Resume Next

    Dim MyDocument As New AutoCAD.AcadDocument
    Set MyDocument = AutoApp.Documents.Open(Fichier)
   MyDocument.Activate
   Set OpenFichier = MyDocument
'   AutoCAD.Visible = False
End Function
Public Function OpenNew()

    Dim MyDocument As AutoCAD.AcadDocument

    Set MyDocument = AutoApp.Documents.Add
    Set OpenNew = MyDocument
End Function

Public Sub SaveAs(Fichier)
On Error Resume Next

'     AutoApp.Documents(0).ActiveLayer = AutoApp.Documents(0).Layers(0)
        AutoApp.Documents(0).PurgeAll
    AutoApp.Documents(0).SaveAs Fichier, acR15_dwg
    If Err Then MsgBox Err.Description
    Err.Clear
    a = Second(Time)
    While Abs(a - Second(Time)) < 3
    DoEvents
    Wend
    AutoApp.Documents(0).Close , False
    
End Sub
Public Sub CloseDocument()
On Error Resume Next
    AutoApp.Documents(0).Close
    On Error GoTo 0
End Sub
Public Function IsConnecteurs(Attributes As Variant) As Boolean
    Dim Table(5) As String
    Dim Trouve As Boolean
    
Table(0) = "DESIGNATION"
Table(1) = "POS"
Table(2) = "N°"
Table(3) = "CODE_APP"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
    IsConnecteurs = True
    
      For i = LBound(Table) To UBound(Table)
      DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsConnecteurs = False
                Exit Function
            End If
      Next i
End Function
Public Function IsComposants(Attributes As Variant) As Boolean
    Dim Table(2) As String
    Dim Trouve As Boolean
    
Table(0) = "DESIGNCOMP"
Table(1) = "NUMCOMP"
Table(2) = "REFCOMP"
    IsComposants = True
    
      For i = LBound(Table) To UBound(Table)
      DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsComposants = False
                Exit Function
            End If
      Next i
End Function
Public Function IsNotas(Attributes As Variant) As Boolean
    Dim Table(0) As String
    Dim Trouve As Boolean
    
Table(0) = "NUMNOTA"

    IsNotas = True
    
      For i = LBound(Table) To UBound(Table)
      DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsNotas = False
                Exit Function
            End If
      Next i
End Function

Public Function IsTor(Attributes As Variant) As Boolean
    Dim Table(0) As String
    Dim Trouve As Boolean
    
Table(0) = "TORDESIGNATION"

    IsTor = True
    
      For i = LBound(Table) To UBound(Table)
      DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = UCase("" & Attributes(i2).TagString) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsTor = False
                Exit Function
            End If
      Next i
End Function

Public Function IsTorDetail(Attributes As Variant) As Boolean
    Dim Table(2) As String
    Dim Trouve As Boolean
    
Table(2) = "TORDESIGNATION"
Table(1) = "TORFILS"
Table(0) = "TORNUM"
    IsTorDetail = True
    
      For i = LBound(Table) To UBound(Table)
      DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
        Debug.Print UCase("" & Attributes(i2).TagString)
            If Table(i) = UCase("" & Attributes(i2).TagString) Then
            
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsTorDetail = False
                Exit Function
            End If
      Next i
End Function
Public Function IsCartoucheClient(Attributes As Variant) As Boolean
    Dim Table(6) As String
    Dim Trouve As Boolean
Table(0) = "DESIGN.1.CART.RENAULT"
Table(1) = "DESGN.1.ANGL.CART.REN"
Table(2) = "REF.PF.CART.RENAULT"
Table(3) = "IND.PF"
Table(4) = "REF.PLAN.INDUSTRIEL"
Table(5) = "IND.PI"
Table(6) = "REF.PIECE.CART.RENAULT"
    IsCartoucheClient = True
    
      For i = LBound(Table) To UBound(Table)
      DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = UCase("" & Attributes(i2).TagString) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsCartoucheClient = False
                Exit Function
            End If
      Next i
End Function
Public Function IsCartoucheEncelade(Attributes As Variant) As Boolean
    Dim Table(6) As String
    Dim Trouve As Boolean
Table(0) = ".NOM.DU.CLIENT"
Table(1) = ".RESPONSABLE.CLIENT"
Table(2) = ".NOM.DU.PROJET"
Table(3) = ".VAGUE"
Table(4) = ".DESIGNATION.LIGNE.1"
Table(5) = ".OPTION.ET.DIVERSITE"
Table(6) = "REFERENCE.PLAN.CLIENT"
    IsCartoucheEncelade = True
    
      For i = LBound(Table) To UBound(Table)
      DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = UCase("" & Attributes(i2).TagString) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsCartoucheEncelade = False
                Exit Function
            End If
      Next i
End Function


Public Function IsNOMBRE_FILS(Attributes As Variant) As Boolean
    Dim Table(0) As String
    
Table(0) = "_FILS"
   IsNOMBRE_FILS = True
    

   For i = LBound(Table) To UBound(Table)
   DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsNOMBRE_FILS = False
                Exit Function
            End If
      Next i
End Function

Public Function IsEpissures(Attributes As Variant) As Boolean
    Dim Table(6) As String
    
Table(0) = "DESIGNATION"
Table(1) = "POS"
Table(2) = "N°"
Table(3) = "CODE_APP"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
Table(6) = "FILG1"


    IsEpissures = True
    

   For i = LBound(Table) To UBound(Table)
   DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsEpissures = False
                Exit Function
            End If
      Next i
      
End Function
Public Function IsTableauFils(Attributes As Variant) As Boolean
    Dim Table(13) As String
    Table(0) = UCase("LIAI")
    Table(1) = UCase("Designation")
    Table(2) = UCase("Fil")
    Table(3) = UCase("SECT")
    Table(4) = UCase("CO")
    Table(5) = UCase("CO")
    Table(6) = UCase("ISO")
    Table(7) = UCase("POS")
    Table(8) = UCase("Con")
    Table(9) = UCase("VOIE")
    Table(10) = UCase("POS")
    Table(11) = UCase("Con")
    Table(12) = UCase("VOIE")
    Table(13) = UCase("LONG")
    IsTableauFils = True
    
    For i = LBound(Table) To UBound(Table)
    DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsTableauFils = False
                Exit Function
            End If
      Next i
    
    
'
'    For i = LBound(Attributes) To UBound(Attributes)
'        If Trim(UCase(Attributes(i).TagString)) <> Trim(Table(i)) Then
'            IsTableauFils = False
'            Exit Function
'        End If
'    Next i
End Function

Public Function IsCriteres(Attributes As Variant) As Boolean
    Dim Table(1) As String
    Table(0) = UCase("REFCRITERE")
    Table(1) = UCase("REFCRITERELIB")
    IsCriteres = True
    
    For i = LBound(Table) To UBound(Table)
    DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsCriteres = False
                Exit Function
            End If
      Next i
    
    
End Function


Public Function IsNoeuds(Attributes As Variant) As Boolean
    Dim Table(1) As String
    Table(0) = UCase("LONG")
    Table(1) = UCase("NOEUD")
    IsNoeuds = True
    
    For i = LBound(Table) To UBound(Table)
    DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsNoeuds = False
                Exit Function
            End If
      Next i
    
    
End Function

Public Function IsRefOption(Attributes As Variant) As Boolean
    Dim Table(0) As String
    Table(0) = UCase("REFOPTION")
   
    IsRefOption = True
    
    For i = LBound(Table) To UBound(Table)
    DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsRefOption = False
                Exit Function
            End If
      Next i
    
    
End Function
Public Function IsEnteteTableauFils(Attributes As Variant) As Boolean
    Dim Table(12) As String
    Table(0) = UCase("LIAI")
    Table(1) = UCase("DESIGNATION")
    Table(2) = UCase("FIL")
    Table(3) = UCase("SECT")
    Table(4) = UCase("CO")
    Table(5) = UCase("ISO")
    Table(6) = UCase("POS")
    Table(7) = UCase("CON")
    Table(8) = UCase("VOIE")
    Table(9) = UCase("POS")
    Table(10) = UCase("CON")
    Table(11) = UCase("VOIE")
    Table(12) = UCase("LONG")
    
    IsEnteteTableauFils = True
    
    For i = LBound(Table) To UBound(Table)
    DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsEnteteTableauFils = False
                Exit Function
            End If
      Next i
    
'
'    For i = LBound(Attributes) To UBound(Attributes)
'        If Trim(UCase(Attributes(i).TagString)) <> Trim(Table(i)) Then
'            IsEnteteTableauFils = False
'            Exit Function
'        End If
'    Next i
End Function

Public Sub LoadDb()
BdDateTable = CherCheInFihier("BdDateTable")
DbNumPlan = CherCheInFihier("Bdnumero")
db = CherCheInFihier("BdAutocable")
DonneesEntreprise = CherCheInFihier("DonneesEntreprise")
DonneesProduction = CherCheInFihier("DonneesProduction")
funOpenDatabase
NmJob = LaodJob
End Sub
Function LaodJob() As Long
Dim Sql As String
Dim Rs As Recordset
If NmJob = 0 Then

Sql = "SELECT [NumErreur]+1 AS Job FROM T_NumErreur WHERE T_NumErreur.LibErreur='Job';"
Set Rs = Con.OpenRecordSet(Sql)
LaodJob = Rs!Job
Sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1 WHERE T_NumErreur.LibErreur='Job';"
Con.Exequte Sql
Set Rs = Con.CloseRecordSet(Rs)

Else
    LaodJob = NmJob
End If
End Function
Public Function PathArchive(PathRacicine As String, Client As String, CleAc As String, Piece As String, MyType As String, Fichier, IdPieces As Long, Indice_Pieces As String, Indice_Plan As String, Version As Long, Optional NoRegistre As Boolean) As String
Dim Fso As New FileSystemObject
Dim Sql As String
Dim Rs As Recordset
Indice_Pieces = Trim("" & Indice_Pieces)
Indice_Plan = Trim("" & Indice_Plan)
Piece = Replace(Piece, "/", "_", 1)
Piece = Replace(Piece, ":", "", 1)
Piece = Replace(Piece, ".", "", 1)
Piece = Piece & "_" & Indice_Pieces
Fichier = Fichier & "_" & Indice_Plan
Fichier = Replace(Fichier, "/", "_", 1)
Fichier = Replace(Fichier, ":", "", 1)
Fichier = Replace(Fichier, ".", "", 1)
If Version > 1 Then
    Piece = "Temp-" & CStr(Version) & "_" & Piece
End If
If Fso.FolderExists(PathRacicine) = False Then
    Fso.CreateFolder PathRacicine
End If
If Client <> "" Then Client = "\" & Client
If Fso.FolderExists(PathRacicine & Client) = False Then
    Fso.CreateFolder PathRacicine & Client
End If

If Fso.FolderExists(PathRacicine & Client & "\PI") = False Then
    Fso.CreateFolder PathRacicine & Client & "\PI"
End If
If Fso.FolderExists(PathRacicine & Client & "\PI\" & CleAc) = False Then
    Fso.CreateFolder PathRacicine & Client & "\PI\" & CleAc
End If
If Fso.FolderExists(PathRacicine & Client & "\PI\" & CleAc & "\16-PI") = False Then
    Fso.CreateFolder PathRacicine & Client & "\PI\" & CleAc & "\16-PI"
End If



If Fso.FolderExists(PathRacicine & Client & "\PI\" & CleAc & "\16-PI\" & Piece) = False Then
    Fso.CreateFolder PathRacicine & Client & "\PI\" & CleAc & "\16-PI\" & Piece
End If
' & "\16-PI\"
Select Case UCase(MyType)
    Case "PL"
        PathArchive = Client & "\PI\" & CleAc & "\16-PI\" & Piece & "\12-PL\"
    Case "OU"
        PathArchive = Client & "\PI\" & CleAc & "\16-PI\" & Piece & "\16-OU\"
    Case "LI"
        PathArchive = Client & "\PI\" & CleAc & "\16-PI\" & Piece & "\14-LI\"
    Case "PI"
End Select
If Fso.FolderExists(PathRacicine & PathArchive) = False Then
    Fso.CreateFolder PathRacicine & PathArchive
End If
PathArchive = PathArchive & Fichier

If NoRegistre = False Then
If NomenclatureOk = True Then
    Sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(MyType) & "AutoCadSave = '" & MyReplace(PathArchive) & "', T_indiceProjet.NbErr = " & NbError & " "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdPieces & ";"
 Else
    Sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(MyType) & "AutoCadSave = '" & MyReplace(PathArchive) & "' "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdPieces & ";"
 End If
    Con.Exequte Sql
End If
    
'        Sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(Mytype) & "AutoCadSave = '" & MyReplace(PathArchive) & "', T_indiceProjet.NbErr = " & NbError & " "
'    Sql = Sql & "WHERE T_indiceProjet.pere=" & IdPieces & ";"
'    Con.Exequte Sql


PathArchive = PathRacicine & PathArchive
End Function

Public Function DelAttribues(Attributes As Variant)
 For i = LBound(Attributes) To UBound(Attributes)
    Attributes(i).TextString = ""
    
 Next
End Function
Public Function IsVignette(Attributes As Variant)
    Dim Table(5) As String
    Table(2) = UCase("DESIGNATION.HAUT")
    Table(0) = UCase("CODE_APP")
    Table(1) = UCase("N°")
    Table(3) = UCase("DESIGNATION.BAS")
    Table(4) = UCase("DESIGNATION.GAUCHE")
    Table(5) = UCase("DESIGNATION.DROITE")
  
    IsVignette = True
     For i = LBound(Table) To UBound(Table)
     DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsVignette = False
                Exit Function
            End If
      Next i
End Function
Public Function IsVignetteEtiquette(Attributes As Variant)
    Dim Table(3) As String
    Dim Trouve As Boolean
    Table(0) = UCase("FIL1")
    Table(1) = UCase("FIL2")
    Table(2) = UCase("FIL3")
    Table(3) = UCase("FIL4")
  
    IsVignetteEtiquette = True
   
    Trouve = False
    For i2 = LBound(Attributes) To UBound(Attributes)
    DoEvents
        If Trim(UCase(Attributes(i2).TagString)) = "N°" Then
           Trouve = True
           Exit For
        End If
     Next i2
        If Trouve = True Then
         IsVignetteEtiquette = False
            Exit Function
        End If
    For i = LBound(Table) To UBound(Table)
    DoEvents
    Trouve = False
    For i2 = LBound(Attributes) To UBound(Attributes)
        If Trim(UCase(Attributes(i2).TagString)) = Trim(Table(i)) Then
           Trouve = True
           Exit For
        End If
     Next i2
        If Trouve = False Then
         IsVignetteEtiquette = False
            Exit Function
        End If
    Next i
End Function

Public Function IsVignetteEPISSURE(Attributes As Variant)
    Dim Table(0) As String
    Table(0) = UCase("EPISSURE")
  
    IsVignetteEPISSURE = True
    
       For i = LBound(Table) To UBound(Table)
       DoEvents
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next i2
         If Trouve = False Then
                  IsVignetteEPISSURE = False
                Exit Function
            End If
      Next i
   
    
    
    
End Function
Public Sub SubLoadFils(IdPieces As Long, MyType As String)
If (bool_Outil_E_Fils = False And MyType = "PL") Or (bool_Outil_L_Fils = False And MyType = "OU") Then Exit Sub
    Dim RsLigne As Recordset
    Dim Sql As String
    Dim Fso As New FileSystemObject
    Dim NbSupprim As Long
   Dim NbFils As Long
'    Dim NewTor As classTor
   
     
    PathBlocs = TableauPath.Item("PathBlocs")
            
         If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
         If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)

    InsertPointLigneTableau_fils(0) = -1096.5549: InsertPointLigneTableau_fils(1) = -76.8274: InsertPointLigneTableau_fils(2) = 0

Sql = "SELECT Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
    Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT,  "
    Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
    Sql = Sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.POS,  "
    Sql = Sql & "Ligne_Tableau_fils.FA, Ligne_Tableau_fils.VOI,  "
    Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.FA2,  "
    Sql = Sql & "Ligne_Tableau_fils.VOI2,Ligne_Tableau_fils.OPTION,  Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.LONG,  "
    Sql = Sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.APP2 " ',  "
'    Sql = Sql & "T_indiceProjet.Id_Pieces "
    Sql = Sql & "FROM Ligne_Tableau_fils "
    Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdPieces & " and Ligne_Tableau_fils.ACTIVER=true "
    Sql = Sql & "ORDER BY Ligne_Tableau_fils.FIL;"

    Set RsLigne = Con.OpenRecordSet(Sql)
   While RsLigne.EOF = False
    NbFils = NbFils + 1
    RsLigne.MoveNext
   Wend
   RsLigne.Requery
EcritureTor RsLigne, MyType

Ecriturefils RsLigne, MyType, NbFils

        SubLoadCirteres IdPieces, MyType
Fin:
    Set Myrange = Nothing
    Set MySheet = Nothing
    ReDim TableauDeConnecteurs(0)
    Set Fso = Nothing
    Set RsLigne = Con.CloseRecordSet(RsLigne)

End Sub

Public Sub SubLoadCirteres(IdPieces As Long, MyType As String)
If (bool_Plan_E_Criteres = False And MyType = "PL") Or (bool_Outil_E_Criteres = False And MyType = "OU") Then Exit Sub

    Dim RsLigne As Recordset
    Dim Sql As String
    Dim NbSupprim As Long
    Dim NewBlock As AcadBlockReference
    Dim Collec As New Collection
    Dim MyColec As New Collection
    Dim Index As Long
'    Dim NewTor As classTor
   
     
    PathBlocs = TableauPath.Item("PathBlocs")
            
         If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
         If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)
InsertPointLigneCritères(1) = -73#
InsertPointLigneCritères(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneCritères(0), 325) '281.7719)
InsertPointLigneCritères(2) = 0

Sql = "SELECT T_Critères.* " ',  "
'    Sql = Sql & "T_indiceProjet.Id_Pieces "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdPieces & "  AND T_Critères.ACTIVER=True "
Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE ;"

Set RsLigne = Con.OpenRecordSet(Sql)
If RsLigne.EOF = False Then
While RsLigne.EOF = False
Index = Index + 1
RsLigne.MoveNext
Wend
RsLigne.Requery
  FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Max = Index
FormBarGrah.ProgressBar1Caption.Caption = " Chargement des Critères"
Set NewBlock = FunInsBlock(PathBlocs & "\RefCriteres.dwg", InsertPointLigneCritères, "")
    While RsLigne.EOF = False
     IncremanteBarGrah FormBarGrah
    InsertPointLigneCritères(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneCritères(1), -3)
            Set NewBlock = FunInsBlock(PathBlocs & "\RefCriteres.dwg", InsertPointLigneCritères, "")
            Att = NewBlock.GetAttributes
            Set Colec = ColectionAttribueConecteur(Att)
            Att(Colec("REFCRITERE")).TextString = "" & RsLigne!CODE_CRITERE
           Att(Colec("REFCRITERELIB")).TextString = "" & RsLigne!CRITERES

        RsLigne.MoveNext
    Wend
End If


End Sub


Public Sub SacnConnecteur(MyType As String)
If (bool_Plan_E_Etiquettes = False And bool_Plan_L_Connecteurs = False And MyType = "PL") Or (bool_Outil_E_Vignettes = False And bool_Outil_E_Connecteurs = False And MyType = "OU") Then Exit Sub

    Dim Index As Long
    Dim NewBlock  As AcadBlockReference
    Dim MyFichier As String
    Dim PathNUMEROFIL As String
     FormBarGrah.ProgressBar1.Value = 0
      PathNUMEROFIL = TableauPath.Item("PathNUMEROFIL") & "\"
                If Left(PathNUMEROFIL, 2) <> "\\" And Left(PathNUMEROFIL, 1) = "\" Then PathNUMEROFIL = TableauPath.Item("PathServer") & PathNUMEROFIL
                If Right(PathNUMEROFIL, 2) = "\\" Then PathNUMEROFIL = Mid(PathNUMEROFIL, 1, Len(PathNUMEROFIL) - 1)
    If UBound(TableauDeConnecteurs) > 0 Then
         FormBarGrah.ProgressBar1.Max = 1 + UBound(TableauDeConnecteurs)
     Else
         FormBarGrah.ProgressBar1.Max = 1 + 1
     End If
     FormBarGrah.ProgressBar1Caption.Caption = " Chargement des vignettes"
    For Index = 1 To UBound(TableauDeConnecteurs)
         IncremanteBarGrah FormBarGrah
        DoEvents
        
        If TableauDeConnecteurs(Index).ConnecteurExiste = True Then
            If TableauDeConnecteurs(Index).indexFile > 0 Then
                TableauDeConnecteurs(Index).TableauFile = TriTableau(TableauDeConnecteurs(Index).TableauFile)
               
                MyFichier = DirNUMEROFIL(PathNUMEROFIL, TableauDeConnecteurs(Index).indexFile)
                If Trim(MyFichier) <> "" Then
                    InsertionPoint = TableauDeConnecteurs(Index).NewVignette.InsertionPoint
                    
                    InsertPointLigneTableau_fils(0) = InsertionPoint(0) + 18.022: InsertPointLigneTableau_fils(1) = InsertionPoint(1): InsertPointLigneTableau_fils(2) = InsertionPoint(2)
'                    Set NewBlock = FunInsBlock(MyFichier, InsertPointLigneTableau_fils, "NF" & CInt(Index))
InStre = RetournInsertEtiquette(CInt(Index), InsertPointLigneTableau_fils)
                     Set NewBlock = FunInsBlock(MyFichier, RetournInsertEtiquette(CInt(Index), InsertPointLigneTableau_fils), "NF" & CInt(Index), RetourneRotationEtiquette(CInt(Index)), RetourneXEtiquette(CInt(Index)), RetourneYEtiquette(CInt(Index)), RetourneZEtiquette(CInt(Index)))
                    AttribueLibV NewBlock, Index
                    InsertPointLigneTableau_fils(1) = InsertionPoint(1)
                    InsertPointLigneTableau_fils(0) = InsertionPoint(0) + 70
                    Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
                    ReseingeTor RetoureCadeApp(Ats), InsertPointLigneTableau_fils
                End If
            End If
        End If
    Next Index

End Sub
Function TriTableau(MyTableau)
    Dim Index As Long
    Dim boolPlus As Boolean
    a = ""
    For Index = 1 To UBound(MyTableau) - 1
        DoEvents
        
        While Val(MyTableau(Index)) > Val(MyTableau(Index + 1))
            z = MyTableau(Index)
            a = MyTableau(Index + 1)
            MyTableau(Index) = a
            MyTableau(Index + 1) = z
            Index = Index - 1
        Wend
    Next Index
    TriTableau = MyTableau

End Function
Function InsertCartoucheEncelad(Index As Long, MyType As String, ParamArray InsertPointLigneTableau_fils())
Dim Coef As Double
Dim MyMod As Double
MyMod = Index Mod 2
 Coef = Index
If MyMod <> 0 Then
   Coef = Coef - 1
End If
Coef = Coef / 2


If MyType = "OU" Then
     InsertPointLigneTableau_fils(0)(0) = 3711.7662: InsertPointLigneTableau_fils(0)(1) = 115.8474: InsertPointLigneTableau_fils(0)(2) = 0
    GoTo Fin
 End If

Select Case Index
        Case 1
            InsertPointLigneTableau_fils(0)(0) = -1187.6736: InsertPointLigneTableau_fils(0)(1) = 1.346: InsertPointLigneTableau_fils(0)(2) = 0

        Case 2
            InsertPointLigneTableau_fils(0)(0) = -1188.0249: InsertPointLigneTableau_fils(0)(1) = -840.7357: InsertPointLigneTableau_fils(0)(2) = 0


        Case Else
            
            If MyMod = 0 Then
            Coef = Coef - 1
                InsertPointLigneTableau_fils(0)(0) = -1188.0249: InsertPointLigneTableau_fils(0)(1) = -840.7357: InsertPointLigneTableau_fils(0)(2) = 0
                InsertPointLigneTableau_fils(0)(0) = InsertPointLigneTableau_fils(0)(0) - (-1188.7956 * Coef)

            Else
                InsertPointLigneTableau_fils(0)(0) = -1187.6736: InsertPointLigneTableau_fils(0)(1) = 1.146: InsertPointLigneTableau_fils(0)(2) = 0
                InsertPointLigneTableau_fils(0)(0) = InsertPointLigneTableau_fils(0)(0) - (-1188.7956 * Coef)
            End If

End Select
Fin:
End Function
Function InsertCartoucheClient(Index As Long, MyType As String, ParamArray InsertPointLigneTableau_fils())
If MyType = "OU" Then
    InsertPointLigneTableau_fils(0)(0) = 3566.7253: InsertPointLigneTableau_fils(0)(1) = 115.8474: InsertPointLigneTableau_fils(0)(2) = 0
    Exit Function
End If
Dim MyMod As Double
Dim Coef As Double
MyMod = Index Mod 2
Coef = Index
If MyMod <> 0 Then
    Coef = Coef - 1
End If
Coef = Coef / 2


Select Case Index
        Case 1
                InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = 126.0753: InsertPointLigneTableau_fils(0)(2) = 0

        Case 2
                 InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = -713.9247: InsertPointLigneTableau_fils(0)(2) = 0


        Case Else
        If MyMod <> 0 Then
            
            aa = -165.0409 - 1022.9591
            InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = 126.0753: InsertPointLigneTableau_fils(0)(2) = 0
            InsertPointLigneTableau_fils(0)(0) = InsertPointLigneTableau_fils(0)(0) - (aa * Coef)
        Else
        Coef = Coef - 1
            aa = -165.0409 - 1022.9591
            InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = -713.9247: InsertPointLigneTableau_fils(0)(2) = 0
            InsertPointLigneTableau_fils(0)(0) = InsertPointLigneTableau_fils(0)(0) - (aa * Coef)
        End If
        
        
End Select
End Function

Public Function ChargeCartoucheEncelade(IdIndiceProjet As Long, MyType As String, NbCartouche As Long, Optional OuOk As Boolean) As Boolean
If (bool_Plan_E_cartouches = False And MyType = "PL") Or (bool_Outil_E_cartouches = False And MyType = "OU") Then Exit Function
  
Dim Fso As New FileSystemObject
Dim Sql As String
Dim Rs As Recordset
Dim Status As String
Dim FichierCartouche As String
Dim Index As Long

Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Cartouche = '" & Replace(MyReplace(RepPlacheClous), TableauPath.Item("PathServer"), "") & "' "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Con.Exequte Sql
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & IdIndiceProjet & ";"

If Left(RepPlacheClous, 2) <> "\\" And Left(RepPlacheClous, 1) = "\" Then RepPlacheClous = TableauPath.Item("PathServer") & RepPlacheClous
If Right(RepPlacheClous, 2) = "\\" Then RepPlacheClous = Mid(RepPlacheClous, 1, Len(RepPlacheClous) - 1)
PathBlocs = TableauPath.Item("PathBlocs")
 If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
 If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)
 PereFilsOk = AtocatOption(IdIndiceProjet)
 Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Function
    Status = "" & Rs!Status
    CartoucheCleient = False
    For Index = 1 To NbCartouche
      InsertCartoucheEncelad Index, MyType, InsertPointLigneTableau_fils
    If MyType = "OU" Then
        FichierCartouche = RepPlacheClous
      
    Else
        If Index = 1 Then
             FichierCartouche = PathBlocs & "\1 CARTOUCHE ENCELADE.dwg"
        Else
         FichierCartouche = PathBlocs & "\CARTOUCHE ENCELADE.dwg"
        End If
    End If
    If Fso.FileExists(FichierCartouche) = False Then
        Set Fso = Nothing
        Exit Function
    End If

    Set NewBlock = FunInsBlock(FichierCartouche, InsertPointLigneTableau_fils, "LeCartouche1E")

    AttClent = NewBlock.GetAttributes

    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    
    AttClent(AttribuCartouche(".BASE.VEHICULE")).TextString = "" & Rs("BaseVehicule")
    AttClent(AttribuCartouche(".NOM.DU.Client")).TextString = "" & Rs!Client
    AttClent(AttribuCartouche(".RESPONSABLE.Client")).TextString = "" & Rs!Responsable
    AttClent(AttribuCartouche(".NOM.DU.Projet")).TextString = "" & Rs!Projet
    AttClent(AttribuCartouche(".VAGUE")).TextString = "" & Rs!Vague
        AttClent(AttribuCartouche(".DESIGNATION.LIGNE.1")).TextString = Replace("" & Rs!Ensemble, Chr(13), "")
        txt = "" & Rs!Equipement
        If PereFilsOk = True Then txt = ""
            AttClent(AttribuCartouche(".OPTION.ET.DIVERSITE")).TextString = txt
        
       
    
       
   
        AttClent(AttribuCartouche("Reference.PLAN.Client")).TextString = "" & Rs!PL
         AttClent(AttribuCartouche("INDICE")).TextString = "" & Rs!PL_Indice
    
   
    AttClent(AttribuCartouche("Reference.PLAN.FONCTIONNEL")).TextString = "" & Rs!RefPF
  
    AttClent(AttribuCartouche("RF2")).TextString = "" & Rs!Ref_PF
   txt = "" & Rs!Pi
        If PereFilsOk = True Then txt = ""
        AttClent(AttribuCartouche("Reference.ENCELADE")).TextString = txt
        txt = "" & Rs!PI_Indice
        If PereFilsOk = True Then txt = ""
        AttClent(AttribuCartouche("RF1")).TextString = txt
     
     AttClent(AttribuCartouche("Reference.OU.ENCELADE")).TextString = "" & Rs!OU
         AttClent(AttribuCartouche("RF3")).TextString = "" & Rs!OU_Indice
    AttClent(AttribuCartouche("DESSINE.PAR")).TextString = "" & Rs!DessineNOM
    AttClent(AttribuCartouche("DESSINELE")).TextString = "" & Rs!DessineDate
    AttClent(AttribuCartouche("VERIFIE.PAR")).TextString = "" & Rs!VerifieNom
    AttClent(AttribuCartouche("VERIFIELE")).TextString = "" & Rs!VerifieDate
    AttClent(AttribuCartouche("APPROUVE.PAR")).TextString = "" & Rs!ApprouveNom
    AttClent(AttribuCartouche("APPROUVELE")).TextString = "" & Rs!ApprouveDate
    AttClent(AttribuCartouche(".MASSE")).TextString = "" & Rs!Masse
    AttClent(AttribuCartouche("ETAT")).TextString = Status
    AttClent(AttribuCartouche("X/X")).TextString = CStr(Index) & "/" & CStr(NbCartouche)
   Next Index

    CartoucheCleient = True
Set Fso = Nothing
End Function
Public Function ChargeCartoucheClient(IdIndiceProjet As Long, MyType As String, NbCartouche As Long, Optional OuOk As Boolean) As Boolean
If (bool_Plan_E_cartouches = False And MyType = "PL") Or (bool_Outil_E_cartouches = False And MyType = "OU") Then Exit Function

Dim Sql As String
Dim FichierCartouche As String
Dim Rs As Recordset
Dim RsCartouche As Recordset
Dim Index As Long
Dim NbCar As Long
LeCartouche = "CARTOUCHE  RENAULT.dwg"
LeCartoucheE = "CARTOUCHE ENCELADE.dwg"
NbCar = 2
If OuOk = False Then NbCar = NbCartouche
'If boolFormClient = False Then Exit Function
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Function
MyCARTOUCHE_Client = Trim("" & Rs!Client)
Sql = "SELECT T_Clients.Formulaire FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(Trim("" & Rs!Client)) & "';"
Set RsCartouche = Con.OpenRecordSet(Sql)
If RsCartouche.EOF = False Then
    LeCartouche = Trim("" & RsCartouche!Formulaire)
     If Left(LeCartouche, 2) <> "\\" And Left(LeCartouche, 1) = "\" Then LeCartouche = TableauPath.Item("PathServer") & LeCartouche
     If Right(LeCartouche, 2) = "\\" Then LeCartouche = Mid(LeCartouche, 1, Len(LeCartouche) - 1)
     
End If
Set RsCartouche = Con.CloseRecordSet(RsCartouche)
If LeCartouche = "" Then Exit Function
    Dim Fso As New FileSystemObject
    CartoucheCleient = False
    For Index = 1 To NbCar
    InsertCartoucheClient Index, MyType, InsertPointLigneTableau_fils
   
    If Fso.FileExists(LeCartouche) = False Then
        Set Fso = Nothing
        Exit Function
    End If
    Set NewBlock = FunInsBlock(LeCartouche, InsertPointLigneTableau_fils, "LeCartouche1")
'    NewBlock.Application.Visible = True
    AttClent = NewBlock.GetAttributes

    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
AttClent(AttribuCartouche("DESIGN.1.CART.RENAULT")).TextString = "" & Rs("Ensemble")
AttClent(AttribuCartouche("MASSE")).TextString = "" & Rs("Masse")

AttClent(AttribuCartouche("REF.PF.CART.RENAULT")).TextString = "" & Rs("RefPF")
AttClent(AttribuCartouche("IND.PF")).TextString = "" & Rs("Ref_PF")
'If OuOk = True Then
'
'        AttClent(AttribuCartouche("REF.PLAN.INDUSTRIEL")).TextString = "" & Rs!OU
'        AttClent(AttribuCartouche("IND.PI")).TextString = "" & Rs("OU_Indice")
'    Else
    
        AttClent(AttribuCartouche("REF.PLAN.INDUSTRIEL")).TextString = "" & Rs!RefP
         AttClent(AttribuCartouche("IND.Pi")).TextString = "" & Rs("Ref_Plan_CLI")
    
'    End If
AttClent(AttribuCartouche("DESIGN.2.CART.RENAULT")).TextString = ""
AttClent(AttribuCartouche("DESGN.1.ANGL.CART.REN")).TextString = ""
AttClent(AttribuCartouche("DESGN.2.ANGL.CART.REN")).TextString = ""

'AttClent(AttribuCartouche("IND.PF")).TextString = ""

 
AttClent(AttribuCartouche("REF.PIECE.CART.RENAULT")).TextString = "" & Rs("RefPieceClient") & "_" & Trim("" & Rs("Ref_Piece_CLI"))
AttClent(AttribuCartouche("SERVICE")).TextString = "" & Rs("Service")
AttClent(AttribuCartouche("UTILISATEURS")).TextString = "" & Rs("Destinataire")



AttClent(AttribuCartouche("REGLEMENT")).TextString = ""
AttClent(AttribuCartouche("NOTE.BE.1")).TextString = ""
AttClent(AttribuCartouche("NOTE.BE.2")).TextString = ""
AttClent(AttribuCartouche("Num.VISA")).TextString = ""
'AttClent(AttribuCartouche("REF.PIECE.CART." & MyCARTOUCHE_Client)).TextString = Trim("" & Rs!RefP) & "_" & Trim("" & Rs!Ref_PF)
AttClent(AttribuCartouche("X/X")).TextString = CStr(Index) & "/" & CStr(NbCartouche)
Next Index

'AttClent(AttribuCartouche("DESGN.1.ANGL.CART.REN")).TextString = ""
    CartoucheCleient = True
    Set Fso = Nothing
End Function
Public Sub Maj(MyControl As Object)
Dim Rs As Recordset
Dim Sql As String
Dim txt As String
txt = MyControl.Text
MyControl.Clear
Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"
MyControl.AddItem ""
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    MyControl.AddItem Trim("" & Rs!Client)
        If UCase(MyControl.List(MyControl.ListCount - 1)) = UCase(txt) Then MyControl.ListIndex = MyControl.ListCount - 1
         
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)

End Sub
Public Function ChercheXls(Myrange, Val, Optional MyRange2, Optional Cherche2 As Boolean) As Long
ChercheXls = 0
 For i = 2 To Myrange.Rows.Count
                If UCase(Trim("" & Myrange(i))) = UCase(Trim("" & Val)) Then
                If Cherche2 = True Then
                    If MyRange2(i) = 1 Then
                         ChercheXls = i
                        Exit For
                    End If
                Else
                        ChercheXls = i
                        Exit For
                End If
                End If
            Next i



End Function
Public Function CherCheInFihier(Cherher As String) As String
Dim FileNumber As Long
Dim MyString As String
FileNumber = FreeFile

  
Open App.Path & "\Autocable.ini" For Input As #FileNumber
Do While Not EOF(FileNumber)    ' Effectue la boucle jusqu'à la fin du fichier.
    Input #FileNumber, MyString ' Lit les données dans deux variables.
    If InStr(1, MyString, Cherher) <> 0 Then
       CherCheInFihier = Right(MyString, Len(MyString) - (Len(Cherher) + InStr(1, MyString, Cherher)))
       Exit Do
    End If
Loop
Close #FileNumber    ' Ferme le fichier
CherCheInFihier = Replace(CherCheInFihier, "§remplaceDate§", Format(Date, "yyyy"))
CherCheInFihier = Trim(CherCheInFihier)
End Function
Public Function RenseigneConnecteurBroches(RangeAttribue As Recordset, MyType As String) As Boolean
If (bool_Plan_E_Connecteurs = False And MyType = "PL") Or (bool_Outil_E_Connecteurs = False And MyType = "OU") Then Exit Function

Dim txtErr As String
    On Error Resume Next
    RenseigneConnecteurBroches = False
'    CollectionTor
    If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).ConnecteurExiste = True Then
        If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).EPISSURE = False Then
            a = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).NewBlock.GetAttributes
            txtErr = "LIAI"
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("LIAI" & RangeAttribue.Fields(9))
            If Err Then
                FunError 5, txtErr & RangeAttribue.Fields(9), Err.Description, "" & RangeAttribue.Fields("APP")
                Err.Clear
            End If
            dd = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Count
            a(IdAttrib).TextString = "" & RangeAttribue.Fields(0)
             txtErr = "FIL"
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("FIL" & RangeAttribue.Fields(9))
            If Trim("" & a(IdAttrib).TextString) <> "" Then
                 txtErr = "MAR"
                IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("MAR" & RangeAttribue.Fields(9))
            End If
            a(IdAttrib).TextString = "" & RangeAttribue.Fields(2)
        Else
            If FunEPISSURE(TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).NewBlock.GetAttributes, "" & RangeAttribue.Fields(2), "" & RangeAttribue.Fields(9), CLng(CollectionCon("" & RangeAttribue.Fields("APP")))) = False Then Exit Function
        End If

    Else
        FunError 3, "" & RangeAttribue.Fields(2), Err.Description, "" & RangeAttribue.Fields("APP")
    End If
    Err.Clear
    If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).ConnecteurExiste = True Then

        If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).EPISSURE = False Then
            a = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).NewBlock.GetAttributes
             txtErr = "LIAI"
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("LIAI" & RangeAttribue.Fields(12))
            If Err Then
                FunError 5, txtErr & RangeAttribue.Fields(12), Err.Description, "" & RangeAttribue.Fields("APP2")
                Err.Clear
            End If
        
            dd = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Count
            a(IdAttrib).TextString = "" & RangeAttribue.Fields(0)
             txtErr = "FIL"
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("FIL" & RangeAttribue.Fields(12))
            If Trim("" & a(IdAttrib).TextString) <> "" Then
             txtErr = "MAR"
                IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("MAR" & RangeAttribue.Fields(12))
            End If
            a(IdAttrib).TextString = "" & RangeAttribue.Fields(2)
        Else
            If FunEPISSURE(TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).NewBlock.GetAttributes, "" & RangeAttribue.Fields(2), "" & RangeAttribue.Fields(12), CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))) = False Then Exit Function

        End If
    Else
        FunError 3, "" & RangeAttribue.Fields(2), Err.Description, "" & RangeAttribue.Fields("APP2")
    End If
    RenseigneConnecteurBroches = True
End Function
Sub AttribueLibV(NewBlock As AcadBlockReference, Index As Long)
    Dim At
    Dim MyAtt As New Collection
    At = NewBlock.GetAttributes
    
    For i = 0 To UBound(At)
        DoEvents
       
        Debug.Print At(i).TagString
        MyAtt.Add CStr(i), At(i).TagString
    Next i
    Set TableauDeConnecteurs(Index).AttribuesFils = MyAtt
    Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
    At(TableauDeConnecteurs(Index).AttribuesFils("DESIGNATION")).TextString = Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")"
    Set MyAtt = Nothing
    i3 = UBound(Ats) - 1
    Dim i2 As Long
    
    For i = 1 To TableauDeConnecteurs(Index).indexFile
        DoEvents
       
        At(TableauDeConnecteurs(Index).AttribuesFils("FIL" & CStr(i))).TextString = TableauDeConnecteurs(Index).TableauFile(i)
    Next i
'    ReseingeTor "aaa"
End Sub
Function RetoureCadeApp(Ats) As String
RetoureCadeApp = ""
For i = 0 To UBound(Ats) - 1
        If Ats(i).TagString = "CODE_APP" Then
        RetoureCadeApp = Ats(i).TextString
        Exit Function
        End If
    Next i
End Function
Function RetournInsertEtiquette(Index As Integer, InsertionPoint)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetournInsertEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).InsertTorTitre
                     Else
                       RetournInsertEtiquette = InsertionPoint
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function
Function RetourneXEtiquette(Index As Integer)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetourneXEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).XScaleFactor
                     Else
                        RetourneXEtiquette = 1
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function
Function RetourneYEtiquette(Index As Integer)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetourneYEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).YScaleFactor
                     Else
                        RetourneYEtiquette = 1
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function
Function RetourneZEtiquette(Index As Integer)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetourneZEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).ZScaleFactor
                     Else
                        RetourneZEtiquette = 1
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function
Function RetourneRotationEtiquette(Index As Integer)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetourneRotationEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).Rotation
                     Else
                        RetourneRotationEtiquette = 0
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function


Sub TestFl()
a = CherCheInFihier("Bdnumero")
End Sub
Public Function LoadComposants(IdIndiceProjet As Long, MyType As String) As Boolean
If (bool_Plan_E_Composants = False And MyType = "PL") Or (bool_Outil_E_Composants = False And MyType = "OU") Then Exit Function

  LoadComposants = False
    Dim RsCompsants As Recordset
    Dim Sql As String
    Dim Myrep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
  
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
   Dim XMin As Double
   Dim YMin As Double
    Dim PathComposantsDefault As String
     Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1
    
      Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathComposants FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathComposants) = "" Then
         PathComposantsDefault = TableauPath.Item("PathComposantsDefault")
   Else
             PathComposantsDefault = RsConnecteur!PathComposants
         If Left(PathComposantsDefault, 2) <> "\\" And Left(PathComposantsDefault, 1) = "\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
             If Right(PathComposantsDefault, 2) = "\\" Then PathComposantsDefault = Mid(PathComposantsDefault, 1, Len(PathComposantsDefault) - 1)
    
    End If
Else
                 PathComposantsDefault = TableauPath.Item("PathComposantsDefault")

End If
If Left(PathComposantsDefault, 2) <> "\\" And Left(PathComposantsDefault, 1) = "\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
If Right(PathComposantsDefault, 2) = "\\" Then PathComposantsDefault = Mid(PathComposantsDefault, 1, Len(PathComposantsDefault) - 1)
' PathComposantsDefault = PathComposantsDefault & "COMPOSANTS\"
 
 Sql = "SELECT Composants.* FROM Composants "
 Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & " AND Composants.ACTIVER=True;"
 Set RsCompsants = Con.OpenRecordSet(Sql)
 While RsCompsants.EOF = False
 On Error Resume Next
                  
                   a = CollectionComp(Trim("C" & RsCompsants!NUMCOMP))
                If Err Then
                    If NUMCOM < RsCompsants!NUMCOMP Then
                         ReDim Preserve TableauComposant(RsCompsants!NUMCOMP)
                         
                         NUMCOM = RsCompsants!NUMCOMP
                    End If
                     CollectionComp.Add RsCompsants("NUMCOMP").Value, Trim("C" & RsCompsants!NUMCOMP)
                End If
 
    
    RsCompsants.MoveNext
 Wend
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = 1 + NUMCOM
 FormBarGrah.ProgressBar1Caption.Caption = " Chargement des Compsants"
 
  RsCompsants.Requery
   XMin = 823.5964: YMin = -954.9939
    For i = 0 To IndexIstC
  
    InsertPointConnecteur(i).InsertPointConnecteur(0) = XMin - (150 * i): InsertPointConnecteur(i).InsertPointConnecteur(1) = YMin - (150 * i): InsertPointConnecteur(i).InsertPointConnecteur(2) = 0
    Next i
 On Error GoTo GesERR
 
 While RsCompsants.EOF = False
  IncremanteBarGrah FormBarGrah
 If TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).PosOkComp = False Then
  InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0), -300)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
                rr = InsertPointHiérarchie(Val(Trim("" & RsCompsants!NUMCOMP)), 0, -1000, 150, -150, 10)
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertComp(0) = rr(0)
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertComp(1) = rr(1)
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertComp(2) = rr(2)
   
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).XScaleFactorComp = 1
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).YScaleFactorComp = 1
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).ZScaleFactorComp = 1
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).RotationComp = 0
 End If

'    PathComposantsDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS"
    Lib1 = PathComposantsDefault & "\" & RsCompsants!Path & "\" & RsCompsants!REFCOMP & ".dwg"
    Lib2 = "" & RsCompsants!REFCOMP
    NumErr = 6
    Set TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).BlockComp = FunInsBlock(PathComposantsDefault & "\" & RsCompsants!Path & "\" & RsCompsants!REFCOMP & ".dwg", TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertComp, "", TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).RotationComp, TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).XScaleFactorComp, TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).YScaleFactorComp, TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).ZScaleFactorComp)
    Err.Clear
    Set Attribues = ColectionAttribueConecteur(TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).BlockComp.GetAttributes)

                Att = TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).BlockComp.GetAttributes
                Lib1 = "DESIGNCOMP"
                Lib2 = "" & RsCompsants!REFCOMP
                NumErr = 7
                Att(Attribues("DESIGNCOMP")).TextString = "" & RsCompsants!DESIGNCOMP
                Err.Clear
                 Lib1 = "NUMCOMP"
                Lib2 = "" & RsCompsants!REFCOMP
                Att(Attribues("NUMCOMP")).TextString = "C" & RsCompsants!NUMCOMP
                Err.Clear
                 Lib1 = "PATHCOMP"
                Lib1 = "NUMCOMP"
                Att(Attribues("PATHCOMP")).TextString = "" & RsCompsants!Path
                  Err.Clear
                 Lib1 = "REFCOMP"
                Lib2 = ""
                Att(Attribues("REFCOMP")).TextString = "" & RsCompsants!REFCOMP
                
                
                
                
     If TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).PosOkDesin = False Then
''  InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0), -300)
'                Nb_L_C = Nb_L_C + 1
'                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertDesing(0) = TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertComp(0) - 41.6692
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertDesing(1) = TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertComp(1) - 13.186
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertDesing(2) = TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertComp(2)
   
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).XScaleFactorDesin = 1
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).YScaleFactorDesin = 1
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).ZScaleFactorDesin = 1
    TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).RotationDesin = 0
 End If

'    PathComposantsDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS"
    Lib1 = PathComposantsDefault & "\" & RsCompsants!Path & "\" & RsCompsants!REFCOMP & ".dwg"
    Lib2 = "" & RsCompsants!REFCOMP
    NumErr = 6
    Set TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).BlockDesing = FunInsBlock(PathComposantsDefault & "\COMP_DESGN.dwg", TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertDesing, "", TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).RotationDesin, TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).XScaleFactorDesin, TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).YScaleFactorDesin, TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).ZScaleFactorDesin)
    Err.Clear
    Set Attribues = ColectionAttribueConecteur(TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).BlockDesing.GetAttributes)

                Att = TableauComposant(CollectionComp("C" & RsCompsants!NUMCOMP)).BlockDesing.GetAttributes
                Lib1 = "DESIGNCOMP"
                Lib2 = "" & RsCompsants!REFCOMP
                NumErr = 7
                Att(Attribues("DESIGNCOMP")).TextString = "" & RsCompsants!DESIGNCOMP
                Err.Clear
                 Lib1 = "NUMCOMP"
                Lib2 = "" & RsCompsants!REFCOMP
                Att(Attribues("NUMCOMP")).TextString = "C" & RsCompsants!NUMCOMP
                Err.Clear
                 Lib1 = "PATHCOMP"
                Lib1 = "NUMCOMP"
                Att(Attribues("PATHCOMP")).TextString = "" & RsCompsants!Path
                  Err.Clear
                 Lib1 = "REFCOMP"
                Lib2 = ""
                Att(Attribues("REFCOMP")).TextString = "" & RsCompsants!REFCOMP
               
    RsCompsants.MoveNext
 Wend
 Exit Function
GesERR:
    FunError NumErr, "" & Lib1, Err.Description, "" & Lib2
 Resume Next
End Function

Public Function LoadNoeuds(IdIndiceProjet As Long, MyType As String) As Boolean
If (bool_Plan_E_Noeuds = False And MyType = "PL") Or (bool_Outil_E_Noeuds = False And MyType = "OU") Then Exit Function
  LoadNoeuds = False
    Dim RsCompsants As Recordset
    Dim Sql As String
    Dim Myrep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
    
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
    Dim XMin As Double
   Dim YMin As Double
    Dim PathNotasDefault As String
     Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1
    
      Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathNotas FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathNotas) = "" Then
         PathNotasDefault = TableauPath.Item("PathNotasDefault")
   Else
             PathNotasDefault = RsConnecteur!PathNotas
         If Left(PathNotasDefault, 2) <> "\\" And Left(PathNotasDefault, 1) = "\" Then PathNotasDefault = TableauPath.Item("PathServer") & PathNotasDefault
         If Right(PathNotasDefault, 2) = "\\" Then PathNotasDefault = Mid(PathNotasDefault, 1, Len(PathNotasDefault) - 1)
    
    End If
Else
                 PathNotasDefault = TableauPath.Item("PathNotasDefault")

End If
PathBlocs = TableauPath.Item("PathBlocs")
 If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
 If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)
' PathNotasDefault = PathNotasDefault & "Nota\"
 
 Sql = "SELECT T_Noeuds.* FROM T_Noeuds "
 Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndiceProjet & " and T_Noeuds.ACTIVER=true "
 Sql = Sql & "order by T_Noeuds.id;"
 Set RsCompsants = Con.OpenRecordSet(Sql)

' Set CollectionNoeuds = New Collection
 While RsCompsants.EOF = False
 On Error Resume Next
                  
                   a = CollectionNoeuds(Trim("N" & RsCompsants!NŒUDS))
                If Err Then
                NUMNOEUDS = NUMNOEUDS + 1
                    Err.Clear
                         ReDim Preserve TableauDeNoeuds(NUMNOEUDS)
                         CollectionNoeuds.Add NUMNOEUDS, Trim("N" & RsCompsants!NŒUDS)
                         
                  
 
                End If
 
    
    RsCompsants.MoveNext
 Wend
   
  RsCompsants.Requery
   XMin = -1337.8928: YMin = 870.4179
    For i = 0 To IndexIstN
  
    InsertNouds(i).InsertPointConnecteur(0) = XMin - (50 * i): InsertNouds(i).InsertPointConnecteur(1) = YMin + (50 * i): InsertNouds(i).InsertPointConnecteur(2) = 0
    Next i
 On Error GoTo GesERR
 
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = 1 + NUMNOEUDS
 FormBarGrah.ProgressBar1Caption.Caption = " Chargement des Noeuds"
 DoEvents
 While RsCompsants.EOF = False
  IncremanteBarGrah FormBarGrah
 If TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).PosOkComp = False Then
  zz = IndexationNoeuds(RsCompsants!Noeuds)
   rr = InsertPointHiérarchie(Val(zz), -100, -1000, -50, -40, 10)
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(0) = rr(0)
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(1) = rr(1)
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(2) = rr(2)
   
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).XScaleFactorComp = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).YScaleFactorComp = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).ZScaleFactorComp = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).RotationComp = 0
     InsertNouds(Nb_L_C).InsertPointConnecteur(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertNouds(Nb_L_C).InsertPointConnecteur(1), -40)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstN + 1 Then
                    Nb_L_C = 0
                    For i = 0 To IndexIstN
                       InsertNouds(i).InsertPointConnecteur(0) = InsertNouds(i).InsertPointConnecteur(0) - 50: InsertNouds(i).InsertPointConnecteur(1) = YMin + (50 * i): InsertNouds(i).InsertPointConnecteur(2) = 0
                     Next
                End If
 End If

    Lib1 = PathBlocs & "\NOEUD.dwg"
    Lib2 = "" & RsCompsants!Noeuds
    NumErr = 6
  
    Set TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockComp = FunInsBlock(Lib1, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp, "", TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).RotationComp, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).XScaleFactorComp, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).YScaleFactorComp, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).ZScaleFactorComp)
    Set Attribues = ColectionAttribueConecteur(TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockComp.GetAttributes)

                Att = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockComp.GetAttributes
                
                NumErr = 7
                
                Att(Attribues("LONG")).TextString = "" & RsCompsants!Longueur
                Att(Attribues("NOEUD")).TextString = "" & RsCompsants!NŒUDS
                Att(Attribues("DIAM")).TextString = "" & RsCompsants!DIAMETRE
                Att(Attribues("HAB")).TextString = "" & RsCompsants!CODE_ENC
                Att(Attribues("CLASSE_T")).TextString = "" & RsCompsants!CLASSE_T
                Err.Clear
                
                
                
 
If TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).PosOkDesin = False Then
 
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertDesing(0) = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(0) + 10
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertDesing(1) = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(1)
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertDesing(2) = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(2)
   
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).XScaleFactorDesin = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).YScaleFactorDesin = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).ZScaleFactorDesin = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).RotationDesin = 0
 End If
If Trim("" & RsCompsants!NŒUDS) <> "AA" Then
'    PathNotasDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\Nota"
'If "" & RsCompsants!NŒUDS = "AA" Then
'    Lib1 = PathBlocs & "\NOEUD_0.dwg"
'Else
    Lib1 = PathBlocs & "\NOEUD_LONG.dwg"
'End If
    Lib2 = "" & RsCompsants!Noeuds
    NumErr = 6
    Set TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockDesing = FunInsBlock(Lib1, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertDesing, "", TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).RotationDesin, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).XScaleFactorDesin, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).YScaleFactorDesin, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).ZScaleFactorDesin)
    Set Attribues = ColectionAttribueConecteur(TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockDesing.GetAttributes)

                Att = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockDesing.GetAttributes
                
                NumErr = 7
                
'                 Lib1 = "NUMNOTA"
'                Lib2 = "" & RsCompsants!Nota
                Att(Attribues("LONG")).TextString = "" & RsCompsants!Longueur
                Att(Attribues("NOEUD")).TextString = "" & RsCompsants!NŒUDS
                Att(Attribues("DIAM")).TextString = "" & RsCompsants!DIAMETRE
                Att(Attribues("HAB")).TextString = "" & RsCompsants!CODE_ENC
                Att(Attribues("CLASSE_T")).TextString = "" & RsCompsants!CLASSE_T
                Err.Clear
 End If
 
                

If TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).PosOkFleche = False Then
 
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertFleche(0) = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(0)
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertFleche(1) = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(1) + 20
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertFleche(2) = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertComp(2)
   
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).XScaleFactorFleche = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).YScaleFactorFleche = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).ZScaleFactorFleche = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).RotationFleche = 0
 End If

'    PathNotasDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\Nota"
'If "" & RsCompsants!NŒUDS = "AA" Then
'    Lib1 = PathBlocs & "\NOEUD_0.dwg"
'Else
If RsCompsants!TORON_PRINCIPAL = True Then

    Lib1 = PathBlocs & "\NOEUD_PRINCIPAL"
Else
     Lib1 = PathBlocs & "\NOEUD_SECONDAIRE" '.dwg"
    End If
    If RsCompsants!Fleche_Droite = False Then
        Lib1 = Lib1 & ".dwg"
    Else
        Lib1 = Lib1 & "1.dwg"
    End If
'End If
    Lib2 = "" & RsCompsants!Noeuds
    NumErr = 6
    Set TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockFleche = FunInsBlock(Lib1, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).InsertFleche, "", TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).RotationFleche, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).XScaleFactorFleche, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).YScaleFactorFleche, TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).ZScaleFactorFleche)
    Set Attribues = ColectionAttribueConecteur(TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockFleche.GetAttributes)

                Att = TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).BlockFleche.GetAttributes
                
                NumErr = 7
                
'                 Lib1 = "NUMNOTA"
'                Lib2 = "" & RsCompsants!Nota
                Att(Attribues("LONG")).TextString = "" & RsCompsants!Longueur
                Att(Attribues("NOEUD")).TextString = "" & RsCompsants!NŒUDS
                Att(Attribues("DIAM")).TextString = "" & RsCompsants!DIAMETRE
                Att(Attribues("HAB")).TextString = "" & RsCompsants!CODE_ENC
                Att(Attribues("CLASSE_T")).TextString = "" & RsCompsants!CLASSE_T
                Att(Attribues("LONG_CUMUL")).TextString = "" & RsCompsants!LONGUEUR_CUMULEE
                Err.Clear

'                 Lib1 = "NOTA"
'                Lib2 = ""
'                Att(TableauDeNoeuds(CollectionNoeuds("N" & RsCompsants!Noeuds)).Attribues("NOTA")).TextString = "" & RsCompsants!NOTA
    RsCompsants.MoveNext
 Wend
 Exit Function
GesERR:
    FunError NumErr, "" & Lib1, Err.Description, "" & Lib2
 Resume Next
End Function

Public Function LoadNotas(IdIndiceProjet As Long, MyType As String) As Boolean
If (bool_Plan_E_Notas = False And MyType = "PL") Or (bool_Outil_E_Notas = False And MyType = "OU") Then Exit Function

  LoadNotas = False
    Dim RsCompsants As Recordset
    Dim Sql As String
    Dim Myrep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
    
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
   Dim XMin As Double
   Dim YMin As Double
    Dim PathNotasDefault As String
     Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1
    
      Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathNotas FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathNotas) = "" Then
         PathNotasDefault = TableauPath.Item("PathNotasDefault")
   Else
             PathNotasDefault = RsConnecteur!PathNotas
         If Left(PathNotasDefault, 2) <> "\\" And Left(PathNotasDefault, 1) = "\" Then PathNotasDefault = TableauPath.Item("PathServer") & PathNotasDefault
         If Right(PathNotasDefault, 2) = "\\" Then PathNotasDefault = Mid(PathNotasDefault, 1, Len(PathNotasDefault) - 1)
    
    End If
Else
                 PathNotasDefault = TableauPath.Item("PathNotasDefault")

End If
If Left(PathNotasDefault, 2) <> "\\" And Left(PathNotasDefault, 1) = "\" Then PathNotasDefault = TableauPath.Item("PathServer") & PathNotasDefault
If Right(PathNotasDefault, 2) = "\\" Then PathNotasDefault = Mid(PathNotasDefault, 1, Len(PathNotasDefault) - 1)
' PathNotasDefault = PathNotasDefault & "Nota\"
 
 Sql = "SELECT Nota.* FROM Nota "
 Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndiceProjet & "  AND Nota.ACTIVER=True "
 Sql = Sql & "order by Nota.NUMNOTA;"
 Set RsCompsants = Con.OpenRecordSet(Sql)
 While RsCompsants.EOF = False
 On Error Resume Next
                    a = ""
                   a = CollectionNota(Trim("N" & RsCompsants!NUMNOTA))
                If Err Then
                
                    If NUMNOTA < RsCompsants!NUMNOTA Then
                         ReDim Preserve TableauDeNotas(RsCompsants!NUMNOTA)
                         
                         NUMNOTA = RsCompsants!NUMNOTA
                    End If
                    CollectionNota.Add RsCompsants("NUMNOTA").Value, Trim("N" & RsCompsants!NUMNOTA)
                End If
 
    
    RsCompsants.MoveNext
 Wend
   
  RsCompsants.Requery
   XMin = -1498.3061: YMin = 482.3797
    For i = 0 To IndexIstC
  
    InsertPointConnecteur(i).InsertPointConnecteur(0) = XMin - (600 * i): InsertPointConnecteur(i).InsertPointConnecteur(1) = YMin + (600 * i): InsertPointConnecteur(i).InsertPointConnecteur(2) = 0
    Next i
 On Error GoTo GesERR
 
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = 1 + NUMNOTA
 FormBarGrah.ProgressBar1Caption.Caption = " Chargement des Notas"
 
 While RsCompsants.EOF = False
  IncremanteBarGrah FormBarGrah
 If TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).NotasExiste = False Then
 rr = InsertPointHiérarchie(Val(Trim("" & RsCompsants!NUMNOTA)), -1300, -800, -600, 600, 4)
  InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(1), -600)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).InsertPointLigneC(0) = rr(0)
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).InsertPointLigneC(1) = rr(1)
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).InsertPointLigneC(2) = rr(2)
   
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).XScaleFactorC = 1
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).YScaleFactorC = 1
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).ZScaleFactorC = 1
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).RotationC = 0
 End If

'    PathNotasDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\Nota"
    Lib1 = PathNotasDefault & "\" & RsCompsants!Nota & ".dwg"
    Lib2 = "" & RsCompsants!Nota
    NumErr = 6
    Set TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).NewBlock = FunInsBlock(PathNotasDefault & "\" & RsCompsants!Nota & ".dwg", TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).InsertPointLigneC, "", TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).RotationC, TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).XScaleFactorC, TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).YScaleFactorC, TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).ZScaleFactorC)
    Set TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).Attribues = ColectionAttribueConecteur(TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).NewBlock.GetAttributes)

                Att = TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).NewBlock.GetAttributes
                
                NumErr = 7
                
                 Lib1 = "NUMNOTA"
                Lib2 = "" & RsCompsants!Nota
                Att(TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).Attribues("NUMNOTA")).TextString = "" & RsCompsants!NUMNOTA
                Err.Clear
'                 Lib1 = "NOTA"
'                Lib2 = ""
'                Att(TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).Attribues("NOTA")).TextString = "" & RsCompsants!NOTA
    RsCompsants.MoveNext
 Wend
 Exit Function
GesERR:
    FunError NumErr, "" & Lib1, Err.Description, "" & Lib2
 Resume Next
End Function

Public Function LoadConnecteur(IdIndiceProjet As Long, MyType As String) As Boolean
If (bool_Plan_E_Connecteurs = False And MyType = "PL") Or (bool_Outil_E_Connecteurs = False And MyType = "OU") Then Exit Function
    LoadConnecteur = False
    Dim RsConnecteur As Recordset
    Dim Sql As String
    Dim Myrep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
  
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
   Dim XMin As Double
   Dim YMin As Double
    Dim PathConnecteursDefault As String
  PathBlocs = TableauPath.Item("PathBlocs")
 If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
 If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)
    Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT T_Clients.Client, T_Clients.PathConnecteurs FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathConnecteurs) = "" Then
         PathConnecteursDefault = TableauPath.Item("PathConnecteursDefault")
   Else
             PathConnecteursDefault = RsConnecteur!PathConnecteurs
         If Left(PathConnecteursDefault, 2) <> "\\" And Left(PathConnecteursDefault, 1) = "\" Then PathConnecteursDefault = TableauPath.Item("PathServer") & PathConnecteursDefault
         If Right(PathConnecteursDefault, 2) = "\\" Then PathConnecteursDefault = Mid(PathConnecteursDefault, 1, Len(PathConnecteursDefault) - 1)
    
    End If
Else
                 PathConnecteursDefault = TableauPath.Item("PathConnecteursDefault")

End If
If Left(PathConnecteursDefault, 2) <> "\\" And Left(PathConnecteursDefault, 1) = "\" Then PathConnecteursDefault = TableauPath.Item("PathServer") & PathConnecteursDefault
If Right(PathConnecteursDefault, 2) = "\\" Then PathConnecteursDefault = Mid(PathConnecteursDefault, 1, Len(PathConnecteursDefault) - 1)



 Sql = "SELECT Connecteurs.CONNECTEUR, Connecteurs.[O/N],  "
Sql = Sql & "Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.PRECO1,  "
Sql = Sql & "Connecteurs.PRECO2  "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & " and Connecteurs.ACTIVER=true "
    Sql = Sql & "ORDER BY Connecteurs.N°;"
    NumErr = 1
    If MyType = "OU" Then
        XMin = 1185.771
        YMin = 1667.3509
    Else
        XMin = 30
        YMin = 870.4179
    End If
    Set RsConnecteur = Con.OpenRecordSet(Sql)
    InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 870.4179: InsertPointLigneTableau_Vignette(2) = 0
    For i = 0 To IndexIstC
  
    InsertPointConnecteur(i).InsertPointConnecteur(0) = XMin + (150 * i): InsertPointConnecteur(i).InsertPointConnecteur(1) = YMin + (150 * i): InsertPointConnecteur(i).InsertPointConnecteur(2) = 0
    Next i
    i = 1
      NbConnecteur = 0
    For i = 1 To CollectionCon.Count
      If NbConnecteur < CollectionCon(i) Then
        NbConnecteur = CollectionCon(i)
      End If
    Next i
    i = 1
    While RsConnecteur.EOF = False
   
    If Trim(UCase("" & RsConnecteur.Fields(0))) <> "NEANT" Then
   
        On Error Resume Next
        NbCol = CLng(CollectionCon("" & RsConnecteur.Fields(3)))
        If Err Then
            Err.Clear
            
           NbConnecteur = NbConnecteur + 1
                CollectionCon.Add NbConnecteur, Trim("" & RsConnecteur.Fields(3))
            
        End If
         
         On Error GoTo 0
     Else
        NbConnecteur = NbConnecteur + 1
     End If
     
  
       RsConnecteur.MoveNext
    Wend
  
    ReDim Preserve TableauDeConnecteurs(NbConnecteur)
     FormBarGrah.ProgressBar1.Value = 0
    If NbConnecteur = 0 Then
     FormBarGrah.ProgressBar1.Max = 1 + 1
    Else
         FormBarGrah.ProgressBar1.Max = 1 + NbConnecteur
    End If
     FormBarGrah.ProgressBar1Caption.Caption = " Chargement des connecteurs"
    If NbConnecteur <> 0 Then
        RsConnecteur.Requery
    End If
      On Error GoTo GesERR
      FormBarGrah.ProgressBar1.Value = 0
    While RsConnecteur.EOF = False
        If FormBarGrah.ProgressBar1.Max = FormBarGrah.ProgressBar1.Value Then
            FormBarGrah.ProgressBar1.Max = FormBarGrah.ProgressBar1.Max + 1
        End If
         IncremanteBarGrah FormBarGrah
        DoEvents
        
        DoEvents


   
    
        If UCase("" & RsConnecteur.Fields(0)) <> "NEANT" Then
        Debug.Print PathConnecteursDefault & "\" & RsConnecteur.Fields(0) & ".dwg"
            If Fso.FileExists(PathConnecteursDefault & "\" & RsConnecteur.Fields(0) & ".dwg") = True Then
                Myrep = PathConnecteursDefault
                Trouve = True
                NumErr = 4
              
                
                TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = True
            Else
                NumErr = 1
                TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = False
                Myrep = ""
                
GesERR:
                Trouve = False
                FunError NumErr, "" & RsConnecteur.Fields(3), Err.Description, "" & RsConnecteur.Fields(0)
              
            End If
            If Trouve = True Then
            If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock = FunInsBlock(Myrep & "\" & RsConnecteur.Fields(0) & ".dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneC, "", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorC)
            Else
                rr = InsertPointHiérarchie(Val(Trim("" & RsConnecteur.Fields(4))), 0, 2000, 200, 300, 10)
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock = FunInsBlock(Myrep & "\" & RsConnecteur.Fields(0) & ".dwg", rr, "")
                InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0), 300)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
             End If
                  If ErrInsert = True Then GoTo EnrSuinant
                If UCase("" & RsConnecteur.Fields(1)) = True Then
                    TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE = True
                    If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then
                     If (bool_Plan_E_Etiquettes = True And MyType = "PL") Or (bool_Outil_E_Vignettes = True And MyType = "OU") Then
                        Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(Myrep & "\EPISSURES.dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneV, "V" & "" & RsConnecteur.Fields(4), TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorV)
                    End If
                    Else
                         If (bool_Plan_E_Etiquettes = True And MyType = "PL") Or (bool_Outil_E_Vignettes = True And MyType = "OU") Then
                            rr = InsertPointHiérarchie(Val(Trim("" & RsConnecteur.Fields(4))), -100, 2000, -135, 400, 10)
                        Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(Myrep & "\EPISSURES.dwg", rr, "V" & "" & RsConnecteur.Fields(4))
                         
'                        InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(rr, 100)
                        NbLignesVignette = NbLignesVignette + 1
                        End If
                    End If
                Else
                    If (bool_Plan_E_Etiquettes = True And MyType = "PL") Or (bool_Outil_E_Vignettes = True And MyType = "OU") Then

                        TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE = False
                        If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then
                            Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(PathBlocs & "\VIGNETTE CONNECTEUR.dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneV, "V" & "" & RsConnecteur.Fields(4), TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorV)
                        Else
                             rr = InsertPointHiérarchie(Val(Trim("" & RsConnecteur.Fields(4))), -100, 2000, -135, 400, 10)
                              
                            Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(PathBlocs & "\VIGNETTE CONNECTEUR.dwg", rr, "V" & "" & RsConnecteur.Fields(4))
                        End If
                            
                            InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 135)
                            NbLignesVignette = NbLignesVignette + 1

                    End If
                End If
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes)
                At = TableauAtribCon(RsConnecteur, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE)
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues

                If (bool_Plan_E_Etiquettes = True And MyType = "PL") Or (bool_Outil_E_Vignettes = True And MyType = "OU") Then
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes)
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE
                End If
                
                
            End If
        End If
        If (bool_Plan_E_Etiquettes = True And MyType = "PL") Or (bool_Outil_E_Vignettes = True And MyType = "OU") Then
        If NbLignesVignette = 11 Then
            InsertPointLigneTableau_Vignette(0) = -1146.1429
            InsertPointLigneTableau_Vignette(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_Vignette(1), -400)
            NbLignesVignette = 0
         End If
        End If
EnrSuinant:
       RsConnecteur.MoveNext
        i = i + 1
    Wend
    LoadConnecteur = True
    Set Fso = Nothing
End Function
Function TableauAtribCon(MyAtrib As Recordset, EPISSURE As Boolean)
    Dim TabAt() As String
    ReDim TabAt(MyAtrib.Fields.Count)
    For Col = 0 To MyAtrib.Fields.Count - 1
    DoEvents
        If (Col = 0) And (EPISSURE = True) Then
            TabAt(Col) = "EPISSURE"
        Else
            TabAt(Col) = "" & MyAtrib.Fields(Col)
        End If
    Next Col
    TableauAtribCon = TabAt
End Function



Public Function PRECO(Var As String, Optional Iis As Boolean) As String
PRECO = Var
PRECO = Replace(UCase(PRECO), "CODE.APP", "CODE_APP")

If InStr(1, UCase(PRECO), "PRECO") <> 0 Then
    PRECO = "PRECO" & Right(PRECO, 1)
    
End If
If Iis = True Then

If (InStr(1, UCase(Var), "PRECO") <> 0) And (InStr(1, UCase(Var), "1") <> 0) Then
    PRECO = "PRECO1"
    
End If
If (InStr(1, UCase(Var), "PRECO") <> 0) And (InStr(1, UCase(Var), "2") <> 0) Then
    PRECO = "PRECO2"
    
End If
    If InStr(1, UCase(PRECO), "FIL") <> 0 Then
        PRECO = "FIL"
    
    End If
    
    If InStr(1, UCase(Var), "_FILS") <> 0 Then
        PRECO = "_FILS"
        
    End If
End If
End Function

Public Function DecodeCode_APP(Code_APP As String) As String
Dim SplitCode_APP
Dim NbUbound As Long
SlitCode_APP = Split(Code_APP, ".")
NbUbound = UBound(SlitCode_APP)
Select Case NbUbound
Case -1
    DecodeCode_APP = vbNullChar
Case 0
    DecodeCode_APP = SlitCode_APP(0)
Case 1
    DecodeCode_APP = SlitCode_APP(0)
Case 2
    DecodeCode_APP = SlitCode_APP(1)
Case Else
     DecodeCode_APP = SlitCode_APP(1)
    End Select
End Function
Sub test()
a = DecodeCode_APP("exzf_120_aa_ee")

a = DecodeCode_APP("exzf_120_aa")
a = DecodeCode_APP("120_aa")
a = DecodeCode_APP("120")
a = DecodeCode_APP("")

End Sub
Public Function funPath()
    Dim MyPath As New Collection
    Dim Rs As Recordset
    Set Rs = Con.OpenRecordSet("SELECT T_Path.* FROM T_Path;")
    While Rs.EOF = False
        MyPath.Add Rs.Fields("PathVar").Value, Rs.Fields("NameVar").Value
        Rs.MoveNext
    Wend
    Set Rs = Con.CloseRecordSet(Rs)
    Set funPath = MyPath
End Function
Function ColectionAttribueConecteur(Attribues) As Collection
    Dim MyAttribue As New Collection
    Dim IndexAt As Long


    IndexAt = 0
    On Error Resume Next
    For IndexAt = 0 To UBound(Attribues)
        
        Debug.Print UCase(Replace(Replace(UCase(Attribues(IndexAt).TagString), "PRECO", "PRECO."), "PRECO..", "PRECO."))
        MyAttribue.Add IndexAt, UCase(Replace(Replace(UCase(Attribues(IndexAt).TagString), "PRECO", "PRECO."), "PRECO..", "PRECO."))
        Set Atr = Nothing
        
     Next
    On Error GoTo 0
    Set ColectionAttribueConecteur = New Collection
    Set ColectionAttribueConecteur = MyAttribue
End Function


Function FunEPISSURE(Attribues, Fil, Valeur, Connecteur As Long) As Boolean
    FunEPISSURE = False
    Dim bollInDif As Boolean
    Dim IbAttribue As Long
    Dim Fils As String
    Dim TouveFil As Boolean
    Dim boolNotExecute As Boolean
    Dim booD As Boolean
    Dim booG As Boolean
    On Error GoTo Fin
Valeur = Valeur & Space(50)
    bollInDif = True
    Fils = "FILG"
    If UCase(Left(Valeur, 1)) = "D" Then
        boolNotExecute = True
        booD = True
    End If
    If UCase(Left(Valeur, 1)) = "G" Then
       
        booG = True
    End If
    Valeur = Trim(Valeur)
    For i = 1 To UBound(Attribues)
        DoEvents
        
        IbAttribue = TableauDeConnecteurs(Connecteur).Attribues.Item(Fils & CStr(i))
        If (Trim("" & Attribues(IbAttribue).TextString) = "") And (boolNotExecute = False) Then
        Exit For
        End If
Retour:
    Next i
    Attribues(IbAttribue).TextString = Fil

    On Error GoTo 0
    FunEPISSURE = True
    Exit Function
Fin:

    If Fils = "FILG" Then
    If booG = True Then
        boolNotExecute = True
    Else
        boolNotExecute = False
     End If
        Fils = "FILD"
        i = 0
        Err.Clear
        GoTo Retour
    End If
    Err.Clear
End Function

 Public Function AtrbNumError() As Long
    Dim Sql As String
    Dim NErr As Long
    Dim RsNumError As Recordset
    Sql = "SELECT T_NumErreur.LibErreur, T_NumErreur.NumErreur "
    Sql = Sql & "FROM T_NumErreur "
    Sql = Sql & "WHERE T_NumErreur.LibErreur='ErrorApp';"
    Set RsNumError = Con.OpenRecordSet(Sql)
    If RsNumError.EOF = False Then
        Sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1;"
        Con.Exequte Sql
        RsNumError.Requery
        AtrbNumError = RsNumError!NumErreur
    End If
End Function
Public Function VersionPices(Pieces As String) As Long
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT  VersionPices.Version FROM VersionPices "
Sql = Sql & "WHERE VersionPices.Pi='" & MyReplace(Pieces) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
Sql = "INSERT INTO VersionPices ( Pi ) VALUES('" & MyReplace(Pieces) & "');"
Con.Exequte Sql
End If
Sql = "UPDATE VersionPices SET VersionPices.Version = [Version] + 1 "
Sql = Sql & "WHERE VersionPices.Pi='" & MyReplace(Pieces) & "';"
Con.Exequte Sql
Rs.Requery
VersionPices = Rs!Version
End Function
Public Function KilVersionXX(Version As String, Archive As String, Optional Kill As Boolean) As Boolean
Dim Fso As New FileSystemObject

Archive2 = Archive
Version2 = Version

Reprise:
SaveVersion2 = Version2
SaveArchive2 = Archive2
a = Split(Version2, "\")
Version2 = ""
For i = LBound(a) To UBound(a) - 1
Version2 = Version2 & a(i) & "\"
Next i
Version2 = Left(Version2, Len(Version2) - 1)



a = Split(Archive2, "\")
Archive2 = ""
For i = LBound(a) To UBound(a) - 1
Archive2 = Archive2 & a(i) & "\"
Next i
Archive2 = Left(Archive2, Len(Archive2) - 1)
Debug.Print Version2
Debug.Print Archive2
Debug.Print SaveVersion2
Debug.Print SaveArchive2
If Version2 <> Archive2 Then GoTo Reprise
If Kill = True Then
    If SaveVersion2 <> SaveArchive2 Then Fso.DeleteFolder SaveVersion2, True
End If

Set Fso = Nothing
KilVersionXX = True
End Function
Public Function funCloseDatabase()
Con.CloseConnection
'ConBaseNum.CloseConnection

End Function
Public Function ReseingeTor(CodeApp As String, InsertTorTitre) As Boolean
On Error GoTo Fin
   Dim PathTorDefault As String
 PathTorDefault = TableauPath.Item("PathTorDefault")
If TableuDeTor(CollectionTor(CodeApp)).Garder = False Then Exit Function
If Left(PathTorDefault, 2) <> "\\" And Left(PathTorDefault, 1) = "\" Then PathTorDefault = TableauPath.Item("PathServer") & PathTorDefault

If Right(PathTorDefault, 2) = "\\" Then PathTorDefault = Mid(PathTorDefault, 1, Len(PathTorDefault) - 1)

If TableuDeTor(CollectionTor(CodeApp)).TorExiste = False Then _
    TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(0) = Val(InsertTorTitre(0)): TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(1) = Val(InsertTorTitre(1)): TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(2) = Val(InsertTorTitre(2))
        Set TableuDeTor(CollectionTor(CodeApp)).NewBlockTorTire = FunInsBlock(PathTorDefault & "\TORDESIGNATION.dwg", TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre, "", TableuDeTor(CollectionTor(CodeApp)).Rotation, TableuDeTor(CollectionTor(CodeApp)).XScaleFactor, TableuDeTor(CollectionTor(CodeApp)).YScaleFactor, TableuDeTor(i).ZScaleFactor)
            Set TableuDeTor(CollectionTor(CodeApp)).Attribues = ColectionAttribueConecteur(TableuDeTor(CollectionTor(CodeApp)).NewBlockTorTire.GetAttributes)
            a = TableuDeTor(CollectionTor(CodeApp)).NewBlockTorTire.GetAttributes
           a(TableuDeTor(CollectionTor(CodeApp)).Attribues("TORDESIGNATION")).TextString = TableuDeTor(CollectionTor(CodeApp)).CodeApp
           For i = 1 To UBound(TableuDeTor(CollectionTor(CodeApp)).Tor)
           
               If TableuDeTor(CollectionTor(CodeApp)).Tor(i).Garder = True Then
               If TableuDeTor(CollectionTor(CodeApp)).Tor(i).TorExiste = False Then
                    If i = 1 Then
                    
                        TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert(0) = TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(0)
                        TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert(1) = TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(1)
                        TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert(2) = TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(2)
                    Else
                        TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert(0) = TableuDeTor(CollectionTor(CodeApp)).Tor(i - i).Insert(0)
                        TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert(1) = TableuDeTor(CollectionTor(CodeApp)).Tor(i - i).Insert(1)
                        TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert(2) = TableuDeTor(CollectionTor(CodeApp)).Tor(i - i).Insert(2)
                    End If
                    TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert(1) = DecalInsertPointLigneTableau_fils_Bas(TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert(1), 5)
              End If
             Set TableuDeTor(CollectionTor(CodeApp)).Tor(i).NewBlockTorDetail = FunInsBlock(PathTorDefault & "\TORDETAIL.dwg", TableuDeTor(CollectionTor(CodeApp)).Tor(i).Insert, "", TableuDeTor(CollectionTor(CodeApp)).Tor(i).Rotation, TableuDeTor(i).Tor(i).XScaleFactor, TableuDeTor(CollectionTor(CodeApp)).YScaleFactor, TableuDeTor(CollectionTor(CodeApp)).Tor(i).ZScaleFactor)
            Set TableuDeTor(CollectionTor(CodeApp)).Attribues = ColectionAttribueConecteur(TableuDeTor(CollectionTor(CodeApp)).Tor(i).NewBlockTorDetail.GetAttributes)
            a = TableuDeTor(CollectionTor(CodeApp)).Tor(i).NewBlockTorDetail.GetAttributes
             a(TableuDeTor(CollectionTor(CodeApp)).Attribues("TORDESIGNATION")).TextString = "" & TableuDeTor(CollectionTor(CodeApp)).CodeApp
            a(TableuDeTor(CollectionTor(CodeApp)).Attribues("TORFILS")).TextString = "" & TableuDeTor(CollectionTor(CodeApp)).Tor(i).TableauFile
            a(TableuDeTor(CollectionTor(CodeApp)).Attribues("TORNUM")).TextString = "" & TableuDeTor(CollectionTor(CodeApp)).Tor(i).TorName
           End If
        Next i
Fin:
Err.Clear
End Function
Public Function funOpenDatabase()

Con.OpenConnetion db
'ConBaseNum.OpenConnetion DbNumPlan
End Function
'Public Function ScanServeur()
'Dim Fso As New FileSystemObject
'Dim PathFavorisreseau As String
'Dim FileNumber As Long
'
'Dim Txt
'txt2 = ""
'i = 0
'FileNumber = FreeFile
'PathFavorisreseau = "c:\Favoris_reseau.Bat"
'Open PathFavorisreseau For Output As FileNumber   ' Ouvre le fichier en lecture.
'    Print #FileNumber, "net View > %1"
'    Print #FileNumber, "cls"
'    Print #FileNumber, "Dir vue.Txt"
'    Print #FileNumber, "pause"
'    Close #FileNumber
'Close #FileNumber
'PathFavorisreseau2 = Environ("USERPROFILE") & "\vue.txt"
'
'Shell PathFavorisreseau & " " & PathFavorisreseau2
'Open PathFavorisreseau For Input As FileNumber   ' Ouvre le fichier en lecture.
'Do While Not EOF(FileNumber)
'i = i + 1 ' Effectue la boucle jusqu'à la fin du fichier.
'    Input #FileNumber, Txt    ' Lit les données dans deux variables.
'    Debug.Print Txt
'    If i > 3 Then
'    If Trim(Txt) = "La commande s'est termin‚e correctement." Then Exit Do
'    txt2 = Mid(Txt, 1, Len(Txt) - InStr(1, Txt, " ")) & vbCrLf
'    Debug.Print txt2
'    End If
'    ' Affiche les données dans la fenêtre Exécution.
'  Loop
'Close #FileNumber   ' Ferme le fichier.
'ScanServeur = Split(txt2, vbCrLf)
'Set Fso = Nothing
'End Function

Public Function ScanServeur()
Dim Fso As New FileSystemObject
Dim PathFavorisreseau As String
Dim FileNumber As Long
Dim txt As String
Dim txt2 As String

txt2 = ""
i = 0
FileNumber = FreeFile
PathFavorisreseau = Environ("USERPROFILE") & "\Favoris_reseau"
If Fso.FolderExists(PathFavorisreseau) = False Then
    Fso.CreateFolder PathFavorisreseau
End If
Open PathFavorisreseau & "\Favoris_reseau.BAT" For Output As FileNumber   ' Ouvre le fichier en lecture.
Print #FileNumber, "net View>" & Chr(34) & PathFavorisreseau & "\Favoris_reseau.Txt" & Chr(34)
'Print #FileNumber, "dir " & PathFavorisreseau & "*.*"
'Print #FileNumber, "PAUSE"
Close #FileNumber

Shell Chr(34) & PathFavorisreseau & "\Favoris_reseau.BAT" & Chr(34)
PathFavorisreseau = PathFavorisreseau & "\Favoris_reseau.Txt"

Open PathFavorisreseau For Input As FileNumber   ' Ouvre le fichier en lecture.
Do While Not EOF(FileNumber)
    Input #FileNumber, txt    ' Lit les données dans deux variables.
    Debug.Print txt
    If InStr(1, txt, Chr(32)) = 0 Then
         txt2 = txt2 & txt & vbCrLf
    Else
        txt2 = txt2 & txt 'Left(Txt, InStr(1, Txt, Chr(32)) - 1) & vbCrLf
    End If
    ' Affiche les données dans la fenêtre Exécution.
  Loop
Close #FileNumber   ' Ferme le fichier.
ScanServeur = Split(txt2, vbCrLf)
Set Fso = Nothing
End Function

Public Sub MajBase(IdIndice As Long)
Dim Sql As String
'***********************************************************************************************************************
'*                                        Supprime les données des tables de travail :                                 *
Sql = "DELETE T_Critères.* "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql

Sql = "DELETE Ligne_Tableau_fils.* "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql

Sql = "DELETE Connecteurs.* "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql

Sql = "DELETE Composants.* "
Sql = Sql & "FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql

Sql = "DELETE Nota.* "
Sql = Sql & "FROM Nota "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql


Sql = "DELETE T_Noeuds.* "
Sql = Sql & "FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql
'***********************************************************************************************************************
'*                                        Enrichie les données des tables de travail :                                 *


Sql = "INSERT INTO T_Critères ( Id_IndiceProjet,ACTIVER,CODE_CRITERE, CRITERES  )  "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet,Xls_Critères.ACTIVER, Xls_Critères.CODE_CRITERE, Xls_Critères.CRITERES  "
Sql = Sql & "FROM Xls_Critères  "
Sql = Sql & "WHERE Xls_Critères.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, ACTIVER,LIAI, DESIGNATION,  "
Sql = Sql & "FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS,  "
Sql = Sql & "[POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2,  "
Sql = Sql & "VOI2, PRECO, [OPTION] ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet,xls_Ligne_Tableau_fils.ACTIVER, xls_Ligne_Tableau_fils.LIAI,  "
Sql = Sql & "xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL,  "
Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,  "
Sql = Sql & "xls_Ligne_Tableau_fils.LONG, xls_Ligne_Tableau_fils.[LONG CP],  "
Sql = Sql & "xls_Ligne_Tableau_fils.COUPE, xls_Ligne_Tableau_fils.POS,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[POS-OUT], xls_Ligne_Tableau_fils.FA,  "
Sql = Sql & "xls_Ligne_Tableau_fils.APP, xls_Ligne_Tableau_fils.VOI,  "
Sql = Sql & "xls_Ligne_Tableau_fils.POS2, xls_Ligne_Tableau_fils.[POS-OUT2],  "
Sql = Sql & "xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.APP2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.VOI2, xls_Ligne_Tableau_fils.PRECO,  "
Sql = Sql & "xls_Ligne_Tableau_fils.OPTION "
Sql = Sql & "FROM xls_Ligne_Tableau_fils "
Sql = Sql & " where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"



Con.Exequte Sql

Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, ACTIVER,CONNECTEUR, [O/N],  "
Sql = Sql & "DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2,  "
Sql = Sql & "[100%] , [OPTION],[Pylone],[Colonne],[Ligne] ) "
Sql = Sql & "SELECT " & IdIndice & "  AS Id_IndiceProjet, Xls_Connecteurs.ACTIVER,Xls_Connecteurs.CONNECTEUR,  "
Sql = Sql & "Xls_Connecteurs.[O/N], Xls_Connecteurs.DESIGNATION,  "
Sql = Sql & "Xls_Connecteurs.CODE_APP, Xls_Connecteurs.N°,  "
Sql = Sql & "Xls_Connecteurs.POS, Xls_Connecteurs.[POS-OUT],  "
Sql = Sql & "Xls_Connecteurs.PRECO1, Xls_Connecteurs.PRECO2,  "
Sql = Sql & "Xls_Connecteurs.[100%] , Xls_Connecteurs.OPTION,Xls_Connecteurs.[Pylone],Xls_Connecteurs.[Colonne],Xls_Connecteurs.[Ligne] "
Sql = Sql & "FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"

Con.Exequte Sql



Sql = "INSERT INTO Composants ( Id_IndiceProjet, ACTIVER,DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Composants.ACTIVER,Xls_Composants.DESIGNCOMP, Xls_Composants.NUMCOMP,   "
Sql = Sql & "Xls_Composants.REFCOMP, Xls_Composants.Path   "
Sql = Sql & "FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"

Con.Exequte Sql

Sql = "INSERT INTO Nota ( Id_IndiceProjet,ACTIVER, NOTA, NUMNOTA ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Nota.ACTIVER,Xls_Nota.NOTA, Xls_Nota.NUMNOTA "
Sql = Sql & "FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"

Con.Exequte Sql

Sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet,Fleche_Droite, ACTIVER, NŒUDS,LONGUEUR,DESIGN_HAB, "
Sql = Sql & "CODE_RSA,CODE_PSA,CODE_ENC,DIAMETRE,CLASSE_T,TORON_PRINCIPAL, LONGUEUR_CUMULEE) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Noeuds.Fleche_Droite, Xls_Noeuds.ACTIVER, "
Sql = Sql & "Xls_Noeuds.NŒUDS,Xls_Noeuds.LONGUEUR,Xls_Noeuds.DESIGN_HAB,Xls_Noeuds.CODE_RSA, "
Sql = Sql & "Xls_Noeuds.CODE_PSA,Xls_Noeuds.CODE_ENC,Xls_Noeuds.DIAMETRE,Xls_Noeuds.CLASSE_T,Xls_Noeuds.TORON_PRINCIPAL, "
Sql = Sql & "Xls_Noeuds.LONGUEUR_CUMULEE  "
Sql = Sql & "FROM Xls_Noeuds "
Sql = Sql & "WHERE Xls_Noeuds.Job=" & NmJob & ";"

Con.Exequte Sql


Sql = "DELETE Xls_Critères.*  FROM Xls_Critères "
Sql = Sql & "where Xls_Critères.Job=" & NmJob & ";"
Con.Exequte Sql


 Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Composants.*  FROM Xls_Composants "
Sql = Sql & "where Xls_Composants.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Nota.*  FROM Xls_Nota "
Sql = Sql & "where Xls_Nota.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Exequte Sql

'***********************************************************************************************************************
'*                                        Attribut les code appareil au tableau de fils :                              *

Sql = "UPDATE (Ligne_Tableau_fils LEFT JOIN Connecteurs ON Ligne_Tableau_fils.FA = Connecteurs.N°)  "
Sql = Sql & "LEFT JOIN Connecteurs AS Connecteurs_1 ON Ligne_Tableau_fils.FA2 = Connecteurs_1.N°  "
Sql = Sql & "SET Ligne_Tableau_fils.APP = [Connecteurs].[CODE_APP], Ligne_Tableau_fils.APP2 = [Connecteurs_1].[CODE_APP] "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & " "
Sql = Sql & "AND Connecteurs.Id_IndiceProjet=" & IdIndice & " "
Sql = Sql & "AND Connecteurs_1.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql
'***********************************************************************************************************************

End Sub
Sub MiseEnPage(MyWorksheet As Worksheet, Myrange As Range, MyLeftHeader As String, _
            MyCenterHeader As String, MyRightHeader As String, MyLeftFooter As String, _
            MyCenterFooter As String, MyRightFooter As String, _
            MyZoom, CellVolet As String, RepeatCol As Boolean, MyxlLandscape As Long, _
            Optional AutoFilterOk As Boolean, Optional NotCouleur As Boolean, Optional MergeOk As Boolean, _
            Optional BottomMargin As Double = 2.5, Optional AutoFit As Boolean = True, Optional ZoneImpression As Boolean = True)
'            MyWorksheet.Application.Visible = True
'
            MyWorksheet.Select
          If Trim(CellVolet) <> "" Then
  MyWorksheet.Range(CellVolet).Select
  End If
            If AutoFilterOk = True Then Myrange.AutoFilter
With MyWorksheet.Range(Myrange(1, 1).Address & ":" & Myrange(1, Myrange.Columns.Count).Address)
        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlContext
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = MergeOk
    End With
    If NotCouleur = False Then _
    MyWorksheet.Range(Myrange(1, 1).Address & ":" & Myrange(1, Myrange.Columns.Count).Address).Interior.ColorIndex = 15
    
    If AutoFit = True Then
        Myrange.ColumnWidth = 120
        Myrange.RowHeight = 120
        Myrange.EntireColumn.AutoFit
        Myrange.EntireRow.AutoFit
    End If
'  MyWorksheet.Application.Visible = True
 If Trim(CellVolet) <> "" Then
  MyWorksheet.Application.ActiveWindow.FreezePanes = True
  End If
           With MyWorksheet.PageSetup
            If ZoneImpression = True Then
                .PrintArea = "A1:" & Myrange(Myrange.Rows.Count, Myrange.Columns.Count).Address
            End If
        .LeftHeader = MyLeftHeader
        .CenterHeader = "&""Arial,Gras""&20&A&""Arial,Normal""&10" & MyCenterHeader
        .RightHeader = MyRightHeader
       
        .TopMargin = MyWorksheet.Application.InchesToPoints(2)
         .LeftMargin = MyWorksheet.Application.InchesToPoints(0.196850393700787)
        .RightMargin = MyWorksheet.Application.InchesToPoints(0.196850393700787)
        .TopMargin = MyWorksheet.Application.InchesToPoints(1.37795275590551)
        .BottomMargin = MyWorksheet.Application.InchesToPoints(BottomMargin / 2.54)  '0.984251968503937)
        .HeaderMargin = MyWorksheet.Application.InchesToPoints(0.196850393700787)
        aa = 0.5 / 2.54
        Debug.Print aa
          .FooterMargin = MyWorksheet.Application.InchesToPoints(0.196850393700787)  '0.984251968503937)
        
        .LeftFooter = MyLeftFooter
        .CenterFooter = MyCenterFooter
        .RightFooter = MyRightFooter
        .Orientation = MyxlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = MyZoom
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
         .CenterHorizontally = True
          .PrintGridlines = False
    End With
         
            MyWorksheet.PageSetup.PrintTitleRows = MyWorksheet.Range(Myrange(1, 1).Address & ":" & Myrange(1, Myrange.Columns.Count).Address).Address
       
        If RepeatCol = True Then _
        MyWorksheet.PageSetup.PrintTitleColumns = MyWorksheet.Range(Myrange(1, 1).Address & ":" & Myrange(1, 1).Address).Address
           
           
           
           
   
End Sub


Public Function AtocatOption(Id_Pieces As Long) As Boolean
If (bool_Plan_E_Options = False And MyType = "PL") Or (bool_Outil_E_Options = False And MyType = "OU") Then Exit Function

Dim Rs As Recordset
Dim RsSelect As Recordset
Dim Sql As String
Dim Index As Long
Dim MyPotionEntete As Collection
Dim Block As AcadBlockReference

AtocatOption = False
Index = 0
Sql = "SELECT T_indiceProjet.*, T_indiceProjet.Id "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Function
Sql = "SELECT T_indiceProjet.*, T_indiceProjet.Id "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Id = " & Id_Pieces & " "
Sql = Sql & "Or T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
AtocatOption = True
While Rs.EOF = False
    Index = Index + 1
    Rs.MoveNext
Wend
  Rs.Requery
   FormBarGrah.ProgressBar1Caption = " Chargement des Options :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = Index * 3
Set MyPotionEntete = New Collection
InsertPointLigneTableau_fils(0) = -1168#: InsertPointLigneTableau_fils(1) = 20#: InsertPointLigneTableau_fils(2) = 0
 InsertPointLigneTableau_fils2(1) = 20#: InsertPointLigneTableau_fils2(2) = 0
InsertPointLigneTableau_fils2(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(0), -24.6369)
While Rs.EOF = False
IncremanteBarGrah FormBarGrah
Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils, "", 0, 0)
 InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), -3)
 aa = Block.GetAttributes
 aa(0).TextString = UCase(Trim("" & Rs!Pi) & "_" & Trim("" & Rs!PI_Indice))
 
 Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils2, "", 0, 0)
 InsertPointLigneTableau_fils2(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils2(1), -3)
 aa = Block.GetAttributes
 aa(0).TextString = UCase(Trim("" & Rs!RefPieceClient))
  Rs.MoveNext
Wend
 Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils, "", 0, 0)
 InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(0), -24.6369)
  Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils2, "", 0, 0)
' Set Block = FunInsBlock("\\10.30.0.5\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\RefOption.dwg", InsertPointLigneTableau_fils2, "", 0, 0)
 InsertPointLigneTableau_fils2(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils2(0), -24.6369)
  aa = Block.GetAttributes
 aa(0).TextString = UCase(Trim("ref client"))
    Rs.Requery
 On Error Resume Next
 While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
 az = Split(Trim("" & Rs!Equipement), ";")
    For i = LBound(aa) To UBound(az) - 1
    aa = MyPotionEntete(Trim("" & az(i)))
    If Err Then
        Err.Clear
'        1134.4342
   
        MyPotionEntete.Add Trim("" & az(i)), Trim("" & az(i))
            Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils2, "", 0, 0)
            InsertPointLigneTableau_fils2(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils2(0), -24.6369)
            aa = Block.GetAttributes
            aa(0).TextString = UCase(Trim("" & az(i)))
    End If
    
    Next
    Rs.MoveNext
Wend
InsertPointLigneTableau_fils(0) = -1168#: InsertPointLigneTableau_fils(1) = 20#: InsertPointLigneTableau_fils(2) = 0
    Rs.Requery
 
While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
 DoEvents
InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(DecalInsertPointLigneTableau_fils_Bas(-1168#, -24.6369), -24.6369)
    aa = MyPotionEntete(Trim("" & Rs!Equipement))
   For i = 1 To MyPotionEntete.Count
   Sql = "SELECT  T_indiceProjet.* FROM T_indiceProjet "
    Sql = Sql & "WHERE ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!Pi) & "_" & Trim("" & Rs!PI_Indice) & "' "
    Sql = Sql & "AND T_indiceProjet.Equipement='" & MyReplace(MyPotionEntete(i)) & "'"
    Sql = Sql & "AND T_indiceProjet.Id=" & Id_Pieces & ") "
    Sql = Sql & "OR ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!Pi) & "_" & Trim("" & Rs!PI_Indice) & "'"
    Sql = Sql & "AND T_indiceProjet.Equipement='" & MyReplace(MyPotionEntete(i)) & "'"
    Sql = Sql & "AND T_indiceProjet.Pere=" & Id_Pieces & ");"
    
    
    
    Sql = "SELECT T_indiceProjet.Id "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!Pi) & "_" & Trim("" & Rs!PI_Indice) & "' "
    Sql = Sql & "AND  [Equipement]  Like '%" & MyReplace(MyPotionEntete(i)) & ";%' "
    Sql = Sql & "AND T_indiceProjet.Id=" & Id_Pieces & ") "
    Sql = Sql & "OR ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!Pi) & "_" & Trim("" & Rs!PI_Indice) & "' "
    Sql = Sql & "AND  [Equipement]  Like '%" & MyReplace(MyPotionEntete(i)) & ";%' "
    Sql = Sql & "AND T_indiceProjet.Pere=" & Id_Pieces & ");"

Set RsSelect = Con.OpenRecordSet(Sql)
       
            Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils, "", 0, 0)
            InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(0), -24.6369)
            aa = Block.GetAttributes
            If RsSelect.EOF = False Then
            
            aa(0).TextString = "X"
            Else
             aa(0).TextString = ""
            End If
    
    Next
    InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), -3)
    Rs.MoveNext
Wend
'MyPotionEntete.Add Block, i1
End Function
Public Sub Racourci(RaccourciName As String, RaccourciCible As String, Extension As String)
Dim Fso As New FileSystemObject
If Fso.FileExists(RaccourciName & ".Lnk") = True Then
     Fso.DeleteFile RaccourciName & ".Lnk"
End If
Set objshell = CreateObject("wscript.shell")
Set objraccourci = objshell.createshortcut(RaccourciName & ".Lnk")
objraccourci.targetpath = RaccourciCible & "." & Extension
objraccourci.Save
Set Fso = Nothing
Set objraccourci = Nothing
End Sub
Sub IncremanteBarGrah(Obj As Object)
If Obj.ProgressBar1.Max = Obj.ProgressBar1.Value Then
            Obj.ProgressBar1.Max = Obj.ProgressBar1.Max + 1
        End If
         Obj.ProgressBar1.Value = Obj.ProgressBar1.Value + 1
         DoEvents
End Sub

Sub EcritureTor(RsLigne As Recordset, MyType As String)
If (bool_Plan_E_Preconisations = False And MyType = "PL") Or (bool_Outil_E_Preconisations = False And MyType = "OU") Then Exit Sub

 While RsLigne.EOF = False
        If Trim("" & RsLigne!PRECO) <> "" Then
        On Error Resume Next
        Set a = Nothing
        DoEvents
            a = ""
            a = CollectionTor("" & RsLigne!App)
            If Err Then
            Err.Clear
                NUMNTORBLOC = NUMNTORBLOC + 1
Reprise:
                CollectionTor.Add NUMNTORBLOC, Trim("" & RsLigne!App)
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!App)))
            End If
                TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CodeApp = Trim("" & RsLigne!App)
                 TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Garder = True
                If TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CodeApp = "" Then
                   GoTo Reprise
                 End If
                  Set a = Nothing
        DoEvents
                a = TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECO)
                If Err Then
            Err.Clear
                TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).NumTor = TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).NumTor + 1
               TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor.Add TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).NumTor, "" & RsLigne!PRECO
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECO))
            End If
            If InStr(1, TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECO)).TableauFile, "" & RsLigne!Fil & " ") = 0 Then
              TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECO)).TableauFile = TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECO)).TableauFile & RsLigne!Fil & " "
               TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECO)).TorName = "" & RsLigne!PRECO
               TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECO)).Garder = True
           End If
            Set a = Nothing
        DoEvents
           On Error Resume Next
        Set a = Nothing
        DoEvents
        a = ""
            a = CollectionTor("" & RsLigne!app2)
            If Err Then
            Err.Clear
                NUMNTORBLOC = NUMNTORBLOC + 1
                CollectionTor.Add NUMNTORBLOC, Trim("" & RsLigne!app2)
Reprise2:
               
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!app2)))
            End If
                TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CodeApp = Trim("" & RsLigne!app2)
                  TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).Garder = True
                If TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CodeApp = "" Then
                   GoTo Reprise2
                 End If
                  Set a = Nothing
        DoEvents
            a = ""
                a = TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CollectionTor("" & RsLigne!PRECO)
                If Err Then
            Err.Clear
                TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).NumTor = TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).NumTor + 1
               TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CollectionTor.Add TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).NumTor, "" & RsLigne!PRECO
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CollectionTor("" & RsLigne!PRECO))
            End If
            If InStr(1, TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CollectionTor("" & RsLigne!PRECO)).TableauFile, "" & RsLigne!Fil & " ") = 0 Then
                 TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CollectionTor("" & RsLigne!PRECO)).Garder = True
              TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CollectionTor("" & RsLigne!PRECO)).TableauFile = TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CollectionTor("" & RsLigne!PRECO)).TableauFile & RsLigne!Fil & " "
               TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!app2))).CollectionTor("" & RsLigne!PRECO)).TorName = "" & RsLigne!PRECO
           End If
            Set a = Nothing
        DoEvents
            On Error GoTo 0
        End If
        
        RsLigne.MoveNext
    Wend

End Sub
Sub Ecriturefils(RsLigne As Recordset, MyType As String, NbFils As Long)
Dim Fso As New FileSystemObject

If (bool_Plan_E_Fils = True And MyType = "PL") Or (bool_Outil_E_Fils = True And MyType = "OU") Then
    
    If Val(NbFils) <> 0 Then
    RsLigne.MoveFirst
    End If
     FormBarGrah.ProgressBar1.Value = 0
    If Val(NbFils) <> 0 Then
     FormBarGrah.ProgressBar1.Max = 1 + NbFils
    Else
         FormBarGrah.ProgressBar1.Max = 1 + 1
    End If
     FormBarGrah.ProgressBar1Caption.Caption = " Chargement de la liste de fils"

    While RsLigne.EOF = False
         IncremanteBarGrah FormBarGrah
        DoEvents
'        AutoApp.Documents(1).ZoomAll

        ReDim Tableau(RsLigne.Fields.Count)
        If UCase(Trim("" & RsLigne.Fields(0))) <> "SUPPRIMER" Then
            For Col = 0 To RsLigne.Fields.Count - 1
                DoEvents
                Tableau(Col) = "" & RsLigne.Fields(Col)
            Next Col
     
            RenseigneConnecteurBroches RsLigne, MyType
            Row = Row + 1
            If NbLignes = 100 Then
                InsertPointLigneTableau_fils(1) = -76.8274: InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_fils(0), 330) '281.7719)
                NbLignes = 0
            End If

            If NbLignes = 0 Then
                If Fso.FileExists(PathBlocs & "\LIGNE TABLEAU DES FILS_RD.dwg") = False Then
                    MsgBox "err"
                End If
                Set NewBlock = FunInsBlock(PathBlocs & "\LIGNE TABLEAU DES FILS_RD.dwg", InsertPointLigneTableau_fils, "E" & CStr(Row))
                InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), 6.8274)
            End If
            If Fso.FileExists(PathBlocs & "\LIGNES TABLEAU DES FILS.dwg") = False Then
                MsgBox "err"
            End If
            Set NewBlock = FunInsBlock(PathBlocs & "\LIGNES TABLEAU DES FILS.dwg", InsertPointLigneTableau_fils, "L" & CInt(Row))
            a = NewBlock.GetAttributes
            InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), 6.8274)

            NbLignes = NbLignes + 1
            On Error GoTo Error1
            Lib1 = RsLigne.Fields(2)
            Lib2 = "" & RsLigne.Fields("APP")
            TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).indexFile = TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).indexFile + 1
            Lib1 = ""
            Lib2 = ""
            ReDim Preserve TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).TableauFile(TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).indexFile)
            TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).TableauFile(TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).indexFile) = RsLigne.Fields(2)
    
            Lib1 = RsLigne.Fields(2)
            Lib2 = "" & RsLigne.Fields("APP2")
            TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).indexFile = TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).indexFile + 1
            Lib1 = ""
            Lib2 = ""
            ReDim Preserve TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).TableauFile(TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).indexFile)

            TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).TableauFile(TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).indexFile) = RsLigne.Fields(2)
    
            a = RsLigne.Fields(0)
            funAttributesLigne_Tableau_fils NewBlock.Name, NewBlock.GetAttributes, Tableau, RsLigne.Fields.Count - 1
            Else
           'NbSupprim = 'NbSupprim + 1
          End If
          InsertPointLigneCritères(0) = InsertPointLigneTableau_fils(0)
          InsertPointLigneCritères(1) = InsertPointLigneTableau_fils(1)
           InsertPointLigneCritères(2) = InsertPointLigneTableau_fils(2)
          RsLigne.MoveNext
     
        Wend
        If NbFils > 0 Then
            Set NewBlock = FunInsBlock(PathBlocs & "\Nombre_fils.dwg", InsertPointLigneTableau_fils, "N1")
            attri = NewBlock.GetAttributes
            attri(0).TextString = NbFils '- NbSupprim
         End If
       Else
          If Val(NbFils) <> 0 Then
    RsLigne.MoveFirst
    End If
     FormBarGrah.ProgressBar1.Value = 0
    If Val(NbFils) <> 0 Then
     FormBarGrah.ProgressBar1.Max = 1 + NbFils
    Else
         FormBarGrah.ProgressBar1.Max = 1 + 1
    End If
     FormBarGrah.ProgressBar1Caption.Caption = " Chargrment de la liste de fils"

    While RsLigne.EOF = False
         IncremanteBarGrah FormBarGrah
        DoEvents
'        AutoApp.Documents(1).ZoomAll

        ReDim Tableau(RsLigne.Fields.Count)
        If UCase(Trim("" & RsLigne.Fields(0))) <> "SUPPRIMER" Then
          
     
            RenseigneConnecteurBroches RsLigne, MyType
           
           
          RsLigne.MoveNext
        End If
        Wend
       
      
       End If
        SacnConnecteur MyType
        Set Fso = Nothing
        Exit Sub
Error1:
    FunError 3, CStr("" & Lib1), CStr("" & Lib2)
Resume Next
End Sub
Public Function NoeuName(Row As Long)
Dim txt As String
Dim Ofset As Long
Dim NbTour As Long
Dim NbTord As Long
Dim txtColone As Long
Dim txtNuberColone As Long

txt = "AA"
txtColone = Len(txt)
txtNuberColone = Len(txt)
Ofset = 0
NbTour = 0
NbTord = 0


For i = 0 To Row - 3
Reprise:
Mid(txt, txtColone, 1) = Chr(Asc(Mid(txt, txtColone, 1)) + 1)
DoEvents
If Asc(Mid(txt, txtColone, 1)) = 91 Then
Mid(txt, txtColone, 1) = "A"
txtColone = txtColone - 1
If txtColone = 0 Then
    txt = txt & "A"
    txtColone = Len(txt)
Else
    GoTo Reprise
End If

End If
   If txtColone <> Len(txt) Then txtColone = Len(txt)



Next

NoeuName = txt
End Function

Public Sub RazFiltreEditExcel(MySpreadsheet As Object)
Dim Myrange
Set Myrange = MySpreadsheet.ActiveSheet.Range("a1").CurrentRegion

If MySpreadsheet.ActiveSheet.AutoFilterMode = True Then
    For i = 1 To Myrange.Columns.Count
    Set aa = MySpreadsheet.ActiveSheet.AutoFilter.Filters(i).Criteria
    aa.ShowAll = True
       
    Next
    MySpreadsheet.ActiveSheet.AutoFilter.Apply

End If
End Sub
Public Function BackUp(Fichier As String, Optional Li As Boolean, Optional MyPathXlsMoins1 As String) As String
BackUp = MyPathXlsMoins1
If Li = True And Bool_Fichier_Li = True Then Exit Function
BackUp = ""
If Fichier = "" Then Exit Function
Dim Fso As New FileSystemObject
If Fso.FileExists(Fichier) = False Then
    Set Fso = Nothing
    Exit Function
End If
Dim Path
Dim PathAs As String
Dim SaveAs As String
Path = Split(Fichier, "\")
PathAs = ""
For i = LBound(Path) To UBound(Path) - 1
PathAs = PathAs & Path(i) & "\"
Debug.Print PathAs
Next
PathAs = PathAs & "Archives"
Debug.Print PathAs
If Fso.FolderExists(PathAs) = False Then
Fso.CreateFolder PathAs
End If
PathAs = PathAs & "\"
SaveAs = Format(Now, "yyyy-mm-dd-h-m-s_") & Path(UBound(Path))
While Fso.FileExists(PathAs & SaveAs) = True
    SaveAs = Format(Date, "yyyy-mm-dd-h-m-s_") & Path(UBound(Path))
Wend
Debug.Print PathAs & SaveAs
Fso.CopyFile Fichier, PathAs & SaveAs
BackUp = PathAs & SaveAs
If Li = True Then
    Bool_Fichier_Li = True
End If
Set Fso = Nothing
End Function
Public Function IndexationNoeuds(NumNoeud As String) As Long
Dim i As Long
Dim txt As String
txt = "AA"
 i = 2
While txt <> NumNoeud
    i = i + 1
    txt = NoeuName(i)
Wend
IndexationNoeuds = i - 1
End Function

Public Function InsertPointHiérarchie(Num As Long, PoseXinit As Double, PoseYinit As Double, DecalX As Double, DecalY As Double, NbBloc As Long)
Dim InsertPoint(0 To 2) As Double
Dim NbTour As Long
Dim NbTour2 As Long
NbTour = 0
InsertPoint(0) = PoseXinit
InsertPoint(1) = PoseYinit
For i = 1 To Num
    If NbTour = NbBloc Then
    NbTour2 = NbTour2 + 1
       InsertPoint(1) = PoseYinit
       InsertPoint(0) = DecalInsertPointLigneTableau_fils_Gauche(PoseXinit, DecalX * NbTour2)
       NbTour = 0
    End If
     If NbTour <> 0 Then
        InsertPoint(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPoint(0), DecalX)
        InsertPoint(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertPoint(1), DecalY)
       
     End If
    NbTour = NbTour + 1
Next
InsertPointHiérarchie = InsertPoint
End Function
'Public Sub FormatExcelPlage(Plage As Range, Couleur As Long, Merge As Boolean, Grille As Boolean, HorizontalAlignment As Long, VerticalAlignment As Long, Optional ZoneImpressionOfset As Long)
'Plage.Interior.ColorIndex = Couleur
'If Merge = True Then Plage.Merge
'    Plage.HorizontalAlignment = HorizontalAlignment 'xlCenter
'    Plage.VerticalAlignment = VerticalAlignment 'xlCenter
'If Grille = True Then
'    Plage.Borders(xlEdgeLeft).LineStyle = xlContinuous
'    Plage.Borders(xlEdgeTop).LineStyle = xlContinuous
'    Plage.Borders(xlEdgeBottom).LineStyle = xlContinuous
'    Plage.Borders(xlEdgeRight).LineStyle = xlContinuous
'    Plage.Borders(xlContinuous).LineStyle = xlContinuous
'End If
'
'
'End Sub

