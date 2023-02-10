Attribute VB_Name = "FonctionsGlobales"
Public Function LoadCalque()
   AutoApp.ActiveDocument.ActiveLayer = AutoApp.ActiveDocument.Layers(0)
'  AutoApp.ActiveDocument.ActiveLayer.Color = acWhite
'  AutoApp.ActiveDocument.ActiveLayer.Linetype = AutoApp.ActiveDocument.Linetypes(0).Name
'   AutoApp.ActiveDocument.ActiveLayer.Lineweight = acLnWtByBlock
    AutoApp.ActiveDocument.ActiveLayer.ViewportDefault = True
  AutoApp.ActiveDocument.PurgeAll
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
 
 
 
 Public Function MyFormat(Mytype As String, MyText As Object, MyLib As String) As Boolean
 MyFormat = True
 If MyText = "" Then Exit Function
 
  Select Case UCase(Mytype)
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
 
Function DecalInsertPointLigneTableau_fils_Bas(y, Ofset)
    DecalInsertPointLigneTableau_fils_Bas = y - Ofset
End Function
Function DecalInsertPointLigneTableau_fils_Gauche(x, Ofset)
    DecalInsertPointLigneTableau_fils_Gauche = x + Ofset
End Function
Public Function funAttributesLigne_Tableau_fils(MyName As String, Attributes, Tableau, Nb, Optional RangeTitre As Recordset, Optional BoolTirte As Boolean, Optional MyColection As Collection, Optional vignette As Boolean, Optional EPISSURE As Boolean)
Dim Designation As String
Dim MyNb As Long
Dim MyNbStart As Long
Dim msgAttib As String
If vignette = True Then
    If EPISSURE = False Then
        Designation = ".HAUT"
        MyNb = 4
        MyNbStart = 2
    Else
        MyNbStart = 3
         Designation = ""
      MyNb = 3
    End If
Else
    Designation = ""
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
            Attributes(MyColection.Item(Replace(RangeTitre.Fields(i).Name, "PRECO", "PRECO.") & Designation)).TextString = Trim("" & Tableau(i))
        Else
            Attributes(MyColection.Item("EPISSURE")).TextString = Trim("" & Tableau(i))

        End If
            Designation = ""
    End If
Next i
Exit Function
MsgError:
FunError 2, RangeTitre.Fields(i).Name, MyName
Resume Next
End Function

Public Function FunInsBlock(PathName, InsertPoint, Name, Optional Rotation As Double, Optional XScaleFactor As Double, Optional YScaleFactor As Double, Optional ZScaleFactor As Double) As AcadBlockReference
On Error GoTo GesERR
ErrInsert = False
    Set FunInsBlock = ActiveDocument.ModelSpace.InsertBlock(InsertPoint, PathName, 1#, 1#, 1#, 0#)
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
    AutoApp.ZoomAll
    
    Exit Function
GesERR:
    FunError 100, "*", PathName & vbCrLf & Err.Description
    ErrInsert = True
   
End Function

Public Function FunInsBlock2(PathName, InsertPoint, Name, Optional Rotation As Double, Optional XScaleFactor As Double, Optional YScaleFactor As Double, Optional ZScaleFactor As Double) As AcadBlockReference
On Error GoTo GesERR
ErrInsert = False
    Set FunInsBlock2 = ActiveDocument.ModelSpace.InsertBlock(InsertPoint, PathName, 1#, 1#, 1#, 0#)
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

    AutoApp.ZoomAll
    
    Exit Function
GesERR:
 Msg = Msg & "***************************************************" & vbCrLf
           Msg = Msg & PathName & vbCrLf
            Msg = Msg & Err.Description & vbCrLf
            Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
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
Dim sql As String
If Trim("" & Lib1) = "" Then Exit Function
If JobError = 0 Then JobError = AtrbNumError
Msg = MsgErreur(NumErr, Lib1, Lib2, Msg)
sql = "INSERT INTO T_Error ( JobError, ValError ) "
sql = sql & "values(" & JobError & ",'" & Msg & "' );"
Con.Exequte sql

End Function
Public Function AfficheErreur(Path As String, Entete)
    Dim NuFichier As Long
    Dim Text
    Dim MyTxtErr
    Dim sql As String
    Dim RsErreur As Recordset
    Dim Fichier As String
    NuFichier = FreeFile
    
    Text = ""
    sql = "SELECT T_Error.ValError FROM T_Error "
    sql = sql & "WHERE T_Error.JobError=" & JobError & ";"
   
    Set RsErreur = Con.OpenRecordSet(sql)
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
    sql = "DELETE T_Error.* FROM T_Error "
    sql = sql & "WHERE T_Error.JobError=" & JobError & ";"
    Con.Exequte sql
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

Public Sub OpenFichier(Fichier)

    Dim MyDocument As AutoCAD.AcadDocument
    Set MyDocument = AutoApp.Documents.Open(Fichier)
    MyDocument.Activate
End Sub
Public Sub OpenNew()

    Dim MyDocument As AutoCAD.AcadDocument

    Set MyDocument = AutoApp.Documents.Add
    MyDocument.Activate
End Sub

Public Sub SaveAs(Fichier)
On Error Resume Next

     AutoApp.ActiveDocument.ActiveLayer = AutoApp.ActiveDocument.Layers(1)
        AutoApp.ActiveDocument.PurgeAll
    AutoApp.ActiveDocument.SaveAs Fichier, acR15_dwg
    If Err Then MsgBox Err.Description
    Err.Clear
    AutoApp.ActiveDocument.Close , False
End Sub
Public Sub CloseDocument()
    AutoApp.ActiveDocument.Close
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

DbNumPlan = CherCheInFihier("Bdnumero")
db = CherCheInFihier("BdAutocable")
DonneesEntreprise = CherCheInFihier("DonneesEntreprise")
DonneesProduction = CherCheInFihier("DonneesProduction")

funOpenDatabase
NmJob = LaodJob
End Sub
Function LaodJob() As Long
Dim sql As String
Dim Rs As Recordset
If NmJob = 0 Then

sql = "SELECT [NumErreur]+1 AS Job FROM T_NumErreur WHERE T_NumErreur.LibErreur='Job';"
Set Rs = Con.OpenRecordSet(sql)
LaodJob = Rs!Job
sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1 WHERE T_NumErreur.LibErreur='Job';"
Con.Exequte sql
Set Rs = Con.CloseRecordSet(Rs)

Else
    LaodJob = NmJob
End If
End Function
Public Function PathArchive(PathRacicine As String, Client As String, CleAc As String, Piece As String, Mytype As String, Fichier, IdPieces As Long, Indice_Pieces As String, Indice_Plan As String, Version As Long) As String
Dim Fso As New FileSystemObject
Dim sql As String
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
If Fso.FolderExists(PathRacicine & Client & "\PI\" & CleAc & "\" & Piece) = False Then
    Fso.CreateFolder PathRacicine & Client & "\PI\" & CleAc & "\" & Piece
End If

Select Case UCase(Mytype)
    Case "PL"
        PathArchive = Client & "\PI\" & CleAc & "\" & Piece & "\12-PL\"
    Case "OU"
        PathArchive = Client & "\PI\" & CleAc & "\" & Piece & "\16-OU\"
    Case "LI"
        PathArchive = Client & "\PI\" & CleAc & "\" & Piece & "\14-LI\"
    Case "PI"
End Select
If Fso.FolderExists(PathRacicine & PathArchive) = False Then
    Fso.CreateFolder PathRacicine & PathArchive
End If
PathArchive = PathArchive & Fichier


    sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(Mytype) & "AutoCadSave = '" & MyReplace(PathArchive) & "', T_indiceProjet.NbErr = " & NbError & " "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdPieces & ";"
    Con.Exequte sql
    
        sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(Mytype) & "AutoCadSave = '" & MyReplace(PathArchive) & "', T_indiceProjet.NbErr = " & NbError & " "
    sql = sql & "WHERE T_indiceProjet.pere=" & IdPieces & ";"
    Con.Exequte sql


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
Public Sub SubLoadFils(IdPieces As Long)
    Dim RsLigne As Recordset
    Dim sql As String
    Dim Fso As New FileSystemObject
    Dim NbSupprim As Long
'    Dim NewTor As classTor
   
     
    PathBlocs = TableauPath.Item("PathBlocs")
            
         If Left(PathBlocs, 2) <> "\\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs

    InsertPointLigneTableau_fils(0) = -1096.5549: InsertPointLigneTableau_fils(1) = -76.8274: InsertPointLigneTableau_fils(2) = 0

sql = "SELECT Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
    sql = sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT,  "
    sql = sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
    sql = sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.POS,  "
    sql = sql & "Ligne_Tableau_fils.FA, Ligne_Tableau_fils.VOI,  "
    sql = sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.FA2,  "
    sql = sql & "Ligne_Tableau_fils.VOI2,Ligne_Tableau_fils.OPTION,  Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.LONG,  "
    sql = sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.APP2 " ',  "
'    Sql = Sql & "T_indiceProjet.Id_Pieces "
    sql = sql & "FROM Ligne_Tableau_fils "
    sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdPieces & " "
    sql = sql & "ORDER BY Ligne_Tableau_fils.FIL;"

    Set RsLigne = Con.OpenRecordSet(sql)
    While RsLigne.EOF = False
        If Trim("" & RsLigne!PRECO) <> "" Then
        On Error Resume Next
        Set a = Nothing
        DoEvents
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
            a = CollectionTor("" & RsLigne!APP2)
            If Err Then
            Err.Clear
                NUMNTORBLOC = NUMNTORBLOC + 1
Reprise2:
                CollectionTor.Add NUMNTORBLOC, Trim("" & RsLigne!APP2)
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2)))
            End If
                TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CodeApp = Trim("" & RsLigne!APP2)
                  TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).Garder = True
                If TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CodeApp = "" Then
                   GoTo Reprise2
                 End If
                  Set a = Nothing
        DoEvents
                a = TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CollectionTor("" & RsLigne!PRECO)
                If Err Then
            Err.Clear
                TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).NumTor = TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).NumTor + 1
               TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CollectionTor.Add TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).NumTor, "" & RsLigne!PRECO
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CollectionTor("" & RsLigne!PRECO))
            End If
            If InStr(1, TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CollectionTor("" & RsLigne!PRECO)).TableauFile, "" & RsLigne!Fil & " ") = 0 Then
                 TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CollectionTor("" & RsLigne!PRECO)).Garder = True
              TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CollectionTor("" & RsLigne!PRECO)).TableauFile = TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CollectionTor("" & RsLigne!PRECO)).TableauFile & RsLigne!Fil & " "
               TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!APP2))).CollectionTor("" & RsLigne!PRECO)).TorName = "" & RsLigne!PRECO
           End If
            Set a = Nothing
        DoEvents
            On Error GoTo 0
        End If
        NbConnecteur = NbConnecteur + 1
        RsLigne.MoveNext
    Wend




    
    If Val(NbConnecteur) <> 0 Then
    RsLigne.MoveFirst
    End If
     FormBarGrah.ProgressBar1.Value = 0
    If Val(NbConnecteur) <> 0 Then
     FormBarGrah.ProgressBar1.Max = 1 + NbConnecteur
    Else
         FormBarGrah.ProgressBar1.Max = 1 + 1
    End If
     FormBarGrah.ProgressBar1Caption.Caption = "Chargement de la liste de fils"

    While RsLigne.EOF = False
         FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        AutoApp.ZoomAll

        ReDim Tableau(RsLigne.Fields.Count)
        If UCase(Trim("" & RsLigne.Fields(0))) <> "SUPPRIMER" Then
            For Col = 0 To RsLigne.Fields.Count - 1
                DoEvents
                Tableau(Col) = "" & RsLigne.Fields(Col)
            Next Col
     
            RenseigneConnecteurBroches RsLigne
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
            If Fso.FileExists(PathBlocs & "\LIGNE TABLEAU DES FILS.dwg") = False Then
                MsgBox "err"
            End If
            Set NewBlock = FunInsBlock(PathBlocs & "\LIGNE TABLEAU DES FILS.dwg", InsertPointLigneTableau_fils, "L" & CInt(Row))
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
           NbSupprim = NbSupprim + 1
          End If
          RsLigne.MoveNext
     
        Wend
        If NbConnecteur > 0 Then
            Set NewBlock = FunInsBlock(PathBlocs & "\Nombre_fils.dwg", InsertPointLigneTableau_fils, "N1")
            attri = NewBlock.GetAttributes
            attri(0).TextString = NbConnecteur - NbSupprim
         End If
       
        SacnConnecteur
Fin:
    Set MyRange = Nothing
    Set MySheet = Nothing
    ReDim TableauDeConnecteurs(0)
    Set Fso = Nothing
    Exit Sub
Error1:
    FunError 3, CStr("" & Lib1), CStr("" & Lib2)
    Resume Next

End Sub

Public Sub SacnConnecteur()
    Dim Index As Long
    Dim NewBlock  As AcadBlockReference
    Dim MyFichier As String
    Dim PathNUMEROFIL As String
     FormBarGrah.ProgressBar1.Value = 0
      PathNUMEROFIL = TableauPath.Item("PathNUMEROFIL") & "\"
                If Left(PathNUMEROFIL, 2) <> "\\" Then PathNUMEROFIL = TableauPath.Item("PathServer") & PathNUMEROFIL
    If UBound(TableauDeConnecteurs) > 0 Then
         FormBarGrah.ProgressBar1.Max = 1 + UBound(TableauDeConnecteurs)
     Else
         FormBarGrah.ProgressBar1.Max = 1 + 1
     End If
     FormBarGrah.ProgressBar1Caption.Caption = "Chargement des vignettes"
    For Index = 1 To UBound(TableauDeConnecteurs)
         FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        
        If TableauDeConnecteurs(Index).ConnecteurExiste = True Then
            If TableauDeConnecteurs(Index).indexFile > 0 Then
                TableauDeConnecteurs(Index).TableauFile = TriTableau(TableauDeConnecteurs(Index).TableauFile)
               
                MyFichier = DirNUMEROFIL(PathNUMEROFIL, TableauDeConnecteurs(Index).indexFile)
                If Trim(MyFichier) <> "" Then
                    InsertionPoint = TableauDeConnecteurs(Index).NewVignette.InsertionPoint
                    
                    InsertPointLigneTableau_fils(0) = InsertionPoint(0) + 18.022: InsertPointLigneTableau_fils(1) = InsertionPoint(1) + 16.041: InsertPointLigneTableau_fils(2) = InsertionPoint(2)
                    Set NewBlock = FunInsBlock(MyFichier, InsertPointLigneTableau_fils, "NF" & CInt(Index))
                    AttribueLibV NewBlock, Index
                    InsertPointLigneTableau_fils(1) = InsertionPoint(1)
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
Function InsertCartoucheEncelad(Index As Long, Mytype As String, ParamArray InsertPointLigneTableau_fils())
If Mytype = "OU" Then
     InsertPointLigneTableau_fils(0)(0) = 3711.7662: InsertPointLigneTableau_fils(0)(1) = 115.8474: InsertPointLigneTableau_fils(0)(2) = 0
    GoTo Fin
 End If

Select Case Index
        Case 1
            InsertPointLigneTableau_fils(0)(0) = -312#: InsertPointLigneTableau_fils(0)(1) = 20: InsertPointLigneTableau_fils(0)(2) = 0

        Case 2
            InsertPointLigneTableau_fils(0)(0) = 876#: InsertPointLigneTableau_fils(0)(1) = 20#: InsertPointLigneTableau_fils(0)(2) = 0

        Case 3
             InsertPointLigneTableau_fils(0)(0) = -312#: InsertPointLigneTableau_fils(0)(1) = -820#: InsertPointLigneTableau_fils(0)(2) = 0

        Case 4
            InsertPointLigneTableau_fils(0)(0) = 876#: InsertPointLigneTableau_fils(0)(1) = -820#: InsertPointLigneTableau_fils(0)(2) = 0


End Select
Fin:
End Function
Function InsertCartoucheClient(Index As Long, Mytype As String, ParamArray InsertPointLigneTableau_fils())
If Mytype = "OU" Then
    InsertPointLigneTableau_fils(0)(0) = 3566.7253: InsertPointLigneTableau_fils(0)(1) = 115.8474: InsertPointLigneTableau_fils(0)(2) = 0
    Exit Function
End If
Select Case Index
        Case 1
                InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = 126.0753: InsertPointLigneTableau_fils(0)(2) = 0

        Case 2
                InsertPointLigneTableau_fils(0)(0) = 1022.9591: InsertPointLigneTableau_fils(0)(1) = 126.0753: InsertPointLigneTableau_fils(0)(2) = 0

        Case 3
                InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = -713.9247: InsertPointLigneTableau_fils(0)(2) = 0

        Case 4
                InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = -713.9247: InsertPointLigneTableau_fils(0)(2) = 0


End Select
End Function

Public Function ChargeCartoucheEncelade(IdIndiceProjet As Long, Mytype As String, NbCartouche As Long, Optional OuOk As Boolean) As Boolean
  
Dim Fso As New FileSystemObject
Dim sql As String
Dim Rs As Recordset
Dim Status As String
Dim FichierCartouche As String
Dim Index As Long

sql = "UPDATE T_indiceProjet SET T_indiceProjet.Cartouche = '" & Replace(MyReplace(RepPlacheClous), TableauPath.Item("PathServer"), "") & "' "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Con.Exequte sql
sql = "SELECT RqCartouche.* "
sql = sql & "FROM RqCartouche "
sql = sql & "WHERE RqCartouche.T_indiceProjet.Id=" & IdIndiceProjet & ";"

If Left(RepPlacheClous, 2) <> "\\" Then RepPlacheClous = TableauPath.Item("PathServer") & RepPlacheClous
PathBlocs = TableauPath.Item("PathBlocs")
 If Left(PathBlocs, 2) <> "\\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = True Then Exit Function
    Status = "" & Rs!Status
    CartoucheCleient = False
    For Index = 1 To NbCartouche
      InsertCartoucheEncelad Index, Mytype, InsertPointLigneTableau_fils
    If Mytype = "OU" Then
        FichierCartouche = RepPlacheClous
      
    Else
        FichierCartouche = PathBlocs & "\CARTOUCHE ENCELADE.dwg"
        
    End If
    If Fso.FileExists(FichierCartouche) = False Then
        Set Fso = Nothing
        Exit Function
    End If

    Set NewBlock = FunInsBlock(FichierCartouche, InsertPointLigneTableau_fils, "LeCartouche1E")

    AttClent = NewBlock.GetAttributes

    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    AttClent(AttribuCartouche(".NOM.DU.Client")).TextString = "" & Rs!Client
    AttClent(AttribuCartouche(".RESPONSABLE.Client")).TextString = "" & Rs!Responsable
    AttClent(AttribuCartouche(".NOM.DU.Projet")).TextString = "" & Rs!Projet
    AttClent(AttribuCartouche(".VAGUE")).TextString = "" & Rs!Vague
        AttClent(AttribuCartouche(".DESIGNATION.LIGNE.1")).TextString = Replace("" & Rs!Ensemble, Chr(13), "")
    AttClent(AttribuCartouche(".OPTION.ET.DIVERSITE")).TextString = "" & Rs!Equipement
    If OuOk = True Then
        AttClent(AttribuCartouche("Reference.PLAN.Client")).TextString = "" & Rs!OU
    Else
        AttClent(AttribuCartouche("Reference.PLAN.Client")).TextString = "" & Rs!PL
    End If
    AttClent(AttribuCartouche("INDICE")).TextString = "" & Rs!PI_Indice
    AttClent(AttribuCartouche("Reference.PLAN.FONCTIONNEL")).TextString = "" & Rs!RefPf
    AttClent(AttribuCartouche("RF2")).TextString = "" ' & Rs!Client
    AttClent(AttribuCartouche("Reference.ENCELADE")).TextString = "" & Rs!Pi
    AttClent(AttribuCartouche("RF1")).TextString = "" & Rs!PI_Indice
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
Public Function ChargeCartoucheClient(IdIndiceProjet As Long, Mytype As String, NbCartouche As Long, Optional OuOk As Boolean) As Boolean
Dim sql As String
Dim FichierCartouche As String
Dim Rs As Recordset
Dim RsCartouche As Recordset
Dim Index As Long
LeCartouche = "CARTOUCHE  RENAULT.dwg"
LeCartoucheE = "CARTOUCHE ENCELADE.dwg"

'If boolFormClient = False Then Exit Function
sql = "SELECT RqCartouche.* "
sql = sql & "FROM RqCartouche "
sql = sql & "WHERE RqCartouche.T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = True Then Exit Function
MyCARTOUCHE_Client = Trim("" & Rs!Client)
sql = "SELECT T_Clients.Formulaire FROM T_Clients "
sql = sql & "WHERE T_Clients.Client='" & MyReplace(Trim("" & Rs!Client)) & "';"
Set RsCartouche = Con.OpenRecordSet(sql)
If RsCartouche.EOF = False Then
    LeCartouche = Trim("" & RsCartouche!Formulaire)
     If Left(LeCartouche, 2) <> "\\" Then LeCartouche = TableauPath.Item("PathServer") & LeCartouche
     
End If
Set RsCartouche = Con.CloseRecordSet(RsCartouche)
If LeCartouche = "" Then Exit Function
    Dim Fso As New FileSystemObject
    CartoucheCleient = False
    For Index = 1 To 2
    InsertCartoucheClient Index, Mytype, InsertPointLigneTableau_fils
   
    If Fso.FileExists(LeCartouche) = False Then
        Set Fso = Nothing
        Exit Function
    End If
    Set NewBlock = FunInsBlock(LeCartouche, InsertPointLigneTableau_fils, "LeCartouche1")
    AttClent = NewBlock.GetAttributes

    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
AttClent(AttribuCartouche("DESIGN.1.CART.RENAULT")).TextString = "" & Rs("Ensemble")
AttClent(AttribuCartouche("MASSE")).TextString = "" & Rs("Masse")
AttClent(AttribuCartouche("IND.Pi")).TextString = "" & Rs("PI_Indice")
AttClent(AttribuCartouche("REF.PF.CART.RENAULT")).TextString = "" & Rs("RefPF")
If OuOk = True Then
   
        AttClent(AttribuCartouche("REF.PLAN.INDUSTRIEL")).TextString = "" & Rs!OU
    Else
    
        AttClent(AttribuCartouche("REF.PLAN.INDUSTRIEL")).TextString = "" & Rs!PL
    
    End If
AttClent(AttribuCartouche("DESIGN.2.CART.RENAULT")).TextString = ""
AttClent(AttribuCartouche("DESGN.1.ANGL.CART.REN")).TextString = ""
AttClent(AttribuCartouche("DESGN.2.ANGL.CART.REN")).TextString = ""

AttClent(AttribuCartouche("IND.PF")).TextString = ""


AttClent(AttribuCartouche("REF.PIECE.CART.RENAULT")).TextString = ""
AttClent(AttribuCartouche("SERVICE")).TextString = ""
AttClent(AttribuCartouche("UTILISATEURS")).TextString = ""

AttClent(AttribuCartouche("REGLEMENT")).TextString = ""
AttClent(AttribuCartouche("NOTE.BE.1")).TextString = ""
AttClent(AttribuCartouche("NOTE.BE.2")).TextString = ""
AttClent(AttribuCartouche("Num.VISA")).TextString = ""
 AttClent(AttribuCartouche("REF.PIECE.CART." & MyCARTOUCHE_Client)).TextString = "" & Rs!RefP
AttClent(AttribuCartouche("X/X")).TextString = CStr(Index) & "/" & CStr(NbCartouche)
Next Index

'AttClent(AttribuCartouche("DESGN.1.ANGL.CART.REN")).TextString = ""
    CartoucheCleient = True
    Set Fso = Nothing
End Function
Public Sub Maj(MyControl As Object)
Dim Rs As Recordset
Dim sql As String

MyControl.Clear
sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
sql = sql & "FROM T_Clients "
sql = sql & "ORDER BY T_Clients.Client;"
MyControl.AddItem ""
Set Rs = Con.OpenRecordSet(sql)
While Rs.EOF = False
    MyControl.AddItem Trim("" & Rs!Client)
        If MyControl.ListCount = 1 Then MyControl.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)

End Sub
Public Function ChercheXls(MyRange, Val) As Long
ChercheXls = 0
 For i = 2 To MyRange.Rows.Count
                If UCase(Trim("" & MyRange(i))) = UCase(Trim("" & Val)) Then
                        ChercheXls = i
                        Exit For
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
Public Function RenseigneConnecteurBroches(RangeAttribue As Recordset) As Boolean
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
Sub TestFl()
a = CherCheInFihier("Bdnumero")
End Sub
Public Function LoadComposants(IdIndiceProjet As Long, Mytype As String) As Boolean
  LoadComposants = False
    Dim RsCompsants As Recordset
    Dim sql As String
    Dim MyRep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
    Dim NbConnecteur As Long
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
    Dim XMin As Long
    Dim YMin As Long
    Dim PathComposantsDefault As String
     sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1
    
      sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    sql = "SELECT  T_Clients.PathComposants FROM T_Clients "
sql = sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathComposants) = "" Then
         PathComposantsDefault = TableauPath.Item("PathComposantsDefault")
   Else
             PathComposantsDefault = RsConnecteur!PathComposants
         If Left(PathComposantsDefault, 2) <> "\\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
    
    End If
Else
                 PathComposantsDefault = RsConnecteur!PathComposants

End If
If Left(PathComposantsDefault, 2) <> "\\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
' PathComposantsDefault = PathComposantsDefault & "COMPOSANTS\"
 
 sql = "SELECT Composants.* FROM Composants "
 sql = sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & ";"
 Set RsCompsants = Con.OpenRecordSet(sql)
 While RsCompsants.EOF = False
 On Error Resume Next
                  
                   a = CollectionComp(Trim("C" & RsCompsants!NUMCOMP))
                If Err Then
                    If NUMCOM < RsCompsants!NUMCOMP Then
                         ReDim Preserve TableauDeComposants(RsCompsants!NUMCOMP)
                         CollectionComp.Add RsCompsants("NUMCOMP").Value, Trim("C" & RsCompsants!NUMCOMP)
                         NUMCOM = RsCompsants!NUMCOMP
                    End If
 
                End If
 
    
    RsCompsants.MoveNext
 Wend
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = 1 + NUMCOM
 FormBarGrah.ProgressBar1Caption.Caption = "Chargement des Compsants"
 
  RsCompsants.Requery
   XMin = 823.5964: YMin = -954.9939
    For i = 0 To IndexIstC
  
    InsertPointConnecteur(i).InsertPointConnecteur(0) = XMin - (150 * i): InsertPointConnecteur(i).InsertPointConnecteur(1) = YMin - (150 * i): InsertPointConnecteur(i).InsertPointConnecteur(2) = 0
    Next i
 On Error GoTo GesERR
 
 While RsCompsants.EOF = False
  FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
 If TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).ComposantsExiste = False Then
  InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0), -300)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
    TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertPointLigneC(0) = InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0)
    TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertPointLigneC(1) = InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(1)
    TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertPointLigneC(2) = InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(2)
   
    TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).XScaleFactorC = 1
    TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).YScaleFactorC = 1
    TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).ZScaleFactorC = 1
    TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).RotationC = 0
 End If

'    PathComposantsDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS"
    Lib1 = PathComposantsDefault & "\" & RsCompsants!Path & "\" & RsCompsants!REFCOMP & ".dwg"
    Lib2 = "" & RsCompsants!REFCOMP
    NumErr = 6
    Set TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).NewBlock = FunInsBlock(PathComposantsDefault & "\" & RsCompsants!Path & "\" & RsCompsants!REFCOMP & ".dwg", TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).InsertPointLigneC, "", TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).RotationC, TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).XScaleFactorC, TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).YScaleFactorC, TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).ZScaleFactorC)
    Err.Clear
    Set TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).Attribues = ColectionAttribueConecteur(TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).NewBlock.GetAttributes)

                Att = TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).NewBlock.GetAttributes
                Lib1 = "DESIGNCOMP"
                Lib2 = "" & RsCompsants!REFCOMP
                NumErr = 7
                Att(TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).Attribues("DESIGNCOMP")).TextString = "" & RsCompsants!DESIGNCOMP
                Err.Clear
                 Lib1 = "NUMCOMP"
                Lib2 = "" & RsCompsants!REFCOMP
                Att(TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).Attribues("NUMCOMP")).TextString = "C" & RsCompsants!NUMCOMP
                Err.Clear
                 Lib1 = "PATHCOMP"
                Lib1 = "NUMCOMP"
                Att(TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).Attribues("PATHCOMP")).TextString = "" & RsCompsants!Path
                  Err.Clear
                 Lib1 = "REFCOMP"
                Lib2 = ""
                Att(TableauDeComposants(CollectionComp("C" & RsCompsants!NUMCOMP)).Attribues("REFCOMP")).TextString = "" & RsCompsants!REFCOMP
               
    RsCompsants.MoveNext
 Wend
 Exit Function
GesERR:
    FunError NumErr, "" & Lib1, Err.Description, "" & Lib2
 Resume Next
End Function

Public Function LoadNotas(IdIndiceProjet As Long, Mytype As String) As Boolean
  LoadNotas = False
    Dim RsCompsants As Recordset
    Dim sql As String
    Dim MyRep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
    Dim NbConnecteur As Long
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
    Dim XMin As Long
    Dim YMin As Long
    Dim PathNotasDefault As String
     sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1
    
      sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    sql = "SELECT  T_Clients.PathNotas FROM T_Clients "
sql = sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathNotas) = "" Then
         PathNotasDefault = TableauPath.Item("PathNotasDefault")
   Else
             PathNotasDefault = RsConnecteur!PathNotas
         If Left(PathNotasDefault, 2) <> "\\" Then PathNotasDefault = TableauPath.Item("PathServer") & PathNotasDefault
    
    End If
Else
                 PathNotasDefault = RsConnecteur!PathNotas

End If
If Left(PathNotasDefault, 2) <> "\\" Then PathNotasDefault = TableauPath.Item("PathServer") & PathNotasDefault
' PathNotasDefault = PathNotasDefault & "Nota\"
 
 sql = "SELECT Nota.* FROM Nota "
 sql = sql & "WHERE Nota.Id_IndiceProjet=" & IdIndiceProjet & " "
 sql = sql & "order by Nota.NUMNOTA;"
 Set RsCompsants = Con.OpenRecordSet(sql)
 While RsCompsants.EOF = False
 On Error Resume Next
                  
                   a = CollectionNota(Trim("N" & RsCompsants!NUMNOTA))
                If Err Then
                
                    If NUMNOTA < RsCompsants!NUMNOTA Then
                         ReDim Preserve TableauDeNotas(RsCompsants!NUMNOTA)
                         CollectionNota.Add RsCompsants("NUMNOTA").Value, Trim("N" & RsCompsants!NUMNOTA)
                         NUMNOTA = RsCompsants!NUMNOTA
                    End If
 
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
 FormBarGrah.ProgressBar1Caption.Caption = "Chargement des Notas"
 
 While RsCompsants.EOF = False
  FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
 If TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).NotasExiste = False Then
  InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(1), -600)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).InsertPointLigneC(0) = InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0)
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).InsertPointLigneC(1) = InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(1)
    TableauDeNotas(CollectionNota("N" & RsCompsants!NUMNOTA)).InsertPointLigneC(2) = InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(2)
   
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



Public Function LoadConnecteur(IdIndiceProjet As Long, Mytype As String) As Boolean
    LoadConnecteur = False
    Dim RsConnecteur As Recordset
    Dim sql As String
    Dim MyRep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
    Dim NbConnecteur As Long
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
    Dim XMin As Long
    Dim YMin As Long
    Dim PathConnecteursDefault As String
  
    sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    sql = "SELECT T_Clients.Client, T_Clients.PathConnecteurs FROM T_Clients "
sql = sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathConnecteurs) = "" Then
         PathConnecteursDefault = TableauPath.Item("PathConnecteursDefault")
   Else
             PathConnecteursDefault = RsConnecteur!PathConnecteurs
         If Left(PathConnecteursDefault, 2) <> "\\" Then PathConnecteursDefault = TableauPath.Item("PathServer") & PathConnecteursDefault
    
    End If
Else
                 PathConnecteursDefault = RsConnecteur!PathConnecteurs

End If
If Left(PathConnecteursDefault, 2) <> "\\" Then PathConnecteursDefault = TableauPath.Item("PathServer") & PathConnecteursDefault



 sql = "SELECT Connecteurs.CONNECTEUR, Connecteurs.[O/N],  "
sql = sql & "Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
sql = sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.PRECO1,  "
sql = sql & "Connecteurs.PRECO2  "
sql = sql & "FROM Connecteurs "
sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & " "
    sql = sql & "ORDER BY Connecteurs.N°;"
    NumErr = 1
    If Mytype = "OU" Then
        XMin = 1185.771
        YMin = 1667.3509
    Else
        XMin = 30
        YMin = 870.4179
    End If
    Set RsConnecteur = Con.OpenRecordSet(sql)
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
     FormBarGrah.ProgressBar1Caption.Caption = "Chargement des connecteurs"
    If NbConnecteur <> 0 Then
        RsConnecteur.Requery
    End If
      On Error GoTo GesERR
      FormBarGrah.ProgressBar1.Value = 0
    While RsConnecteur.EOF = False
        
         FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        
        DoEvents


   
    
        If UCase("" & RsConnecteur.Fields(0)) <> "NEANT" Then
            If Fso.FileExists(PathConnecteursDefault & "\" & RsConnecteur.Fields(0) & ".dwg") = True Then
                MyRep = PathConnecteursDefault
                Trouve = True
                NumErr = 4
              
                
                TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = True
            Else
                NumErr = 1
                TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = False
                MyRep = ""
                
GesERR:
                Trouve = False
                FunError NumErr, "" & RsConnecteur.Fields(3), Err.Description, "" & RsConnecteur.Fields(0)
            End If
            If Trouve = True Then
            If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock = FunInsBlock(MyRep & "\" & RsConnecteur.Fields(0) & ".dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneC, "", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorC)
            Else
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock = FunInsBlock(MyRep & "\" & RsConnecteur.Fields(0) & ".dwg", InsertPointConnecteur(Nb_L_C).InsertPointConnecteur, "")
                InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0), 300)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
             End If
                  If ErrInsert = True Then GoTo EnrSuinant
                If UCase("" & RsConnecteur.Fields(1)) = True Then
                    TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE = True
                    If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then

                    Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(MyRep & "\EPISSURES.dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneV, "V" & "" & RsConnecteur.Fields(4), TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorV)
                    Else
                        Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(MyRep & "\EPISSURES.dwg", InsertPointLigneTableau_Vignette, "V" & "" & RsConnecteur.Fields(4))

                        InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 100)
                        NbLignesVignette = NbLignesVignette + 1
                    End If
                Else
                    TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE = False
                    If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then

                    Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(MyRep & "\VIGNETTE CONNECTEUR.dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneV, "V" & "" & RsConnecteur.Fields(4), TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorV)
                    Else
                    Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(MyRep & "\VIGNETTE CONNECTEUR.dwg", InsertPointLigneTableau_Vignette, "V" & "" & RsConnecteur.Fields(4))

                    InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 100)
                    NbLignesVignette = NbLignesVignette + 1
                    End If
                End If
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes)
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes)

                At = TableauAtribCon(RsConnecteur, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE)
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE
                
            End If
        End If
        If NbLignesVignette = 11 Then
            InsertPointLigneTableau_Vignette(0) = -1146.1429
            InsertPointLigneTableau_Vignette(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_Vignette(1), -40)
            NbLignesVignette = 0
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
    While IndexAt < UBound(Attribues) + 1
        
        Debug.Print UCase(Replace(Replace(UCase(Attribues(IndexAt).TagString), "PRECO", "PRECO."), "PRECO..", "PRECO."))
        MyAttribue.Add IndexAt, UCase(Replace(Replace(UCase(Attribues(IndexAt).TagString), "PRECO", "PRECO."), "PRECO..", "PRECO."))
        Set Atr = Nothing
        
        IndexAt = IndexAt + 1
    Wend
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
    On Error GoTo Fin

    bollInDif = True
    Fils = "FILG"
    For i = 1 To UBound(Attribues)
        DoEvents
        
        IbAttribue = TableauDeConnecteurs(Connecteur).Attribues.Item(Fils & CStr(i))
        If Trim("" & Attribues(IbAttribue).TextString) = "" Then
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
        Fils = "FILD"
        i = 0
        Err.Clear
        GoTo Retour
    End If
    Err.Clear
End Function

 Public Function AtrbNumError() As Long
    Dim sql As String
    Dim NErr As Long
    Dim RsNumError As Recordset
    sql = "SELECT T_NumErreur.LibErreur, T_NumErreur.NumErreur "
    sql = sql & "FROM T_NumErreur "
    sql = sql & "WHERE T_NumErreur.LibErreur='ErrorApp';"
    Set RsNumError = Con.OpenRecordSet(sql)
    If RsNumError.EOF = False Then
        sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1;"
        Con.Exequte sql
        RsNumError.Requery
        AtrbNumError = RsNumError!NumErreur
    End If
End Function
Public Function VersionPices(Pieces As String) As Long
Dim Rs As Recordset
Dim sql As String
sql = "SELECT  VersionPices.Version FROM VersionPices "
sql = sql & "WHERE VersionPices.Pi='" & MyReplace(Pieces) & "';"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = True Then
sql = "INSERT INTO VersionPices ( Pi ) VALUES('" & MyReplace(Pieces) & "');"
Con.Exequte sql
End If
sql = "UPDATE VersionPices SET VersionPices.Version = [Version] + 1 "
sql = sql & "WHERE VersionPices.Pi='" & MyReplace(Pieces) & "';"
Con.Exequte sql
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
ConBaseNum.CloseConnection

End Function
Public Function ReseingeTor(CodeApp As String, InsertTorTitre) As Boolean
On Error GoTo Fin
   Dim PathTorDefault As String
 PathTorDefault = TableauPath.Item("PathTorDefault")
If TableuDeTor(CollectionTor(CodeApp)).Garder = False Then Exit Function
If Left(PathTorDefault, 2) <> "\\" Then PathTorDefault = TableauPath.Item("PathServer") & PathTorDefault
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
Dim sql As String
sql = "DELETE Ligne_Tableau_fils.* "
sql = sql & "FROM Ligne_Tableau_fils "
sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte sql

sql = "DELETE Connecteurs.* "
sql = sql & "FROM Connecteurs "
sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte sql

sql = "DELETE Composants.* "
sql = sql & "FROM Composants "
sql = sql & "WHERE Composants.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte sql

sql = "DELETE Nota.* "
sql = sql & "FROM Nota "
sql = sql & "WHERE Nota.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte sql



sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION,  "
sql = sql & "FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS,  "
sql = sql & "[POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2,  "
sql = sql & "VOI2, PRECO, [OPTION] ) "
sql = sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI,  "
sql = sql & "xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL,  "
sql = sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT,  "
sql = sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,  "
sql = sql & "xls_Ligne_Tableau_fils.LONG, xls_Ligne_Tableau_fils.[LONG CP],  "
sql = sql & "xls_Ligne_Tableau_fils.COUPE, xls_Ligne_Tableau_fils.POS,  "
sql = sql & "xls_Ligne_Tableau_fils.[POS-OUT], xls_Ligne_Tableau_fils.FA,  "
sql = sql & "xls_Ligne_Tableau_fils.APP, xls_Ligne_Tableau_fils.VOI,  "
sql = sql & "xls_Ligne_Tableau_fils.POS2, xls_Ligne_Tableau_fils.[POS-OUT2],  "
sql = sql & "xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.APP2,  "
sql = sql & "xls_Ligne_Tableau_fils.VOI2, xls_Ligne_Tableau_fils.PRECO,  "
sql = sql & "xls_Ligne_Tableau_fils.OPTION "
sql = sql & "FROM xls_Ligne_Tableau_fils "
sql = sql & " where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"



Con.Exequte sql

sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR, [O/N],  "
sql = sql & "DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2,  "
sql = sql & "[100%] )  "
sql = sql & "SELECT " & IdIndice & "  AS Id_IndiceProjet, Xls_Connecteurs.CONNECTEUR,  "
sql = sql & "Xls_Connecteurs.[O/N], Xls_Connecteurs.DESIGNATION,  "
sql = sql & "Xls_Connecteurs.CODE_APP, Xls_Connecteurs.N°,  "
sql = sql & "Xls_Connecteurs.POS, Xls_Connecteurs.[POS-OUT],  "
sql = sql & "Xls_Connecteurs.PRECO1, Xls_Connecteurs.PRECO2,  "
sql = sql & "Xls_Connecteurs.[100%] "
sql = sql & "FROM Xls_Connecteurs "
sql = sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"

Con.Exequte sql



sql = "INSERT INTO Composants ( Id_IndiceProjet, DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
sql = sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Composants.DESIGNCOMP, Xls_Composants.NUMCOMP,   "
sql = sql & "Xls_Composants.REFCOMP, Xls_Composants.Path   "
sql = sql & "FROM Xls_Composants "
sql = sql & "WHERE Xls_Composants.Job=" & NmJob & ";"

Con.Exequte sql

sql = "INSERT INTO Nota ( Id_IndiceProjet, NOTA, NUMNOTA ) "
sql = sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Nota.NOTA, Xls_Nota.NUMNOTA "
sql = sql & "FROM Xls_Nota "
sql = sql & "WHERE Xls_Nota.Job=" & NmJob & ";"

Con.Exequte sql

 sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
sql = sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Composants.*  FROM Xls_Composants "
sql = sql & "where Xls_Composants.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Nota.*  FROM Xls_Nota "
sql = sql & "where Xls_Nota.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
sql = sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Exequte sql
End Sub
