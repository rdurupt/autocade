Attribute VB_Name = "FonctionsGlobales"
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
On erro GoTo MsgError
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

Public Function FunInsBlock(PathName, InsertPoint, Name) As AcadBlockReference
On Error GoTo GesERR
ErrInsert = False
    ' Create the attribute definition object in model space
    Set FunInsBlock = ActiveDocument.ModelSpace.InsertBlock(InsertPoint, PathName, 1#, 1#, 1#, 0#)
   Dim layerObj As AcadLayer
   a = FunInsBlock.Name
  E = FunInsBlock.GetAttributes
'   If Trim("" & Name) <> "" Then
'    ActiveDocument.Blocks(FunInsBlock.Name).Name = Name
'    Else
'    Debug.Print ActiveDocument.Blocks(FunInsBlock.Name).Name
'    End If
    ZoomAll
    Exit Function
GesERR:
    FunError 6, "*", PathName & vbCrLf & Err.Description
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
Public Function FunError(NumErr As Long, Lib1 As String, MSG As String, Optional Lib2 As String)
'Dim Msg As String
Dim Sql As String
If Trim("" & Lib1) = "" Then Exit Function
If JobError = 0 Then JobError = AtrbNumError
MSG = MsgErreur(NumErr, Lib1, Lib2, MSG)
Sql = "INSERT INTO T_Error ( JobError, ValError ) "
Sql = Sql & "values(" & JobError & ",'" & MSG & "' );"
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
MyReplace = Replace(MyReplace, vbCrLf, " ;")
End Function
Public Sub OpenFichier(Fichier)

    Dim MyDocument As AutoCAD.AcadDocument

    Set MyDocument = AutoApp.Documents.Open(Fichier)
    MyDocument.Activate
End Sub
Public Sub SaveAs(Fichier)
On Error Resume Next
    AutoApp.ActiveDocument.SaveAs Fichier, acR15_dwg
    If Err Then MsgBox Err.Description
    Err.Clear
End Sub
Public Sub CloseDocument()
    AutoApp.ActiveDocument.Close
End Sub
Public Function IsConnecteurs(Attributes As Variant) As Boolean
    Dim Table(6) As String
    Dim Trouve As Boolean
Table(0) = "DESIGNATION"
Table(1) = "POS"
Table(2) = "N°"
Table(3) = "CODE_APP"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
Table(6) = "FIL"
    IsConnecteurs = True
    
      For i = LBound(Table) To UBound(Table)
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
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
    
Table(0) = "NOMBRE_FILS"
   IsNOMBRE_FILS = True
    

   For i = LBound(Table) To UBound(Table)
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
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
Table(6) = "FILG"


    IsEpissures = True
    

   For i = LBound(Table) To UBound(Table)
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
    For i = LBound(Attributes) To UBound(Attributes)
        If Trim(UCase(Attributes(i).TagString)) <> Trim(Table(i)) Then
            IsTableauFils = False
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
'    Table(13) = UCase("LONG")
    IsEnteteTableauFils = True
    For i = LBound(Attributes) To UBound(Attributes)
        If Trim(UCase(Attributes(i).TagString)) <> Trim(Table(i)) Then
            IsEnteteTableauFils = False
            Exit Function
        End If
    Next i
End Function

Public Function PathArchive(PathRacicine As String, Client As String, CleAc As String, Fichier) As String
Dim Fso As New FileSystemObject
Dim Sql As String
Dim Rs As Recordset
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
If Fso.FolderExists(PathRacicine & Client & "\PI\" & CleAc & "\12-PL") = False Then
    Fso.CreateFolder PathRacicine & Client & "\PI\" & CleAc & "\12-PL"
End If
PathArchive = PathRacicine & Client & "\PI\" & CleAc & "\12-PL\"

    Sql = "SELECT DISTINCT  T_indiceProjet.Id "
    Sql = Sql & "FROM T_Projet INNER JOIN (T_indiceProjet INNER JOIN Connecteurs ON T_indiceProjet.Id = Connecteurs.Id_IndiceProjet) ON T_Projet.id = T_indiceProjet.IdProjet "
    Sql = Sql & "WHERE T_Projet.Projet='" & varProjet & "' "
    Sql = Sql & "AND T_indiceProjet.Li='" & CartoucheEncelade.CombLi.List(CartoucheEncelade.CombLi.ListIndex, 1) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.AutoCadSaveAs = '" & MyReplace(Client & "\PI\" & CleAc & "\12-PL\" & Fichier) & "' "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & Rs!Id & ";"
    Con.Exequte Sql
End If
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
    For i = LBound(Attributes) To UBound(Attributes)
        If Trim(UCase(Attributes(i).TagString)) <> Trim(Table(i)) Then
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
    For i = LBound(Attributes) To UBound(Attributes)
        If Trim(UCase(Attributes(i).TagString)) <> Trim(Table(i)) Then
            IsVignetteEPISSURE = False
            Exit Function
        End If
    Next i
End Function
Public Sub MyAccess()
    Dim RsLigne As Recordset
    Dim Sql As String
    Dim Fso As New FileSystemObject
    InsertPointLigneTableau_fils(0) = -1096.5549: InsertPointLigneTableau_fils(1) = -76.8274: InsertPointLigneTableau_fils(2) = 0
    Sql = "SELECT Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.POS, Ligne_Tableau_fils.FA, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.APP2 "
    Sql = Sql & "FROM T_Projet INNER JOIN (T_indiceProjet INNER JOIN Ligne_Tableau_fils ON T_indiceProjet.Id = Ligne_Tableau_fils.Id_IndiceProjet) ON T_Projet.id = T_indiceProjet.IdProjet "
    Sql = Sql & "WHERE T_Projet.Projet='" & varProjet & "' "
    Sql = Sql & "AND T_indiceProjet.LI='" & MyReplace(CartoucheEncelade.CombLi.List(CartoucheEncelade.CombLi.ListIndex, 1)) & "';"

    Set RsLigne = Con.OpenRecordSet(Sql)
    While RsLigne.EOF = False
        NbConnecteur = NbConnecteur + 1
        RsLigne.MoveNext
    Wend
    If Val(NbConnecteur) <> 0 Then
    RsLigne.MoveFirst
    End If
    Menu.ProgressBar1.Value = 0
    If Val(NbConnecteur) <> 0 Then
    Menu.ProgressBar1.Max = NbConnecteur
    Else
        Menu.ProgressBar1.Max = 1
    End If
    Menu.ProgressBar1Caption.Caption = "Chargement de la liste de fils"

    While RsLigne.EOF = False
        Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
        DoEvents
        AutoApp.ZoomAll
a = CLng(CollectionCon("" & RsLigne.Fields("APP")))
a = CLng(CollectionCon("" & RsLigne.Fields("APP2")))
        ReDim Tableau(RsLigne.Fields.Count)
        For Col = 0 To RsLigne.Fields.Count - 1
            DoEvents
            Tableau(Col) = "" & RsLigne.Fields(Col)
        Next Col
        RenseigneConnecteurBroches RsLigne
        Row = Row + 1
        If NbLignes = 100 Then
            InsertPointLigneTableau_fils(1) = -76.8274: InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_fils(0), 281.7719)
            NbLignes = 0
        End If

        If NbLignes = 0 Then
            If Fso.FileExists(TableauPath.Item("PathBlocs") & "LIGNE TABLEAU DES FILS_RD.dwg") = False Then
                MsgBox "err"
            End If
            Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & "LIGNE TABLEAU DES FILS_RD.dwg", InsertPointLigneTableau_fils, "E" & CStr(Row))
            InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), 6.8274)
        End If
        If Fso.FileExists(TableauPath.Item("PathBlocs") & "LIGNE TABLEAU DES FILS.dwg") = False Then
            MsgBox "err"
        End If
        Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & "LIGNE TABLEAU DES FILS.dwg", InsertPointLigneTableau_fils, "L" & CInt(Row))
        a = NewBlock.GetAttributes
        InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), 6.8274)

        NbLignes = NbLignes + 1
        On Error GoTo Error1
        Lib1 = RsLigne.Fields(2)
        Lib2 = RsLigne.Fields(8)
        TableauDeConnecteurs(RsLigne.Fields(8)).indexFile = TableauDeConnecteurs(RsLigne.Fields(8)).indexFile + 1
        Lib1 = ""
        Lib2 = ""
        ReDim Preserve TableauDeConnecteurs(RsLigne.Fields(8)).TableauFile(TableauDeConnecteurs(RsLigne.Fields(8)).indexFile)
        TableauDeConnecteurs(RsLigne.Fields(8)).TableauFile(TableauDeConnecteurs(RsLigne.Fields(8)).indexFile) = RsLigne.Fields(2)

        Lib1 = RsLigne.Fields(2)
        Lib2 = RsLigne.Fields(11)

        TableauDeConnecteurs(RsLigne.Fields(11)).indexFile = TableauDeConnecteurs(RsLigne.Fields(11)).indexFile + 1
        Lib1 = ""
        Lib2 = ""
        ReDim Preserve TableauDeConnecteurs(RsLigne.Fields(11)).TableauFile(TableauDeConnecteurs(RsLigne.Fields(11)).indexFile)

        TableauDeConnecteurs(RsLigne.Fields(11)).TableauFile(TableauDeConnecteurs(RsLigne.Fields(11)).indexFile) = RsLigne.Fields(2)

        a = RsLigne.Fields(0)
        funAttributesLigne_Tableau_fils NewBlock.Name, NewBlock.GetAttributes, Tableau, RsLigne.Fields.Count - 1
        RsLigne.MoveNext
    Wend
    If NbConnecteur > 0 Then
        Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & "Nombre_fils.dwg", InsertPointLigneTableau_fils, "N1")
        attri = NewBlock.GetAttributes
        attri(0).TextString = NbConnecteur
     End If
   
    SacnConnecteur
'    SaveAs "c:\RD\rd"
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
    Menu.ProgressBar1.Value = 0
    If UBound(TableauDeConnecteurs) > 0 Then
        Menu.ProgressBar1.Max = UBound(TableauDeConnecteurs)
     Else
        Menu.ProgressBar1.Max = 1
     End If
    Menu.ProgressBar1Caption.Caption = "Chargement des vignettes"
    For Index = 1 To UBound(TableauDeConnecteurs)
        Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
        DoEvents
        AutoApp.ZoomAll
        If TableauDeConnecteurs(Index).ConnecteurExiste = True Then
            If TableauDeConnecteurs(Index).indexFile > 0 Then
                TableauDeConnecteurs(Index).TableauFile = TriTableau(TableauDeConnecteurs(Index).TableauFile)
                MyFichier = DirNUMEROFIL(TableauPath.Item("PathNUMEROFIL"), TableauDeConnecteurs(Index).indexFile)
                If Trim(MyFichier) <> "" Then
                    InsertionPoint = TableauDeConnecteurs(Index).NewVignette.InsertionPoint
                    InsertPointLigneTableau_fils(0) = InsertionPoint(0) + 18.022: InsertPointLigneTableau_fils(1) = InsertionPoint(1) + 16.041: InsertPointLigneTableau_fils(2) = InsertionPoint(2)
                    Set NewBlock = FunInsBlock(MyFichier, InsertPointLigneTableau_fils, "NF" & CInt(Index))
                    AttribueLibV NewBlock, Index
                Else
                
                End If
            End If
        End If
    Next Index

End Sub
Public Function ChargeCartoucheEncelade(MyCARTOUCHE_Encelade) As Boolean

    Dim Fso As New FileSystemObject
    Dim Sql As String
Dim Rs As Recordset
Dim Status As String
Sql = "SELECT T_Status.Status "
Sql = Sql & "FROM T_Status INNER JOIN T_indiceProjet ON T_Status.Id = T_indiceProjet.IdStatus "
Sql = Sql & "WHERE T_indiceProjet.Li='" & MyReplace(CartoucheEncelade.CombLi.List(CartoucheEncelade.CombLi.ListIndex, 1)) & " ';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then Status = Trim("" & Rs!Status)
    CartoucheCleient = False
    InsertPointLigneTableau_fils(0) = -312#: InsertPointLigneTableau_fils(1) = 20: InsertPointLigneTableau_fils(2) = 0
    If Fso.FileExists(TableauPath.Item("PathBlocs") & LeCartoucheE) = False Then
        Set Fso = Nothing
        Exit Function
    End If
    
    Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & LeCartoucheE, InsertPointLigneTableau_fils, "LeCartouche1E")
    AttClent = NewBlock.GetAttributes

    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    For iClient = 0 To NbContolClient
        ControlTag = Split(MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient)).tag, ";")
        a = MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient))
        AttClent(AttribuCartouche(UCase(ControlTag(0)))).TextString = Replace(MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient)), vbCrLf, Right(vbCrLf, 1))
    Next iClient
    AttClent(AttribuCartouche("ETAT")).TextString = Status
    AttClent(AttribuCartouche("X/X")).TextString = "1/4"
    
    InsertPointLigneTableau_fils(0) = 876#: InsertPointLigneTableau_fils(1) = 20#: InsertPointLigneTableau_fils(2) = 0
    Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & LeCartoucheE, InsertPointLigneTableau_fils, "LeCartouche2E")
    AttClent = NewBlock.GetAttributes
    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    For iClient = 0 To NbContolClient
        ControlTag = Split(MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient)).tag, ";")
        a = MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient))
        AttClent(AttribuCartouche(UCase(ControlTag(0)))).TextString = Replace(MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient)), vbCrLf, Right(vbCrLf, 1))
    Next iClient
    
    AttClent(AttribuCartouche("ETAT")).TextString = Status
    AttClent(AttribuCartouche("X/X")).TextString = "2/4"

    
    InsertPointLigneTableau_fils(0) = 876#: InsertPointLigneTableau_fils(1) = -820#: InsertPointLigneTableau_fils(2) = 0
    Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & LeCartoucheE, InsertPointLigneTableau_fils, "LeCartouche3E")
    AttClent = NewBlock.GetAttributes
    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    For iClient = 0 To NbContolClient
        ControlTag = Split(MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient)).tag, ";")
        a = MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient))
        AttClent(AttribuCartouche(UCase(ControlTag(0)))).TextString = Replace(MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient)), vbCrLf, Right(vbCrLf, 1))
    Next iClient
    
    AttClent(AttribuCartouche("ETAT")).TextString = Status
    AttClent(AttribuCartouche("X/X")).TextString = "4/4"


    InsertPointLigneTableau_fils(0) = -312#: InsertPointLigneTableau_fils(1) = -820#: InsertPointLigneTableau_fils(2) = 0
    Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & LeCartoucheE, InsertPointLigneTableau_fils, "LeCartouche4E")
    AttClent = NewBlock.GetAttributes
    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    For iClient = 0 To NbContolClient
        ControlTag = Split(MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient)).tag, ";")
        a = MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient))
        AttClent(AttribuCartouche(UCase(ControlTag(0)))).TextString = Replace(MyCARTOUCHE_Encelade.Controls("txt" & CStr(iClient)), vbCrLf, Right(vbCrLf, 1))
    Next iClient
    
    AttClent(AttribuCartouche("ETAT")).TextString = Status
    AttClent(AttribuCartouche("X/X")).TextString = "3/4"
    CartoucheCleient = True
    Set Fso = Nothing
End Function
Public Function ChargeCartoucheClient(MyCARTOUCHE_Client) As Boolean

If boolFormClient = False Then Exit Function
    Dim Fso As New FileSystemObject
    CartoucheCleient = False
    InsertPointLigneTableau_fils(0) = -165.0409: InsertPointLigneTableau_fils(1) = 126.0753: InsertPointLigneTableau_fils(2) = 0
    If Fso.FileExists(TableauPath.Item("PathBlocs") & LeCartouche) = False Then
        Set Fso = Nothing
        Exit Function
    End If
    Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & LeCartouche, InsertPointLigneTableau_fils, "LeCartouche1")
    AttClent = NewBlock.GetAttributes

    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    For iClient = 0 To NbContolClient
        ControlTag = Split(MyCARTOUCHE_Client.Controls("txt" & CStr(iClient)).tag, ";")
        a = MyCARTOUCHE_Client.Controls("txt" & CStr(iClient))
        AttClent(AttribuCartouche(UCase(ControlTag(0)))).TextString = Replace(MyCARTOUCHE_Client.Controls("txt" & CStr(iClient)), vbCrLf, Right(vbCrLf, 1))
    Next iClient
    AttClent(AttribuCartouche("X/X")).TextString = "1/4"

    InsertPointLigneTableau_fils(0) = 1022.9591: InsertPointLigneTableau_fils(1) = 126.0753: InsertPointLigneTableau_fils(2) = 0
    Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & LeCartouche, InsertPointLigneTableau_fils, "LeCartouche2")
    AttClent = NewBlock.GetAttributes
    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)

    For iClient = 0 To NbContolClient
        ControlTag = Split(MyCARTOUCHE_Client.Controls("txt" & CStr(iClient)).tag, ";")
        a = MyCARTOUCHE_Client.Controls("txt" & CStr(iClient))
        AttClent(AttribuCartouche(UCase(ControlTag(0)))).TextString = Replace(MyCARTOUCHE_Client.Controls("txt" & CStr(iClient)), vbCrLf, Right(vbCrLf, 1))
    Next iClient

    AttClent(AttribuCartouche("X/X")).TextString = "2/4"

    InsertPointLigneTableau_fils(0) = 1022.9591: InsertPointLigneTableau_fils(1) = -713.9247: InsertPointLigneTableau_fils(2) = 0
    Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & LeCartouche, InsertPointLigneTableau_fils, "LeCartouche3")
    AttClent = NewBlock.GetAttributes
    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    For iClient = 0 To NbContolClient
        ControlTag = Split(MyCARTOUCHE_Client.Controls("txt" & CStr(iClient)).tag, ";")
        a = MyCARTOUCHE_Client.Controls("txt" & CStr(iClient))
        AttClent(AttribuCartouche(UCase(ControlTag(0)))).TextString = Replace(MyCARTOUCHE_Client.Controls("txt" & CStr(iClient)), vbCrLf, Right(vbCrLf, 1))
    Next iClient
    
    AttClent(AttribuCartouche("X/X")).TextString = "4/4"
    
    InsertPointLigneTableau_fils(0) = -165.0409: InsertPointLigneTableau_fils(1) = -713.9247: InsertPointLigneTableau_fils(2) = 0
    Set NewBlock = FunInsBlock(TableauPath.Item("PathBlocs") & LeCartouche, InsertPointLigneTableau_fils, "LeCartouche4")
    AttClent = NewBlock.GetAttributes
    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    
    For iClient = 0 To NbContolClient
        ControlTag = Split(MyCARTOUCHE_Client.Controls("txt" & CStr(iClient)).tag, ";")
        a = MyCARTOUCHE_Client.Controls("txt" & CStr(iClient))
        AttClent(AttribuCartouche(UCase(ControlTag(0)))).TextString = Replace(MyCARTOUCHE_Client.Controls("txt" & CStr(iClient)), vbCrLf, Right(vbCrLf, 1))
    Next iClient
    
    AttClent(AttribuCartouche("X/X")).TextString = "3/4"
    CartoucheCleient = True
    Set Fso = Nothing
End Function
Public Sub Maj(MyControl As Object)
Dim Rs As Recordset
Dim Sql As String
Con.OpenConnetion db
MyControl.Clear
Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"
MyControl.AddItem ""
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    MyControl.AddItem Trim("" & Rs!Client)
        If MyControl.ListCount = 1 Then MyControl.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Con.CloseConnection
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
  Set AutoApp = ThisDrawing.Application

Open AutoApp.ActiveDocument.Path & "\Autocable.ini" For Input As #FileNumber
Do While Not EOF(FileNumber)    ' Effectue la boucle jusqu'à la fin du fichier.
    Input #FileNumber, MyString ' Lit les données dans deux variables.
    If InStr(1, MyString, Cherher) <> 0 Then
       CherCheInFihier = Right(MyString, Len(MyString) - (Len(Cherher) + InStr(1, MyString, Cherher)))
       Exit Do
    End If
Loop
Close #FileNumber    ' Ferme le fichier
CherCheInFihier = Replace(CherCheInFihier, "§remplaceDate§", Format(Date, "yyyy"))
End Function
Public Function RenseigneConnecteurBroches(RangeAttribue As Recordset) As Boolean
    On Error Resume Next
    RenseigneConnecteurBroches = False
    If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).ConnecteurExiste = True Then
        If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).EPISSURE = False Then
            a = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).NewBlock.GetAttributes
            
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("LIAI" & RangeAttribue.Fields(9))
            If Err Then
                FunError 5, "FIL" & RangeAttribue.Fields(9), Err.Description, CStr(RangeAttribue.Fields(8))
                Err.Clear
            End If
            dd = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Count
            a(IdAttrib).TextString = RangeAttribue.Fields(0)
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("FIL" & RangeAttribue.Fields(9))
            If Trim("" & a(IdAttrib).TextString) <> "" Then
                IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("MAR" & RangeAttribue.Fields(9))
            End If
            a(IdAttrib).TextString = RangeAttribue.Fields(2)
        Else
            If FunEPISSURE(TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).NewBlock.GetAttributes, RangeAttribue.Fields(2), RangeAttribue.Fields(9), CLng(CollectionCon("" & RangeAttribue.Fields("APP")))) = False Then Exit Function
        End If

    Else
        FunError 3, RangeAttribue.Fields(2), Err.Description, RangeAttribue.Fields(8)
    End If
    If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).ConnecteurExiste = True Then

        If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).EPISSURE = False Then
            a = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).NewBlock.GetAttributes
            
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("LIAI" & RangeAttribue.Fields(12))
            If Err Then
                FunError 5, "FIL" & RangeAttribue.Fields(12), Err.Description, CStr(RangeAttribue.Fields(11))
                Err.Clear
            End If
        
            dd = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Count
            a(IdAttrib).TextString = RangeAttribue.Fields(0)
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("FIL" & RangeAttribue.Fields(12))
            If Trim("" & a(IdAttrib).TextString) <> "" Then
                IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("MAR" & RangeAttribue.Fields(12))
            End If
            a(IdAttrib).TextString = RangeAttribue.Fields(2)
        Else
            If FunEPISSURE(TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).NewBlock.GetAttributes, RangeAttribue.Fields(2), RangeAttribue.Fields(12), CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))) = False Then Exit Function

        End If
    Else
        FunError 3, RangeAttribue.Fields(2), Err.Description, RangeAttribue.Fields(11)
    End If
    RenseigneConnecteurBroches = True
End Function
Sub AttribueLibV(NewBlock As AcadBlockReference, Index As Long)
    Dim At
    Dim MyAtt As New Collection
    At = NewBlock.GetAttributes
    
    For i = 0 To UBound(At)
        DoEvents
        AutoApp.ZoomAll
        Debug.Print At(i).TagString
        MyAtt.Add CStr(i), At(i).TagString
    Next i
    Set TableauDeConnecteurs(Index).AttribuesFils = MyAtt
    Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
    At(TableauDeConnecteurs(Index).AttribuesFils("DESIGNATION")).TextString = Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")"
    Set MyAtt = Nothing
    i3 = UBound(At) - 1
    Dim i2 As Long
    For i2 = 0 To i3
        Debug.Print At(i2).TagString
    Next i2
    For i = 1 To TableauDeConnecteurs(Index).indexFile
        DoEvents
        AutoApp.ZoomAll
        At(TableauDeConnecteurs(Index).AttribuesFils("FIL" & CStr(i))).TextString = TableauDeConnecteurs(Index).TableauFile(i)
    Next i
End Sub
Sub TestFl()
a = CherCheInFihier("Bdnumero")
End Sub
