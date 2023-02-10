Attribute VB_Name = "ModifierPlan"
Public Function ModifierUnOtil(IdIndiceProjet As Long) As Boolean
  Dim Sql As String
     Dim PathPl As String
    Dim Rs As Recordset
    Dim PathDessin As String
    Dim Fso As New FileSystemObject
 
    Dim Collec As New Collection
    Dim Etiquettes As New Collection
    Dim MyFichier As String
     Dim NewBlockV  As AcadBlockReference
     Set CollectionCon = Nothing
     Set CollectionCon = New Collection
     Set CollectionComp = Nothing
     Set CollectionComp = New Collection
    Set CollectionNota = Nothing
    Set CollectionNota = New Collection
    Set CollectionNoeuds = Nothing
    Set CollectionNoeuds = New Collection
     NbLignesVignette = 0
     ModifierUnOtil = False
     
   NbConnecteur = 0
       InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 790.0288: InsertPointLigneTableau_Vignette(2) = 0

Sql = "SELECT T_indiceProjet.PlAutoCadSave, T_indiceProjet.PlAutoCadSaveas "
Sql = Sql & "FROM T_Projet INNER JOIN (T_Pieces INNER  "
Sql = Sql & "JOIN T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces)  "
Sql = Sql & "ON T_Projet.id = T_Pieces.IdProjet "
Sql = Sql & "WHERE T_Projet.id=" & IdProjet & "  "
Sql = Sql & "AND T_Pieces.Id=" & IdPieces & ";"

Sql = "SELECT T_indiceProjet.ouAutoCadSave,  "
Sql = Sql & "T_indiceProjet.ouAutoCadSaveas "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"


    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
        Exit Function
     Else
        If Trim("" & Rs!OuAutoCadSaveAs) <> "" Then
            MyFichier = "" & Rs!OuAutoCadSaveAs
        End If
        If Trim("" & Rs!OuAutoCadSave) <> "" Then
            MyFichier = "" & Rs!OuAutoCadSave
        End If
        If Trim("" & MyFichier) = "" Then Exit Function
    End If
    
  MyFichier = Replace(MyFichier, ".dwg", "")
    PathDessin = PathArchiveAutocad & Trim("" & MyFichier) & ".dwg"
     If Fso.FileExists(PathDessin) = False Then Exit Function
         NbLignes = 0
'    'Set AutoApp = ThisDrawing.Application
 OpenFichier PathDessin
 
       FormBarGrah.ProgressBar1Caption = " Scanne des Connecteurs:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
  For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
  
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
                    If IsConnecteurs(Attributes) = True Then
  
  Set Collec = ColectionAttribueConecteur(Attributes)
  
  Debug.Print Attributes(Collec("N°")).TextString
  On Error Resume Next
  a = CollectionCon(Attributes(Collec("CODE_APP")).TextString)
  If Err Then
  Err.Clear
  NbConnecteur = NbConnecteur + 1
  CollectionCon.Add NbConnecteur, Attributes(Collec("CODE_APP")).TextString
  On Error GoTo 0
End If
  ReDim Preserve TableauDeConnecteurs(NbConnecteur)
  
   Set TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).NewBlock = BlocRef
    Set TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).Attribues = Collec
    TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).ConnecteurExiste = True
     TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).Kill = True
     Debug.Print TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).NewBlock.Name
  
            TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).EPISSURE = IsEpissures(Attributes)
 
            End If
            End If
        End If
    Next i
    
    
    FormBarGrah.ProgressBar1Caption = " Scanne des Vignettes :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
  For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
                a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes

                Set Collec = ColectionAttribueConecteur(Attributes)

                If IsVignette(Attributes) = True Then
                    Set Collec = ColectionAttribueConecteur(Attributes)
                    On Error Resume Next
                     a = CollectionCon(Attributes(Collec("CODE_APP")).TextString)
                    If Err Then
                        Err.Clear
                        NbConnecteur = NbConnecteur + 1
                        CollectionCon.Add NbConnecteur, Attributes(Collec("CODE_APP")).TextString
                        On Error GoTo 0
                        ReDim Preserve TableauDeConnecteurs(NbConnecteur)
                    End If
  
                    Debug.Print Attributes(Collec("N°")).TextString
                        Set TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).NewVignette = BlocRef
                         Set TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).AttribuesVignette = Collec
                        DelAttribues Attributes
                        NbLignesVignette = NbLignesVignette + 1
                        InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
                Else
                    If IsVignetteEPISSURE(Attributes) = True Then
                    
                     Set Collec = ColectionAttribueConecteur(Attributes)
                        For i2 = 1 To UBound(TableauDeConnecteurs)
                        
                            If TableauDeConnecteurs(i2).EPISSURE = True Then
                                a = TableauDeConnecteurs(i2).NewBlock.GetAttributes
                                b = a(TableauDeConnecteurs(i2).Attribues("CODE_APP")).TextString
                                If a(TableauDeConnecteurs(i2).Attribues("CODE_APP")).TextString = Attributes(Collec("EPISSURE")).TextString Then
                                    Set TableauDeConnecteurs(i2).NewVignette = BlocRef
                                     Set TableauDeConnecteurs(i2).AttribuesVignette = Collec
                                    NbLignesVignette = NbLignesVignette + 1
                                    InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
                                End If
                            End If
                        Next i2
                  
                         
                    End If
                End If
            End If
      End If
      If NbLignesVignette = 15 Then
            InsertPointLigneTableau_Vignette(0) = -1146.1429
            InsertPointLigneTableau_Vignette(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_Vignette(1), 40)
            NbLignesVignette = 0
        End If
    Next i
     For i2 = 1 To UBound(TableauDeConnecteurs)
                        
                            If TableauDeConnecteurs(i2).EPISSURE = True Then
                                a = TableauDeConnecteurs(i2).NewBlock.GetAttributes
                                 DelAttribues a
                            End If
    Next i2
    
    
      FormBarGrah.ProgressBar1Caption = " Scanne des Etiquettes :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
  For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
                 a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes

                Set Collec = ColectionAttribueConecteur(Attributes)

                If IsVignetteEtiquette(Attributes) = True Then
               
                 Set NewBlockV = BlocRef
                Etiquettes.Add NewBlockV
                Set NewBlockV = Nothing
                   
                      
                End If
              
            End If
      End If
      If i > AutoApp.Documents(0).ModelSpace.Count - 1 Then Exit For
    Next i
   For i = 1 To Etiquettes.Count

    Set BlocRef = Etiquettes(i)
     BlocRef.Delete
   Next i
   
   FormBarGrah.ProgressBar1Caption = " Scanne des Composants :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
     NUMCOM = 0
  For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
        IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
                Attributes = BlocRef.GetAttributes
                
                Set Collec = ColectionAttribueConecteur(Attributes)
                
                If IsComposants(Attributes) = True Then
                    On Error Resume Next
                    
                    Set b = ColectionAttribueConecteur(Attributes)
                    a = ""
                    a = CollectionComp(b("NUMCOMP"))
                    If Err Then
                        If NUMCOM < CInt(Mid(Attributes(b("NUMCOMP")).TextString, 2, Len(Attributes(b("NUMCOMP")).TextString) - 1)) Then
                         NUMCOM = CInt(Mid(Attributes(b("NUMCOMP")).TextString, 2, Len(Attributes(b("NUMCOMP")).TextString) - 1))
                            ReDim Preserve TableauDeComposants(NUMCOM)
                           
                           
                        End If
                    End If
                     CollectionComp.Add CInt(Mid(Attributes(b("NUMCOMP")).TextString, 2, Len(Attributes(b("NUMCOMP")).TextString) - 1)), UCase(Attributes(b("NUMCOMP")).TextString)
                    Set TableauDeComposants(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString))).NewBlock = BlocRef
                    TableauDeComposants(CollectionComp(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString)))).InsertPointLigneC(0) = BlocRef.InsertPointLigne(0)
                    TableauDeComposants(CollectionComp(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString)))).InsertPointLigneC(1) = BlocRef.InsertPointLigne(1)
                    TableauDeComposants(CollectionComp(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString)))).InsertPointLigneC(2) = BlocRef.InsertPointLigne(2)
                    
                    Set TableauDeComposants(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString))).Attribues = ColectionAttribueConecteur(Attributes)
                    
                    TableauDeComposants(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString))).RotationC = BlocRef.Rotation
                    
                    TableauDeComposants(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString))).XScaleFactorC = BlocRef.XScaleFactor
                    TableauDeComposants(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString))).YScaleFactorC = BlocRef.YScaleFactor
                    TableauDeComposants(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString))).ZScaleFactorC = BlocRef.ZScaleFactor
                    TableauDeComposants(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString))).ComposantsExiste = True
                    TableauDeComposants(CollectionComp(UCase(Attributes(b("NUMCOMP")).TextString))).Kill = True
                    Set BlocRef = Nothing
                
                
                End If
            
            End If
        End If
        If i > AutoApp.Documents(0).ModelSpace.Count - 1 Then Exit For
    Next i
   For i = 1 To UBound(TableauDeComposants)
InsertionPoint = TableauDeComposants(i).NewBlock.InsertionPoint
 TableauDeComposants(i).InsertPointLigneC(0) = InsertionPoint(0)
 TableauDeComposants(i).InsertPointLigneC(1) = InsertionPoint(1)
 TableauDeComposants(i).InsertPointLigneC(2) = InsertionPoint(2)
 Set BlocRef = TableauDeComposants(i).NewBlock
        TableauDeComposants(i).RotationC = BlocRef.Rotation
        
         TableauDeComposants(i).XScaleFactorC = BlocRef.XScaleFactor
         TableauDeComposants(i).YScaleFactorC = BlocRefYScaleFactor
       TableauDeComposants(i).ZScaleFactorC = TBlocRefZScaleFactor
   
     BlocRef.Delete
   Next i
   

    FormBarGrah.ProgressBar1Caption = " Scanne des Notas :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
     NUMNOTA = 0
  For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
        IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
                Attributes = BlocRef.GetAttributes
                
                Set Collec = ColectionAttribueConecteur(Attributes)
                
                If IsNotas(Attributes) = True Then
                    On Error Resume Next
                    
                    Set b = ColectionAttribueConecteur(Attributes)
                    a = ""
                    a = CollectionComp(b("NUMNOTA"))
                    If Err Then
                        If NUMNOTA < Attributes(b("NUMNOTA")).TextString Then
                         NUMNOTA = Attributes(b("NUMNOTA")).TextString
                            ReDim Preserve TableauDeNotas(NUMNOTA)
                                                
                        End If
                    End If
                    CollectionNota.Add Attributes(b("NUMNOTA")).TextString, "N" & UCase(Attributes(b("NUMNOTA")).TextString)
                    Set TableauDeNotas(CollectionNota(UCase("N" & Attributes(b("NUMNOTA")).TextString))).NewBlock = BlocRef
                    TableauDeNotas(CollectionNota(CollectionComp("N" & UCase(Attributes(b("NUMNOTA")).TextString)))).InsertPointLigneC(0) = BlocRef.InsertPointLigne(0)
                    TableauDeNotas(CollectionNota(CollectionComp("N" & UCase(Attributes(b("NUMNOTA")).TextString)))).InsertPointLigneC(1) = BlocRef.InsertPointLigne(1)
                    TableauDeNotas(CollectionNota(CollectionComp("N" & UCase(Attributes(b("NUMNOTA")).TextString)))).InsertPointLigneC(2) = BlocRef.InsertPointLigne(2)
                    
                    Set TableauDeNotas(CollectionNota(UCase("N" & Attributes(b("NUMNOTA")).TextString))).Attribues = ColectionAttribueConecteur(Attributes)
                    
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(b("NUMNOTA")).TextString))).RotationC = BlocRef.Rotation
                    
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(b("NUMNOTA")).TextString))).XScaleFactorC = BlocRef.XScaleFactor
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(b("NUMNOTA")).TextString))).YScaleFactorC = BlocRef.YScaleFactor
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(b("NUMNOTA")).TextString))).ZScaleFactorC = BlocRef.ZScaleFactor
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(b("NUMNOTA")).TextString))).NotasExiste = True
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(b("NUMNOTA")).TextString))).Kill = True
                    Set BlocRef = Nothing
                
                
                End If
            
            End If
        End If
        If i > AutoApp.Documents(0).ModelSpace.Count - 1 Then Exit For
    Next i
   For i = 1 To UBound(TableauDeNotas)
InsertionPoint = TableauDeNotas(i).NewBlock.InsertionPoint
 TableauDeNotas(i).InsertPointLigneC(0) = InsertionPoint(0)
 TableauDeNotas(i).InsertPointLigneC(1) = InsertionPoint(1)
 TableauDeNotas(i).InsertPointLigneC(2) = InsertionPoint(2)
 Set BlocRef = TableauDeNotas(i).NewBlock
        TableauDeNotas(i).RotationC = BlocRef.Rotation
        
         TableauDeNotas(i).XScaleFactorC = BlocRef.XScaleFactor
         TableauDeNotas(i).YScaleFactorC = BlocRefYScaleFactor
       TableauDeNotas(i).ZScaleFactorC = TBlocRefZScaleFactor
   
     BlocRef.Delete
   Next i
   
   
   
   
     FormBarGrah.ProgressBar1Caption = " Scanne Tableau des Fils:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
       
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            a = BlocRef.Name
            If BlocRef.HasAttributes Then
                Attributes = BlocRef.GetAttributes
                If (UBound(Attributes) = 15) Or (UBound(Attributes) = 14) Or (UBound(Attributes) = 13) Or (UBound(Attributes) = 12) Then
                    If IsTableauFils(Attributes) = True Then
                         Set NewBlockV = BlocRef
                        Etiquettes.Add NewBlockV
                        Set NewBlockV = Nothing
                     Else
                        If IsEnteteTableauFils(Attributes) = True Then
                             Set NewBlockV = BlocRef
                            Etiquettes.Add NewBlockV
                            Set NewBlockV = Nothing
                        End If
                    End If
                    
                    Else
                    If IsNOMBRE_FILS(Attributes) = True Then
                         Set NewBlockV = BlocRef
                Etiquettes.Add NewBlockV
                Set NewBlockV = Nothing
                    End If
                End If
            End If
        End If
    Next i

   For i = 1 To Etiquettes.Count

    Set BlocRef = Etiquettes(i)
     BlocRef.Delete
   Next i
   
    FormBarGrah.ProgressBar1Caption = " Scanne des cartouches :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
         IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
        
        If Entity.ObjectName = "AcDbBlockReference" Then
        Set BlocRef = Entity
       
            If BlocRef.HasAttributes Then
            a = BlocRef.Name
                Attributes = BlocRef.GetAttributes
                
                If IsCartoucheEncelade(Attributes) = True Then
                    Set NewBlockV = BlocRef
                    Etiquettes.Add NewBlockV
                    Set NewBlockV = Nothing
                Else
                    If IsCartoucheClient(Attributes) = True Then
                        Set NewBlockV = BlocRef
                        Etiquettes.Add NewBlockV
                        Set NewBlockV = Nothing
                    End If
                End If
            
            End If
        End If
    Next i

   For i = 1 To Etiquettes.Count

    Set BlocRef = Etiquettes(i)
     BlocRef.Delete
   Next i
  
  
  AutoApp.Documents(0).PurgeAll
   ModifierUnOtil = True
   Exit Function
  
End Function
Public Function ModifierUnPlan(IdIndiceProjet As Long, Mytype As String) As Boolean
    Dim Sql As String
     Dim PathPl As String
    Dim Rs As Recordset
    Dim PathDessin As String
    Dim Fso As New FileSystemObject
    Dim NewBlock  As AcadBlockReference
    Dim Collec As New Collection
    Dim Etiquettes As New Collection
    Dim BoolBlocValider As Boolean
'    Dim Collec As Collection
    Dim MyFichier As String
     Dim BlocRef  As AcadBlockReference
     Set CollectionCon = Nothing
     Set CollectionCon = New Collection
     Set CollectionComp = Nothing
     Set CollectionComp = New Collection
    Set CollectionNota = Nothing
    Set CollectionNota = New Collection
     Set CollectionNoeuds = Nothing
     Set CollectionNoeuds = New Collection
     Set CollectionFils = Nothing
     Set CollectionFils = New Collection
     Set CollectionEtiquettes = Nothing
     Set CollectionEtiquettes = New Collection
     Set RefOption = Nothing
     Set RefOption = New Collection
     Set RefCriteres = Nothing
     Set RefCriteres = New Collection
     Set CollectionChartouche = Nothing
    Set CollectionChartouche = New Collection
    ReDim TableauDeNoeuds(0)
        NUMNETT = 0
      NUMCOM = 0
      NUMNOTA = 0
      NUMNTOR = 0
    NUMNOEUDS = 0
     NbLignesVignette = 0
     ModifierUnPlan = False
     
   NbConnecteur = 0
       InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 790.0288: InsertPointLigneTableau_Vignette(2) = 0

Sql = "SELECT T_indiceProjet." & Mytype & "AutoCadSave, T_indiceProjet." & Mytype & "AutoCadSaveas "
Sql = Sql & "FROM T_Projet INNER JOIN (T_Pieces INNER  "
Sql = Sql & "JOIN T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces)  "
Sql = Sql & "ON T_Projet.id = T_Pieces.IdProjet "
Sql = Sql & "WHERE T_Projet.id=" & IdProjet & "  "
Sql = Sql & "AND T_Pieces.Id=" & IdPieces & ";"

Sql = "SELECT T_indiceProjet." & Mytype & "AutoCadSave,  "
Sql = Sql & "T_indiceProjet." & Mytype & "AutoCadSaveas "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"


    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
        Exit Function
     Else
        If Trim("" & Rs(Mytype & "AutoCadSaveas")) <> "" Then
            MyFichier = "" & Rs(Mytype & "AutoCadSaveas")
        End If
        If Trim("" & Rs(Mytype & "AutoCadSave")) <> "" Then
            MyFichier = "" & Rs(Mytype & "AutoCadSave")
        End If
        If Trim("" & MyFichier) = "" Then Exit Function
    End If
    
    Set TableauPath = funPath
    MyFichier = Replace(MyFichier, ".dwg", "")
    PathDessin = TableauPath("PathArchiveAutocad") & Trim("" & MyFichier) & ".dwg"
    If Left(PathDessin, 2) <> "\\" And Left(PathDessin, 1) = "\" Then PathDessin = TableauPath.Item("PathServer") & PathDessin
     If Fso.FileExists(PathDessin) = False Then Exit Function
         NbLignes = 0
'    'Set AutoApp = ThisDrawing.Application
'AutoApp.Visible = True
BackUp PathDessin
OpenFichier PathDessin
If Mytype = "PL" Then
    FormBarGrah.ProgressBar1Caption = " Scanne le Plan :"
Else
    FormBarGrah.ProgressBar1Caption = " Scanne l'Outil :"
End If
'    FormBarGrah.ProgressBar1Caption = " Scanne des cartouches :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
     LoadCalque
     For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
  
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
        If BlocRef.HasAttributes Then
             Set Collec = ColectionAttribueConecteur(BlocRef.GetAttributes)
            aa = BlocRef.Name
            
'             If "8200192000" = aa Then
'                MsgBox ""
'             End If
                BoolBlocValider = LectureNoeuds(Mytype, BlocRef, Collec)
            
            If BoolBlocValider = False Then
            If LectureNotas(Mytype, BlocRef, Collec) = True Then BoolBlocValider = True
            End If
            
             If BoolBlocValider = False Then
                If LectureFils(Mytype, BlocRef) = True Then BoolBlocValider = True
            End If
            If BoolBlocValider = False Then
                If LectureOptions(Mytype, BlocRef) = True Then BoolBlocValider = True
            End If
             If BoolBlocValider = False Then
                If LectureComposants(Mytype, BlocRef, Collec) = True Then BoolBlocValider = True
            End If
            
            If BoolBlocValider = False Then
                If LectureCritères(Mytype, BlocRef) = True Then BoolBlocValider = True
            End If
            
           
            
            If BoolBlocValider = False Then
                If LectureEtiquettes(Mytype, BlocRef, Collec) = True Then BoolBlocValider = True
            End If
           
            
            If BoolBlocValider = False Then
                If LecturePréconisations(Mytype, BlocRef, Collec) = True Then BoolBlocValider = True
            End If
            
            If BoolBlocValider = False Then
                If LectureCartouches(Mytype, BlocRef) = True Then BoolBlocValider = True
            End If
          
            
           If BoolBlocValider = False Then
                If LectureConnecteur(Mytype, BlocRef, Collec) = True Then BoolBlocValider = True
                If LectureVignettes(Mytype, BlocRef, Collec) = True Then BoolBlocValider = True
            End If
        End If
    End If
   Next
   
  
   
   KillConnecteur
   KillFils
   KillEtiquettes
   KillComposant
   KillNotas
   KillPreco
   KillOption
   KillCriteres
   KillCartouche
   KillNoeuds


 


   ModifierUnPlan = True
  
  
End Function
Sub KillConnecteur()
 On Error Resume Next
 For i = 1 To UBound(TableauDeConnecteurs)
    
    If TableauDeConnecteurs(i).ConnecteurExiste = True Then

        TableauDeConnecteurs(i).PosOk = True

         InsertionPoint = TableauDeConnecteurs(i).NewBlock.InsertionPoint
        TableauDeConnecteurs(i).InsertPointLigneC(0) = InsertionPoint(0)
        TableauDeConnecteurs(i).InsertPointLigneC(1) = InsertionPoint(1)
        TableauDeConnecteurs(i).InsertPointLigneC(2) = 1

        InsertionPoint = TableauDeConnecteurs(i).NewVignette.InsertionPoint
        TableauDeConnecteurs(i).InsertPointLigneV(0) = InsertionPoint(0)
        TableauDeConnecteurs(i).InsertPointLigneV(1) = InsertionPoint(1)
        TableauDeConnecteurs(i).InsertPointLigneV(2) = 1
        TableauDeConnecteurs(i).RotationV = 0
        TableauDeConnecteurs(i).RotationC = 0
        TableauDeConnecteurs(i).RotationV = TableauDeConnecteurs(i).NewVignette.Rotation
        TableauDeConnecteurs(i).RotationC = TableauDeConnecteurs(i).NewBlock.Rotation

        TableauDeConnecteurs(i).XScaleFactorC = TableauDeConnecteurs(i).NewBlock.XScaleFactor
        TableauDeConnecteurs(i).YScaleFactorC = TableauDeConnecteurs(i).NewBlock.YScaleFactor
        TableauDeConnecteurs(i).ZScaleFactorC = TableauDeConnecteurs(i).NewBlock.ZScaleFactor

        TableauDeConnecteurs(i).XScaleFactorV = TableauDeConnecteurs(i).NewVignette.XScaleFactor
        TableauDeConnecteurs(i).YScaleFactorV = TableauDeConnecteurs(i).NewVignette.YScaleFactor
        TableauDeConnecteurs(i).ZScaleFactorV = TableauDeConnecteurs(i).NewVignette.ZScaleFactor

        TableauDeConnecteurs(i).NewVignette.Delete
        Set TableauDeConnecteurs(i).NewVignette = Nothing
        TableauDeConnecteurs(i).NewBlock.Delete
        Set TableauDeConnecteurs(i).NewBlock = Nothing
        TableauDeConnecteurs(i).ConnecteurExiste = False

        
    End If

      DoEvents
  Next i
  AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillComposant()
   On Error Resume Next
   For i = 1 To CollectionComp.Count
   If TableauComposant(i).PosOkDesin = True Then
    InsertionPoint = TableauComposant(i).BlockDesing.InsertionPoint
    TableauComposant(i).InsertDesing(0) = InsertionPoint(0)
    TableauComposant(i).InsertDesing(1) = InsertionPoint(1)
    TableauComposant(i).InsertDesing(2) = InsertionPoint(2)
    Set BlocRef = TableauComposant(i).BlockDesing
     TableauComposant(i).XScaleFactorDesin = BlocRef.XScaleFactor
   TableauComposant(i).YScaleFactorDesin = BlocRef.YScaleFactor
   TableauComposant(i).ZScaleFactorDesin = BlocRef.ZScaleFactor
   TableauComposant(i).RotationDesin = BlocRef.Rotation
   BlocRef.Delete
    End If
        If TableauComposant(i).PosOkComp = True Then
    InsertionPoint = TableauComposant(i).BlockComp.InsertionPoint
    TableauComposant(i).InsertComp(0) = InsertionPoint(0)
    TableauComposant(i).InsertComp(1) = InsertionPoint(1)
    TableauComposant(i).InsertComp(2) = InsertionPoint(2)
    Set BlocRef = TableauComposant(i).BlockComp
   TableauComposant(i).XScaleFactorComp = BlocRef.XScaleFactor
   TableauComposant(i).YScaleFactorComp = BlocRef.YScaleFactor
   TableauComposant(i).ZScaleFactorComp = BlocRef.ZScaleFactor
   TableauComposant(i).RotationComp = BlocRef.Rotation
    BlocRef.Delete
   End If
'InsertionPoint = 'TableauDeComposants(i).NewBlock.InsertionPoint
' 'TableauDeComposants(i).InsertPointLigneC(0) = InsertionPoint(0)
' 'TableauDeComposants(i).InsertPointLigneC(1) = InsertionPoint(1)
' 'TableauDeComposants(i).InsertPointLigneC(2) = InsertionPoint(2)
' Set BlocRef = 'TableauDeComposants(i).NewBlock
'        'TableauDeComposants(i).RotationC = BlocRef.Rotation
'
'         'TableauDeComposants(i).XScaleFactorC = BlocRef.XScaleFactor
'         'TableauDeComposants(i).YScaleFactorC = BlocRef.YScaleFactor
'       'TableauDeComposants(i).ZScaleFactorC = BlocRef.ZScaleFactor
   
     
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillCartouche()
On Error Resume Next
   For i = 1 To CollectionChartouche.Count

    Set BlocRef = CollectionChartouche(i)
     BlocRef.Delete
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillCriteres()
On Error Resume Next
  For i = 1 To RefCriteres.Count

    Set BlocRef = RefCriteres(i)
     BlocRef.Delete
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillNoeuds()
On Error Resume Next
   For i = 1 To UBound(TableauDeNoeuds)
    If TableauDeNoeuds(i).PosOkComp = True Then
        Set NewBlock = TableauDeNoeuds(i).BlockComp
        InsertPoint = NewBlock.InsertionPoint
        TableauDeNoeuds(i).InsertComp(0) = InsertPoint(0)
        TableauDeNoeuds(i).InsertComp(1) = InsertPoint(1)
        TableauDeNoeuds(i).InsertComp(2) = InsertPoint(2)
        TableauDeNoeuds(i).RotationComp = NewBlock.Rotation
        TableauDeNoeuds(i).XScaleFactorComp = NewBlock.XScaleFactor
        TableauDeNoeuds(i).YScaleFactorComp = NewBlock.YScaleFactor
        TableauDeNoeuds(i).ZScaleFactorComp = NewBlock.ZScaleFactor
         NewBlock.Delete
    End If
     If TableauDeNoeuds(i).PosOkDesin = True Then
        Set NewBlock = TableauDeNoeuds(i).BlockDesing
        InsertPoint = NewBlock.InsertionPoint
        TableauDeNoeuds(i).InsertDesing(0) = InsertPoint(0)
        TableauDeNoeuds(i).InsertDesing(1) = InsertPoint(1)
        TableauDeNoeuds(i).InsertDesing(2) = InsertPoint(2)
        TableauDeNoeuds(i).RotationDesin = NewBlock.Rotation
        TableauDeNoeuds(i).XScaleFactorDesin = NewBlock.XScaleFactor
        TableauDeNoeuds(i).YScaleFactorDesin = NewBlock.YScaleFactor
        TableauDeNoeuds(i).ZScaleFactorDesin = NewBlock.ZScaleFactor
         NewBlock.Delete
    End If
     If TableauDeNoeuds(i).PosOkFleche = True Then
        Set NewBlock = TableauDeNoeuds(i).BlockFleche
        InsertPoint = NewBlock.InsertionPoint
        TableauDeNoeuds(i).InsertFleche(0) = InsertPoint(0)
        TableauDeNoeuds(i).InsertFleche(1) = InsertPoint(1)
        TableauDeNoeuds(i).InsertFleche(2) = InsertPoint(2)
        TableauDeNoeuds(i).RotationFleche = NewBlock.Rotation
        TableauDeNoeuds(i).XScaleFactorFleche = NewBlock.XScaleFactor
        TableauDeNoeuds(i).YScaleFactorFleche = NewBlock.YScaleFactor
        TableauDeNoeuds(i).ZScaleFactorFleche = NewBlock.ZScaleFactor
         NewBlock.Delete
    End If
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillOption()
On Error Resume Next
   For i = 1 To RefOption.Count

    Set BlocRef = RefOption(i)
     BlocRef.Delete
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillPreco()
 On Error Resume Next
   For i = 1 To CollectionTor.Count
   On Error Resume Next
        InsertionPoint = TableuDeTor(i).NewBlockTorTire.InsertionPoint
        TableuDeTor(i).InsertTorTitre(0) = InsertionPoint(0)
        TableuDeTor(i).InsertTorTitre(1) = InsertionPoint(1)
        TableuDeTor(i).InsertTorTitre(2) = InsertionPoint(2)
        Set BlocRef = TableuDeTor(i).NewBlockTorTire
        TableuDeTor(i).Rotation = BlocRef.Rotation
       
        
        TableuDeTor(i).XScaleFactor = BlocRef.XScaleFactor
        TableuDeTor(i).YScaleFactor = BlocRef.YScaleFactor
        TableuDeTor(i).ZScaleFactor = BlocRef.ZScaleFactor
        
        BlocRef.Delete
         
         
         For i2 = 1 To TableuDeTor(i).NumTor
                InsertionPoint = TableuDeTor(i).Tor(i2).NewBlockTorDetail.InsertionPoint
                TableuDeTor(i).Tor(i2).Insert(0) = InsertionPoint(0)
                TableuDeTor(i).Tor(i2).Insert(1) = InsertionPoint(1)
                TableuDeTor(i).Tor(i2).Insert(2) = InsertionPoint(2)
                Set BlocRef = TableuDeTor(i).Tor(i2).NewBlockTorDetail
                TableuDeTor(i).Tor(i2).Rotation = BlocRef.Rotation
                
                
                TableuDeTor(i).Tor(i2).XScaleFactor = BlocRef.XScaleFactor
                TableuDeTor(i).Tor(i2).YScaleFactor = BlocRef.YScaleFactor
                TableuDeTor(i).Tor(i2).ZScaleFactor = BlocRef.ZScaleFactor
                
                BlocRef.Delete
         Next i2
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillNotas()
On Error Resume Next
   For i = 1 To CollectionNota.Count
InsertionPoint = TableauDeNotas(i).NewBlock.InsertionPoint
 TableauDeNotas(i).InsertPointLigneC(0) = InsertionPoint(0)
 TableauDeNotas(i).InsertPointLigneC(1) = InsertionPoint(1)
 TableauDeNotas(i).InsertPointLigneC(2) = InsertionPoint(2)
 Set BlocRef = TableauDeNotas(i).NewBlock
        TableauDeNotas(i).RotationC = BlocRef.Rotation
        
         TableauDeNotas(i).XScaleFactorC = BlocRef.XScaleFactor
         TableauDeNotas(i).YScaleFactorC = BlocRef.YScaleFactor
       TableauDeNotas(i).ZScaleFactorC = BlocRef.ZScaleFactor
   
     BlocRef.Delete
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillEtiquettes()
On Error Resume Next
   For i = 1 To UBound(TableuEtiquettes)
    Set BlocRef = TableuEtiquettes(i).NewBlockTorTire
    Insert = BlocRef.InsertionPoint
    TableuEtiquettes(i).InsertTorTitre(0) = Insert(0)
    TableuEtiquettes(i).InsertTorTitre(1) = Insert(1)
    TableuEtiquettes(i).InsertTorTitre(2) = Insert(2)
    TableuEtiquettes(i).Rotation = BlocRef.Rotation
    TableuEtiquettes(i).YScaleFactor = BlocRef.YScaleFactor
    TableuEtiquettes(i).ZScaleFactor = BlocRef.ZScaleFactor
     BlocRef.Delete
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Sub KillFils()
On Error Resume Next
   For i = 1 To CollectionFils.Count

    Set BlocRef = CollectionFils(i)
     BlocRef.Delete
   Next i
   AutoApp.Documents(0).PurgeAll
   On Error GoTo 0
End Sub
Function EnteteCartouche()
    Dim txt
    Dim txt2
    Dim Mysapce
    Mysapce = Space(65)
    txt = "******************************************************************" & vbCrLf
    txt = txt & "* Listes des erreurs survenues lors de l'exécution de la macro : *" & vbCrLf
    txt = txt & "* Créer un Plan                                                  *" & vbCrLf
    txt2 = "* Projet : " & varProjet & " Indice : " & varIndice
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
    txt = txt & "******************************************************************" & vbCrLf
    txt = txt & vbCrLf
    EnteteCartouche = txt
End Function
Function LectureConnecteur(MyOption As String, BlocRef As AcadBlockReference, Collec As Collection) As Boolean

  Attributes = BlocRef.GetAttributes
  If UBound(Attributes) < 6 Then Exit Function
LectureConnecteur = IsConnecteurs(Attributes)
If (bool_Plan_L_Connecteurs = False And MyOption = "PL") Or (bool_Outil_L_Connecteurs = False And MyOption = "OU") Then Exit Function
'AutoApp.Documents(0).PurgeAll
'  FormBarGrah.ProgressBar1Caption = " Scanne des Connecteurs:"
'     FormBarGrah.ProgressBar1.Value = 0
'     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
'  For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
  
'     IncremanteBarGrah FormBarGrah
'    DoEvents
'        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
'        If Entity.ObjectName = "AcDbBlockReference" Then
'            Set BlocRef = Entity
            
         
                
If LectureConnecteur = True Then
  
  
  Debug.Print Attributes(Collec("N°")).TextString
  On Error Resume Next
  a = ""
  a = CollectionCon(Attributes(Collec("CODE_APP")).TextString)
  If Err Then
  Err.Clear
  NbConnecteur = NbConnecteur + 1
  CollectionCon.Add NbConnecteur, Attributes(Collec("CODE_APP")).TextString
  On Error GoTo Fin
End If
  ReDim Preserve TableauDeConnecteurs(NbConnecteur)
  
   Set TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).NewBlock = BlocRef
    Set TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).Attribues = Collec
    TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).ConnecteurExiste = True
     TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).Kill = True
     Debug.Print TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).NewBlock.Name
  
            TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).EPISSURE = IsEpissures(Attributes)
 
            End If
            
'        End If
'Reprise:
'    Next i
    

'
'
''
''     For i2 = 1 To UBound(TableauDeConnecteurs)
''
''                            If TableauDeConnecteurs(i2).EPISSURE = True Then
''                                A = TableauDeConnecteurs(i2).NewBlock.GetAttributes
''                                 DelAttribues A
''                            End If
''    Next i2
'
'     For i = 1 To UBound(TableauDeConnecteurs)
'     On Error Resume Next
'    If TableauDeConnecteurs(i).ConnecteurExiste = True Then
'
'        TableauDeConnecteurs(i).PosOk = True
'
'         InsertionPoint = TableauDeConnecteurs(i).NewBlock.InsertionPoint
'        TableauDeConnecteurs(i).InsertPointLigneC(0) = InsertionPoint(0)
'        TableauDeConnecteurs(i).InsertPointLigneC(1) = InsertionPoint(1)
'        TableauDeConnecteurs(i).InsertPointLigneC(2) = 1
'
'        InsertionPoint = TableauDeConnecteurs(i).NewVignette.InsertionPoint
'        TableauDeConnecteurs(i).InsertPointLigneV(0) = InsertionPoint(0)
'        TableauDeConnecteurs(i).InsertPointLigneV(1) = InsertionPoint(1)
'        TableauDeConnecteurs(i).InsertPointLigneV(2) = 1
'        TableauDeConnecteurs(i).RotationV = 0
'        TableauDeConnecteurs(i).RotationC = 0
'        TableauDeConnecteurs(i).RotationV = TableauDeConnecteurs(i).NewVignette.Rotation
'        TableauDeConnecteurs(i).RotationC = TableauDeConnecteurs(i).NewBlock.Rotation
'
'        TableauDeConnecteurs(i).XScaleFactorC = TableauDeConnecteurs(i).NewBlock.XScaleFactor
'        TableauDeConnecteurs(i).YScaleFactorC = TableauDeConnecteurs(i).NewBlock.YScaleFactor
'        TableauDeConnecteurs(i).ZScaleFactorC = TableauDeConnecteurs(i).NewBlock.ZScaleFactor
'
'        TableauDeConnecteurs(i).XScaleFactorV = TableauDeConnecteurs(i).NewVignette.XScaleFactor
'        TableauDeConnecteurs(i).YScaleFactorV = TableauDeConnecteurs(i).NewVignette.YScaleFactor
'        TableauDeConnecteurs(i).ZScaleFactorV = TableauDeConnecteurs(i).NewVignette.ZScaleFactor
'
'        TableauDeConnecteurs(i).NewVignette.Delete
'        Set TableauDeConnecteurs(i).NewVignette = Nothing
'        TableauDeConnecteurs(i).NewBlock.Delete
'        Set TableauDeConnecteurs(i).NewBlock = Nothing
'        TableauDeConnecteurs(i).ConnecteurExiste = False
'
'        On Error GoTo 0
'    End If
'
'      DoEvents
'  Next i
'  AutoApp.Documents(0).PurgeAll
'
    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
    Resume Next
    
End Function
Function LectureEtiquettes(MyOption As String, BlocRef As AcadBlockReference, Collec As Collection) As Boolean
On Error GoTo Fin
 Attributes = BlocRef.GetAttributes
LectureEtiquettes = IsVignetteEtiquette(Attributes)
If LectureEtiquettes = False Then Exit Function
If (bool_Plan_L_Etiquettes = False And MyOption = "PL") Or (bool_Outil_L_Etiquettes = False And MyOption = "OU") Then Exit Function
                 a = BlocRef.Name
         
              
                If LectureEtiquettes = True Then
               
               a = ""
               On Error Resume Next
                        a = CollectionEtiquettes(Trim("E" & Attributes(Collec("DESIGNATION")).TextString))
                    If Err Then
                        Err.Clear
                        NUMNETT = NUMNETT + 1
                        CollectionEtiquettes.Add NUMNETT, Trim("E" & Attributes(Collec("DESIGNATION")).TextString)
                        ReDim Preserve TableuEtiquettes(NUMNETT)
                        On Error GoTo 0
                    End If
               
               
               
               
               Set TableuEtiquettes(CollectionEtiquettes(Trim("E" & Attributes(Collec("DESIGNATION")).TextString))).NewBlockTorTire = BlocRef
              
                   TableuEtiquettes(CollectionEtiquettes(Trim("E" & Attributes(Collec("DESIGNATION")).TextString))).PosOk = True
                End If
              
    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
    Resume Next
End Function
Function LectureComposants(MyOption As String, BlocRef As AcadBlockReference, Collec As Collection) As Boolean
On Error GoTo Fin
aa = BlocRef.Name
Attributes = BlocRef.GetAttributes
If UBound(Attributes) <> 3 Then Exit Function
 LectureComposants = IsComposants(Attributes)
If (bool_Plan_L_Composants = False And MyOption = "PL") Or (bool_Outil_L_Composants = False And MyOption = "OU") Then Exit Function

                Attributes = BlocRef.GetAttributes
                
              
                
                If LectureComposants = True Then
                  rrr = BlocRef.Name
                    
                   
                    a = ""
                    On Error Resume Next
                    a = CollectionComp(Attributes(Collec("NUMCOMP")).TextString)
                    If Err Then
                    Err.Clear
                        If NUMCOM < CInt(Mid(Attributes(Collec("NUMCOMP")).TextString, 2, Len(Attributes(Collec("NUMCOMP")).TextString) - 1)) Then
                         NUMCOM = CInt(Mid(Attributes(Collec("NUMCOMP")).TextString, 2, Len(Attributes(Collec("NUMCOMP")).TextString) - 1))
'
                            ReDim Preserve TableauComposant(NUMCOM)
                           
                        End If
                    End If
                   
                    CollectionComp.Add CInt(Mid(Attributes(Collec("NUMCOMP")).TextString, 2, Len(Attributes(Collec("NUMCOMP")).TextString) - 1)), UCase(Attributes(Collec("NUMCOMP")).TextString)
                    If UCase(rrr) = "COMP_DESGN" Then
                        Set TableauComposant(CollectionComp(UCase(Attributes(Collec("NUMCOMP")).TextString))).BlockDesing = BlocRef
                        TableauComposant(CollectionComp(UCase(Attributes(Collec("NUMCOMP")).TextString))).PosOkDesin = True
                    
                    Else
                        Set TableauComposant(CollectionComp(UCase(Attributes(Collec("NUMCOMP")).TextString))).BlockComp = BlocRef
                        TableauComposant(CollectionComp(UCase(Attributes(Collec("NUMCOMP")).TextString))).PosOkComp = True
                    End If
'                    Set BlocRef = Nothing
                 On Error GoTo Fin
                
                End If
            
'            End If
'        End If
'        If i > AutoApp.Documents(0).ModelSpace.Count - 1 Then Exit For
'Reprise:
'    Next i
 
    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
    Resume Next
End Function
Function LectureNotas(MyOption As String, BlocRef As AcadBlockReference, Collec As Collection) As Boolean
On erroro GoTo Fin
aaa = BlocRef.Name
 Attributes = BlocRef.GetAttributes
 If UBound(Attributes) <> 0 Then Exit Function
 LectureNotas = IsNotas(Attributes)
  If LectureNotas = False Then Exit Function
If (bool_Plan_L_Notas = False And MyOption = "PL") Or (bool_Outil_L_Notas = False And MyOption = "OU") Then Exit Function
'AutoApp.Documents(0).PurgeAll
' FormBarGrah.ProgressBar1Caption = " Scanne des Notas :"
'     FormBarGrah.ProgressBar1.Value = 0
'     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
'     NUMNOTA = 0
'  For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
'        IncremanteBarGrah FormBarGrah
'        DoEvents
'        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
'        If Entity.ObjectName = "AcDbBlockReference" Then
'            Set BlocRef = Entity
'            If BlocRef.HasAttributes Then
          
                Attributes = BlocRef.GetAttributes
                
                              
                If LectureNotas = True Then
                   
                    
                    
                    On Error Resume Next
                    a = ""
                    a = CollectionComp(Collec("NUMNOTA"))
                    If Err Then
                        If NUMNOTA < Attributes(Collec("NUMNOTA")).TextString Then
                         NUMNOTA = Attributes(Collec("NUMNOTA")).TextString
                            ReDim Preserve TableauDeNotas(NUMNOTA)
                           
                        End If
                    End If
                    CollectionNota.Add CInt(Attributes(Collec("NUMNOTA")).TextString), "N" & UCase(Attributes(Collec("NUMNOTA")).TextString)

                    Set TableauDeNotas(CollectionNota(UCase("N" & Attributes(Collec("NUMNOTA")).TextString))).NewBlock = BlocRef
                    TableauDeNotas(CollectionNota(CollectionComp("N" & UCase(Attributes(Collec("NUMNOTA")).TextString)))).InsertPointLigneC(0) = BlocRef.InsertPointLigne(0)
                    TableauDeNotas(CollectionNota(CollectionComp("N" & UCase(Attributes(Collec("NUMNOTA")).TextString)))).InsertPointLigneC(1) = BlocRef.InsertPointLigne(1)
                    TableauDeNotas(CollectionNota(CollectionComp("N" & UCase(Attributes(Collec("NUMNOTA")).TextString)))).InsertPointLigneC(2) = BlocRef.InsertPointLigne(2)
                    
                    Set TableauDeNotas(CollectionNota(UCase("N" & Attributes(Collec("NUMNOTA")).TextString))).Attribues = ColectionAttribueConecteur(Attributes)
                    
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(Collec("NUMNOTA")).TextString))).RotationC = BlocRef.Rotation
                    
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(Collec("NUMNOTA")).TextString))).XScaleFactorC = BlocRef.XScaleFactor
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(Collec("NUMNOTA")).TextString))).YScaleFactorC = BlocRef.YScaleFactor
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(Collec("NUMNOTA")).TextString))).ZScaleFactorC = BlocRef.ZScaleFactor
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(Collec("NUMNOTA")).TextString))).NotasExiste = True
                    TableauDeNotas(CollectionNota(UCase("N" & Attributes(Collec("NUMNOTA")).TextString))).Kill = True
'                    Set BlocRef = Nothing
                
                
                End If
            
'            End If
'        End If
'        If i > AutoApp.Documents(0).ModelSpace.Count - 1 Then Exit For
'Reprise:
'    Next i
  

    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
   Err.Clear
   Resume Next
End Function
Function LecturePréconisations(MyOption As String, BlocRef As AcadBlockReference, Collec As Collection) As Boolean
On Error Resume Next
  Attributes = BlocRef.GetAttributes
   LecturePréconisations = IsTor(Attributes)
  
   
If (bool_Plan_L_Preconisations = False And MyOption = "PL") Or (bool_Outil_L_Preconisations = False And MyOption = "OU") Then Exit Function
'AutoApp.Documents(0).PurgeAll
'   FormBarGrah.ProgressBar1Caption = " Scanne des Préconisations :"
'     FormBarGrah.ProgressBar1.Value = 0
'     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
'
'  For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
'        IncremanteBarGrah FormBarGrah
'        DoEvents
'        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
'        If Entity.ObjectName = "AcDbBlockReference" Then
'            Set BlocRef = Entity
'            If BlocRef.HasAttributes Then
          aa = BlocRef.Name
                Attributes = BlocRef.GetAttributes
                
               
                
If LecturePréconisations = True Then

    
    a = ""
    a = CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)
    If Err Then
        NUMNTORBLOC = NUMNTORBLOC + 1
        CollectionTor.Add NUMNTORBLOC, Attributes(Collec("TORDESIGNATION")).TextString
        ReDim Preserve TableuDeTor(NUMNTORBLOC)
        TableuDeTor(NUMNTORBLOC).CodeApp = Attributes(Collec("TORDESIGNATION")).TextString
        If IsTorDetail(Attributes) = False Then
            TableuDeTor(NUMNTORBLOC).TorExiste = True
            Set TableuDeTor(NUMNTORBLOC).NewBlockTorTire = BlocRef
        End If
        
    
    
    End If
    If IsTorDetail(Attributes) = True Then
    
        a = ""
        a = TableuDeTor(Attributes(Collec("TORDESIGNATION")).TextString).CollectionTor(Collec("TORNUM")).TextString
        If Err Then
            TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).NumTor = TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).NumTor + 1
            TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).CollectionTor.Add TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).NumTor, Attributes(Collec("TORNUM")).TextString
            
            ReDim Preserve TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).Tor(TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).NumTor)
            TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).Tor(TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).NumTor).TorExiste = True
            '                         TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).Tor(TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).NumTor).TorName = Attributes(Collec("TORNUM")).TextString
            
            Set TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).Tor(TableuDeTor(CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)).NumTor).NewBlockTorDetail = BlocRef
        End If
    
    End If
    
    '                    Set BlocRef = Nothing


End If
            
'            End If
'        End If
'        If i > AutoApp.Documents(0).ModelSpace.Count - 1 Then Exit For
'Reprise:
'    Next i
'
   
  
   On Error GoTo 0
    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
   Resume Next
End Function
Function LectureOptions(MyOption As String, BlocRef As AcadBlockReference) As Boolean
On Error GoTo Fin
 Attributes = BlocRef.GetAttributes
If UBound(Attributes) < 2 Then Exit Function
LectureOptions = IsRefOption(Attributes)
If LectureOptions = False Then Exit Function
If (bool_Plan_L_Options = False And MyOption = "PL") Or (bool_Outil_L_Options = False And MyOption = "OU") Then Exit Function
'AutoApp.Documents(0).PurgeAll
'    FormBarGrah.ProgressBar1Caption = " Scanne des Options :"
'     FormBarGrah.ProgressBar1.Value = 0
'     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
'    Set RefOption = Nothing
'    Set RefOption = New Collection
'    For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
'        IncremanteBarGrah FormBarGrah
'        DoEvents
'        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
'
'        If Entity.ObjectName = "AcDbBlockReference" Then
'            Set BlocRef = Entity
'            A = BlocRef.Name
'            If BlocRef.HasAttributes Then
            
               
                    If LectureOptions = True Then
                        
                        RefOption.Add BlocRef
'                        Set NewBlockV = Nothing
                    End If
            
               
'         End If
'      End If
'Reprise:
'    Next i

   On Error GoTo 0
    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
   Resume Next
End Function
Function LectureFils(MyOption As String, BlocRef As AcadBlockReference) As Boolean
On Error GoTo Fin
aa = UCase(BlocRef.Name)
If (UCase(BlocRef.Name) = "LIGNES TABLEAU DES FILS") Or (UCase(BlocRef.Name) = "LIGNE TABLEAU DES FILS") Or (UCase(BlocRef.Name) = "LIGNE TABLEAU DES FILS_RD") Or (UCase(BlocRef.Name) = "NOMBRE_FILS") Then
   
Else
 Exit Function
End If
Attributes = BlocRef.GetAttributes

LectureFils = IsTableauFils(Attributes)
 If LectureFils = False Then LectureFils = IsEnteteTableauFils(Attributes)
 If LectureFils = False Then LectureFils = IsNOMBRE_FILS(Attributes)
 aa = BlocRef.Name
If (bool_Plan_L_Fils = False And MyOption = "PL") Or (bool_Outil_L_Fils = False And MyOption = "OU") Then Exit Function
'AutoApp.Documents(0).PurgeAll
'     FormBarGrah.ProgressBar1Caption = " Scanne Tableau des Fils:"
'     FormBarGrah.ProgressBar1.Value = 0
'     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
'
'    For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
'     IncremanteBarGrah FormBarGrah
'    DoEvents
'        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
'
'        If Entity.ObjectName = "AcDbBlockReference" Then
'            Set BlocRef = Entity
'            A = BlocRef.Name
'            If BlocRef.HasAttributes Then
                
                    
    If LectureFils = True Then
        Set NewBlockV = BlocRef
        CollectionFils.Add NewBlockV
        Set NewBlockV = Nothing
    End If
'            End If
'        End If
'Reprise:
'    Next i

    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
    Resume Next
End Function
Function LectureCritères(MyOption As String, BlocRef As AcadBlockReference) As Boolean
If UCase(BlocRef.Name) = UCase("RefCriteres") Then
Else
Exit Function
End If
Attributes = BlocRef.GetAttributes

LectureCritères = IsCriteres(Attributes)
On Error GoTo Fin
If (bool_Plan_L_Criteres = False And MyOption = "PL") Or (bool_Outil_L_Criteres = False And MyOption = "OU") Then Exit Function
'AutoApp.Documents(0).PurgeAll
'    FormBarGrah.ProgressBar1Caption = " Scanne des Critères :"
'     FormBarGrah.ProgressBar1.Value = 0
'     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
'    Set Etiquettes = Nothing
'    Set Etiquettes = New Collection
'    For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
'        IncremanteBarGrah FormBarGrah
'        DoEvents
'        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
'
'        If Entity.ObjectName = "AcDbBlockReference" Then
'            Set BlocRef = Entity
'            A = BlocRef.Name
'            If BlocRef.HasAttributes Then
            
                
               
                    If LectureCritères = True Then
'                        Set NewBlockV = BlocRef
                        RefCriteres.Add BlocRef
'                        Set NewBlockV = Nothing
                    End If
            
                
'         End If
'      End If
'Reprise:
'    Next i
On Error GoTo 0
 
    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
    Resume Next
End Function

Function LectureNoeuds(MyOption As String, BlocRef As AcadBlockReference, Collec As Collection) As Boolean

On Error GoTo Fin
 Attributes = BlocRef.GetAttributes
 aa = BlocRef.Name
 If (UCase(BlocRef.Name) = "NOEUD_SECONDAIRE1") Or (UCase(BlocRef.Name) = "NOEUD_SECONDAIRE") Or (UCase(BlocRef.Name) = "NOEUD_LONG") Or (UCase(BlocRef.Name) = "NOEUD_PRINCIPAL1") Or (UCase(BlocRef.Name) = "NOEUD_PRINCIPAL") Or (UCase(BlocRef.Name) = "NOEUD") Then
 Else
 Exit Function
 End If

LectureNoeuds = IsNoeuds(Attributes)
If (bool_Plan_L_Noeuds = False And MyOption = "PL") Or (bool_Outil_L_Noeuds = False And MyOption = "OU") Then Exit Function
'AutoApp.Documents(0).PurgeAll
'    FormBarGrah.ProgressBar1Caption = " Scanne des Noeuds :"
'     FormBarGrah.ProgressBar1.Value = 0
'     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.Documents(0).ModelSpace.Count
'
'    For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
'        IncremanteBarGrah FormBarGrah
'        DoEvents
'        Set Entity = AutoApp.Documents(0).ModelSpace.Item(i)
'
'        If Entity.ObjectName = "AcDbBlockReference" Then
'            Set BlocRef = Entity
'            A = BlocRef.Name
'            If BlocRef.HasAttributes Then
            
                Attributes = BlocRef.GetAttributes
               
                    If LectureNoeuds = True Then
                     
                    
                        On Error Resume Next
                        a = ""
                        a = CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString))
                    If Err Then
                        Err.Clear
                        NUMNOEUDS = NUMNOEUDS + 1
                        CollectionNoeuds.Add NUMNOEUDS, Trim("N" & Attributes(Collec("NOEUD")).TextString)
                        ReDim Preserve TableauDeNoeuds(NUMNOEUDS)
                        On Error GoTo Fin
                    End If
                    
'
'     PathBlocs & "\NOEUD_PRINCIPAL.dwg"
'Else
'     Lib1 = PathBlocs & "\NOEUD_SECONDAIRE.dwg"
'    End If

                Select Case BlocRef.Name
                    Case "NOEUD_LONG"
                    Set TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).BlockDesing = BlocRef
                     TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).PosOkDesin = True
                Case "NOEUD"
                       Set TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).BlockComp = BlocRef
                        TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).PosOkComp = True
                 Case "NOEUD_0"
                       Set TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).BlockComp = BlocRef
                        TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).PosOkComp = True
                 Case "NOEUD_PRINCIPAL"
                       Set TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).BlockFleche = BlocRef
                        TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).PosOkFleche = True
                Case "NOEUD_SECONDAIRE"
                       Set TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).BlockFleche = BlocRef
                        TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).PosOkFleche = True
                 Case "NOEUD_PRINCIPAL1"
                       Set TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).BlockFleche = BlocRef
                        TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).PosOkFleche = True
                Case "NOEUD_SECONDAIRE1"
                       Set TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).BlockFleche = BlocRef
                        TableauDeNoeuds(Val(CollectionNoeuds(Trim("N" & Attributes(Collec("NOEUD")).TextString)))).PosOkFleche = True
            End Select
                
         End If
'      End If
'Reprise:
'    Next i

   On Error GoTo 0
    Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
    Resume Next
End Function
Function LectureCartouches(MyOption As String, BlocRef As AcadBlockReference) As Boolean
On Error GoTo Fin
 Attributes = BlocRef.GetAttributes
 aa = BlocRef.Name
 aa = UBound(Attributes)
 If UBound(Attributes) < 15 And UBound(Attributes) > 21 Then Exit Function
               LectureCartouches = IsCartoucheEncelade(Attributes)
                   
                If LectureCartouches = False Then LectureCartouches = IsCartoucheClient(Attributes)
                  
If (bool_Plan_L_cartouches = False And MyOption = "PL") Or (bool_Outil_L_cartouches = False And MyOption = "OU") Then Exit Function
           
                Attributes = BlocRef.GetAttributes
                
                If LectureCartouches = True Then
                    Set NewBlockV = BlocRef
                    CollectionChartouche.Add NewBlockV
                    Set NewBlockV = Nothing
              
                End If
            

   On Error GoTo 0
 Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
   Resume Next
End Function
Function LectureVignettes(MyOption As String, BlocRef As AcadBlockReference, Collec As Collection) As Boolean
On Error GoTo Fin
Attributes = BlocRef.GetAttributes
If UBound(Attributes) > 5 Then
    Exit Function
Else
    If UBound(Attributes) <> 5 Then
        If UBound(Attributes) > 0 Then Exit Function
     End If
End If
LectureVignettes = IsVignette(Attributes)
If ((bool_Plan_L_Vignettes = False Or bool_Plan_L_Connecteurs = False) And MyOption = "PL") Or ((bool_Outil_L_Vignettes = False Or bool_Outil_L_Connecteurs = False) And MyOption = "OU") Then Exit Function
            If BlocRef.HasAttributes Then
                a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes

                

                If LectureVignettes = True Then
                   
                  a = ""
                     a = CollectionCon(Attributes(Collec("CODE_APP")).TextString)
                    If Err Then
                        Err.Clear
                        NbConnecteur = NbConnecteur + 1
                        CollectionCon.Add NbConnecteur, Attributes(Collec("CODE_APP")).TextString
                       
                        ReDim Preserve TableauDeConnecteurs(NbConnecteur)
                    End If
  
                    Debug.Print Attributes(Collec("N°")).TextString
                        Set TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).NewVignette = BlocRef
                         Set TableauDeConnecteurs(Val(CollectionCon(Attributes(Collec("CODE_APP")).TextString))).AttribuesVignette = Collec
'                        DelAttribues Attributes
                        NbLignesVignette = NbLignesVignette + 1
                        InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
                Else
                    If IsVignetteEPISSURE(Attributes) = True Then
                        LectureVignettes = True
                    
                        For i2 = 1 To UBound(TableauDeConnecteurs)
                        
                            If TableauDeConnecteurs(i2).EPISSURE = True Then
                               a = TableauDeConnecteurs(i2).NewBlock.GetAttributes
                                b = a(TableauDeConnecteurs(i2).Attribues("CODE_APP")).TextString
                                If a(TableauDeConnecteurs(i2).Attribues("CODE_APP")).TextString = Attributes(Collec("EPISSURE")).TextString Then
                                    Set TableauDeConnecteurs(i2).NewVignette = BlocRef
                                     Set TableauDeConnecteurs(i2).AttribuesVignette = Collec
                                    NbLignesVignette = NbLignesVignette + 1
                                    InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
                                End If
                            End If
                        Next i2
                  
                         
                    End If
                End If
            End If
'      End If
      If NbLignesVignette = 15 Then
            InsertPointLigneTableau_Vignette(0) = -1146.1429
            InsertPointLigneTableau_Vignette(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_Vignette(1), 40)
            NbLignesVignette = 0
        End If
'Reprise:
'    Next i
'    AutoApp.Documents(0).PurgeAll
    On Error GoTo 0
     Exit Function
Fin:
    FunError 100, "", FormBarGrah.ProgressBar1Caption & vbCrLf & Err.Description, ""
    Err.Clear
    Resume Next
End Function
