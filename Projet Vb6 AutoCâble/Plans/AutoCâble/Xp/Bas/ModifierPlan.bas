Attribute VB_Name = "ModifierPlan"
Public Function ModifierUnOtil(IdIndiceProjet As Long) As Boolean
  Dim sql As String
     Dim PathPl As String
    Dim Rs As Recordset
    Dim PathDessin As String
    Dim Fso As New FileSystemObject
    Dim NbConnecteur As Long
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
     NbLignesVignette = 0
     ModifierUnOtil = False
     
   NbConnecteur = 0
       InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 790.0288: InsertPointLigneTableau_Vignette(2) = 0

sql = "SELECT T_indiceProjet.PlAutoCadSave, T_indiceProjet.PlAutoCadSaveas "
sql = sql & "FROM T_Projet INNER JOIN (T_Pieces INNER  "
sql = sql & "JOIN T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces)  "
sql = sql & "ON T_Projet.id = T_Pieces.IdProjet "
sql = sql & "WHERE T_Projet.id=" & IdProjet & "  "
sql = sql & "AND T_Pieces.Id=" & IdPieces & ";"

sql = "SELECT T_indiceProjet.ouAutoCadSave,  "
sql = sql & "T_indiceProjet.ouAutoCadSaveas "
sql = sql & "FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"


    Set Rs = Con.OpenRecordSet(sql)
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
    
  
    PathDessin = PathArchiveAutocad & Trim("" & MyFichier) & ".dwg"
     If Fso.FileExists(PathDessin) = False Then Exit Function
         NbLignes = 0
'    Set AutoApp = ThisDrawing.Application
 OpenFichier PathDessin
 
       FormBarGrah.ProgressBar1Caption = "Lecture des Connecteurs:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
  
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
    If NbConnecteur = 0 Then
    AutoApp.ActiveDocument.Close , False
    Exit Function
   End If
    
    FormBarGrah.ProgressBar1Caption = "Lecture des Vignettes :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
    
    
      FormBarGrah.ProgressBar1Caption = "Lecture des Etiquettes :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
      If i > AutoApp.ActiveDocument.ModelSpace.Count - 1 Then Exit For
    Next i
   For i = 1 To Etiquettes.Count

    Set BlocRef = Etiquettes(i)
     BlocRef.Delete
   Next i
   
   FormBarGrah.ProgressBar1Caption = "Lecture des Composants :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
     NUMCOM = 0
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
        FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
        If i > AutoApp.ActiveDocument.ModelSpace.Count - 1 Then Exit For
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
   

    FormBarGrah.ProgressBar1Caption = "Lecture des Notas :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
     NUMNOTA = 0
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
        FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
        If i > AutoApp.ActiveDocument.ModelSpace.Count - 1 Then Exit For
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
   
   
   
   
     FormBarGrah.ProgressBar1Caption = "Lecture Tableau des Fils:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
       
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
   
    FormBarGrah.ProgressBar1Caption = "Lecture des cartouches :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
         FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
        
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
  For i = 1 To UBound(TableauDeConnecteurs)
    If TableauDeConnecteurs(i).ConnecteurExiste = True Then
    On Error Resume Next
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
        
        On Error GoTo 0
    End If
    DoEvents
  Next i
  AutoApp.ActiveDocument.PurgeAll
   ModifierUnOtil = True
   Exit Function
  
End Function
Public Function ModifierUnPlan(IdIndiceProjet As Long, Mytype As String) As Boolean
    Dim sql As String
     Dim PathPl As String
    Dim Rs As Recordset
    Dim PathDessin As String
    Dim Fso As New FileSystemObject
    Dim NbConnecteur As Long
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
     
     NbLignesVignette = 0
     ModifierUnPlan = False
     
   NbConnecteur = 0
       InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 790.0288: InsertPointLigneTableau_Vignette(2) = 0

sql = "SELECT T_indiceProjet." & Mytype & "AutoCadSave, T_indiceProjet." & Mytype & "AutoCadSaveas "
sql = sql & "FROM T_Projet INNER JOIN (T_Pieces INNER  "
sql = sql & "JOIN T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces)  "
sql = sql & "ON T_Projet.id = T_Pieces.IdProjet "
sql = sql & "WHERE T_Projet.id=" & IdProjet & "  "
sql = sql & "AND T_Pieces.Id=" & IdPieces & ";"

sql = "SELECT T_indiceProjet." & Mytype & "AutoCadSave,  "
sql = sql & "T_indiceProjet." & Mytype & "AutoCadSaveas "
sql = sql & "FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"


    Set Rs = Con.OpenRecordSet(sql)
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
    
    PathDessin = TableauPath("PathArchiveAutocad") & Trim("" & MyFichier) & ".dwg"
    If Left(PathDessin, 2) <> "\\" Then PathDessin = TableauPath.Item("PathServer") & PathDessin
     If Fso.FileExists(PathDessin) = False Then Exit Function
         NbLignes = 0
'    Set AutoApp = ThisDrawing.Application
AutoApp.Visible = True
 OpenFichier PathDessin
 LoadCalque
       FormBarGrah.ProgressBar1Caption = "Lecture des Connecteurs:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
  
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
    If NbConnecteur = 0 Then
    AutoApp.ActiveDocument.Close , False
    Exit Function
   End If
    
    FormBarGrah.ProgressBar1Caption = "Lecture des Vignettes :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
    
    
      FormBarGrah.ProgressBar1Caption = "Lecture des Etiquettes :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
      If i > AutoApp.ActiveDocument.ModelSpace.Count - 1 Then Exit For
    Next i
   For i = 1 To Etiquettes.Count

    Set BlocRef = Etiquettes(i)
     BlocRef.Delete
   Next i
   
   
   
   
     FormBarGrah.ProgressBar1Caption = "Lecture des Composants :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
     NUMCOM = 0
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
        FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
        If i > AutoApp.ActiveDocument.ModelSpace.Count - 1 Then Exit For
    Next i
   For i = 1 To CollectionComp.Count
InsertionPoint = TableauDeComposants(i).NewBlock.InsertionPoint
 TableauDeComposants(i).InsertPointLigneC(0) = InsertionPoint(0)
 TableauDeComposants(i).InsertPointLigneC(1) = InsertionPoint(1)
 TableauDeComposants(i).InsertPointLigneC(2) = InsertionPoint(2)
 Set BlocRef = TableauDeComposants(i).NewBlock
        TableauDeComposants(i).RotationC = BlocRef.Rotation
        
         TableauDeComposants(i).XScaleFactorC = BlocRef.XScaleFactor
         TableauDeComposants(i).YScaleFactorC = BlocRef.YScaleFactor
       TableauDeComposants(i).ZScaleFactorC = BlocRef.ZScaleFactor
   
     BlocRef.Delete
   Next i
   
   
   
   


   
   
    FormBarGrah.ProgressBar1Caption = "Lecture des Notas :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
     NUMNOTA = 0
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
        FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
                    CollectionNota.Add CInt(Attributes(b("NUMNOTA")).TextString), "N" & UCase(Attributes(b("NUMNOTA")).TextString)

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
        If i > AutoApp.ActiveDocument.ModelSpace.Count - 1 Then Exit For
    Next i
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
   
   
   
   
   FormBarGrah.ProgressBar1Caption = "Lecture des Préconisations :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
     NUMNOTA = 0
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
        FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
          
                Attributes = BlocRef.GetAttributes
                
                Set Collec = ColectionAttribueConecteur(Attributes)
                
                If IsTor(Attributes) = True Then
                    On Error Resume Next
                    
                    a = ""
                    a = CollectionTor(Attributes(Collec("TORDESIGNATION")).TextString)
                    If Err Then
                     NUMNTOR = NUMNTOR + 1
                     CollectionTor.Add NUMNTOR, Attributes(Collec("TORDESIGNATION")).TextString
                    ReDim Preserve TableuDeTor(NUMNTOR)
                      TableuDeTor(NUMNTOR).CodeApp = Attributes(Collec("TORDESIGNATION")).TextString
                          If IsTorDetail(Attributes) = False Then
                            TableuDeTor(NUMNTOR).TorExiste = True
                            Set TableuDeTor(NUMNTOR).NewBlockTorTire = BlocRef
                          End If
                            ReDim Preserve TableauDeNotas(NUMNOTA)
                           
                      On Error GoTo 0
                    End If
                    If IsTorDetail(Attributes) = True Then
                        On Error Resume Next
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
                        On Error GoTo 0
                    End If

                    Set BlocRef = Nothing
                
                
                End If
            
            End If
        End If
        If i > AutoApp.ActiveDocument.ModelSpace.Count - 1 Then Exit For
    Next i
   
   
   
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
   
   
   
   
   
   
   
   
   
   
   
   
   
     FormBarGrah.ProgressBar1Caption = "Lecture Tableau des Fils:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
       
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
   
    FormBarGrah.ProgressBar1Caption = "Lecture des cartouches :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
         FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
        DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
        
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
  For i = 1 To UBound(TableauDeConnecteurs)
    If TableauDeConnecteurs(i).ConnecteurExiste = True Then
    On Error Resume Next
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
        
        On Error GoTo 0
    End If
    DoEvents
  Next i
' For i = 0 To AutoApp.ActiveDocument.Layers
'    Debug.Print AutoApp.ActiveDocument.Lineweight(0)
' Next i
 


   ModifierUnPlan = True
   Exit Function
  
End Function

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
