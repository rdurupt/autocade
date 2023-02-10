Attribute VB_Name = "ModifierPlan"
Public Function ModifierUnPlan() As Boolean
    Dim Sql As String
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
     NbLignesVignette = 0
     ModifierUnPlan = False
     Con.OpenConnetion db
   NbConnecteur = 0
       InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 790.0288: InsertPointLigneTableau_Vignette(2) = 0

    Sql = "SELECT T_indiceProjet.AutoCadSave,T_indiceProjet.AutoCadSaveAs "
    Sql = Sql & "FROM T_Projet INNER JOIN T_indiceProjet ON T_Projet.id = T_indiceProjet.IdProjet "
    Sql = Sql & "WHERE T_Projet.Projet='" & varProjet & "' "
    Sql = Sql & "AND T_indiceProjet.Li='" & MyReplace(CartoucheEncelade.CombLi.List(CartoucheEncelade.CombLi.ListIndex, 1)) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
        Exit Function
     Else
        If Trim("" & Rs!AutoCadSave) <> "" Then
            MyFichier = "" & Rs!AutoCadSave
        End If
        If Trim("" & Rs!AutoCadSaveAs) <> "" Then
            MyFichier = "" & Rs!AutoCadSaveAs
        End If
        If Trim("" & MyFichier) = "" Then Exit Function
    End If
    
    Set TableauPath = funPath
    PathDessin = TableauPath("PathArciveAutocad") & Trim("" & MyFichier) & ".dwg"
     If Fso.FileExists(PathDessin) = False Then Exit Function
         NbLignes = 0
    Set AutoApp = ThisDrawing.Application
 OpenFichier PathDessin
 
      Menu.ProgressBar1Caption = "Lecture des Connecteurs:"
    Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1.Max = AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
  
    Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
'                If UBound(Attributes) = 13 Then
                    If IsConnecteurs(Attributes) = True Then
  
  Set Collec = ColectionAttribueConecteur(Attributes)
   
  Debug.Print Attributes(Collec("N°")).TextString
  If NbConnecteur < Val(Attributes(Collec("N°")).TextString) Then NbConnecteur = Val(Attributes(Collec("N°")).TextString)
  CollectionCon.Add Val(Attributes(Collec("N°")).TextString), Attributes(Collec("CODE_APP")).TextString

  ReDim Preserve TableauDeConnecteurs(NbConnecteur)
  
   Set TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).NewBlock = BlocRef
    Set TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).Attribues = Collec
    TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).ConnecteurExiste = True
'    TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).Trouve = True
     TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).Kill = True
  DelAttribues Attributes
                    Else
                        If IsEpissures(Attributes) = True Then
                        
                          Set Collec = ColectionAttribueConecteur(Attributes)
'                            If UCase(Attributes(Collec("CODE_APP")).TextString) = "E3FB1" Then
'  MsgBox ""
'  End If

                        If NbConnecteur < Val(Attributes(Collec("N°")).TextString) Then NbConnecteur = Val(Attributes(Collec("N°")).TextString)
    Debug.Print Attributes(Collec("N°")).TextString
CollectionCon.Add Val(Attributes(Collec("N°")).TextString), Attributes(Collec("CODE_APP")).TextString
                            ReDim Preserve TableauDeConnecteurs(NbConnecteur)
                            Set TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).NewBlock = BlocRef
                            Set TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).Attribues = Collec
'                            TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).Trouve = True
                                TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).ConnecteurExiste = True
                                 TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).Kill = True
                            TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).EPISSURE = True
                          'DelAttribues Attributes
'                            EnrichirBaseConnecteurEpissure Attributes, BlocRef.Name
                        End If
                    End If
'                End If
            End If
        End If
    Next i
    If NbConnecteur = 0 Then
    AutoApp.ActiveDocument.Close , False
    Exit Function
   End If
    
   Menu.ProgressBar1Caption = "Lecture des Vignettes :"
    Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1.Max = AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
    Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
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
                     If NbConnecteur < Val(Attributes(Collec("N°")).TextString) Then
                        NbConnecteur = Val(Attributes(Collec("N°")).TextString)
                    End If

  ReDim Preserve TableauDeConnecteurs(NbConnecteur)
                    Debug.Print Attributes(Collec("N°")).TextString
'                    If NbConnecteur < Val(Attributes(Collec("N°")).TextString) Then NbConnecteur = Val(Attributes(Collec("N°")).TextString)
                        Set TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).NewVignette = BlocRef
                         Set TableauDeConnecteurs(Val(Attributes(Collec("N°")).TextString)).AttribuesVignette = Collec
                        DelAttribues Attributes
                        NbLignesVignette = NbLignesVignette + 1
                        InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
'                    End If
                Else
                    If IsVignetteEPISSURE(Attributes) = True Then
                    
                     Set Collec = ColectionAttribueConecteur(Attributes)
'                    If NbConnecteur < Val(Attributes(Collec("N°")).TextString) Then NbConnecteur = Val(Attributes(Collec("N°")).TextString)
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
                         DelAttribues Attributes
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
    
    
     Menu.ProgressBar1Caption = "Lecture des Etiquettes :"
    Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1.Max = AutoApp.ActiveDocument.ModelSpace.Count
  For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
    Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
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
                   
'                 NewBlockV.Delete
                   
'                    Set Collec = ColectionAttribueConecteur(Attributes)
'                    If NbConnecteur < Val(Attributes(Collec("N°")).TextString) Then NbConnecteur = Val(Attributes(Collec("N°")).TextString)
                      
                End If
              
            End If
      End If
      If i > AutoApp.ActiveDocument.ModelSpace.Count - 1 Then Exit For
    Next i
   For i = 1 To Etiquettes.Count

    Set BlocRef = Etiquettes(i)
     BlocRef.Delete
   Next i
   
    Menu.ProgressBar1Caption = "Lecture Tableau des Fils:"
    Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1.Max = AutoApp.ActiveDocument.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
    Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
       
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            a = BlocRef.Name
            If BlocRef.HasAttributes Then
                Attributes = BlocRef.GetAttributes
                If (UBound(Attributes) = 13) Or (UBound(Attributes) = 12) Then
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
   
   Menu.ProgressBar1Caption = "Lecture des cartouches :"
    Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1.Max = AutoApp.ActiveDocument.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
        Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
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
   LoadConnecteur
   MyAccess
   ChargeCartoucheClient MyCARTOUCHE_Client
    ChargeCartoucheEncelade CartoucheEncelade
   
      
'          SaveAs PathArchive(TableauPath.Item("PathArciveAutocad"), CartoucheEncelade.txt1.List(CartoucheEncelade.txt1.ListIndex, 0), CartoucheEncelade.CleAc, CartoucheEncelade.CombLi.Text) & CartoucheEncelade.CombLi.Text
If CartoucheEncelade.txt1.ListIndex = -1 Then
     PathPl = PathArchive(TableauPath.Item("PathArciveAutocad"), "", CartoucheEncelade.CleAc, CartoucheEncelade.txt15 & "_" & CartoucheEncelade.txt16)
Else
    PathPl = PathArchive(TableauPath.Item("PathArciveAutocad"), CartoucheEncelade.txt1.List(CartoucheEncelade.txt1.ListIndex, 0), CartoucheEncelade.CleAc, CartoucheEncelade.txt15 & "_" & CartoucheEncelade.txt16)
End If
    SaveAs PathPl & CartoucheEncelade.txt15 & "_" & CartoucheEncelade.txt16
        If boolFormClient = True Then Unload MyCARTOUCHE_Client
    Unload CartoucheEncelade
    Set MyCARTOUCHE_Client = Nothing
        DoEvents
         AfficheErreur PathPl, EnteteCartouche
         Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1Caption.Caption = "Fin du traitement"
MsgBox "Fin du traitement"
Con.CloseConnection
Unload Menu
      ModifierUnPlan = True
  
End Function
Function LoadConnecteur() As Boolean
    LoadConnecteur = False
    Dim RsConnecteur As Recordset
    Dim Sql As String
    Dim MyRep As String
    Dim Trouve As Boolean
    Dim NbLignesVignette As Long
    Dim NbConnecteur As Long
    Dim Fso As New FileSystemObject
     Dim RefNull  As AcadBlockReference
    Dim NumErr As Long
    Sql = "SELECT Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, "
    Sql = Sql & "Connecteurs.CODE_APP, Connecteurs.N°, Connecteurs.POS, Connecteurs.PRECO1, Connecteurs.PRECO2 "
    Sql = Sql & "FROM T_Projet INNER JOIN (T_indiceProjet INNER JOIN Connecteurs ON T_indiceProjet.Id = Connecteurs.Id_IndiceProjet) ON T_Projet.id = T_indiceProjet.IdProjet "
    Sql = Sql & "WHERE T_Projet.Projet='" & varProjet & "' "
    Sql = Sql & "AND T_indiceProjet.Li='" & MyReplace(CartoucheEncelade.CombLi.List(CartoucheEncelade.CombLi.ListIndex, 1)) & " ';"
    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)



    InsertPointConnecteur(0) = 100: InsertPointConnecteur(1) = 100: InsertPointConnecteur(2) = 0
    i = 1
    NbConnecteur = UBound(TableauDeConnecteurs)
    While "" & RsConnecteur.EOF = False
     On Error Resume Next
    CollectionCon.Add CLng("" & RsConnecteur.Fields(4)), "" & RsConnecteur.Fields(3)
    Err.Clear
    If CLng("" & RsConnecteur.Fields(4)) > NbConnecteur Then
    NbConnecteur = CLng("" & RsConnecteur.Fields(4))
   
    On Error GoTo 0
    End If
      
       RsConnecteur.MoveNext
    Wend
    If NbConnecteur > UBound(TableauDeConnecteurs) Then
    ReDim Preserve TableauDeConnecteurs(NbConnecteur)
    End If
    Menu.ProgressBar1.Value = 0
    If NbConnecteur = 0 Then
    Menu.ProgressBar1.Max = 1
    Else
        Menu.ProgressBar1.Max = NbConnecteur
    End If
    Menu.ProgressBar1Caption.Caption = "Chargement des connecteurs"
    If NbConnecteur <> 0 Then
        RsConnecteur.MoveFirst
    End If
     On Error GoTo GesERR
    While RsConnecteur.EOF = False
        Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
       
        DoEvents
        
        DoEvents

        If UCase("" & RsConnecteur.Fields(0)) <> "NEANT" Then
             If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = True Then
                If Trim(UCase(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.Name)) <> UCase(Trim("" & RsConnecteur.Fields(0))) Then
                    If Fso.FileExists(TableauPath.Item("PathConnecteurs" & LeCient) & "" & RsConnecteur.Fields(0) & ".dwg") = True Then
                        MyRep = TableauPath.Item("PathConnecteurs" & LeCient)
                        Trouve = True
                        NumErr = 4
                       
                        
'                        TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = True
                    Else
                        NumErr = 1
                        TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = False
                        MyRep = ""
                    
GesERR:
                        Trouve = False
                        FunError NumErr, "" & RsConnecteur.Fields(4), Err.Description, "" & RsConnecteur.Fields(0)
                        a = RsConnecteur!N°
                         TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = False
                    End If
                     If Trouve = True Then
                        InsertionPoint = TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.InsertionPoint
                        InsertPointConnecteur(0) = InsertionPoint(0): InsertPointConnecteur(1) = InsertionPoint(1): InsertPointConnecteur(2) = InsertionPoint(2)
                        TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.Delete
                        Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock = FunInsBlock(MyRep & "" & RsConnecteur.Fields(0) & ".dwg", InsertPointConnecteur, "")
                          If ErrInsert = True Then GoTo EnrSuinant
                        If UCase("" & RsConnecteur.Fields(1)) = "O" Then
                            TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE = True
                            InsertionPoint = TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.InsertionPoint
                            InsertPointConnecteur(0) = InsertionPoint(0): InsertPointConnecteur(1) = InsertionPoint(1): InsertPointConnecteur(2) = InsertionPoint(2)
                            TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.Delete

                            Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(MyRep & "EPISSURES.dwg", InsertPointConnecteur, "V" & "" & RsConnecteur.Fields(4))
                        Else
                            TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE = False
                                        InsertionPoint = TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.InsertionPoint
                            InsertPointConnecteur(0) = InsertionPoint(0): InsertPointConnecteur(1) = InsertionPoint(1): InsertPointConnecteur(2) = InsertionPoint(2)
                            TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.Delete

                            Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(MyRep & "VIGNETTE CONNECTEUR.dwg", InsertPointConnecteur, "V" & "" & RsConnecteur.Fields(4))
                             
                        End If
                    End If
                End If
                If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = True Then
                     TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Kill = False
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes)
                On Error Resume Next
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes)
                On Error GoTo 0
 At = TableauAtribCon(RsConnecteur, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE)
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues
                On Error Resume Next
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE
                On Error GoTo 0
                End If
               
             Else
              If Fso.FileExists(TableauPath.Item("PathConnecteurs" & LeCient) & "" & RsConnecteur.Fields(0) & ".dwg") = True Then
                MyRep = TableauPath.Item("PathConnecteurs" & LeCient)
                Trouve = True
                NumErr = 4
               
                
                TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = True
            Else
                NumErr = 1
                TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = False
                MyRep = ""
                

                Trouve = False
               GoTo GesERR
            End If
            If Trouve = True Then
                InsertPointConnecteur(0) = 100: InsertPointConnecteur(1) = 100: InsertPointConnecteur(2) = 1
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock = FunInsBlock(MyRep & "" & RsConnecteur.Fields(0) & ".dwg", InsertPointConnecteur, "")
                If ErrInsert = True Then GoTo EnrSuinant
                If UCase("" & RsConnecteur.Fields(1)) = "O" Then
                    TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE = True
                    Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(MyRep & "EPISSURES.dwg", InsertPointLigneTableau_Vignette, "V" & "" & RsConnecteur.Fields(4))
                    InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
                    NbLignesVignette = NbLignesVignette + 1
                Else
                    TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE = False
                    Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(MyRep & "VIGNETTE CONNECTEUR.dwg", InsertPointLigneTableau_Vignette, "V" & "" & RsConnecteur.Fields(4))
                    InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
                    NbLignesVignette = NbLignesVignette + 1
                End If
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes)
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes)

                At = TableauAtribCon(RsConnecteur, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE)
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).EPISSURE
                
            End If
             End If
        
      
               
               
               
              
        End If
         If NbLignesVignette = 15 Then
            InsertPointLigneTableau_Vignette(0) = -1146.1429
            InsertPointLigneTableau_Vignette(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_Vignette(1), 40)
            NbLignesVignette = 0
       
        End If
EnrSuinant:
       RsConnecteur.MoveNext
        i = i + 1
    Wend
    For i = LBound(TableauDeConnecteurs) To UBound(TableauDeConnecteurs)
        If TableauDeConnecteurs(i).Kill = True Then
             TableauDeConnecteurs(i).NewBlock.Delete
            TableauDeConnecteurs(i).NewVignette.Delete
             Set BlocRef = TableauDeConnecteurs(i).NewVignette
            Set TableauDeConnecteurs(i).NewBlock = Nothing
            Set TableauDeConnecteurs(i).NewVignette = Nothing
            TableauDeConnecteurs(i).ConnecteurExiste = False
            TableauDeConnecteurs(i).EPISSURE = False
            Set BlocRef = Nothing
        End If
    Next i
    LoadConnecteur = True
    Set Fso = Nothing
End Function

Function EnteteCartouche()
    Dim Txt
    Dim txt2
    Dim Mysapce
    Mysapce = Space(65)
    Txt = "******************************************************************" & vbCrLf
    Txt = Txt & "* Listes des erreurs survenues lors de l'exécution de la macro : *" & vbCrLf
    Txt = Txt & "* Créer un Plan                                                  *" & vbCrLf
    txt2 = "* Projet : " & varProjet & " Indice : " & varIndice
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
    Txt = Txt & "******************************************************************" & vbCrLf
    Txt = Txt & vbCrLf
    EnteteCartouche = Txt
End Function
