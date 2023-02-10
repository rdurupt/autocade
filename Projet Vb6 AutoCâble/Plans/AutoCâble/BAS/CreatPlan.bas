Attribute VB_Name = "CreatPlan"

Public Sub subDessinerPlan()
    Dim Rs As Recordset
    Dim PathPl As String

    If MsgBox("Voulez vous exécuter la Macro", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    If ModifierUnPlan = True Then Exit Sub
'     Con.OpenConnetion db
    Set TableauPath = funPath
    Dim Tableau() As String
    Dim NbFil As Long
    Dim NbLignes As Long
    Dim Fso As New FileSystemObject
    NbLignes = 0
    Set AutoApp = ThisDrawing.Application
    If Fso.FileExists(TableauPath.Item("PathPlantVierge") & "PLAN VIERGE.dwg") = False Then
        MsgBox "Err"
        Exit Sub
    End If
    OpenFichier TableauPath.Item("PathPlantVierge") & "PLAN VIERGE.dwg"
    InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 790.0288: InsertPointLigneTableau_Vignette(2) = 0

    If LoadConnecteur = False Then GoTo Fin

    ChargeCartoucheClient MyCARTOUCHE_Client
    ChargeCartoucheEncelade CartoucheEncelade
    
    InsertPointConnecteur(0) = 100: InsertPointConnecteur(1) = 100: InsertPointConnecteur(2) = 0
   
    MyAccess
aa = CartoucheEncelade.TextBox7 & CartoucheEncelade.CleAc & "_" & CartoucheEncelade.Annee & CartoucheEncelade.txt20 & "_" & CartoucheEncelade.txt16
    
    PathPl = PathArchive(TableauPath.Item("PathArciveAutocad"), CartoucheEncelade.txt1.List(CartoucheEncelade.txt1.ListIndex, 0), CartoucheEncelade.CleAc, CartoucheEncelade.txt15 & "_" & CartoucheEncelade.txt16)
    SaveAs PathPl & CartoucheEncelade.txt15 & "_" & CartoucheEncelade.txt16
     If boolFormClient = True Then Unload MyCARTOUCHE_Client
    Unload CartoucheEncelade
    Set MyCARTOUCHE_Client = Nothing
        DoEvents

Fin:
'    CloseDocument
'    Set AutoApp = Nothing
    ReDim TableauDeConnecteurs(0)
    AfficheErreur PathPl, EnteteCartouche
    
    Con.CloseConnection
    Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1Caption.Caption = "Fin du traitement"
MsgBox "Fin du traitement"
Unload Menu
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








Sub CopyFile()
    For i = 1 To 10
        DoEvents
        FileCopy PathConstructionModelNUMEROFIL & "Copie de NUMEROFIL40.dwg", PathConstructionModelNUMEROFIL & "NUMEROFIL" & CStr(i * 4) & ".dwg"
    Next i
    For i = 11 To 20
        DoEvents
        FileCopy PathConstructionModelNUMEROFIL & "c_NUMEROFIL80.dwg", PathConstructionModelNUMEROFIL & "NUMEROFIL" & CStr(i * 4) & ".dwg"
    Next i
End Sub

Function funPath()
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


Public Function ValideChampsTexte(Formulaire, NbChamps As Long) As Boolean
    Dim MyTag
    ValideChampsTexte = False
    For i = 0 To NbChamps
        DoEvents
        MyTag = Split(Formulaire.Controls("txt" & CStr(i)).tag, ";")
    
        If Trim("" & Formulaire.Controls("txt" & CStr(i))) = "" Then
            If UCase(MyTag(2)) = "QRY" Then
                MsgBox "Valeur de : " & MyTag(1) & " obligatoire", vbExclamation
                Formulaire.Controls("txt" & CStr(i)).SetFocus
                Exit Function
            End If
        Else
    
            Select Case UCase(MyTag(3))
                    Case "DATE"
                        If Not IsDate(Formulaire.Controls("txt" & CStr(i))) Then
                            MsgBox "Vous devez saisir une date pour : " & MyTag(1), vbExclamation
                            Formulaire.Controls("txt" & CStr(i)) = ""
                            Formulaire.Controls("txt" & CStr(i)).SetFocus
                            
                            Exit Function
                        Else
                            Formulaire.Controls("txt" & CStr(i)) = Format(Formulaire.Controls("txt" & CStr(i)), "dd/mm/yyyy")
                        End If
                    Case "ENT"
                        If Not IsNumeric(Formulaire.Controls("txt" & CStr(i))) Then
                            MsgBox "Vous devez saisir un nombre entier pour : " & MyTag(1), vbExclamation
                            Formulaire.Controls("txt" & CStr(i)) = ""
                            Formulaire.Controls("txt" & CStr(i)).SetFocus
                            Exit Function
                        Else
                            If (InStr(1, (Formulaire.Controls("txt" & CStr(i))), ",") <> 0) Or (InStr(1, (Formulaire.Controls("txt" & CStr(i))), ".") <> 0) Then
                                MsgBox "Vous devez saisir un nombre entier pour : " & MyTag(1), vbExclamation
                                Formulaire.Controls("txt" & CStr(i)) = ""
                                Formulaire.Controls("txt" & CStr(i)).SetFocus
                                Exit Function
                            End If
                        End If
                    Case "DBL"
                        If Not IsNumeric(Formulaire.Controls("txt" & CStr(i))) Then
                            Formulaire.Controls("txt" & CStr(i)) = Replace(Formulaire.Controls("txt" & CStr(i)), ".", ",")
                        End If
                        If Not IsNumeric(Formulaire.Controls("txt" & CStr(i))) Then
                            MsgBox "Vous devez saisir un nombre à virgule pour : " & MyTag(1), vbExclamation
                            Formulaire.Controls("txt" & CStr(i)) = ""
                            Formulaire.Controls("txt" & CStr(i)).SetFocus
                            Exit Function
                        End If
            End Select
    
        End If
    Next i
    ValideChampsTexte = True
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

Function LoadConnecteur() As Boolean
    LoadConnecteur = False
    Dim RsConnecteur As Recordset
    Dim Sql As String
    Dim MyRep As String
    Dim Trouve As Boolean
   
    Dim NbConnecteur As Long
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Set CollectionCon = Nothing
    Set CollectionCon = New Collection
    
    Sql = "SELECT Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, "
    Sql = Sql & "Connecteurs.CODE_APP, Connecteurs.N°, Connecteurs.POS, Connecteurs.PRECO1, Connecteurs.PRECO2 "
    Sql = Sql & "FROM T_Projet INNER JOIN (T_indiceProjet INNER JOIN Connecteurs ON T_indiceProjet.Id = Connecteurs.Id_IndiceProjet) ON T_Projet.id = T_indiceProjet.IdProjet "
    Sql = Sql & "WHERE T_Projet.Projet='" & varProjet & "' "
    Sql = Sql & "AND T_indiceProjet.Li='" & MyReplace(CartoucheEncelade.CombLi.List(CartoucheEncelade.CombLi.ListIndex, 1)) & "';"
    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)



    InsertPointConnecteur(0) = 100: InsertPointConnecteur(1) = 100: InsertPointConnecteur(2) = 0
    i = 1
    While "" & RsConnecteur.EOF = False
    If Trim(UCase("" & RsConnecteur.Fields(0))) <> "NEANT" Then
     CollectionCon.Add CLng("" & RsConnecteur.Fields(4)), "" & RsConnecteur.Fields(3)
     End If
    If CLng("" & RsConnecteur.Fields(4)) > NbConnecteur Then
        NbConnecteur = CLng("" & RsConnecteur.Fields(4))
    End If
       RsConnecteur.MoveNext
    Wend
    ReDim TableauDeConnecteurs(NbConnecteur)
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
            If Fso.FileExists(TableauPath.Item("PathConnecteurs" & LeCient) & "" & RsConnecteur.Fields(0) & ".dwg") = True Then
                MyRep = TableauPath.Item("PathConnecteurs" & LeCient)
                Trouve = True
                NumErr = 4
              
                
                TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).ConnecteurExiste = True
            Else
                NumErr = 1
                TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).ConnecteurExiste = False
                MyRep = ""
                
GesERR:
                Trouve = False
                FunError NumErr, "" & RsConnecteur.Fields(4), Err.Description, "" & RsConnecteur.Fields(0)
            End If
            If Trouve = True Then
           
                Set TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewBlock = FunInsBlock(MyRep & "" & RsConnecteur.Fields(0) & ".dwg", InsertPointConnecteur, "")
                  If ErrInsert = True Then GoTo EnrSuinant
                If UCase("" & RsConnecteur.Fields(1)) = "O" Then
                    TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).EPISSURE = True
                    Set TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewVignette = FunInsBlock(MyRep & "EPISSURES.dwg", InsertPointLigneTableau_Vignette, "V" & "" & RsConnecteur.Fields(4))
                    InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
                    NbLignesVignette = NbLignesVignette + 1
                Else
                    TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).EPISSURE = False
                    Set TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewVignette = FunInsBlock(MyRep & "VIGNETTE CONNECTEUR.dwg", InsertPointLigneTableau_Vignette, "V" & "" & RsConnecteur.Fields(4))
                    InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), 80)
                    NbLignesVignette = NbLignesVignette + 1
                End If
                Set TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).Attribues = ColectionAttribueConecteur(TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewBlock.GetAttributes)
                Set TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).AttribuesVignette = ColectionAttribueConecteur(TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewVignette.GetAttributes)

                At = TableauAtribCon(RsConnecteur, TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).EPISSURE)
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewBlock.Name, TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewBlock.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).Attribues
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewVignette.Name, TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).NewVignette.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).AttribuesVignette, True, TableauDeConnecteurs(CLng("" & RsConnecteur.Fields(4))).EPISSURE
                
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
Function ColectionAttribueConecteur(Attribues) As Collection
    Dim MyAttribue As New Collection
    Dim IndexAt As Long


    IndexAt = 0
    On Error Resume Next
    While IndexAt < UBound(Attribues) + 1
        Debug.Print Attribues(IndexAt).TagString
        Debug.Print Replace(Replace(UCase(Attribues(IndexAt).TagString), "PRECO", "PRECO."), "PRECO..", "PRECO.")
        MyAttribue.Add IndexAt, Replace(Replace(UCase(Attribues(IndexAt).TagString), "PRECO", "PRECO."), "PRECO..", "PRECO.")
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

