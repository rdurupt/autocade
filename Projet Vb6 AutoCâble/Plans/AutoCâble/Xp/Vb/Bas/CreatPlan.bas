Attribute VB_Name = "CreatPlan"

Public Sub subDessinerPlan(IdIndiceProjet As Long)
If boolAutoCAD = False Then Exit Sub
NotSaveRacourci = False
    Dim Rs As Recordset
    Dim Rs2 As Recordset
    Dim PathPl As String
    Dim Sql As String
    Set TableauPath = funPath
    Dim PathPlantVierge As String
    Set TableauPath = funPath
    Set CollectionTor = Nothing
    Set CollectionTor = New Collection
'
     NUMCOM = 0
    NUMNOTA = 0
    NUMNTOR = 0
     NUMNTORBLOC = 0
    PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
     If Left(PathArchiveAutocad, 2) <> "\\" And Left(PathArchiveAutocad, 1) = "\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad
  If Right(PathArchiveAutocad, 2) = "\\" Then PathArchiveAutocad = Mid(PathArchiveAutocad, 1, Len(MyPath) - 1)
    Sql = "SELECT T_indiceProjet.PL , T_indiceProjet.NbCartouche FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    
    Set Rs = Con.OpenRecordSet(Sql)
    NbCartouche = Rs!NbCartouche
     PathPlantVierge = TableauPath.Item("PathPlantVierge")
     If Left(PathPlantVierge, 2) <> "\\" And Left(PathPlantVierge, 1) = "\" Then PathPlantVierge = TableauPath.Item("PathServer") & PathPlantVierge
      If Right(PathPlantVierge, 2) = "\\" Then PathPlantVierge = Mid(PathPlantVierge, 1, Len(PathPlantVierge) - 1)

If Rs.EOF = True Then Exit Sub
    If bool_Plan_Ouvrir = False Then Exit Sub
    NbError = 0
    AutoApp.Visible = True
    If ModifierUnPlan(IdIndiceProjet, "PL") = False Then
    Set TableauPath = funPath
    Dim Tableau() As String
    Dim NbFil As Long
    Dim NbLignes As Long
    Dim fso As New FileSystemObject
    ReDim TableuDeTor(0)
    ReDim TableuEtiquettes(0)
    NbLignes = 0
    
   
    OpenNew
  
End If
LoadCalque
    LoadConnecteur IdIndiceProjet, "PL"
LoadComposants IdIndiceProjet, "PL"
LoadNotas IdIndiceProjet, "PL"
LoadNoeuds IdIndiceProjet, "PL"
    ChargeCartoucheEncelade IdIndiceProjet, "PL", NbCartouche
    ChargeCartoucheClient IdIndiceProjet, "PL", NbCartouche
    SubLoadFils IdIndiceProjet, "PL"
    
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
        PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version)
        
    ExporteXls PathPl, IdIndiceProjet
    
     If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(PathArchiveAutocad, "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
       Racourci "" & PathPl2, PathPl, "XLS"
    End If
    PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "PL", Rs.Fields("PL"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("PL_Indice"), Rs!Version)
   
    End If
    SaveAs PathPl
    If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(PathArchiveAutocad, "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "PL", Rs2.Fields("PL"), IdFils, Rs.Fields("PI_Indice"), Rs2.Fields("PL_Indice"), Rs2!Version)
       Racourci "" & PathPl2, PathPl, "dwg"
    End If
        DoEvents

Fin:
    ReDim TableauDeConnecteurs(0)
    AfficheErreur PathPl, EnteteCartouche("" & Rs.Fields("Ensemble"), "" & Rs.Fields("PL_Indice"), "" & Rs.Fields("PL"))
    
    
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1Caption.Caption = " Fin du traitement"
'     FormBarGrah.SetFocus
 AutoApp.Visible = False
 EporteSynthese
 Sql = "SELECT T_indiceProjet.CleAc FROM T_indiceProjet WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
 Set Rs = Con.OpenRecordSet(Sql)
 If Rs.EOF = False Then EporteSynthese Rs!CleAc

MenuShow = False
End Sub

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

Function EnteteCartouche(varProjet As String, varIndice As String, Plan As String)
    Dim txt
    Dim txt2
    Dim Mysapce
    Mysapce = Space(78)
          txt = "             ******************************************************************" & vbCrLf
    txt = txt & "             * Listes des erreurs survenues lors de l'exécution de la macro : *" & vbCrLf
    txt = txt & "             * Créer un Plan                                                  *" & vbCrLf
         txt2 = "             * Projet : " & Replace(varProjet, vbCrLf, " ")
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * Plan : " & Plan & " Indice : " & varIndice
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             *"
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * Nombre d'erreur(s) : " & NbError
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
    txt = txt & "             ******************************************************************" & vbCrLf
    txt = txt & vbCrLf
    Debug.Print txt
    EnteteCartouche = txt
End Function

