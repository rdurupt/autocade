Attribute VB_Name = "CreatPlan"

Sub subDessinerPlan(IdIndiceProjet As Long)
bool_MiseEnPage = True
On Error Resume Next
GetAutocad
AutoApp.Visible = True
If bool_Plan_Ouvrir = False Then Exit Sub
If boolAutoCAD = False Then Exit Sub
NotSaveRacourci = False
    Dim Rs As Recordset
    Dim Rs2 As Recordset
    Dim PathPl As String
    Dim Sql As String
    Set TableauPath = funPath
   ' Dim ModelAC As String
'    Set TableauPath = funPath
    Set CollectionTor = Nothing
    Set CollectionTor = New Collection
'
     NUMCOM = 0
    NUMNOTA = 0
    NUMNTOR = 0
     NUMNTORBLOC = 0
    PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
    PathArchiveAutocad = DefinirChemienComplet(TableauPath.Item("PathServer"), PathArchiveAutocad)

'     If Left(PathArchiveAutocad, 2) <> "\\" And Left(PathArchiveAutocad, 1) = "\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad
'  If Right(PathArchiveAutocad, 2) = "\\" Then PathArchiveAutocad = Mid(PathArchiveAutocad, 1, Len(MyPath) - 1)
    Sql = "SELECT T_indiceProjet.PL , T_indiceProjet.NbCartouche FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    
    Set Rs = Con.OpenRecordSet(Sql)
    NbCartouche = Rs!NbCartouche
'     ModelAC = TableauPath.Item("ModelAC")
'     ModelAC = DefinirChemienComplet(TableauPath.Item("PathServer"), ModelAC)
'     If Left(ModelAC, 2) <> "\\" And Left(ModelAC, 1) = "\" Then ModelAC = TableauPath.Item("PathServer") & ModelAC
'      If Right(ModelAC, 2) = "\\" Then ModelAC = Mid(ModelAC, 1, Len(ModelAC) - 1)

If Rs.EOF = True Then Exit Sub
    If bool_Plan_Ouvrir = False Then Exit Sub
    NbError = 0
    If IsServeur = False Then
        AutoApp.Visible = True
    End If
    If ModifierUnPlan(IdIndiceProjet, "PL") = False Then
    Set TableauPath = funPath
    Dim tableau() As String
    Dim NbFil As Long
    Dim NbLignes As Long
    Dim Fso As New FileSystemObject
    ReDim TableuDeTor(0)
    ReDim TableuEtiquettes(0)
    NbLignes = 0
    
   
   AdcFileName = OpenNew
    Sql = "UPDATE T_Job SET T_Job.AutocadDoc = '" & AdcFileName & " ' "
    Sql = Sql & "WHERE T_Job.Job= " & Command & ";"
    Con.Execute Sql
    Con.Execute Sql
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
    PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "PL", Rs.Fields("PL"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("PL_Indice"), Rs!Version)
     SaveAs PathPl
        PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version)
        
    ExporteXls PathPl, IdIndiceProjet
    
     If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(TableauPath.Item("PathArchiveAutocad"), "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
       Racourci "" & PathPl2, PathPl, "XLS"
    End If
    
   
    End If
   ' Restart.Show vbmodal
   
    If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(TableauPath.Item("PathArchiveAutocad"), "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "PL", Rs2.Fields("PL"), IdFils, Rs.Fields("PI_Indice"), Rs2.Fields("PL_Indice"), Rs2!Version)
       Racourci "" & PathPl2, PathPl, "dwg"
    End If
        DoEvents

Fin:
    ReDim TableauDeConnecteurs(0)
    
    AfficheErreur PathPl, EnteteCartouche("" & Rs.Fields("Ensemble"), "" & Rs.Fields("PL_Indice"), "" & Rs.Fields("PL"))
    
    
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1Caption.Caption = " Fin du traitement"
'     FormBarGrah.SetFocus
 'AutoApp.Visible = False
 EporteSynthese "SyntG"
 Sql = "SELECT T_indiceProjet.CleAc FROM T_indiceProjet WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
 Set Rs = Con.OpenRecordSet(Sql)
 If Rs.EOF = False Then EporteSynthese "Synt", Rs!CleAc
bool_MiseEnPage = False
MenuShow = False
End Sub

Sub CopyFile()
    For I = 1 To 10
        DoEvents
        FileCopy PathConstructionModelNUMEROFIL & "Copie de NUMEROFIL40.dwg", PathConstructionModelNUMEROFIL & "NUMEROFIL" & CStr(I * 4) & ".dwg"
    Next I
    For I = 11 To 20
        DoEvents
        FileCopy PathConstructionModelNUMEROFIL & "c_NUMEROFIL80.dwg", PathConstructionModelNUMEROFIL & "NUMEROFIL" & CStr(I * 4) & ".dwg"
    Next I
End Sub

Function EnteteCartouche(varProjet As String, varIndice As String, Plan As String)
    Dim Txt
    Dim txt2
    Dim Mysapce
    Mysapce = Space(78)
          Txt = "             ******************************************************************" & vbCrLf
    Txt = Txt & "             * Listes des erreurs survenues lors de l'exécution de la macro : *" & vbCrLf
    Txt = Txt & "             * Créer un Plan                                                  *" & vbCrLf
         txt2 = "             * Projet : " & Replace(varProjet, vbCrLf, " ")
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * Plan : " & Plan & " Indice : " & varIndice
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             *"
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * Nombre d'erreur(s) : " & NbError
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
    Txt = Txt & "             ******************************************************************" & vbCrLf
    Txt = Txt & vbCrLf
    Debug.Print Txt
    EnteteCartouche = Txt
End Function

