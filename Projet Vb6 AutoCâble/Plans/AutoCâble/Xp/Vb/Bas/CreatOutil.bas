Attribute VB_Name = "CreatOutil"
Public Sub subDessinerOtil(IdIndiceProjet As Long)
If boolAutoCAD = False Then Exit Sub
NotSaveRacourci = False
    Dim rs As Recordset
    Dim Rs2 As Recordset
    Dim PathPl As String
    Dim sql As String
    Set CollectionTor = Nothing
    Set CollectionTor = New Collection
    ReDim TableuDeTor(0)
    ReDim TableuEtiquettes(0)
     NUMCOM = 0
    NUMNOTA = 0
    NUMNTOR = 0
    NUMNTORBLOC = 0
   
    Set TableauPath = funPath
    sql = "SELECT T_indiceProjet.ou FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    
    Set rs = Con.OpenRecordSet(sql)

If rs.EOF = True Then Exit Sub
    If bool_Outil_Ouvrir = False Then Exit Sub
    NbError = 0
    AutoApp.Visible = True
    If ModifierUnPlan(IdIndiceProjet, "OU") = False Then
    Set TableauPath = funPath
    Dim Tableau() As String
    Dim NbFil As Long
    Dim NbLignes As Long
    Dim Fso As New FileSystemObject
    
    NbLignes = 0
'    'Set AutoApp = ThisDrawing.Application
    
     OpenNew
  
End If
LoadCalque
     LoadConnecteur IdIndiceProjet, "OU"
     LoadNoeuds IdIndiceProjet, "OU"
    LoadComposants IdIndiceProjet, "OU"
LoadNotas IdIndiceProjet, "OU"
    ChargeCartoucheClient IdIndiceProjet, "OU", 1
    ChargeCartoucheEncelade IdIndiceProjet, "OU", 1
    SubLoadFils IdIndiceProjet, "OU"
    
    
 

    

    
   
sql = "SELECT RqCartouche.* "
sql = sql & "FROM RqCartouche "
 sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set rs = Con.OpenRecordSet(sql)
If rs.EOF = False Then
        PathPl = PathArchive(PathArchiveAutocad, "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "Li", rs.Fields("Li"), IdIndiceProjet, rs.Fields("PI_Indice"), rs.Fields("LI_Indice"), rs!Version)
    ExporteXls PathPl, IdIndiceProjet
    If IdFils <> 0 Then
        sql = "SELECT RqCartouche.* "
        sql = sql & "FROM RqCartouche "
        sql = sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(sql)
         PathPl2 = PathArchive(PathArchiveAutocad, "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", rs.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
       Racourci "" & PathPl2, PathPl, "XLS"
    End If
    PathPl = PathArchive(PathArchiveAutocad, "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "OU", rs.Fields("OU"), IdIndiceProjet, rs.Fields("PI_Indice"), rs.Fields("OU_Indice"), rs!Version)

    End If
    SaveAs PathPl
     If IdFils <> 0 Then
        sql = "SELECT RqCartouche.* "
        sql = sql & "FROM RqCartouche "
        sql = sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(sql)
         PathPl2 = PathArchive(PathArchiveAutocad, "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "OU", rs.Fields("OU"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("OU_Indice"), Rs2!Version)
       Racourci "" & PathPl2, PathPl, "dwg"
    End If
        DoEvents

Fin:
    ReDim TableauDeConnecteurs(0)
    AfficheErreur PathPl, EnteteCartouche("" & rs.Fields("Ensemble"), "" & rs.Fields("OU_Indice"), "" & rs.Fields("OU"))
    
  EporteSynthese
 sql = "SELECT T_indiceProjet.CleAc FROM T_indiceProjet WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
 Set rs = Con.OpenRecordSet(sql)
 If rs.EOF = False Then EporteSynthese rs!CleAc
  
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1Caption.Caption = " Fin du traitement"
     AutoApp.Visible = False
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

Function EnteteCartouche(varProjet As String, varIndice As String, Outils As String)
    Dim txt
    Dim txt2
    Dim Mysapce
     Mysapce = Space(78)
          txt = "             ******************************************************************" & vbCrLf
    txt = txt & "             * Listes des erreurs survenues lors de l'exécution de la macro : *" & vbCrLf
    txt = txt & "             * Créer un Outil                                                 *" & vbCrLf
         txt2 = "             * Projet1 : " & Replace(varProjet, vbCrLf, " ")
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * Outil : " & Outils & " Indice : " & Trim(varIndice)
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             *            "
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * Nombre d'erreur(s) : " & NbError
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
    txt = txt & "             ******************************************************************" & vbCrLf
    txt = txt & vbCrLf
    Debug.Print txt
    EnteteCartouche = txt
End Function


