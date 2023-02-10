Attribute VB_Name = "CreatPlan"

Public Sub subDessinerPlan(IdIndiceProjet As Long)
    Dim Rs As Recordset
    Dim PathPl As String
    Dim sql As String
    Set TableauPath = funPath
    Dim PathPlantVierge As String
    Set TableauPath = funPath
    Set CollectionTor = Nothing
    Set CollectionTor = New Collection
     NUMCOM = 0
    NUMNOTA = 0
    NUMNTOR = 0
     NUMNTORBLOC = 0
    PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
     If Left(PathArchiveAutocad, 2) <> "\\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad
  
    sql = "SELECT T_indiceProjet.PL FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    
    Set Rs = Con.OpenRecordSet(sql)
     PathPlantVierge = TableauPath.Item("PathPlantVierge")
     If Left(PathPlantVierge, 2) <> "\\" Then PathPlantVierge = TableauPath.Item("PathServer") & PathPlantVierge
NbError = 0
If Rs.EOF = True Then Exit Sub
    If MsgBox("Voulez-vous exécuter la Macro Créer/Modifier" & vbCrLf & Rs!PL, vbQuestion + vbYesNo, "Auto-Câble") = vbNo Then Exit Sub
    If ModifierUnPlan(IdIndiceProjet, "PL") = False Then
    Set TableauPath = funPath
    Dim Tableau() As String
    Dim NbFil As Long
    Dim NbLignes As Long
    Dim Fso As New FileSystemObject
    ReDim TableuDeTor(0)
    NbLignes = 0
'    Set AutoApp = ThisDrawing.Application
    
    If Fso.FileExists(PathPlantVierge) = False Then
        MsgBox "Err"
        Exit Sub
    End If
    OpenFichier PathPlantVierge
  
End If
LoadCalque
    If LoadConnecteur(IdIndiceProjet, "PL") = False Then GoTo Fin
LoadComposants IdIndiceProjet, "PL"
LoadNotas IdIndiceProjet, "PL"
    ChargeCartoucheClient IdIndiceProjet, "PL", 4
    ChargeCartoucheEncelade IdIndiceProjet, "PL", 4
   
    SubLoadFils IdIndiceProjet
sql = "SELECT RqCartouche.* "
sql = sql & "FROM RqCartouche "
 sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
        PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version)
    ExporteXls PathPl, IdIndiceProjet
    PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "PL", Rs.Fields("PL"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("PL_Indice"), Rs!Version)

    End If
    SaveAs PathPl
        DoEvents

Fin:
    ReDim TableauDeConnecteurs(0)
    AfficheErreur PathPl, EnteteCartouche
    
    
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1Caption.Caption = "Fin du traitement"
MsgBox "Fin du traitement" & vbCrLf & NbError & " Erreur(s) détectée(s) !"
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

