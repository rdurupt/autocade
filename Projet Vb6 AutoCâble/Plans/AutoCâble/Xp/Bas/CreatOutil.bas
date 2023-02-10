Attribute VB_Name = "CreatOutil"
Public Sub subDessinerOtil(IdIndiceProjet As Long)
    Dim Rs As Recordset
    Dim PathPl As String
    Dim sql As String
    Set CollectionTor = Nothing
    Set CollectionTor = New Collection
    ReDim TableuDeTor(0)
     NUMCOM = 0
    NUMNOTA = 0
    NUMNTOR = 0
    NUMNTORBLOC = 0
   
    Set TableauPath = funPath
    sql = "SELECT T_indiceProjet.ou FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    
    Set Rs = Con.OpenRecordSet(sql)
NbError = 0
If Rs.EOF = True Then Exit Sub
    If MsgBox("Voulez-vous exécuter la Macro Créer/Modifier" & vbCrLf & Rs!OU, vbQuestion + vbYesNo, "Auto-Câble") = vbNo Then Exit Sub
    If ModifierUnPlan(IdIndiceProjet, "OU") = False Then
    Set TableauPath = funPath
    Dim Tableau() As String
    Dim NbFil As Long
    Dim NbLignes As Long
    Dim Fso As New FileSystemObject
    
    NbLignes = 0
    Set AutoApp = ThisDrawing.Application
    
     OpenNew
  
End If
LoadCalque
    If LoadConnecteur(IdIndiceProjet, "OU") = False Then GoTo Fin
    LoadComposants IdIndiceProjet, "OU"
LoadNotas IdIndiceProjet, "OU"
    ChargeCartoucheClient IdIndiceProjet, "OU", 1
    ChargeCartoucheEncelade IdIndiceProjet, "OU", 1
    
   
   
    SubLoadFils IdIndiceProjet
sql = "SELECT RqCartouche.* "
sql = sql & "FROM RqCartouche "
 sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
        PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version)
'
    ExporteXls PathPl, IdIndiceProjet
    PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "OU", Rs.Fields("OU"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("OU_Indice"), Rs!Version)

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


