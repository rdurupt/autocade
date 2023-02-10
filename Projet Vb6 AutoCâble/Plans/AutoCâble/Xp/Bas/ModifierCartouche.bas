Attribute VB_Name = "ModifierCartouche"
Public Function ModifierUnCartouche(IdIndiceProjet As Long, Optional Approb As Boolean) As Boolean
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
     If Approb = False Then
        If MsgBox("Voulez vous apporter les modifications du Cartouche" & _
            vbCrLf & "sur les différents plans", vbQuestion + vbYesNo, "Modification Cartouche :") = vbNo Then Exit Function
      End If
      
      Set TableauPath = funPath
     NbLignesVignette = 0
     ModifierUnCartouche = False
     


sql = "SELECT T_indiceProjet.Cartouche FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(sql)
RepPlacheClous = "" & Rs!Cartouche
If Left(RepPlacheClous, 2) <> "\\" Then RepPlacheClous = TableauPath.Item("PathServer") & RepPlacheClous

PlanchClous = Rs!Cartouche

Set Rs = Con.CloseRecordSet(Rs)
   NbConnecteur = 0
       InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 790.0288: InsertPointLigneTableau_Vignette(2) = 0
  sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set Rs = Con.OpenRecordSet(sql)
    LeCient = UCase(Trim("" & Rs!Client))
Select Case UCase(Trim("" & Rs!Client))
    Case "RENAULT"
        boolFormClient = True
        MyCARTOUCHE_Client = "RENAULT"
    Case Else
        
         MyCARTOUCHE_Client = "RENAULT"
End Select

sql = "SELECT T_indiceProjet.PlAutoCadSave,  "
sql = sql & "T_indiceProjet.PlAutoCadSaveas "
sql = sql & "FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"


    Set Rs = Con.OpenRecordSet(sql)
    If Rs.EOF = True Then
        MsgBox "Plan introuvable", vbQuestion, "Modification Cartouche :"
      GoTo OU
     Else
        If Trim("" & Rs!plAutoCadSaveas) <> "" Then
            MyFichier = "" & Rs!plAutoCadSaveas
        End If
        If Trim("" & Rs!plAutoCadSave) <> "" Then
            MyFichier = "" & Rs!plAutoCadSave
        End If
        If Trim("" & MyFichier) = "" Then
             MsgBox "Plan introuvable", vbQuestion, "Modification Cartouche :"
            GoTo OU
        End If
    End If
    
    Set TableauPath = funPath
    PathDessin = TableauPath("PathArchiveAutocad") & Trim("" & MyFichier) & ".dwg"
     If Fso.FileExists(PathDessin) = False Then
        MsgBox "Plan : " & Trim("" & MyFichier) & ".dwg introuvable", vbQuestion, "Modification Cartouche :"
        GoTo OU
     End If
         NbLignes = 0
    Set AutoApp = ThisDrawing.Application
 OpenFichier PathDessin
 
   
 
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
   
    Etiquettes(i).Delete
Next i
    ChargeCartoucheClient IdIndiceProjet, "PL", 4
    ChargeCartoucheEncelade IdIndiceProjet, "PL", 4

  AutoApp.ActiveDocument.PurgeAll
  If boolValideMOD = True Then
            sql = "SELECT RqCartouche.* "
            sql = sql & "FROM RqCartouche "
            sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
            Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = False Then
            PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "PL", Rs.Fields("PL"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("PL_Indice"), Rs!Version)
             SaveAs PathPl
             
        End If

  Else
   SaveAs PathDessin
   End If
OU:
   

  
  sql = "SELECT T_indiceProjet.ouAutoCadSave,  "
sql = sql & "T_indiceProjet.ouAutoCadSaveas "
sql = sql & "FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
    Set Rs = Con.OpenRecordSet(sql)
    If Rs.EOF = True Then
         MsgBox "Outil introuvable", vbQuestion, "Modification Cartouche :"
       GoTo Fin
     Else
        If Trim("" & Rs!OuAutoCadSaveAs) <> "" Then
            MyFichier = "" & Rs!OuAutoCadSaveAs
        End If
        If Trim("" & Rs!OuAutoCadSave) <> "" Then
            MyFichier = "" & Rs!OuAutoCadSave
        End If
        If Trim("" & MyFichier) = "" Then
            MsgBox "Outil introuvable", vbQuestion, "Modification Cartouche :"
            GoTo Fin
        End If
    End If
    
 
    PathDessin = TableauPath("PathArchiveAutocad") & Trim("" & MyFichier) & ".dwg"
     If Fso.FileExists(PathDessin) = False Then
         MsgBox "Outil : " & Trim("" & MyFichier) & ".dwg introuvable", vbQuestion, "Modification Cartouche :"
         GoTo Fin
     End If
         NbLignes = 0
    Set AutoApp = ThisDrawing.Application
 OpenFichier PathDessin
 
   
 
   
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
   
    Etiquettes(i).Delete
Next i


    ChargeCartoucheClient IdIndiceProjet, "OU", 1
    ChargeCartoucheEncelade IdIndiceProjet, "OU", 1
     
     If boolValideMOD = True Then
            sql = "SELECT RqCartouche.* "
            sql = sql & "FROM RqCartouche "
            sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
            Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = False Then
            PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "OU", Rs.Fields("OU"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("PL_Indice"), Rs!Version)
             SaveAs PathPl
             KilVersionXX PathDessin, PathPl, True
             ExporteXls PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "LI", Rs.Fields("LI"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("PL_Indice"), Rs!Version), IdIndiceProjet
        End If

  Else
    SaveAs PathDessin
End If

   ModifierUnCartouche = True
  End
Fin:
    FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1Caption.Caption = "Fin du traitement"
End Function

