Attribute VB_Name = "ModifierCartouche"
Function ModifierUnCartouche(IdIndiceProjet As Long, Optional Approb As Boolean) As Boolean
    Dim Sql As String
     Dim PathPl As String
    Dim Rs As Recordset
    Dim PathDessin As String
    Dim Fso As New FileSystemObject
    
   If boolAutoCAD = False Then Exit Function

    Dim Collec As New Collection
    Dim Etiquettes As New Collection
    Dim RefOption As New Collection
    Dim MyFichier As String
     Dim NewBlockV  As Object
     Set CollectionCon = Nothing
     Set CollectionCon = New Collection
     If Approb = False Then
        If MsgBox("Voulez vous apporter les modifications du Cartouche" & _
            vbCrLf & "sur les différents plans", vbQuestion + vbYesNo, "Modification Cartouche :") = vbNo Then Exit Function
      End If
      
      Set TableauPath = funPath
     NbLignesVignette = 0
     ModifierUnCartouche = False
     


Sql = "SELECT T_indiceProjet.Cartouche,T_indiceProjet.NbCartouche FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
RepPlacheClous = "" & Rs!Cartouche
NbCartouche = Rs!NbCartouche
RepPlacheClous = DefinirChemienComplet(TableauPath.Item("PathServer"), RepPlacheClous)
'If Left(RepPlacheClous, 2) <> "\\" And Left(RepPlacheClous, 1) = "\" Then RepPlacheClous = TableauPath.Item("PathServer") & RepPlacheClous

PlanchClous = Rs!Cartouche

Set Rs = Con.CloseRecordSet(Rs)
 
       InsertPointLigneTableau_Vignette(0) = -1146.1429: InsertPointLigneTableau_Vignette(1) = 790.0288: InsertPointLigneTableau_Vignette(2) = 0
  Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set Rs = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & Rs!Client))
Select Case UCase(Trim("" & Rs!Client))
    Case "RENAULT"
        boolFormClient = True
        MyCARTOUCHE_Client = "RENAULT"
    Case Else
        
         MyCARTOUCHE_Client = "RENAULT"
End Select

Sql = "SELECT T_indiceProjet.PlAutoCadSave,  "
Sql = Sql & "T_indiceProjet.PlAutoCadSaveas "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"


    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
        MsgBox "Plan introuvable", vbQuestion, "Modification Cartouche :"
      GoTo OU
     Else
        If Trim("" & Rs!PlAutoCadSaveAs) <> "" Then
            MyFichier = "" & Rs!PlAutoCadSaveAs
        End If
        If Trim("" & Rs!PlAutoCadSave) <> "" Then
            MyFichier = "" & Rs!PlAutoCadSave
        End If
        If Trim("" & MyFichier) = "" Then
             MsgBox "Plan introuvable", vbQuestion, "Modification Cartouche :"
            GoTo OU
        End If
    End If
    
    Set TableauPath = funPath
    MyFichier = Replace(MyFichier, ".dwg", "")
    PathDessin = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), MyFichier)
    PathDessin = PathDessin & ".dwg"
    
     If Fso.FileExists(PathDessin) = False Then
        MsgBox "Plan : " & Trim("" & MyFichier) & ".dwg introuvable", vbQuestion, "Modification Cartouche :"
        GoTo OU
     End If
         NbLignes = 0
'    'Set AutoApp = ThisDrawing.Application
SecuFill PathDessin, False
AdcFileName = OpenFichier(PathDessin)
 
   
 
    FormBarGrah.ProgressBar1Caption = " Scanne des cartouches :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
         IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
       
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
    Next I
For I = 1 To Etiquettes.Count
   
    Etiquettes(I).Delete
Next I
Set RefAcCorrective = Nothing
Set RefAcCorrective = New Collection
'FormBarGrah.Visible = True
  FormBarGrah.ProgressBar1Caption = " Scanne des Actions Corrective :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
        IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            a = BlocRef.Name
            If BlocRef.HasAttributes Then
            
                Attributes = BlocRef.GetAttributes
                If UBound(Attributes) < 2 Then
                    If IsActionCorrective(Attributes) = True Then
                        Set NewBlockV = BlocRef
                        RefAcCorrective.Add NewBlockV
                        Set NewBlockV = Nothing
                    End If
            
                End If
         End If
      End If
    Next I

   For I = 1 To RefAcCorrective.Count

    Set BlocRef = RefAcCorrective(I)
     BlocRef.Delete
   Next I
  
  
  FormBarGrah.ProgressBar1Caption = " Scanne des Options :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
    
        IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            a = BlocRef.Name
            If BlocRef.HasAttributes Then
            
                Attributes = BlocRef.GetAttributes
                If UBound(Attributes) < 2 Then
                    If IsRefOption(Attributes) = True Then
                        Set NewBlockV = BlocRef
                        RefOption.Add NewBlockV
                        Set NewBlockV = Nothing
                    End If
            
                End If
         End If
      End If
    Next I

   For I = 1 To RefOption.Count

    Set BlocRef = RefOption(I)
     BlocRef.Delete
   Next I
  
  
    ChargeCartoucheClient IdIndiceProjet, "PL", NbCartouche
    ChargeCartoucheEncelade IdIndiceProjet, "PL", NbCartouche
  DocAutoCad.PurgeAll
'  If boolValideMOD = True Then
'            Sql = "SELECT RqCartouche.* "
'            Sql = Sql & "FROM RqCartouche "
'            Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
'            Set rs = Con.OpenRecordSet(Sql)
'            If rs.EOF = False Then
'            PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "PL", rs.Fields("PL"), IdIndiceProjet, rs.Fields("PI_Indice"), rs.Fields("PL_Indice"), rs!Version)
'             SaveAs PathPl
'
'        End If
'
'  Else
    Sql = "SELECT RqCartouche.* "
            Sql = Sql & "FROM RqCartouche "
            Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = False Then
                SaveAs PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "PL", Rs.Fields("PL"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("PL_Indice"), Rs!Version)
'            End If
   End If
OU:
   

  
  Sql = "SELECT T_indiceProjet.ouAutoCadSave,  "
Sql = Sql & "T_indiceProjet.ouAutoCadSaveas "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
         MsgBox "Outil introuvable", vbQuestion, "Modification Cartouche :"
       GoTo Fin
     Else
        If Trim("" & Rs!OuAutoCadSaveAs) <> "" Then
            MyFichier = "" & Rs!OuAutoCadSaveAs
        End If
        If Trim("" & Rs!OUAutoCadSave) <> "" Then
            MyFichier = "" & Rs!OUAutoCadSave
        End If
        If Trim("" & MyFichier) = "" Then
            MsgBox "Outil introuvable", vbQuestion, "Modification Cartouche :"
            GoTo Fin
        End If
    End If
    
 MyFichier = Replace(MyFichier, ".dwg", "")
    PathDessin = DefinirChemienComplet(TableauPath("PathArchiveAutocad"), MyFichier)
    PathDessin = PathDessin & ".dwg"
     If Fso.FileExists(PathDessin) = False Then
         MsgBox "Outil : " & Trim("" & MyFichier) & ".dwg introuvable", vbQuestion, "Modification Cartouche :"
         GoTo Fin
     End If
         NbLignes = 0
    'Set AutoApp = ThisDrawing.Application
    SecuFill PathDessin, False
 AdcFileName = OpenFichier(PathDessin)
 
   
 Set Etiquettes = Nothing
 Set Etiquettes = New Collection
   
    FormBarGrah.ProgressBar1Caption = " Scanne des cartouches :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
         IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        
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
    Next I
For I = 1 To Etiquettes.Count
   
    Etiquettes(I).Delete
Next I
Set RefAcCorrective = Nothing
Set RefAcCorrective = New Collection
  FormBarGrah.ProgressBar1Caption = " Scanne des Actions Corrective :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
        IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            a = BlocRef.Name
            If BlocRef.HasAttributes Then
            
                Attributes = BlocRef.GetAttributes
                If UBound(Attributes) < 2 Then
                    If IsActionCorrective(Attributes) = True Then
                        Set NewBlockV = BlocRef
                        RefAcCorrective.Add NewBlockV
                        Set NewBlockV = Nothing
                    End If
            
                End If
         End If
      End If
    Next I

   For I = 1 To RefAcCorrective.Count

    Set BlocRef = RefAcCorrective(I)
     BlocRef.Delete
   Next I
  
   Set RefOption = Nothing
 Set RefOption = New Collection
 
  FormBarGrah.ProgressBar1Caption = " Scanne des Options :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    Set Etiquettes = Nothing
    Set Etiquettes = New Collection
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
        IncremanteBarGrah FormBarGrah
        DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            a = BlocRef.Name
            If BlocRef.HasAttributes Then
            
                Attributes = BlocRef.GetAttributes
                If UBound(Attributes) < 2 Then
                    If IsRefOption(Attributes) = True Then
                        Set NewBlockV = BlocRef
                        RefOption.Add NewBlockV
                        Set NewBlockV = Nothing
                    End If
            
                End If
         End If
      End If
    Next I

   For I = 1 To RefOption.Count

    Set BlocRef = RefOption(I)
     BlocRef.Delete
   Next I

    ChargeCartoucheClient IdIndiceProjet, "OU", 1
    ChargeCartoucheEncelade IdIndiceProjet, "OU", 1
'
'            PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version)
'    ExporteXls PathPl, IdIndiceProjet
    
'
'  Else
     Sql = "SELECT RqCartouche.* "
            Sql = Sql & "FROM RqCartouche "
            Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = False Then
    SaveAs PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "OU", Rs.Fields("OU"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("OU_Indice"), Rs!Version)
'    If Fso.FileExists(DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & rs![LiAutoCadSave]) & ".XLS") = True Then
'      Fso.CopyFile DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & rs![LiAutoCadSave]) & ".XLS", PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "LI", rs.Fields("LI"), IdIndiceProjet, rs.Fields("PI_Indice"), rs.Fields("LI_Indice"), rs!Version) & ".XLS"
'    End If
    End If
'End If


 If boolValideMOD = True Then
'            Sql = "SELECT RqCartouche.* "
'            Sql = Sql & "FROM RqCartouche "
'            Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
'            Set rs = Con.OpenRecordSet(Sql)
Rs.Requery
            If Rs.EOF = False Then
'            PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "OU", rs.Fields("OU"), IdIndiceProjet, rs.Fields("PI_Indice"), rs.Fields("PL_Indice"), rs!Version)
'             SaveAs PathPl
'             If Fso.FileExists(DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & rs![LiAutoCadSave]) & ".XLS") = True Then
'                Fso.CopyFile DefinirChemienComplet(TableauPath("PathArchiveAutocad"), "" & rs![LiAutoCadSave]) & ".XLS", PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "LI", rs.Fields("LI"), IdIndiceProjet, rs.Fields("PI_Indice"), rs.Fields("PL_Indice"), rs!Version) & ".XLS"
'              End If
             ExporteXls PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "LI", Rs.Fields("LI"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("PL_Indice"), Rs!Version), IdIndiceProjet
'             KilVersionXX PathDessin, PathPl, True
            End If
        End If
   ModifierUnCartouche = True

Fin:
    FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1Caption.Caption = " Fin du traitement"
  
End Function



