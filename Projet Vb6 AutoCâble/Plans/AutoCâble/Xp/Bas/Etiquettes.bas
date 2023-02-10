Attribute VB_Name = "Etiquettes"
Option Explicit
Dim ClipJointExiste As Collection
Function IsClpJointExiste(Value As String) As Boolean
Dim Txt As String
On Error Resume Next
    Txt = ClipJointExiste(Value)
    If Err Then
        Err.Clear
        ClipJointExiste.Add Value, Value
        Exit Function
    End If
On Error GoTo 0
IsClpJointExiste = True
End Function
Public Sub GenairEtiquette2(Id_IndiceProjet As Long, Options As String, SurOption As Boolean, SurFournisseur As Boolean)
Dim Sql As String
Dim Rs As Recordset
Dim PathModelWord As String
Dim MyEtiquette As ClsEtiqette
Dim tableau
Dim tableau2
Dim tableau3
Dim I As Long
Dim RefJoint As String
Dim TableauJoint
Dim RefBouchon As String
Dim TableauBouchon
Dim refVerrou As String
Dim TableauVerrou
Dim RefCapot As String
Dim TableauCapot
Dim saveFamille As String
Dim PathPl As String
Dim PI As String
Dim BarrGraphCoun As Long
Dim SaveApp As String
Dim EditEtiquette As Boolean
Dim CloseWher As String
Dim Equippement As String
Dim Ensemble As String
If Left(CloseWher, 1) <> ";" Then CloseWher = ";" & CloseWher
If Right(CloseWher, 1) <> ";" Then CloseWher = CloseWher & ";"
 CloseWher = " AND (';' + [Options] + ';' Like '%;TOUS;%' or ';' + [Options] + ';' Like '%;ALL;%' Or ';' + [Options] + ';' Like '%" & Options + "%')  "
Set MyWord = CreateObject("Word.Application")
'MyWord.Visible = True
PathModelWord = TableauPath.Item("PathModelWord")
PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
Set MyWordDoc = WordNewDocApp(PathModelWord, MyWord)
          
          
'MyWordDoc.Application.Visible = True
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE RqCartouche.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
    MyWord.Quit False
    Exit Sub
End If
   PI = "" & Rs!PI & "_" & Rs!PI_Indice
   Ensemble = "" & Rs!Ensemble
PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), Id_IndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version, True)

PathModelWord = TableauPath.Item("PathModelWordMarc")
PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
Set MyWordDoc2 = WordNewDocApp(PathModelWord, MyWord)
Sql = "SELECT T_Nomenclature.CODE_APP, T_Nomenclature.CONNECTEUR, T_Nomenclature.[Famille Lib] ,  "
Sql = Sql & "T_Nomenclature.[Joint Four Réf] , T_indiceProjet.Ensemble, [PI] & '_' & [PI_Indice] AS PIE,  "
Sql = Sql & "T_Nomenclature.DESIGNATION, T_Nomenclature.[Alvé Réf Fourr], T_Nomenclature.[Ref Capot], T_Nomenclature.[Ref Verrou], T_Nomenclature.[Bouch Réf Four] "
Sql = Sql & "FROM T_indiceProjet INNER JOIN T_Nomenclature ON T_indiceProjet.Id = T_Nomenclature.Id_IndiceProjet "
Sql = Sql & "Where T_Nomenclature.Id_IndiceProjet = " & Id_IndiceProjet & " " '& " and T_Nomenclature.CONNECTEUR='7703297954'"
Sql = Sql & "ORDER BY T_Nomenclature.CONNECTEUR;"

Sql = "SELECT NomeclatureConnecteurs.App, NomeclatureConnecteurs.Designation, NomeclatureConnecteurs.Connecteur AS RefConnecteur,  "
Sql = Sql & "NomeclatureConnecteurs.Bouchon AS RefBouchon, NomeclatureConnecteurs.Capot AS RefCapot, NomeclatureConnecteurs.Verrou AS RefVerrou,  "
Sql = Sql & "NomeclatureConnecteurs.Options, NomeclatureConnecteurs.Clip AS RefClip, Count(NomeclatureConnecteurs.Clip) AS CompteDeClip,  "
Sql = Sql & "NomeclatureConnecteurs.Joint AS RefJont, Count(NomeclatureConnecteurs.Joint) AS CompteDeJoint "
Sql = Sql & "FROM NomeclatureConnecteurs "








Sql = "SELECT NomeclatureConnecteurs.App, NomeclatureConnecteurs.Designation,   "
Sql = Sql & "NomeclatureConnecteurs.Connecteur, NomeclatureConnecteurs.Connecteur_Four,   "
Sql = Sql & "NomeclatureConnecteurs.Bouchon, NomeclatureConnecteurs.Capot,   "
Sql = Sql & "NomeclatureConnecteurs.Capot_Four, NomeclatureConnecteurs.Verrou,   "
Sql = Sql & "NomeclatureConnecteurs.Verrout_Four, NomeclatureConnecteurs.Options,   "
Sql = Sql & "NomeclatureConnecteurs.Clip , Count(NomeclatureConnecteurs.Clip) AS CompteDeClip,   "
Sql = Sql & "NomeclatureConnecteurs.ClipFour as [Famille], NomeclatureConnecteurs.Joint,   "
Sql = Sql & "Count(NomeclatureConnecteurs.Joint) AS CompteDeJoint,   "
Sql = Sql & "NomeclatureConnecteurs.JointFour  as [Ref Joint]  "
Sql = Sql & "FROM NomeclatureConnecteurs  "

If SurFournisseur = False Then
    Sql = "SELECT MyFrom.App, MyFrom.Designation, MyFrom.Connecteur AS RefConnecteur, MyFrom.Bouchon AS RefBouchon, MyFrom.Capot AS RefCapot,   "
    Sql = Sql & "MyFrom.Verrou AS RefVerrou, MyFrom.Clip AS RefClip, Count(MyFrom.Clip) AS CompteDeClip, MyFrom.Joint AS [Ref Joint],   "
    Sql = Sql & "Count(MyFrom.Joint) AS CompteDeJoint "

Else
    
    Sql = "SELECT MyFrom.App, MyFrom.Designation, MyFrom.Connecteur_Four AS RefConnecteur, MyFrom.BouchonFour AS RefBouchon,  "
    Sql = Sql & "MyFrom.Capot_Four AS RefCapot, MyFrom.Verrout_Four AS RefVerrou, MyFrom.ClipFour AS RefClip, Count(MyFrom.Clip) AS  "
    Sql = Sql & "CompteDeClip, MyFrom.JointFour AS [Ref Joint], Count(MyFrom.Joint) AS CompteDeJoint "
   
End If


    Sql = Sql & "FROM (SELECT NomeclatureConnecteurs.Liaison, NomeclatureConnecteurs.App, NomeclatureConnecteurs.Designation,  "
    Sql = Sql & "NomeclatureConnecteurs.Connecteur, NomeclatureConnecteurs.Connecteur_Four, NomeclatureConnecteurs.Clip,  "
    Sql = Sql & "NomeclatureConnecteurs.ClipFour, NomeclatureConnecteurs.Joint, NomeclatureConnecteurs.JointFour, NomeclatureConnecteurs.Bouchon,  "
    Sql = Sql & "NomeclatureConnecteurs.BouchonFour, NomeclatureConnecteurs.Capot, NomeclatureConnecteurs.Capot_Four, NomeclatureConnecteurs.Verrou,  "
    Sql = Sql & "NomeclatureConnecteurs.Verrout_Four, NomeclatureConnecteurs.Options "
    Sql = Sql & "FROM NomeclatureConnecteurs "
   Sql = Sql & "Where NomeclatureConnecteurs.Id_IndiceProjet = " & Id_IndiceProjet & "  "
    If SurOption = True Then Sql = Sql & CloseWher
    Sql = Sql & "GROUP BY NomeclatureConnecteurs.Liaison, NomeclatureConnecteurs.App, NomeclatureConnecteurs.Designation,  "
    Sql = Sql & "NomeclatureConnecteurs.Connecteur, NomeclatureConnecteurs.Connecteur_Four, NomeclatureConnecteurs.Clip,  "
    Sql = Sql & "NomeclatureConnecteurs.ClipFour, NomeclatureConnecteurs.Joint, NomeclatureConnecteurs.JointFour,  "
    Sql = Sql & "NomeclatureConnecteurs.Bouchon, NomeclatureConnecteurs.BouchonFour, NomeclatureConnecteurs.Capot,  "
    Sql = Sql & "NomeclatureConnecteurs.Capot_Four, NomeclatureConnecteurs.Verrou, NomeclatureConnecteurs.Verrout_Four,  "
    Sql = Sql & "NomeclatureConnecteurs.Options "
    
     Sql = Sql & ") AS MyFrom "





If SurFournisseur = False Then
   
    Sql = Sql & "GROUP BY MyFrom.App, MyFrom.Designation, MyFrom.Connecteur, MyFrom.Bouchon, MyFrom.Capot,MyFrom.Verrou, MyFrom.Clip, MyFrom.Joint  "
    
Else
     
   
    Sql = Sql & "GROUP BY MyFrom.App, MyFrom.Designation, MyFrom.Connecteur_Four, MyFrom.BouchonFour,  "
    Sql = Sql & "MyFrom.Capot_Four, MyFrom.Verrout_Four, MyFrom.ClipFour, MyFrom.JointFour "
End If

Sql = Sql & " ORDER BY MyFrom.App;"
MyWord.Visible = True
Set Rs = Con.OpenRecordSet(Sql)
BarrGraphCoun = 0
While Rs.EOF = False
    BarrGraphCoun = BarrGraphCoun + 1
    Rs.MoveNext
Wend
Rs.Requery
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Max = BarrGraphCoun + 1
FormBarGrah.ProgressBar1Caption.Caption = "Générer étiquettes"
While Rs.EOF = False
IncremanteBarGrah FormBarGrah
IncrmentServer FormBarGrah, "Etiquettes"
If SaveApp <> "" & Rs!App Then
    Set ClipJointExiste = Nothing
    Set ClipJointExiste = New Collection
    If SaveApp <> "" Then
    On Error GoTo Fin
        For I = MyEtiquette.TableMin To MyEtiquette.TableMax
        '      MyEtiquette.RetournEtiquette I
            CreerEtiquette MyEtiquette.RetournEtiquette(I)
        Next
Fin:
    End If
    SaveApp = "" & Rs!App
    Set MyEtiquette = Nothing
    Set MyEtiquette = New ClsEtiqette
    MyEtiquette.PrpareEtiqet "" & Rs!App, "" & Rs!DESIGNATION
    MyEtiquette.RenseigneChamp "PI", "" & PI
    MyEtiquette.RenseigneChamp "Ensemble", "" & Ensemble
End If
For I = 0 To Rs.Fields.Count - 1
    Select Case UCase(Rs(I).Name)  '= UCase("Alvé Réf") Then
        Case UCase("RefClip")
        If Trim("" & Rs(I)) <> "" Then
            If IsClpJointExiste(Rs(I).Value) = False Then
                MyEtiquette.RenseigneChamp "" & Rs(I).Name, "" & Rs(I).Value & "(" & Rs!CompteDeClip & ")", True
             End If
        End If
        
     Case UCase("Ref Joint")
        If Trim("" & Rs(I)) <> "" Then
             If IsClpJointExiste(Rs(I).Value) = False Then
                MyEtiquette.RenseigneChamp "" & Rs(I).Name, "" & Rs(I).Value & "(" & Rs!CompteDeJoint & ")", True
             End If
        End If
'    Case UCase("Ref Joint")
'        If Trim("" & rs(I)) <> "" Then
'            MyEtiquette.RenseigneChamp "" & rs(I).Name, "" & rs(I).Value & "(" & rs!CompteDeJoint & ")"
'        End If
    Case Else
        On Error Resume Next
            MyEtiquette.RenseigneChamp "" & Rs(I).Name, "" & Rs(I), True
       
    End Select
Next






    Rs.MoveNext
Wend
On Error GoTo Fin2
For I = MyEtiquette.TableMin To MyEtiquette.TableMax
        '      MyEtiquette.RetournEtiquette I
            CreerEtiquette MyEtiquette.RetournEtiquette(I)
        Next
Fin2:
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1Caption.Caption = "Traitement terminé"
IncrmentServer FormBarGrah, ""
If SurOption = True Then
    MyWordSaveAs "" & PathPl, True, "_" & Replace(Options, ";", "_")
Else
    MyWordSaveAs "" & PathPl, True
End If
End Sub
Public Sub GenairEtiquette(Id_IndiceProjet As Long)
Dim Sql As String
Dim Rs As Recordset
Dim PathModelWord As String
Dim MyEtiquette As ClsEtiqette
Dim tableau
Dim tableau2
Dim tableau3
Dim I As Long
Dim RefJoint As String
Dim TableauJoint
Dim RefBouchon As String
Dim TableauBouchon
Dim refVerrou As String
Dim TableauVerrou
Dim RefCapot As String
Dim TableauCapot
Dim saveFamille As String
Dim PathPl As String
Dim BarrGraphCoun As Long
Set MyWord = CreateObject("Word.Application")
'MyWord.Visible = True
PathModelWord = TableauPath.Item("PathModelWord")
PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
Set MyWordDoc = WordNewDocApp(PathModelWord, MyWord)
          
          

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE RqCartouche.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
    MyWord.Quit False
    Exit Sub
End If
' MyWord.Visible = True
PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), Id_IndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version, True)

PathModelWord = TableauPath.Item("PathModelWordMarc")
PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
Set MyWordDoc2 = WordNewDocApp(PathModelWord, MyWord)
Sql = "SELECT T_Nomenclature.CODE_APP, T_Nomenclature.CONNECTEUR, T_Nomenclature.[Famille Lib] ,  "
Sql = Sql & "T_Nomenclature.[Joint Four Réf] , T_indiceProjet.Ensemble, [PI] + '_' + [PI_Indice] AS PIE,  "
Sql = Sql & "T_Nomenclature.DESIGNATION, T_Nomenclature.[Alvé Réf Fourr], T_Nomenclature.[Ref Capot], T_Nomenclature.[Ref Verrou], T_Nomenclature.[Bouch Réf Four] "
Sql = Sql & "FROM T_indiceProjet INNER JOIN T_Nomenclature ON T_indiceProjet.Id = T_Nomenclature.Id_IndiceProjet "
Sql = Sql & "Where T_Nomenclature.Id_IndiceProjet = " & Id_IndiceProjet & " " '& " and T_Nomenclature.CONNECTEUR='7703297954'"
Sql = Sql & "ORDER BY T_Nomenclature.CONNECTEUR;"



Set Rs = Con.OpenRecordSet(Sql)
BarrGraphCoun = 0
While Rs.EOF = False
    BarrGraphCoun = BarrGraphCoun + 1
    Rs.MoveNext
Wend
Rs.Requery
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Max = BarrGraphCoun + 1
While Rs.EOF = False
FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
Set MyEtiquette = New ClsEtiqette
'MyWord.Visible = True
MyEtiquette.PrpareEtiqet "" & Rs!Code_APP, "" & Rs!DESIGNATION
MyEtiquette.RenseigneChamp "PI", "" & Rs!PIE
For I = 0 To Rs.Fields.Count - 1
    MyEtiquette.RenseigneChamp "" & Rs(I).Name, "" & Rs(I).Value
Next

'MyEtiquette.RenseigneChamp "" & Rs2(I).Name, "" & Rs2(I).Value
'
'    MyEtiquette.RenseigneChamp "Connecteur", "" & Rs!Connecteur
'
'      MyEtiquette.RenseigneChamp "Ensemble", "" & Rs!Ensemble
tableau = Split("" & Rs![Famille Lib], Chr(10))
tableau2 = Split("" & Rs![Alvé Réf Fourr], Chr(10))
tableau3 = ""
saveFamille = ""
For I = 0 To UBound(tableau)
    If Trim("" & tableau(I)) <> "" Then
        If Trim("" & tableau(I)) <> saveFamille Then
            saveFamille = tableau(I)
            If Trim("" & tableau3) = "" Then
                tableau3 = tableau(I) & ": "
            Else
                tableau3 = tableau3 & " ;" & tableau(I) & ": "
            End If
              
        End If
          tableau3 = tableau3 & tableau2(I) & "(_____), "
    End If
'    tableauAlve2(TbAlve(TableauAlve(3, I)), 1) = tableauAlve2(TbAlve(TableauAlve(3, I)), 1) & "" & TableauAlve(5, I) & "(_____), "
Next
If Right("" & tableau3, 1) = ", " Then tableau3 = Left(tableau3, Len("" & tableau3) - 2)
 MyEtiquette.RenseigneChamp "Alvé Réf", "" & tableau3
 MyEtiquette.RenseigneChamp "FAMILLE", "" & tableau3
  
 TableauJoint = Split("" & Rs![Joint Four Réf] & Chr(10), Chr(10))
 RefJoint = ""
 For I = 0 To UBound(TableauJoint) - 1
    If Trim("" & TableauJoint(I)) <> "" Then
        RefJoint = RefJoint & TableauJoint(I) & "(_____), "
    End If
 Next
  TableauBouchon = Split("" & Rs![Bouch Réf Four] & Chr(10), Chr(10))
  RefBouchon = ""
  For I = 0 To UBound(TableauBouchon) - 1
    If Trim("" & TableauBouchon(I)) <> "" Then
        RefBouchon = RefBouchon & TableauBouchon(I) & "(_____), "
    End If
 Next
 If Right(RefBouchon, 2) = ", " Then RefBouchon = Left(RefBouchon, Len(RefBouchon) - 2)
 MyEtiquette.RenseigneChamp "Bouchon", "" & RefBouchon
 
 
 
 
 TableauVerrou = Split("" & Rs![Ref Verrou] & Chr(10), Chr(10))
  refVerrou = ""
  For I = 0 To UBound(TableauVerrou) - 1
    If Trim("" & TableauVerrou(I)) <> "" Then
        refVerrou = refVerrou & TableauVerrou(I) & "(_____), "
    End If
 Next
 If Right(refVerrou, 2) = ", " Then refVerrou = Left(refVerrou, Len(refVerrou) - 2)
 MyEtiquette.RenseigneChamp "Verrou", "" & refVerrou
 
 TableauCapot = Split("" & Rs![Ref Capot] & Chr(10), Chr(10))
  RefCapot = ""
  For I = 0 To UBound(TableauCapot) - 1
    If Trim("" & TableauCapot(I)) <> "" Then
        RefCapot = RefCapot & TableauCapot(I) & "(_____), "
    End If
 Next
 If Right(refVerrou, 2) = ", " Then refVerrou = Left(refVerrou, Len(refVerrou) - 2)
 MyEtiquette.RenseigneChamp "Capot", "" & RefCapot
 '

If Right(RefJoint, 2) = ", " Then RefJoint = Left(RefJoint, Len(RefJoint) - 2)
Debug.Print RefJoint
MyEtiquette.RenseigneChamp "Ref Joint", "" & RefJoint
On Error GoTo Fin
 For I = MyEtiquette.TableMin To MyEtiquette.TableMax
'      MyEtiquette.RetournEtiquette I
        CreerEtiquette MyEtiquette.RetournEtiquette(I)
      Next
Fin:
Set MyEtiquette = Nothing
    Rs.MoveNext
Wend
MyWordSaveAs "" & PathPl, True
End Sub

