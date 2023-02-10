Attribute VB_Name = "PrepareNomenclature"
Option Explicit
Dim MyColectionClsNomenclature As Collection
Dim BarrGraphCoun As Long
Dim NuInit  As String
Sub Generer_NomenclatuerFinal(Id_IndiceProjet As Long)
Dim Sql As String
Dim Rs As Recordset
Dim RecordCount As Long
Dim I As Long
Dim Client As String
Dim AppColection As New Collection
Dim Connecteur As ClsNomanclatureGenerer
Dim RsBouchon As Recordset
'Dim RsIdMenu As Recordset
Dim TableauBouchonQTS() As Long
Dim TableauBouchonLib() As String
Dim TableauBouchonFourLib() As String
Dim TableauOption() As String
Dim ColecBouchon As New Collection
Dim ChampCli As String
Dim ChampReff As String
Dim ChampQuantiteEncelade As String
Dim ChampQuantite As String
Dim Prixderevient As String
Dim RefCaddyPrixU As String
Dim RefCaddyDesignation As String
Dim RsIdMenu As Recordset
Dim RefCaddy As String
Dim RefFour As String
'LoadDb
Set TableauPath = funPath
Sql = "DELETE NomenclaturFinal.* , NomenclaturFinal.Id_IndiceProjet FROM NomenclaturFinal "
Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & Id_IndiceProjet & ";"

Con.Execute Sql

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Client = Trim("" & Rs!Client)
Else
    Client = "RENAULT"
End If '


Sql = "INSERT INTO NomenclaturFinal ( Designation, Ref, RefFour, Qts, Id_IndiceProjet, ISO, TEINT, TEINT2, SECT ) "
Sql = Sql & "SELECT Nomenclature2.Designation, Nomenclature2.Ref, Nomenclature2.RefFour, Sum(Nomenclature2.Qts)  "
Sql = Sql & "AS SommeDeQts, Nomenclature2.Id_IndiceProjet, Nomenclature2.ISO, Nomenclature2.TEINT,  "
'Sql = Sql & "Nomenclature2.TEINT2, Nomenclature2.SECT "
'Sql = Sql & "FROM Nomenclature2 "
'Sql = Sql & "GROUP BY Nomenclature2.Designation, Nomenclature2.Ref,  "
'Sql = Sql & "Nomenclature2.RefFour, Nomenclature2.Id_IndiceProjet, Nomenclature2.ISO,  "
'Sql = Sql & "Nomenclature2.TEINT, Nomenclature2.TEINT2, Nomenclature2.SECT "
'Sql = Sql & "HAVING Nomenclature2.Id_IndiceProjet=" & Id_IndiceProjet & ";"

Sql = "INSERT INTO NomenclaturFinal ( Designation, Ref, RefFour, Qts, Id_IndiceProjet, ISO, TEINT, TEINT2, SECT, Options ) "
Sql = Sql & "SELECT Nomenclature2.Designation, Nomenclature2.Ref, Nomenclature2.RefFour, Sum(Nomenclature2.Qts) AS SommeDeQts,  "
Sql = Sql & "Nomenclature2.Id_IndiceProjet, Nomenclature2.ISO, Nomenclature2.TEINT, Nomenclature2.TEINT2, Nomenclature2.SECT, Nomenclature2.Options "
Sql = Sql & "FROM Nomenclature2 "
Sql = Sql & "GROUP BY Nomenclature2.Designation, Nomenclature2.Ref, Nomenclature2.RefFour, Nomenclature2.Id_IndiceProjet,  "
Sql = Sql & "Nomenclature2.ISO, Nomenclature2.TEINT, Nomenclature2.TEINT2, Nomenclature2.SECT, Nomenclature2.Options "
Sql = Sql & "HAVING Nomenclature2.Id_IndiceProjet=" & Id_IndiceProjet & ""

Con.Execute Sql

Dim Clip
Dim Clip2
Dim ClipFour1
Dim ClipFour2
Sql = "SELECT  NomenclaturFinal.Famille, NomenclaturFinal.Ref,  "
Sql = Sql & "NomenclaturFinal.RefFour "
Sql = Sql & "FROM NomenclaturFinal "
Sql = Sql & "WHERE  NomenclaturFinal.Id_IndiceProjet=" & Id_IndiceProjet & " ;"
Set Rs = Con.OpenRecordSet(Sql)
Dim MySplit
While Rs.EOF = False
    MySplit = Split("" & Rs!Ref & "§", "§")
    Rs!Ref = Replace(Trim("" & MySplit(0)), Chr(10), "")
    MySplit = Split("" & Rs!RefFour & "§", "§")
    Rs!RefFour = Replace(Trim("" & MySplit(0)), Chr(10), "")
    Rs.Update
    Rs.MoveNext
Wend


Sql = "SELECT  NomenclaturFinal.Famille, NomenclaturFinal.Ref,  "
Sql = Sql & "NomenclaturFinal.RefFour "
Sql = Sql & "FROM NomenclaturFinal "
Sql = Sql & "WHERE NomenclaturFinal.Designation='Clips'  "
Sql = Sql & "AND NomenclaturFinal.Id_IndiceProjet=" & Id_IndiceProjet & " ;"
Set Rs = Con.OpenRecordSet(Sql)

While Rs.EOF = False
    Clip = Split("" & Rs!Ref & ":", ":")
    ClipFour1 = Split("" & Rs!RefFour & ":", ":")
    Rs!Famille = Replace(Trim("" & Clip(0)), Chr(10), "")
    Rs!Ref = Replace(Trim("" & Clip(1)), Chr(10), "")
    Rs!RefFour = Replace(Trim("" & ClipFour1(1)), Chr(10), "")
    Rs.Update
    Rs.MoveNext
Wend

Dim Lib_Menu As String
Dim RsMenuEboutique As Recordset
Sql = "SELECT Menu.libelle, Menu.Base FROM Menu IN '"
Sql = Sql & TableauPath("Eb_menu")
Sql = Sql & "' WHERE Menu.PasVisible=False;"
Set RsMenuEboutique = Con.OpenRecordSet(Sql)
While RsMenuEboutique.EOF = False
    ChampReff = EboutiqueEboutiqueGetDefault("ChampReff", "txt1", TableauPath("" & RsMenuEboutique!libelle))
    ChampQuantite = EboutiqueEboutiqueGetDefault("ChampQuantite", "txt11", TableauPath("" & RsMenuEboutique!libelle))
    ChampQuantiteEncelade = EboutiqueEboutiqueGetDefault("ChampQuantiteEncelade", "TXT81", TableauPath("" & RsMenuEboutique!libelle))
    Prixderevient = EboutiqueEboutiqueGetDefault("Prixderevient", "txt47", TableauPath("" & RsMenuEboutique!libelle))
    RefCaddyPrixU = EboutiqueEboutiqueGetDefault("RefCaddyPrixU", "txt9", TableauPath("" & RsMenuEboutique!libelle))
    RefCaddy = EboutiqueEboutiqueGetDefault("RefCaddy", "txt3", TableauPath("" & RsMenuEboutique!libelle))
    RefFour = EboutiqueEboutiqueGetDefault("RefFour", "lst9", TableauPath("" & RsMenuEboutique!libelle))
    RefCaddyDesignation = EboutiqueEboutiqueGetDefault("RefCaddyDesignation", "mem1", TableauPath("" & RsMenuEboutique!libelle))
    Lib_Menu = "" & RsMenuEboutique!libelle

    Sql = "UPDATE NomenclaturFinal INNER JOIN (SELECT con_contacts.ContactID AS Id_Produit, con_contacts." & RefCaddy & " AS RefFour, " & RefFour & ".CatName AS Fournisseur,trim('' & con_contacts." & ChampReff & ") As Ref, con_contacts." & Prixderevient & " as [Prix Revient],  "
    Sql = Sql & "con_contacts." & RefCaddyPrixU & " as [Prix de vente], con_contacts." & ChampQuantite & " AS [Qt disponible], con_contacts." & ChampQuantiteEncelade & " AS [Quantite Encelade],   "
    Sql = Sql & "(Val('' & [" & ChampQuantite & "])+Val('' & [" & ChampQuantiteEncelade & "])) AS QtsTotal "
    Sql = Sql & "FROM con_contacts INNER JOIN " & RefFour & " ON con_contacts." & RefFour & " = " & RefFour & ".CatID IN '"
    Sql = Sql & TableauPath("" & RsMenuEboutique!libelle)
    
    Sql = Sql & "') AS MyFrom ON NomenclaturFinal.Ref = MyFrom.Ref SET   NomenclaturFinal.RefFour = [MyFrom].[RefFour], NomenclaturFinal.Fournisseur = [MyFrom].[Fournisseur],NomenclaturFinal.Lib_Menu= '" & Lib_Menu & "', NomenclaturFinal.Id_Produit = [MyFrom].[Id_Produit], NomenclaturFinal.Prix_Revient = Val('' & [Prix Revient])* "
    Sql = Sql & "Val('' & [Qts]), NomenclaturFinal.Prix_Vente = Val('' & [Prix de vente])*Val('' & [Qts]),  "
    Sql = Sql & "NomenclaturFinal.Qts_Appro = [QtsTotal]-[qts], NomenclaturFinal.Qts_Encelade = [Quantite Encelade], "
    Sql = Sql & " NomenclaturFinal.Qts_E_Boutique = [Qt disponible] "
    Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & Id_IndiceProjet & ";"

    Con.Execute Sql
    RsMenuEboutique.MoveNext
Wend
Sql = "UPDATE NomenclaturFinal SET NomenclaturFinal.Fournisseur = Null "
Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & Id_IndiceProjet & " "
Sql = Sql & "AND NomenclaturFinal.Fournisseur='(Sélectionner)';"
Con.Execute Sql

Sql = "UPDATE NomenclaturFinal SET NomenclaturFinal.Qts_Appro = 0 "
Sql = Sql & "WHERE NomenclaturFinal.Qts_Appro>0  "
Sql = Sql & "AND NomenclaturFinal.Id_IndiceProjet=" & Id_IndiceProjet & ";"
Con.Execute Sql
Sql = "UPDATE NomenclaturFinal SET NomenclaturFinal.Qts_Appro = Abs([Qts_Appro]) "
Sql = Sql & "WHERE NomenclaturFinal.Qts_Appro<0  "
Sql = Sql & "AND NomenclaturFinal.Id_IndiceProjet=" & Id_IndiceProjet & ";"

Con.Execute Sql
Sql = "UPDATE NomenclaturFinal SET NomenclaturFinal.Qts_Appro = [qts] "
Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & Id_IndiceProjet & "  "
Sql = Sql & "AND NomenclaturFinal.Qts_Encelade=0  "
Sql = Sql & "AND NomenclaturFinal.Qts_E_Boutique=0;"
Con.Execute Sql
End Sub
Sub Generer_Nomenclatuer2(Id_IndiceProjet As Long)
Dim Sql As String
Dim Rs As Recordset
Dim RecordCount As Long
Dim I As Long
Dim Client As String
Dim clsIso As New PreparISO
Dim Bouchon1
Dim Bouchon2
Dim BouchonFour1
Dim BouchonFour2
Dim NbLigne As Long
'Dim AppColection As New Collection
'Dim Connecteur As ClsNomanclatureGenerer
'LoadDb
Set TableauPath = funPath

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Client = Trim("" & Rs!Client)
Else
    Client = "RENAULT"
End If
Sql = "Drop table  Temp_ISO_" & NmJob & ";"
Con.Execute Sql
clsIso.IsTableCrate = False

Sql = "SELECT Ligne_Tableau_fils.LIAI, Val('' & [SECT]) AS [SECTion], Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.Id_IndiceProjet "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "GROUP BY Ligne_Tableau_fils.LIAI, Val('' & [SECT]),  "
Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.Id_IndiceProjet,  "
Sql = Sql & "Ligne_Tableau_fils.ACTIVER "
Sql = Sql & "HAVING Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & "  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True;"

Set Rs = Con.OpenRecordSet(Sql)
NbLigne = 0
While Rs.EOF = False
    NbLigne = NbLigne + 1
    Rs.MoveNext
Wend
Rs.Requery
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Prépare Nomenclature 4:"
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
     IncrmentServer FormBarGrah, ""
    clsIso.DefinIso Rs
    Rs.MoveNext
Wend

Sql = "DELETE Nomenclature2.* FROM Nomenclature2 "
Sql = Sql & "WHERE Nomenclature2.Id_IndiceProjet=" & Id_IndiceProjet & ";"
Con.Execute Sql

Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Ref, RefFour, App, Options ) "
Sql = Sql & "SELECT Connecteurs.Id_IndiceProjet, 'connecteurs' AS Designation, Connecteurs.CONNECTEUR,  "
Sql = Sql & "Connecteurs.RefConnecteurFour, Connecteurs.CODE_APP, Connecteurs.OPTION "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "GROUP BY Connecteurs.Id_IndiceProjet, 'connecteurs', Connecteurs.CONNECTEUR,  "
Sql = Sql & "Connecteurs.RefConnecteurFour, Connecteurs.CODE_APP, Connecteurs.OPTION, Connecteurs.ACTIVER "
Sql = Sql & "HAVING Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & "  "
Sql = Sql & "AND Connecteurs.ACTIVER=True "
Sql = Sql & "ORDER BY Connecteurs.CODE_APP;"

Con.Execute Sql

Sql = "UPDATE Nomenclature2 INNER JOIN "
Sql = Sql & "(SELECT con_contacts.txt3, lst9.CatName "
Sql = Sql & "FROM con_contacts INNER JOIN lst9 ON con_contacts.lst9 = lst9.CatID IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyFrom ON Nomenclature2.RefFour = MyFrom.txt3 SET Nomenclature2.Fournisseur = [CatName] "
Sql = Sql & "WHERE (Nomenclature2.Id_IndiceProjet=" & Id_IndiceProjet & ";"
Con.Execute Sql

Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Ref, RefFour, Options, App ) "
Sql = Sql & "SELECT Connecteurs.Id_IndiceProjet, 'Capot' AS Designation, Connecteurs.ReFCapot,  "
Sql = Sql & "Connecteurs.ReFCapotFour, Connecteurs.OPTION, Connecteurs.CODE_APP "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "GROUP BY Connecteurs.Id_IndiceProjet, 'Capot', Connecteurs.ReFCapot, Connecteurs.ReFCapotFour,  "
Sql = Sql & "Connecteurs.OPTION, Connecteurs.CODE_APP, Connecteurs.ACTIVER "
Sql = Sql & "HAVING Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & "  "
Sql = Sql & "AND Connecteurs.ReFCapot<>''  "
Sql = Sql & "AND Connecteurs.ACTIVER=True "
Sql = Sql & "ORDER BY Connecteurs.CODE_APP;"
Con.Execute Sql
Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Ref, RefFour, Options, App ) "
Sql = Sql & "SELECT Connecteurs.Id_IndiceProjet, 'Bouchon' AS Designation, Connecteurs.RefBouchon,  "
Sql = Sql & "Connecteurs.RefBouchonFour, Connecteurs.OPTION, Connecteurs.CODE_APP "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "GROUP BY Connecteurs.Id_IndiceProjet, 'Bouchon', Connecteurs.RefBouchon, Connecteurs.RefBouchonFour,  "
Sql = Sql & "Connecteurs.OPTION, Connecteurs.CODE_APP, Connecteurs.ACTIVER "
Sql = Sql & "HAVING Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & " "
Sql = Sql & "AND Connecteurs.RefBouchon<>'' "
Sql = Sql & "AND Connecteurs.ACTIVER=True "
Sql = Sql & "ORDER BY Connecteurs.CODE_APP;"
Con.Execute Sql

Sql = "SELECT Nomenclature2.Designation, Nomenclature2.Ref, Nomenclature2.RefFour, Nomenclature2.Qts "
Sql = Sql & "FROM Nomenclature2 "
Sql = Sql & "WHERE Nomenclature2.Designation='Bouchon'  "
Sql = Sql & "AND Nomenclature2.Id_IndiceProjet=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
NbLigne = 0
While Rs.EOF = False
    NbLigne = NbLigne + 1
    Rs.MoveNext
Wend
Rs.Requery
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 
 FormBarGrah.ProgressBar1Caption.Caption = " Prépare Nomenclature 5:"
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
IncrmentServer FormBarGrah, ""
Bouchon1 = Split("" & Rs!Ref & "(", "(")
Bouchon2 = Split("" & Bouchon1(1) & ")", ")")
BouchonFour1 = Split("" & Rs!RefFour & "(", "(")
BouchonFour2 = Split("" & BouchonFour1(1) & ")", ")")
Rs!Ref = Bouchon1(0)
Rs!RefFour = BouchonFour1(0)
Rs!Qts = Val("" & Bouchon2(0))
Rs.Update
    Rs.MoveNext
Wend

Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Ref, RefFour, Options, App )  "
Sql = Sql & "SELECT Connecteurs.Id_IndiceProjet, 'Verrou' AS Designation, Connecteurs.RefVerrou,   "
Sql = Sql & "Connecteurs.RefVerrouFour, Connecteurs.OPTION, Connecteurs.CODE_APP  "
Sql = Sql & "FROM Connecteurs  "
Sql = Sql & "GROUP BY Connecteurs.Id_IndiceProjet, 'Verrou', Connecteurs.RefVerrou, Connecteurs.RefVerrouFour,   "
Sql = Sql & "Connecteurs.OPTION, Connecteurs.CODE_APP, Connecteurs.ACTIVER  "
Sql = Sql & "HAVING Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & "   "
Sql = Sql & "AND Connecteurs.RefVerrou<>''   "
Sql = Sql & "AND Connecteurs.ACTIVER=True "
Sql = Sql & "ORDER BY Connecteurs.CODE_APP;"

Con.Execute Sql
'
'Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Longueur, [Longueur Total], App, LIAI, ISO, SECT,  "
'Sql = Sql & "Options, TEINT, TEINT2 ) "
'Sql = Sql & "SELECT Ligne_Tableau_fils.Id_IndiceProjet, 'Fils' AS Designation,  "
'Sql = Sql & "Ligne_Tableau_fils.LONG, Val('' & [LONG CP])+Val('' & [Long_Add])+Val('' & [Long_Add2]) AS Expr4,  "
'Sql = Sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.ISO,  "
'Sql = Sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.TEINT,  "
'Sql = Sql & "Ligne_Tableau_fils.TEINT2 "
'Sql = Sql & "FROM Ligne_Tableau_fils "
'Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, 'Fils', Ligne_Tableau_fils.LONG,  "
'Sql = Sql & "Val('' & [LONG CP])+Val('' & [Long_Add])+Val('' & [Long_Add2]), Ligne_Tableau_fils.APP,  "
'Sql = Sql & "Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.SECT,  "
'Sql = Sql & "Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
'Sql = Sql & "Ligne_Tableau_fils.ACTIVER "
'Sql = Sql & "HAVING Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & "  "
'Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True "
'Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"

'Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Longueur, [Longueur Total], App,  "
'Sql = Sql & "LIAI, ISO, SECT, Options, TEINT, TEINT2, Ref, RefFour ) "
'Sql = Sql & "SELECT Ligne_Tableau_fils.Id_IndiceProjet, 'Fils' AS Designation, Ligne_Tableau_fils.LONG,  "
'Sql = Sql & "Val('' & [LONG CP])+Val('' & [Long_Add])+Val('' & [Long_Add2]) AS Expr4, Ligne_Tableau_fils.APP,  "
'Sql = Sql & "Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.OPTION,  "
'Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Temp_ISO_" & NmJob & ".Ref, Temp_ISO_" & NmJob & ".RefFour "
'Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Temp_ISO_" & NmJob & " ON Ligne_Tableau_fils.LIAI = Temp_ISO_" & NmJob & ".LIAI "
'Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, 'Fils', Ligne_Tableau_fils.LONG, Val('' & [LONG CP])+ "
'Sql = Sql & "Val('' & [Long_Add])+Val('' & [Long_Add2]), Ligne_Tableau_fils.APP, Ligne_Tableau_fils.LIAI,  "
'Sql = Sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.TEINT,  "
'Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ACTIVER, Temp_ISO_" & NmJob & ".Ref, Temp_ISO_" & NmJob & ".RefFour "
'Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & "   "
'Sql = Sql & "And Ligne_Tableau_fils.Activer = True "
'Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"


Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Longueur, [Longueur Total], App, LIAI, ISO, SECT, Options, TEINT, TEINT2,   "
Sql = Sql & "Ref, RefFour, Qts )  "
Sql = Sql & "SELECT Ligne_Tableau_fils.Id_IndiceProjet, 'Fils' AS Designation, Ligne_Tableau_fils.LONG,   "
Sql = Sql & "Val('' & [LONG CP])+Val('' & [Long_Add])+Val('' & [Long_Add2]) AS Expr4, Ligne_Tableau_fils.APP,   "
Sql = Sql & "Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.OPTION,   "
Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Temp_ISO_" & NmJob & ".Ref, Temp_ISO_" & NmJob & ".RefFour,   "
Sql = Sql & "Val('' & [LONG CP])+Val('' & [Long_Add])+Val('' & [Long_Add2]) AS Expr1  "
Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Temp_ISO_" & NmJob & " ON Ligne_Tableau_fils.LIAI = Temp_ISO_" & NmJob & ".LIAI  "
Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, 'Fils', Ligne_Tableau_fils.LONG, Val('' & [LONG CP])+  "
Sql = Sql & "Val('' & [Long_Add])+Val('' & [Long_Add2]), Ligne_Tableau_fils.APP, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.ISO,   "
Sql = Sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,   "
Sql = Sql & "Temp_ISO_" & NmJob & ".Ref, Temp_ISO_" & NmJob & ".RefFour, Ligne_Tableau_fils.ACTIVER, Val('' & [LONG CP])+Val('' & [Long_Add])+  "
Sql = Sql & "Val('' & [Long_Add2])  "
Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & "   "
Sql = Sql & "And Ligne_Tableau_fils.Activer = True "
Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"

'" & NmJob & "
Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Longueur, [Longueur Total], App, Voie,  "
Sql = Sql & "App2, Voie2, LIAI, ISO, SECT, Options, TEINT, TEINT2, Ref, RefFour, Qts ) "
Sql = Sql & "SELECT Ligne_Tableau_fils.Id_IndiceProjet, 'Fils' AS Designation, Ligne_Tableau_fils.LONG, Val('' & [LONG CP])+ "
Sql = Sql & "Val('' & [Long_Add])+Val('' & [Long_Add2]) AS Expr4, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2,  "
Sql = Sql & "Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.LIAI, First(Ligne_Tableau_fils.ISO) AS PremierDeISO, Ligne_Tableau_fils.SECT,  "
Sql = Sql & "Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, First(Temp_ISO_" & NmJob & ".Ref) AS PremierDeRef,  "
Sql = Sql & "First(Temp_ISO_" & NmJob & ".RefFour) AS PremierDeRefFour, Val('' & [LONG CP])+Val('' & [Long_Add])+Val('' & [Long_Add2]) AS Expr1 "
Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Temp_ISO_" & NmJob & " ON Ligne_Tableau_fils.LIAI = Temp_ISO_" & NmJob & ".LIAI "
Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, 'Fils', Ligne_Tableau_fils.LONG, Val('' & [LONG CP])+Val('' & [Long_Add])+ "
Sql = Sql & "Val('' & [Long_Add2]), Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
Sql = Sql & "Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils.TEINT2, Val('' & [LONG CP])+Val('' & [Long_Add])+Val('' & [Long_Add2]), Ligne_Tableau_fils.ACTIVER "
Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & "  "
Sql = Sql & "And Ligne_Tableau_fils.Activer = True "
Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"

Con.Execute Sql
Sql = "Drop table  Temp_ISO_" & NmJob & ";"
Con.Execute Sql

Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, App, Ref, RefFour, Options, LIAI ) "
Sql = Sql & "SELECT Ligne_Tableau_fils.Id_IndiceProjet, 'Clips' AS Designation, Ligne_Tableau_fils.APP,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four], Ligne_Tableau_fils.OPTION,  "
Sql = Sql & "Ligne_Tableau_fils.LIAI "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, 'Clips', Ligne_Tableau_fils.APP,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four], Ligne_Tableau_fils.OPTION,  "
Sql = Sql & "Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.ACTIVER "
Sql = Sql & "HAVING Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & "   "
Sql = Sql & "AND Ligne_Tableau_fils.[Ref Clip]<>''  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True "
Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"

Con.Execute Sql

Sql = "SELECT Nomenclature2.* FROM Nomenclature2 "
Sql = Sql & "WHERE Nomenclature2.Id_IndiceProjet=" & Id_IndiceProjet & " "
Sql = Sql & "AND Nomenclature2.Designation='Fils';"
Set Rs = Con.OpenRecordSet(Sql)

Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Ref, RefFour, App, Options, Designation, LIAI ) "
Sql = Sql & "SELECT Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.[Ref Clip2],  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip Four2], Ligne_Tableau_fils.APP, Ligne_Tableau_fils.OPTION,  "
Sql = Sql & "'Clips' AS Designation, Ligne_Tableau_fils.LIAI "
Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN Nomenclature2 ON (Ligne_Tableau_fils.LIAI = Nomenclature2.LIAI)  "
Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Nomenclature2.Id_IndiceProjet)  "
Sql = Sql & "AND (Ligne_Tableau_fils.[Ref Clip2] = Nomenclature2.Ref) AND (Ligne_Tableau_fils.APP = Nomenclature2.App) "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & "  "
Sql = Sql & "AND Ligne_Tableau_fils.[Ref Clip2]<>''  "
Sql = Sql & "AND Nomenclature2.Designation Is Null  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True "
Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"


Con.Execute Sql



Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, App, Ref, RefFour, LIAI, Options ) "
Sql = Sql & "SELECT Ligne_Tableau_fils.Id_IndiceProjet, 'Joint' AS Designation, Ligne_Tableau_fils.APP,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint four], Ligne_Tableau_fils.LIAI, "
Sql = Sql & " Ligne_Tableau_fils.OPTION "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, 'Joint', Ligne_Tableau_fils.APP,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint four], Ligne_Tableau_fils.LIAI,  "
Sql = Sql & "Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.OPTION "
Sql = Sql & "HAVING  Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & "   "
Sql = Sql & "AND Ligne_Tableau_fils.[Ref Joint]<>''  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True;"

Con.Execute Sql
Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Ref, RefFour, App, Options, Designation, LIAI ) "
Sql = Sql & "SELECT Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.[Ref Joint2],  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint Four2], Ligne_Tableau_fils.APP, Ligne_Tableau_fils.OPTION,  "
Sql = Sql & "'Joint' AS Designation, Ligne_Tableau_fils.LIAI "
Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN Nomenclature2 ON (Ligne_Tableau_fils.[Ref Joint2] =  "
Sql = Sql & "Nomenclature2.Ref) AND (Ligne_Tableau_fils.APP = Nomenclature2.App)  "
Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Nomenclature2.Id_IndiceProjet)  "
Sql = Sql & "AND Ligne_Tableau_fils.LIAI = Nomenclature2.LIAI "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & "  "
Sql = Sql & "AND Ligne_Tableau_fils.[Ref Joint2]<>''  "
Sql = Sql & "AND Nomenclature2.Designation Is Null  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True "
Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"
Con.Execute Sql

Sql = "INSERT INTO Nomenclature2 ( Id_IndiceProjet, Designation, Ref ) "
Sql = Sql & "SELECT Composants.Id_IndiceProjet, Composants.Path AS Designation, Composants.REFCOMP "
Sql = Sql & "FROM Composants "
Sql = Sql & "GROUP BY Composants.Id_IndiceProjet, Composants.Path, Composants.REFCOMP, Composants.ACTIVER "
Sql = Sql & "HAVING Composants.Id_IndiceProjet=" & Id_IndiceProjet & " "
Sql = Sql & "AND Composants.Path<>''  "
Sql = Sql & "AND Composants.ACTIVER=True "
Sql = Sql & "ORDER BY Composants.Path, Composants.REFCOMP; "

Con.Execute Sql

Set clsIso = Nothing
End Sub
Sub Generer_Nomenclatuer(Id_IndiceProjet As Long)
Dim Sql As String
Dim Rs As Recordset
Dim RecordCount As Long
Dim I As Long
Dim Client As String
Dim AppColection As New Collection
Dim Connecteur As ClsNomanclatureGenerer
Dim SplitConnecteur
Dim SpliPath
'LoadDb
On Error Resume Next
Dim NbLigne As Long
Set TableauPath = funPath

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Client = Trim("" & Rs!Client)
Else
    Client = "RENAULT"
End If
Sql = "SELECT Connecteurs.*  FROM Connecteurs  "
Sql = Sql & "Where Connecteurs.Id_IndiceProjet = " & Id_IndiceProjet & " "
'sql = sql & "and (Connecteurs.CODE_APP='98644-2001') "
'sql = sql & "or Connecteurs.CODE_APP='1337.AH' "
'sql = sql & "or Connecteurs.CODE_APP='20-12A' "
'sql = sql & "or Connecteurs.CODE_APP='147.AA' "
'sql = sql & "or Connecteurs.CODE_APP='24-12A' "
'sql = sql & "or Connecteurs.CODE_APP='887.AA' "
'sql = sql & "or Connecteurs.CODE_APP='242.AA' "
'sql = sql & "or Connecteurs.CODE_APP='34-12A' "
'sql = sql & "or Connecteurs.CODE_APP='ENT-1' "
''sql = sql & "or Connecteurs.CODE_APP='NF-1A' "
'
'
''sql = sql & "or Connecteurs.CODE_APP='1337.AH' "
'sql = sql & ") "
Sql = Sql & "ORDER BY Connecteurs.CODE_APP;"
Set AppColection = Nothing
Set AppColection = New Collection
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    NbLigne = NbLigne + 1
    Rs.MoveNext
Wend
Rs.Requery

 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Prépare Nomenclature :"
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
     IncrmentServer FormBarGrah, ""
        Set Connecteur = New ClsNomanclatureGenerer
        Connecteur.App = "" & Rs!Code_APP
        
        SplitConnecteur = "" & Rs!Connecteur & "$CN"
        SplitConnecteur = Split(SplitConnecteur, "$CN")
        Connecteur.FourConnecteur = "" & Rs!RefConnecteurFour
        Connecteur
        Connecteur.DESIGNATION = "" & Rs!DESIGNATION
        Connecteur.IntiConnecteur "" & SplitConnecteur(0), "" & Rs!refVerrou, "" & Rs!RefVerrouFour, "" & Rs!RefCapot
        Connecteur.InitConnecteur "" & Rs!RefBouchon, "" & Rs!RefBouchonFour, "" & Rs!RefCapot, "" & Rs!ReFCapotFour, "" & Rs!refVerrou, "" & Rs!RefVerrouFour
        AppColection.Add Connecteur, "" & Rs!Code_APP
        
        Set Connecteur = Nothing
    Rs.MoveNext
Wend
Sql = "SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils "
Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & " "
'sql = sql & "AND (Ligne_Tableau_fils.APP='120.AA'  "
'sql = sql & "or  Ligne_Tableau_fils.APP2='120.AA')  "
Sql = Sql & "ORDER BY Ligne_Tableau_fils.LIAI;"


'sql = "SELECT Ligne_Tableau_fils.* "
'sql = sql & "FROM Ligne_Tableau_fils "
'sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=808  "
'sql = sql & "AND (Ligne_Tableau_fils.APP='120.AA'  "
'sql = sql & "or  Ligne_Tableau_fils.APP2='120.AA)  "
'sql = sql & "or Ligne_Tableau_fils.APP='20-12A' "
'sql = sql & "or Ligne_Tableau_fils.APP2='20-12A') "
'sql = sql & "ORDER BY Ligne_Tableau_fils.LIAI;"



Set Rs = Con.OpenRecordSet(Sql)
NbLigne = 0
While Rs.EOF = False
    NbLigne = NbLigne + 1
    Rs.MoveNext
Wend
Rs.Requery

 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Prépare Nomenclature 2 :"
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
     IncrmentServer FormBarGrah, ""
     ComteurPass = ComteurPass + 1
On Error GoTo MoveNext
'If (Trim("" & Rs!App) = "120.AA" And Trim("" & Rs!VOI) = "G2") Or (Trim("" & Rs!App2) = "120.AA" And Trim("" & Rs!VOI2) = "G2") Then
'MsgBox ""
'End If
        AppColection("" & Rs!App).AjouterCritaire "" & Rs!Option
'        AppColection("" & Rs!App).renseigneVoies "" & Rs![Ref Connecteur], "" & Rs!App, "" & Rs!Option
        AppColection("" & Rs!App).AjouterLstVoie "" & Rs!VOI, "" & "" & Rs!Option
        AppColection("" & Rs!App).initLiaison "" & Rs!VOI, "" & Rs!Option, "" & Rs!Liai
        AppColection("" & Rs!App).IniJoint Rs
        
        
        'L As Double, C As Double, L_CP As Double, L_add As Double, S As Double, ISO As String, Couleur As String, Liseret As String, Voie As String, Critaire As String)
        AppColection("" & Rs!App).RenseigneLongeur Val(Replace("" & Rs!LONG, ",", ".")), Val(Replace("" & Rs!Coupe, ",", ".")), _
        Val(Replace("" & Rs![LONG CP], ",", ".")), Val(Replace("" & Rs!Long_Add, ",", ".")), Val(Replace("" & Rs!SECT, ",", ".")), _
         "" & Rs!ISO, "" & Rs!TEINT, "" & Rs!TEINT2, "" & Rs!VOI, "" & Rs!Option
        
        AppColection("" & Rs!App2.Value).AjouterCritaire "" & Rs!Option
'        AppColection("" & Rs!App2).renseigneVoies "" & Rs![Ref Connecteur2], "" & Rs!App2, "" & Rs!Option
        AppColection("" & Rs!App2).AjouterLstVoie "" & Rs!VOI2, "" & "" & Rs!Option
        AppColection("" & Rs!App2).initLiaison "" & Rs!VOI2, "" & Rs!Option, "" & Rs!Liai
        AppColection("" & Rs!App2).IniJoint Rs
        
        
        AppColection("" & Rs!App2).RenseigneLongeur Val(Replace("" & Rs!LONG, ",", ".")), Val(Replace("" & Rs!Coupe, ",", ".")), _
        Val(Replace("" & Rs![LONG CP], ",", ".")), Val(Replace("" & Rs!Long_Add2, ",", ".")), Val(Replace("" & Rs!SECT, ",", ".")), _
         "" & Rs!ISO, "" & Rs!TEINT, "" & Rs!TEINT2, "" & Rs!VOI2, "" & Rs!Option
        
'        AppColection("" & rs!App2).RenseigneLongeur Val(Replace("" & rs!Long_Add2, ",", ".")), "" & rs!VOI2, "" & rs!Option
    
        
'        AppColection ("" & rs!Code_APP)
    Rs.MoveNext
Wend
Dim RM As New ReyRecordsetMaker
Dim rr(24, 4) As String

 initChampRecordset rr
'AddField "Nom", FT_VarChar, 30
'    AddField "Naiss", FT_VarChar, 10
Dim Rs2 As Recordset

'Rs2.CursorType = adOpenDynamic
'Rs2.Fields.Append "toto", adInteger
'Rs2.AddNew "toto", 10

'RM.CreatFilds rr
Set Rs = RM.Rs
'Set RM.Rs = Rs

'Set rs = RM.Recordset
'Set Rs = AppColection(1).RetourneRecordset(RM, Rs, rr, Id_IndiceProjet)
Dim a As Long
'Rs.MoveFirst
'    While Not Rs.EOF
'        For a = 0 To Rs.Fields.Count - 1
'            MsgBox Rs.Fields(a).Name & " = " & Rs.Fields(a).Value, vbOKOnly, Rs.AbsolutePosition
'        Next a
'        Rs.MoveNext
'    Wend
'a = AppColection("120.AC")
For I = 1 To AppColection.Count

    AppColection(I).RetourneRecordset RM, rr
'    Set RM.rs = RM.Recordset
Next
'Set RM.Rs = RM.Recordset
Sql = "DELETE NomeclatureConnecteurs.* FROM NomeclatureConnecteurs "
Sql = Sql & "WHERE NomeclatureConnecteurs.Id_IndiceProjet=" & Id_IndiceProjet & ";"
Con.Execute Sql

Sql = "SELECT NomeclatureConnecteurs.* FROM NomeclatureConnecteurs "
Sql = Sql & "WHERE NomeclatureConnecteurs.Id_IndiceProjet=" & Id_IndiceProjet & " ;"
Set Rs2 = Con.OpenRecordSet(Sql)

'
'Set Rs = RM.Recordset
'Rs.MoveFirst
On Error GoTo Fin
NbLigne = AppColection.Count
'While Rs.EOF = False
'    NbLigne = NbLigne + 1
'    Rs.MoveNext
'Wend
'Rs.MoveFirst

 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Prépare Nomenclature 3:"
 Dim J As Long
 Dim L As Long
 For I = 1 To AppColection.Count
   IncremanteBarGrah FormBarGrah
IncrmentServer FormBarGrah, ""
    For L = 1 To AppColection(I).NewRsUpdate.Count
        Rs2.AddNew
         Rs2(1) = Id_IndiceProjet
        For J = 0 To 24
        
          Rs2(Replace(Replace(AppColection(I).NewRsUpdate(L).RetournName(J), "[", ""), "]", "")).Value = AppColection(I).NewRsUpdate(L).RetournValue(J)
          Debug.Print AppColection(I).NewRsUpdate(L).RetournName(J)
          Debug.Print AppColection(I).NewRsUpdate(L).RetournValue(J)
'           If J = 23 Then
'         MsgBox ""
'         End If
        Next
         Rs2.Update
    Next
 Next
'While Rs.EOF = False
'     IncremanteBarGrah FormBarGrah
'IncrmentServer FormBarGrah, ""
''        If Trim("" & Rs.Fields("Liaison").Value) <> "" Then
'            Rs2.AddNew
'            For a = 0 To Rs.Fields.Count - 1
''            If "" & Rs.Fields(a).Value = "120.AC" Then
''                MsgBox ""
''            End If
'                Debug.Print Rs2(Rs.Fields(a).Name).Name & " : " & Rs.Fields(a).Name & " : " & Rs.Fields(a).value
'                Rs2(1) = Id_IndiceProjet
'
'                Rs2(Rs.Fields(a).Name).value = Rs.Fields(a).value
'    '            MsgBox rs.Fields(A).Name & " = " & rs.Fields(A).Value, vbOKOnly, rs.AbsolutePosition
'            Next a
'            On Error GoTo 0
'            Rs2.Update
''         Else
''         MsgBox ""
''        End If
'        Rs.MoveNext
'    Wend
'    MsgBox ""
    GoTo Fin
MoveNext:
Resume Next
Fin:
On Error GoTo 0
End Sub
Function initChampRecordset(rr)
Dim I As Long
I = 0

rr(I, 0) = "App"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Designation"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Connecteur"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Connecteur_Four"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Liaison"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Voie"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Long_Add"
rr(I, 1) = FT_Decimal
rr(I, 2) = 10.2
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "SECT"
rr(I, 1) = FT_Decimal
rr(I, 2) = 10.2
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "LONG"
rr(I, 1) = FT_Decimal
rr(I, 2) = 10.2
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "COUPE"
rr(I, 1) = FT_Decimal
rr(I, 2) = 10.2
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "[LONG CP]"
rr(I, 1) = FT_Decimal
rr(I, 2) = 10.2
rr(I, 3) = 0


I = I + 1
rr(I, 0) = "TEINT"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "TEINT2"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0


I = I + 1
rr(I, 0) = "Famille"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Bouchon"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
I = I + 1
rr(I, 0) = "BouchonFour"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
I = I + 1
rr(I, 0) = "Capot"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Capot_Four"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0


I = I + 1
rr(I, 0) = "Verrou"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Verrout_Four"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Options"
rr(I, 1) = FT_VarChar
rr(I, 2) = 25
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Clip"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "ClipFour"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "Joint"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0

I = I + 1
rr(I, 0) = "JointFour"
rr(I, 1) = FT_VarChar
rr(I, 2) = 255
rr(I, 3) = 0
'I = I + 1
'rr(I, 0) = "ID"
'rr(I, 1) = FT_Integer
'rr(I, 2) = 0
'rr(I, 3) = 1
 initChampRecordset = rr
End Function
Sub PreparationNomenclatuer(Id_IndiceProjet As Long)
Dim Sql As String
Dim Rs As Recordset
Dim RecordCount As Long
Dim I As Long
Dim Client As String
'LoadDb
Set TableauPath = funPath

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Client = Trim("" & Rs!Client)
Else
    Client = "RENAULT"
End If

Set MyColectionClsNomenclature = New Collection

Sql = "SELECT Connecteurs.* FROM Connecteurs "
Sql = Sql & "Where Connecteurs.Id_IndiceProjet = " & Id_IndiceProjet & " and Connecteurs.ACTIVER=True "
Sql = Sql & "ORDER BY Connecteurs.CODE_APP;"

Set Rs = Con.OpenRecordSet(Sql)
BarrGraphCoun = 0
While Rs.EOF = False
    BarrGraphCoun = BarrGraphCoun + 1
    Rs.MoveNext
Wend
Rs.Requery

FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Max = BarrGraphCoun + 1
FormBarGrah.ProgressBar1Caption = "Initialisation " & NuInit
 NuInit = CStr(Val("" & NuInit) + 1)
While Rs.EOF = False
    FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
    addClsNomanclature "" & Rs![Connecteur], "" & Rs!Code_APP, Client
   MyColectionClsNomenclature("" & Rs!Code_APP).IsEpisure = Rs("O/N")
    Rs.MoveNext
Wend

Sql = "SELECT Ligne_Tableau_fils.* "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & "  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True "
Sql = Sql & "ORDER BY Ligne_Tableau_fils.[Ref Connecteur];"
Set Rs = Con.OpenRecordSet(Sql)
BarrGraphCoun = 0
While Rs.EOF = False
    BarrGraphCoun = BarrGraphCoun + 1
    Rs.MoveNext
Wend
Rs.Requery
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Max = BarrGraphCoun + 1
FormBarGrah.ProgressBar1Caption = "Initialisation " & NuInit
NuInit = CStr(Val("" & NuInit) + 1)
On Error Resume Next
While Rs.EOF = False
    FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    IncrmentServer FormBarGrah, ""
    DoEvents
'    addClsNomanclature "" & rs![Ref Connecteur], "" & rs!App, Client
'    addClsNomanclature "" & rs![Ref Connecteur2], "" & rs!App2, Client
 MyColectionClsNomenclature("" & Rs!App).IsEpisure = MyColectionClsNomenclature("" & Rs!App).IsEpisure
 
 If MyColectionClsNomenclature("" & Rs!App).IsEpisure = False Then
   
    
    MyColectionClsNomenclature("" & Rs!App).SubSection "" & Rs!VOI, Val(Replace("" & Rs!SECT, ",", "."))
 Else
    MyColectionClsNomenclature("" & Rs!App).SubSection "D1", Val(Replace("" & Rs!SECT, ",", "."))
 End If
 If MyColectionClsNomenclature("" & Rs!App2).IsEpisure = False Then
    MyColectionClsNomenclature("" & Rs!App2).SubSection "" & Rs!VOI2, Val(Replace("" & Rs!SECT, ",", "."))
 Else
    MyColectionClsNomenclature("" & Rs!App2).SubSection "D1", Val(Replace("" & Rs!SECT, ",", "."))
 End If
    MyColectionClsNomenclature("" & Rs!App).IniJoint Rs
    MyColectionClsNomenclature("" & Rs!App2).IniJoint Rs
    
Rs.MoveNext
Wend
'For I = 1 To MyColectionClsNomenclature.Count
'    MyColectionClsNomenclature(I).DelBouchon
'Next
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Max = MyColectionClsNomenclature.Count

For I = 1 To MyColectionClsNomenclature.Count
    FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    IncrmentServer FormBarGrah, ""
    DoEvents
    MyColectionClsNomenclature(I).initFilsDirection Id_IndiceProjet
    MyColectionClsNomenclature(I).RendeignePrix
    MyColectionClsNomenclature(I).InitCip
    MyColectionClsNomenclature(I).ChoixClip
Next
Rs.Requery
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Max = BarrGraphCoun + 1
FormBarGrah.ProgressBar1Caption = "Mise à jours"
While Rs.EOF = False
    FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    IncrmentServer FormBarGrah, ""
    DoEvents
    MyColectionClsNomenclature("" & Rs!App).MajTableauFils Rs
    MyColectionClsNomenclature("" & Rs!App2).MajTableauFils Rs
    MyColectionClsNomenclature("" & Rs!App).MajConnecteur Rs
    MyColectionClsNomenclature("" & Rs!App2).MajConnecteur Rs
    Rs.MoveNext
Wend
 

End Sub
Function addClsNomanclature(Connecteur As String, App As String, Client As String) As ClsNomanclature
On Error Resume Next
Dim Rs As Recordset
Dim Sql As String
Dim ChampCli As String
Dim ChampReff As String
Dim SplitConnecteur

Set addClsNomanclature = New ClsNomanclature
SplitConnecteur = Split(Connecteur & "§", "§")
addClsNomanclature.Connecteur = "" & SplitConnecteur(0)
ChampCli = GetDefault(Client, "txt1")
ChampReff = GetDefault("Fournisseur", "txt3")
addClsNomanclature.ChampCli = ChampCli
addClsNomanclature.ChampReff = ChampReff
Sql = "SELECT con_contacts." & ChampReff & " "
Sql = Sql & "FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "' "

Sql = Sql & "WHERE con_contacts." & ChampCli & "='" & SplitConnecteur(0) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    addClsNomanclature.FourConnecteur = "" & Rs(ChampReff)
End If
Set Rs = Con.CloseRecordSet(Rs)


addClsNomanclature.App = App
addClsNomanclature.renseigneVoies Connecteur, App
addClsNomanclature.RendeigneConnecteur Connecteur
MyColectionClsNomenclature.Add addClsNomanclature, App
 
Set addClsNomanclature = Nothing
End Function
