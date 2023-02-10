Attribute VB_Name = "NomenclatureMajStock"
Option Explicit

Public Sub MajStock(Id_Projet As Long, Id_Fils As Long, NbPice As Long, FormBarGrah As Object)
Dim Sql As String
Dim Rs As Recordset
Dim Equipement As String
Dim TableauEquipement
Dim TableauOption
Dim Index_Equipement As Long
Dim CloseLiK As String
Dim RsChampQuantite As Recordset
Dim QtsTotalEboutique As Long
Dim QtsTotalEncelade As Long
Dim QtsTotalTotal As Long
Dim Prix_evient As Double
Dim Prix_vente As Double
Dim RefCaddyDesignation As String
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Id_Projet & ";"
Set Rs = Con.OpenRecordSet(Sql)
Dim Client As String
If Rs.EOF = False Then
    Client = Trim("" & Rs!Client)
Else
    Client = "RENAULT"
End If '
Sql = "SELECT T_indiceProjet.Equipement FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_Fils & ";"
Set Rs = Con.OpenRecordSet(Sql)
Equipement = Trim("" & Rs!Equipement)
TableauEquipement = Split(Equipement & ";", ";")
CloseLiK = " ("
For Index_Equipement = LBound(TableauEquipement) To UBound(TableauEquipement)
    If Trim("" & TableauEquipement(Index_Equipement)) <> "" Then
        TableauOption = Split(TableauEquipement(Index_Equipement) & "_", "_")
        CloseLiK = CloseLiK & "NomenclaturFinal.Options Like  '%" & TableauOption(0) & "%' or "
    
    End If
Next

CloseLiK = CloseLiK & "NomenclaturFinal.Options Is Null or NomenclaturFinal.Options Like '%ALL%' "
CloseLiK = CloseLiK & "or NomenclaturFinal.Options Like '%TOUS%' "
CloseLiK = CloseLiK & ") AND NomenclaturFinal.Id_IndiceProjet=" & Id_Projet & " "
Sql = "SELECT  NomenclaturFinal.* FROM NomenclaturFinal WHERE "
Sql = Sql & CloseLiK & "AND  NomenclaturFinal.Lib_Menu Is Not Null;"
Set Rs = Con.OpenRecordSet(Sql)
Dim ChampReff As String
ChampReff = GetDefault(Client, "txt1")
Dim ChampQuantite As String
ChampQuantite = GetDefault("ChampQuantite", "txt11")
Dim ChampQuantiteEncelade As String
ChampQuantiteEncelade = GetDefault("ChampQuantiteEncelade", "TXT81")
Dim Prixderevient As String
Prixderevient = GetDefault("Prixderevient", "txt47")
Dim RefCaddyPrixU As String
RefCaddyPrixU = GetDefault("RefCaddyPrixU", "txt9")

Dim NbEnregistrement As Long
Dim PathBase As String
Dim numDevis As String
Dim Id_Caddie As Long
Dim TVA As String
Dim Remise As Double
Dim Qts As Long
Dim MyModulo As Long
Dim Id_User As Long
Dim Id_Socete As Long
Id_Socete = EboutiqueSocieteId
Id_User = EboutiqueUserId
Remise = Val(GetRemiseEboutique())
Remise = Remise * (1 / 100)
TVA = GetDefaultEboutique("TVA", "19.6")
Sql = "DELETE t_Autocable_caddie.* FROM t_Autocable_caddie "
Sql = Sql & "WHERE t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & ";"
Con.Execute Sql

NbEnregistrement = 0
While Rs.EOF = False
    NbEnregistrement = NbEnregistrement + 1
    Rs.MoveNext
Wend
Rs.Requery
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbEnregistrement
FormBarGrah.ProgressBar1Caption.Caption = " Décrémente le stock"
IncrmentServer FormBarGrah, ""
While Rs.EOF = False
    IncremanteBarGrah FormBarGrah
    IncrmentServer FormBarGrah, ""
    If Trim("" & Rs!Lib_Menu) <> "" Then
    
    ChampReff = EboutiqueEboutiqueGetDefault("ChampReff", "txt1", TableauPath("" & Rs!Lib_Menu))
    ChampQuantite = EboutiqueEboutiqueGetDefault("ChampQuantite", "txt11", TableauPath("" & Rs!Lib_Menu))
    ChampQuantiteEncelade = EboutiqueEboutiqueGetDefault("ChampQuantiteEncelade", "TXT81", TableauPath("" & Rs!Lib_Menu))
    Prixderevient = EboutiqueEboutiqueGetDefault("Prixderevient", "txt47", TableauPath("" & Rs!Lib_Menu))
    RefCaddyPrixU = EboutiqueEboutiqueGetDefault("RefCaddyPrixU", "txt9", TableauPath("" & Rs!Lib_Menu))
    RefCaddyDesignation = EboutiqueEboutiqueGetDefault("RefCaddyDesignation", "mem1", TableauPath("" & Rs!Lib_Menu))


        Sql = "SELECT  con_contacts." & RefCaddyDesignation & ", con_contacts.ContactID AS Id_Produit,con_contacts." & ChampReff & " As Ref, con_contacts." & Prixderevient & " as [Prix Revient],  "
        Sql = Sql & "con_contacts." & RefCaddyPrixU & " as [Prix de vente], con_contacts." & ChampQuantite & " AS [Qt disponible], con_contacts." & ChampQuantiteEncelade & " AS [Quantite Encelade],   "
        Sql = Sql & "(Val('' & [" & ChampQuantite & "])+Val('' & [" & ChampQuantiteEncelade & "])) AS QtsTotal FROM con_contacts IN '"
        Sql = Sql & TableauPath("" & Rs!Lib_Menu)
        Sql = Sql & "' WHERE con_contacts.ContactID=" & Rs!id_produit

         Set RsChampQuantite = Con.OpenRecordSet(Sql)
    
        If RsChampQuantite.EOF = False Then
            QtsTotalEboutique = Val(Trim("" & RsChampQuantite![Qt disponible]))
            QtsTotalEncelade = Val(Trim("" & RsChampQuantite![Quantite Encelade]))
            QtsTotalTotal = Val(Trim("" & RsChampQuantite!QtsTotal))
            Prix_evient = Val(Trim("" & RsChampQuantite![Prix Revient]))
            Prix_vente = Val(Trim("" & RsChampQuantite![Prix de vente]))
    
       
            Sql = "INSERT INTO t_Autocable_caddie ( Id_IndiceProjet,"
            Sql = Sql & "Id_Menu,"
            Sql = Sql & "id_produit,"
            Sql = Sql & "RefProduit,"
            Sql = Sql & "TypePiece,"
            Sql = Sql & "qte_produit,"
            Sql = Sql & "PrixU_produit,"
            Sql = Sql & "Designation,"
            Sql = Sql & "TVA,"
            Sql = Sql & "PrixRevient,"
            Sql = Sql & "Remise,"
            Sql = Sql & "QtsEboutique,"
            Sql = Sql & "QtsEncelade,"
            Sql = Sql & "QtsComande,"
            Sql = Sql & "Id_User,"
            Sql = Sql & "id_r_social ) "
           
            Sql = Sql & "VALUES ( "
            Sql = Sql & Id_Projet & ", "
            Sql = Sql & RetournMenuId("" & Rs!Lib_Menu) & ", "
            Sql = Sql & Rs!id_produit & ","
            
            If Trim("" & Rs!refFour) = "" Then
                Sql = Sql & "NULL,"
            Else
                Sql = Sql & "'" & Rs!refFour & "', "
            End If
            Sql = Sql & "'" & Rs!Lib_Menu & "', "
'           Designation ModuloYes
            Qts = Rs!Qts * NbPice
            If UCase("" & Rs!DESIGNATION) = UCase("Fils") Then
                MyModulo = Qts Mod 1000
                Qts = Qts - MyModulo
                Qts = Qts * (1 / 1000)
                If MyModulo <> 0 Then Qts = Qts + 1
            End If
            Sql = Sql & Replace(Qts, ",", ".", 1) & ","
            Sql = Sql & Replace(Prix_vente, ",", ".", 1) & ",'"
            Sql = Sql & Replace(MyReplace("" & RsChampQuantite(RefCaddyDesignation)), vbCrLf, Chr(10)) & "',"
            Sql = Sql & TVA & ","
            Sql = Sql & Replace(Prix_evient, ",", ".", 1) & ","
            Sql = Sql & Replace(Remise, ",", ".") & ","
            Sql = Sql & QtsTotalEboutique & ","
            Sql = Sql & QtsTotalEncelade & ","
            Sql = Sql & "0,"
            Sql = Sql & Id_User & ","
            Sql = Sql & Id_Socete & ");"
            Con.Execute Sql
         
            Sql = "UPDATE t_Autocable_caddie SET t_Autocable_caddie.QtsEboutique = [QtsEboutique]-[qte_produit] "
            Sql = Sql & "WHERE t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & " "
            Sql = Sql & " AND t_Autocable_caddie.Flinguer=False;"
            Con.Execute Sql
        
            Sql = "UPDATE t_Autocable_caddie SET t_Autocable_caddie.QtsEboutique = 0, t_Autocable_caddie.QtsEncelade = [QtsEncelade]+[QtsEboutique] "
            Sql = Sql & "WHERE t_Autocable_caddie.QtsEboutique<0 "
            Sql = Sql & "AND t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & " "
            Sql = Sql & " AND t_Autocable_caddie.Flinguer=False;"
            Con.Execute Sql
            
            Sql = "UPDATE t_Autocable_caddie SET t_Autocable_caddie.QtsComande = [QtsComande]-[QtsEncelade], t_Autocable_caddie.QtsEncelade = 0 "
            Sql = Sql & "WHERE t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & " "
            Sql = Sql & "AND t_Autocable_caddie.QtsEncelade<0  "
            Sql = Sql & " AND t_Autocable_caddie.Flinguer=False;"
            Con.Execute Sql
            
            Sql = "UPDATE t_Autocable_caddie SET t_Autocable_caddie.qte_produit = [qte_produit]-[QtsComande] "
            Sql = Sql & "WHERE t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & " "
            Sql = Sql & " AND t_Autocable_caddie.Flinguer=False;"
            Con.Execute Sql
        
'            Sql = "UPDATE ( con_contacts." & RefCaddyPrixU & " as [Prix de vente], con_contacts." & ChampQuantite & "  "
'            Sql = Sql & "AS [Qt disponible], con_contacts." & ChampQuantiteEncelade & "  "
'            Sql = Sql & "AS [Quantite Encelade],   "
'            Sql = Sql & "FROM con_contacts IN '"
'            Sql = Sql & TableauPath("" & Rs!Lib_Menu)
'            Sql = Sql & "' WHERE con_contacts.ContactID=" & Rs!id_produit & " ) AS  "
'            Sql = Sql & "MyFrom INNER JOIN t_Autocable_caddie ON MyFrom.Id_Produit =  "
'            Sql = Sql & "t_Autocable_caddie.Id_IndiceProjet SET MyFrom.[Qt disponible] = [ChampQuantite],  "
'            Sql = Sql & "MyFrom.[Quantite Encelade] = [ChampQuantiteEncelade] "
'            Sql = Sql & "WHERE t_Autocable_caddie.Flinguer=False "
'            Sql = Sql & "AND t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & "; "
        
        
            Sql = "UPDATE (SELECT con_contacts.*  "
            Sql = Sql & "FROM con_contacts IN '"
            Sql = Sql & TableauPath("" & Rs!Lib_Menu)
            Sql = Sql & "') AS MyFRom INNER JOIN t_Autocable_caddie ON MyFRom.ContactID =  "
            Sql = Sql & "t_Autocable_caddie.id_produit SET MyFRom." & ChampQuantite & " = [QtsEboutique],  "
            Sql = Sql & "MyFRom." & ChampQuantiteEncelade & " = [QtsEncelade] "
            Sql = Sql & "WHERE t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & "  "
            Sql = Sql & "AND t_Autocable_caddie.Flinguer=False;"
            Con.Execute Sql
        
            Sql = "UPDATE t_Autocable_caddie SET t_Autocable_caddie.Flinguer=true "
            Sql = Sql & "WHERE t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & " "
            Sql = Sql & " AND t_Autocable_caddie.Flinguer=False;"
            Con.Execute Sql
       
 
        End If
 
    End If
    Rs.MoveNext
Wend

Dim IdAvoir As Long
IdAvoir = NumCommadEboutiqueID

numDevis = NumeroChrono("CD", "CD_")
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1Caption.Caption = "Création de la commande N° : " & numDevis
IncrmentServer FormBarGrah, ""
Sql = "INSERT INTO T_Devis ( Id_User,id_caddie, id_r_social, Id_Menu, "
            Sql = Sql & "id_produit,  "
            Sql = Sql & "qte_produit,  "
            Sql = Sql & "TypePiece,  "
            Sql = Sql & "RefProduit, "
            Sql = Sql & "numDevis, "
            Sql = Sql & "PrixU_produit, "
            Sql = Sql & "Designation, "
            Sql = Sql & "Id_Avoire,  "
            Sql = Sql & "TVA,"
            Sql = Sql & "Prixderevientient, "
            Sql = Sql & "Remise ) IN '"
            Sql = Sql & TableauPath("Eb_Menu")
            Sql = Sql & "' "
            Sql = Sql & "(SELECT  "
            Sql = Sql & "t_Autocable_caddie.Id_User, "
            Sql = Sql & "t_Autocable_caddie.id_caddie, "
            Sql = Sql & "t_Autocable_caddie.id_r_social, "
            Sql = Sql & "t_Autocable_caddie.Id_Menu, "
            Sql = Sql & "t_Autocable_caddie.id_produit, "
            Sql = Sql & "t_Autocable_caddie.RefProduit, "
            Sql = Sql & "t_Autocable_caddie.qte_produit, "
            Sql = Sql & "t_Autocable_caddie.TypePiece, "
            Sql = Sql & "'" & numDevis & "' as numDevis, "
            Sql = Sql & "t_Autocable_caddie.PrixU_produit, "
            Sql = Sql & "t_Autocable_caddie.Designation,  "
            Sql = Sql & "" & IdAvoir & " as Id_Avoire ,  "
            Sql = Sql & "t_Autocable_caddie.TVA,  "
            Sql = Sql & "t_Autocable_caddie.Prixderevientient,  "
            Sql = Sql & "t_Autocable_caddie.Remise "
            Sql = Sql & "FROM t_Autocable_caddie "
            Sql = Sql & "WHERE t_Autocable_caddie.qte_produit>0 "
            Sql = Sql & "AND t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & "); "
            
            Sql = "INSERT INTO T_Devis (  "
            Sql = Sql & "id_caddie,  "
            Sql = Sql & "Id_Menu,  "
            Sql = Sql & "id_produit,  "
            Sql = Sql & "Id_User,  "
            Sql = Sql & "id_r_social,  "
            Sql = Sql & "date_modif,  "
            Sql = Sql & "qte_produit,  "
            Sql = Sql & "QtsDispo,"
            Sql = Sql & "TypePiece,  "
            Sql = Sql & "RefProduit,  "
            Sql = Sql & "PrixU_produit,  "
            Sql = Sql & "Designation,  "
            Sql = Sql & "TVA,  "
            Sql = Sql & "PrixRevient,  "
            Sql = Sql & "Remise,  "
            Sql = Sql & "numDevis,  "
            Sql = Sql & "Id_Avoire ) IN '"
           Sql = Sql & TableauPath("Eb_Menu")
            Sql = Sql & "'  "
            Sql = Sql & "SELECT t_Autocable_caddie.id_caddie,  "
            Sql = Sql & "t_Autocable_caddie.Id_Menu,  "
            Sql = Sql & "t_Autocable_caddie.id_produit,  "
            Sql = Sql & "t_Autocable_caddie.Id_User,  "
            Sql = Sql & "t_Autocable_caddie.id_r_social,  "
            Sql = Sql & "t_Autocable_caddie.date_modif,  "
            Sql = Sql & "[t_Autocable_caddie].qte_produit + [t_Autocable_caddie].[QtsComande],  "
            Sql = Sql & "t_Autocable_caddie.qte_produit,  "
            Sql = Sql & "t_Autocable_caddie.TypePiece,  "
            Sql = Sql & "t_Autocable_caddie.RefProduit,  "
            Sql = Sql & "t_Autocable_caddie.PrixU_produit,  "
            Sql = Sql & "t_Autocable_caddie.Designation,  "
            Sql = Sql & "t_Autocable_caddie.TVA,  "
            Sql = Sql & "t_Autocable_caddie.PrixRevient,  "
            Sql = Sql & "t_Autocable_caddie.Remise,  "
            Sql = Sql & "'" & numDevis & "' AS numDevis,  "
            Sql = Sql & IdAvoir & " AS Id_Avoire "
            Sql = Sql & "FROM t_Autocable_caddie "
            Sql = Sql & "WHERE ([t_Autocable_caddie].[qte_produit] + [t_Autocable_caddie].[QtsComande])>0  "
            Sql = Sql & "AND t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & ";"

            
            
            
            Con.Execute Sql
            Sql = "INSERT INTO T_Frais_Transport ( numDevis, ModeTransport,TVA )  IN '"
            Sql = Sql & TableauPath("Eb_Menu")
            Sql = Sql & "'  "
            Sql = Sql & "VALUES ( '" & numDevis & "', '" & MyReplace(GetDefaultEboutique("LivraisonComtoire", "Je désire  que ma commande me soit remise au comptoir.")) & "'," & TVA & ");"
             Con.Execute Sql
             Dim RsTotalCom As Recordset
             
            Sql = "SELECT T_Avoire.Debit FROM T_Avoire IN '"
            Sql = Sql & TableauPath("Eb_Menu")
            Sql = Sql & "'  "
            Sql = Sql & "WHERE T_Avoire.IdAvoire=" & IdAvoir & ";"
            Dim RsAvoir As Recordset
            Dim EbDebit As Double
             Set RsAvoir = Con.OpenRecordSet(Sql)
             If RsAvoir.EOF = False Then
                    EbDebit = RsAvoir!Debit
             End If
             
             FormBarGrah.ProgressBar1Caption.Caption = "Mise a jour de l'avenant"

             IncrmentServer FormBarGrah, ""
             Set RsAvoir = Con.CloseRecordSet(RsAvoir)
             Sql = "SELECT Sum(MyFrom.Expr1) AS EbDebit  "
            Sql = Sql & "FROM (select  ([qte_produit]+[QtsComande])*[PrixU_produit]*(1-[Remise])* "
            Sql = Sql & "(1+([TVA]/100)) AS Expr1 "
            Sql = Sql & "FROM t_Autocable_caddie "
            Sql = Sql & "WHERE t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & ") AS MyFrom;"
            
            Set RsAvoir = Con.OpenRecordSet(Sql)
            If RsAvoir.EOF = False Then
                EbDebit = EbDebit + Val(Replace(Trim("" & RsAvoir!EbDebit), ",", "."))
            End If
            Set RsAvoir = Con.CloseRecordSet(RsAvoir)
            Sql = "UPDATE T_Avoire IN '"
            Sql = Sql & TableauPath("Eb_Menu")
            Sql = Sql & "'  SET T_Avoire.Debit = " & Replace(EbDebit, ",", ".") & " "
            Sql = Sql & "WHERE T_Avoire.IdAvoire=" & IdAvoir & ";"
            
            

            
            
            Con.Execute Sql

            Sql = "DELETE t_Autocable_caddie.* FROM t_Autocable_caddie "
            Sql = Sql & "WHERE t_Autocable_caddie.Id_IndiceProjet=" & Id_Projet & ";"
            Con.Execute Sql
            Dim txtDevis As String
            Dim NouContacter As String
            NouContacter = EboutiqueEboutiqueGetDefault("email", "", TableauPath("Eb_Menu"))
            txtDevis = CrerDevis(numDevis, NouContacter)
            
'            Sql = "SELECT T_Message_Mail.Sujet,T_Message_Mail.Body, T_Users.Email "
'Sql = Sql & "FROM T_Users INNER JOIN (T_Message_Mail INNER JOIN T_Destinataire ON T_Message_Mail.Id = "
'Sql = Sql & "T_Destinataire.Id_Message) ON T_Users.Id = T_Destinataire.Id_Useur "
'Sql = Sql & "WHERE T_Message_Mail.Routine='" & MyReplace(Routine) & " ' "
'Sql = Sql & "AND T_Users.Email Is Not Null;"
'Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
'    While Rs.EOF = False
'        Destinataire = Destinataire & Rs!EMail & ";"
'        Sujet = Rs!Sujet
'        Body = ReplaceHtml(Rs!Body)
'        Rs.MoveNext
'    Wend
'    Destinataire = Left(Destinataire, Len(Destinataire) - 1)
Dim RsSmtp As Recordset
Sql = "SELECT [PI] & '_'& [PI_Indice] AS Pièce, NomenclaturFinal.Designation, NomenclaturFinal.Famille,  "
Sql = Sql & "NomenclaturFinal.Fournisseur, NomenclaturFinal.Ref, NomenclaturFinal.RefFour, NomenclaturFinal.Qts * " & NbPice & " as [Qts a Fournir],  "
Sql = Sql & "NomenclaturFinal.ISO, NomenclaturFinal.TEINT, NomenclaturFinal.TEINT2, NomenclaturFinal.SECT,  "
Sql = Sql & "NomenclaturFinal.Qts_Encelade, NomenclaturFinal.Qts_E_Boutique, NomenclaturFinal.Qts_Appro* " & NbPice & " as [Qts Appro],  "
Sql = Sql & "NomenclaturFinal.Prix_Revient, NomenclaturFinal.Prix_Vente, NomenclaturFinal.Options, NomenclaturFinal.Commentaires "
Sql = Sql & "FROM T_indiceProjet INNER JOIN NomenclaturFinal ON T_indiceProjet.Id = NomenclaturFinal.Id_IndiceProjet WHERE "
Sql = Sql & CloseLiK & "AND  NomenclaturFinal.Lib_Menu Is Null;"
Set Rs = Con.OpenRecordSet(Sql)
Dim MyFichier As String
Dim IndexFieldes As Long
Dim RRRR As String
Dim MyPasPice As String
MyFichier = ""
If Rs.EOF = False Then
MyPasPice = Rs.Fields(0)
    For IndexFieldes = 0 To Rs.Fields.Count - 1
        MyFichier = MyFichier & Rs.Fields(IndexFieldes).Name & ";"
    Next
    MyFichier = Left(MyFichier, Len(MyFichier) - 1)
 Dim toto
    toto = Rs.GetString(, , ";", "¤")


'    While RsBase.EOF = False
toto = Replace(toto, Chr(13), "")
toto = Replace(toto, Chr(10), " ")
toto = Replace(toto, "¤", vbCrLf)
        MyFichier = MyFichier & vbCrLf & toto
MyPasPice = Environ("USERPROFILE") & "\Mes Documents\Pas_Touvé_" & MyPasPice & ".csv"
' Crée le nom du fichier.
    Open MyPasPice For Output As #1
    Print #1, MyFichier    ' Écrit le texte.
    Close #1    ' Ferme le fichier.
    

End If

Sql = "SELECT t_Devis.numDevis, t_Devis.TypePiece, t_Devis.RefProduit,  "
Sql = Sql & "t_Devis.Designation, [qte_produit]-[QtsDispo] AS [QTS à Commander],  "
Sql = Sql & "t_Devis.PrixRevient, t_Devis.PrixU_produit, [Remise]*100 AS [Remise %],  "
Sql = Sql & "t_Devis.TVA "
Sql = Sql & "FROM t_Devis IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' "
Sql = Sql & "WHERE [qte_produit]-[QtsDispo]>0 "
Sql = Sql & "AND t_Devis.numDevis='" & numDevis & "';"
Set Rs = Con.OpenRecordSet(Sql)
Dim MyPiceCammade As String
MyFichier = ""
If Rs.EOF = False Then
MyPiceCammade = numDevis
    For IndexFieldes = 0 To Rs.Fields.Count - 1
        MyFichier = MyFichier & Rs.Fields(IndexFieldes).Name & ";"
    Next
    MyFichier = Left(MyFichier, Len(MyFichier) - 1)

    toto = Rs.GetString(, , ";", "¤")


'    While RsBase.EOF = False
toto = Replace(toto, Chr(13), "")
toto = Replace(toto, Chr(10), " ")
toto = Replace(toto, "¤", vbCrLf)
        MyFichier = MyFichier & vbCrLf & toto
MyPiceCammade = Environ("USERPROFILE") & "\Mes Documents\A_Cammander_" & MyPiceCammade & ".csv"
' Crée le nom du fichier.
    Open MyPiceCammade For Output As #1
    Print #1, MyFichier    ' Écrit le texte.
    Close #1    ' Ferme le fichier.
    

End If

Dim Pj As String
If Trim("" & txtDevis) <> "" Then
    If Trim("" & MyPasPice) <> "" Then Pj = MyPasPice
    If Trim("" & MyPiceCammade) <> "" Then
        If Trim("" & Pj) <> "" Then Pj = Pj & ";"
        Pj = Pj & MyPiceCammade
    End If
    Set RsSmtp = Con.OpenRecordSet("SELECT T_Serveur_Smtp.* FROM T_Serveur_Smtp;")
     FormBarGrah.ProgressBar1Caption.Caption = "Envoi   Mail : " & NouContacter

             IncrmentServer FormBarGrah, ""
    MailEnvoi RsSmtp!SMTP, RsSmtp!Authentification, RsSmtp!Utilisatuer, RsSmtp!PassWord, RsSmtp!Port, 15, RsSmtp!Messagerie, NouContacter, "", "Cammade AutoCâble N° : " & numDevis & " Pour un quantité de " & NbPice & " pièce(s)", txtDevis, Pj
    Set RsSmtp = Con.CloseRecordSet(RsSmtp)
Dim Fso As New FileSystemObject
If Trim("" & MyPasPice) <> "" Then Fso.DeleteFile Trim("" & MyPasPice)
If Trim("" & MyPiceCammade) <> "" Then Fso.DeleteFile Trim("" & MyPiceCammade)
Set Fso = Nothing
End If
            
End Sub
