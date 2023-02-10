VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Portal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
        
        '             border,   headerFG,  headerBG,  headrHiFg, hdrHiBg ,itmFgColor, itmBgColor, itmHiFgColor, itmHiBgColor
     Const MenuColor = "'#FFFFFF', '#000080', '#C0F0C0', '#FF0000', '#C0C0F0', '#000080', '#C0C0FF','#FFFFFF', '#000080' "
    '            <---------- HEADER ----------------><---------- ITEMS----------------------->
     Const MenuFonts = " 'Verdana', 'plain', 'bold', 'xx-small', 'Verdana', 'plain', 'bold', 'xx-small' "
    '           BorderSize,Height,SepSize
     Const BorderSize = "1, 3, 1 "
     Const ImageSet = "'','.bouton_on.jpg','bouton_off.jpg'"
    
    Const ForReading = 1
    Const ForWriting = 2
    Const ForAppending = 8
    Const GlobalStyle = "<link rel=STYLESHEET href=""PMainStyle1.css"" type=""text/css""> "

    
    Public url_base As String
    Public Application  As Application
    Public Request As Request
    Public Session As Session
    Public Response As Response
    Public Server As Server
    Public My_Conn
    Public FormCaption As String
    Public FlagHTMLEditor As Boolean
    Public res As String
    
    
Public Function GroupSelect() As String

    CheckConnexion
    CheckAdminRights

    pr ("<html>")
    pr ("<head>")
    pr ("<link rel=STYLESHEET href='PMainStyle1.css' type='text/css'>")
    pr ("<script src='Portal_Cat_tree.js'></script>")
    pr ("<script>")
    pr ("USETEXTLINKS = 0")
    pr ("indexOfEntries = new Array")
    pr ("nEntries = 0")
    pr ("doc = document")
    pr ("browserVersion = 0")
    pr ("selectedFolder = 0")
    pr ("</script>")
    pr ("<script>")
    pr ("USETEXTLINKS = 1")
    pr ("</script>")
    pr ("</head>")
    pr ("<body background='Background.asp' topmargin='7' >")
    pr ("<script>")

    '=========================================
    ' Chargement de la liste des groupes issues de dbp_Groups
    '=========================================
    SQL = "SELECT * FROM dbp_Groups WHERE (Cid=" & Session("candidat_APPLICATION") & ") ORDER BY GroupID;"
    Set Group = Server.CreateObject("ADODB.Recordset")
    Set subGroup = Server.CreateObject("ADODB.Recordset")
    Group.Open SQL, My_Conn
    '=========================================
    Response.Write "foldersTree = gFld('[Add new group]','../portal_ASP/Portal.ASP?mode=GroupAddNew target=CatMain')" & vbCrLf
    Do Until Group.EOF
         If Group("GroupParent") = 0 Then
            pr ("tv1  = insFld(foldersTree, gFld(""" & Group("GroupDescr") & """,""../Portal_ASP/Portal.asp?mode=GroupEdit&group=" & Group("GroupId") & " target=CatMain""))")
            SQL2 = "SELECT * FROM dbp_Groups WHERE ( (Cid=" & Session("candidat_APPLICATION") & ") AND (GroupParent=" & Group("GroupId") & ")) ORDER BY GroupID;"
            subGroup.Open SQL2, My_Conn
            Do Until subGroup.EOF
                pr ("tv2  = insFld( tv1 , gLnk(""" & subGroup("GroupDescr") & """,""../Portal_ASP/Portal.asp?mode=GroupEdit&group=" & subGroup("GroupId") & " target=CatMain""))")
                subGroup.MoveNext
            Loop
            subGroup.Close
        End If
    Group.MoveNext
    Loop
Group.Close
Set Group = Nothing
Set subGroup = Nothing
My_Conn.Close
Set My_Conn = Nothing

pr ("</script>")
pr ("<center><table border=0 cellpadding=0 cellspacing=0 width=142 height=22>")
pr ("<tr><td background='../portal_image/Cart_petit.gif' align=center><font id='THeading'><b>Groups</td></tr>")
pr ("</table></center>")
pr ("<br>")
pr ("<script>")
pr ("initializeDocument()")
pr ("</script>")
pr ("</body>")
pr ("</html>")

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                OnStartPage
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub OnStartPage(PassedScriptingContext As ScriptingContext)

    Set MyScriptingContext = PassedScriptingContext
    Set Request = MyScriptingContext.Request
    Set Response = MyScriptingContext.Response
    Set Application = MyScriptingContext.Application
    Set Server = MyScriptingContext.Server
    Set Session = MyScriptingContext.Session

End Sub

Public Sub Main()
    

    '**************************  License Information ******
    Session("candidat_IsEvaluation") = "False"
    Session("candidat_LicensedTo") = "Euxia Web Solutions"
    Session("candidat_ProductID") = "EUXIA PORTAL-00/1798"
    Session("candidat_DateLicensed") = "2/09/2000"
    Session("candidat_accessToSetting") = "true"
    '******************************************************
    
        
    Select Case UCase(Request("mode"))
        Case "MENU"
            Call navBuildMenu
        
        Case "TOPICS"
            Call topicView(0)
        Case "TOPICSADDNEW"
            Call topicAddNew(0)
        Case "TOPICSEDIT"
            Call topicEdition(0)
        Case "TOPICSDELETE"
            Call topicDelete(0)
        Case "TOPICSUPDATE"
            Call topicUpdate(0)
        
        Case "CATEGORYADDNEW"
            Call catAddNew
        Case "CATEGORYSELECT"
            Call catSelect
        Case "CATUPDATE"
            Call catUpdate(0)
        Case "CATEDITION"
            Call catEdition(0)
        Case "CATDELETE"
            Call catDelete(0)
        
        Case "USERADDNEW"
            Call userAddNew
        Case "USERUPDATE"
            Call userUpdate(0)
        Case "USEREDIT"
        Case "USERDELETE"
        
        Case "GROUPADDNEW"
            Call groupAddNew
        Case "GROUPEDIT"
            Call groupEdition
        Case "GROUPDELETE"
            Call groupDelete
        Case "GROUPUPDATE"
            Call groupUpdate
        Case "GROUPSELECT"
            Call GroupSelect
        Case "GROUPPERMUPDATE"
            Call GroupPermUpdate
        Case "GROUPPERMUPDATE2"
            Call GroupPermUpdate2
        Case "GROUPPERMEDIT"
            Call GroupPermEdit
        Case "GROUPPERMEDIT2"
            Call GroupPermEdit2
        Case "GROUPMAIL"
            Call GroupMail
        Case "DISPLAY"
            Call display_htm
        Case "QUICK"
            Call display_quick
        Case "SEARCH"
            Call search
        Case "RELATEDLINK"
            Call rel_seek
        Case "LOGIN"
            Call checkUser
        Case "STAND"
            Call stand
        Case "RESERVE_STAND"
            Call reserve_stand
        Case "RESERVE_CONGRES"
            Call reserve_congres
            
     
    End Select
    
   'pr ("Fonction " & Request("mode") & " inconnue ")

End Sub
Public Function navBuildMenu() As String
Dim my_menu As String
Dim my_page As String

CheckConnexion
'définition des paramètres d'affichage pour les niveaux d'arborescence
my_menu = "font_level0 = ""Arial"";bold_level0 = false;italic_level0 = false;taille_level0 = 1;couleur_level0 = ""#6B7073"";font_level1 = ""Arial"";bold_level1 = true;" & vbCrLf
my_menu = my_menu & "italic_level1 = true;taille_level1 = 1;couleur_level1 = ""#6B7073"";font_level2 = ""Arial"";bold_level2 = false;italic_level2 = false;taille_level2 = 1;couleur_level2 = ""#000000"";target = ""contenu"";" & vbCrLf
Session("candidat_secteur_id") = Request.QueryString("catid")
'construction du menu
Call arbre_menu(CLng(Request.QueryString("catid")), Request.QueryString("catrep"), 2, "1")
my_menu = my_menu & res

my_page = "<html><head><title>gauche</title></head>" & vbCrLf
my_page = my_page & "<body bgcolor=""#ffffff"" background=""../portal_html" & Request.QueryString("catrep") & "bgmenu.jpg"" topmargin=""0"" leftmargin=""0"" marginheight=""0"" marginwidth=""0"">"
my_page = my_page & "<table cellspacing=""0"" cellpadding=""0"" border=""0"">"
my_page = my_page & "<tr><td><table cellspacing=""0"" cellpadding=""0"" border=""0""><tr>"
my_page = my_page & "<td valign=""top"" >"
my_page = my_page & "<table cellspacing=""0"" cellpadding=""0"" border=""0""  >"
my_page = my_page & "<tr valign=""top"">"
my_page = my_page & "<td ><br>"
my_page = my_page & "<br><br><br>"
my_page = my_page & "<script src=""../portal_java/menu.js""></script>" & vbCrLf

my_page = my_page & "<script language=""Javascript"">" & my_menu

my_page = my_page & "renderSitemap( ""OUTLINE"", false, ""#000000"", ""#ffffff"", ""#6B7073"", ""../portal_html" & Request.QueryString("catrep") & """ );"
my_page = my_page & "</script>"

my_page = my_page & "</td></tr></table></td></tr></table></td></tr></table></body></html>"
My_Conn.Close
pr (my_page)
End Function
Sub arbre_menu(my_parent, my_rep, my_level, my_parent_sort)
Dim SQL1 As String
Dim rs1
        
        SQL1 = "SELECT dbp_CatUserList.CatId, dbp_Categories.CatCaption, dbp_Categories.Catdescription, dbp_Categories.CatPath,dbp_Categories.Parent, dbp_Categories.CatHasChilds, dbp_Categories.CatLevel, dbp_Categories.CatSort, dbp_Categories.CatExtURL, dbp_Categories.Catstyle, dbp_Categories.Catx  "
        SQL1 = SQL1 & "FROM ((dbp_Clients INNER JOIN dbp_Users ON dbp_Clients.CId = dbp_Users.CId) INNER JOIN dbp_CatUserList ON dbp_Users.UserId = dbp_CatUserList.UserId) "
        SQL1 = SQL1 & "INNER JOIN dbp_Categories ON dbp_CatUserList.CatId = dbp_Categories.CatId "
        SQL1 = SQL1 & "WHERE (((dbp_Categories.Parent)=" & my_parent & ") AND ((dbp_Categories.CatLevel)=" & my_level & ") AND ((dbp_Clients.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_Users.UserId)=" & Session("candidat_USERId") & ")) "
        SQL1 = SQL1 & "ORDER BY dbp_Categories.CatSort;"
       ' Debug.Print SQL1
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open SQL1, My_Conn
        
        Do While Not rs1.EOF
            If rs1("catexturl") = "" And rs1("catstyle") = 3 Then
                my_href = "../portal_asp/portal.asp?mode=topics&Category=" & rs1("CatId") & "&CatStyle=3"
            Else
                If InStr(1, rs1("catexturl"), "http") = 0 And InStr(1, rs1("catexturl"), "../") = 0 Then
                    my_href = "." & my_rep & rs1("catexturl")
                Else
                    my_href = rs1("catexturl")
                End If
            End If
            my_href = "../portal_html/droite.Asp?catid=" & rs1("CatId") & "&urltogo=" & Replace(my_href, "&", "µ")
            If my_level = 2 Then
                res = res & "menu=""" & rs1("catsort") & """;text = """ & rs1("catdescription") & """;image_collapsed = ""menu" & rs1("catsort") & ".gif"";image_extended = ""menu" & rs1("catsort") & "b.gif"";hauteur=" & rs1("catx") & ";url = """ & my_href & """;target = ""droite"";item_last = false;lecture();" & vbCrLf
            Else
                res = res & "menu=""" & my_parent_sort & "." & rs1("catsort") & """;text = """ & rs1("catdescription") & """;image_collapsed = null;image_extended = null;url = """ & my_href & """;target = ""droite"";item_last = false;lecture();" & vbCrLf
            End If
            If rs1("cathaschilds") > 0 Then
                If my_level = 2 Then
                    my_parent_sort2 = rs1("catsort")
                Else
                    my_parent_sort2 = my_parent_sort & "." & rs1("catsort")
                End If
                my_id = CLng(rs1("catid"))
                my_level2 = my_level + 1
                Call arbre_menu(my_id, my_rep, my_level2, my_parent_sort2)
            End If
             res = res & vbCrLf
       
        rs1.MoveNext
        Loop
        rs1.Close

End Sub

Public Function navSetCurrentFilter(ByVal currentfilter As Integer) As String
End Function
Public Function catAddNew() As String
    
    Dim i, s
    CheckConnexion
    CheckAdminRights

    '*************************************************************
    ' Formulaire de saisie
    '*************************************************************
    FormCaption = "Ajouter une catégorie au portail"
    InitForm "ADM_CAT_ADD_REP", "../Portal_Asp/Portal.asp?mode=CatUpdate&Category=0"

    ' Variables Session

    CId = Session("candidat_APPLICATION")

    CatOwner = 0
    

    CreateField "CATid", "ID de la Categorie (laisser à 0 pour la création)", 0

    CreateField "CATCAPTION", "Titre (requis)", ""
    CreateTextArea "CATDESCRIPTION", "Description", "", 100, 40, 5
    
    InitSelectList "CATStyle", "Type de catégorie", 1, 1
    AddSelectList 9, "Rubrique Principale, niv. arbre = 1, affichée sur home page+menu haut", 9
    AddSelectList 8, "Rubrique Principale, niv. arbre = 1, affichée sur menu haut", 0
    AddSelectList 1, "Rubrique type Outils, niv. arbre = 1", 0
    AddSelectList 6, "Elément de menu gauche, 1er niv", 0
    AddSelectList 3, "Elément de menu gauche, 1er niv, type News", 0
    AddSelectList 7, "Elément de menu gauche, 1er niv, raccourci sur menu haut", 0
    AddSelectList 5, "Elément de menu gauche, 2eme niv ou supérieur", 0
    'AddSelectList 2, "???????", 0
    'AddSelectList 4, "?????????", 0
    EndSelectList
    CreateField "CATSORT", "Ordre de tri numérique (requis)", "    "
    CreateTextArea "cathtmlcontent", "Page HTML", "", 100, 80, 40
    
    InitSelectList "TYPEURL", "Type d'URL (création automatique)", 1, 1
    AddSelectList 0, "Autre (remplir le champ URL)", 0
    AddSelectList 1, "Page html", 0
    AddSelectList 2, "News", 0
    AddSelectList 3, "Moteur de recherche d'emploi", 0
    AddSelectList 4, "Publication d'offres d'emploi (profil DRH)", 0
    AddSelectList 5, "Gestion des offres d'emploi (profil ADMIN)", 0
    AddSelectList 6, "Contact/HelpDesk", 0
    AddSelectList 7, "Résultats des Sondages", 0
    AddSelectList 8, "Résumé de tous les évènements", 0
    i = 9
    'Recherche des events
    SQL = "select event_title from dbe_events order by event_title"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open SQL, My_Conn
    Do Until rs1.EOF
        AddSelectList i, "Evènement '" & rs1("event_title") & "'", 0
        i = i + 1
     rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    'Recherche des faqs
    SQL = "select name from dbfaq_faq order by name"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open SQL, My_Conn
    Do Until rs1.EOF
        AddSelectList i, "FAQ '" & rs1("name") & "'", 0
        i = i + 1
     rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    'Recherche des sondages
    SQL = "select pollname from dbpoll_poll order by pollname"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open SQL, My_Conn
    Do Until rs1.EOF
        AddSelectList i, "Sondage '" & rs1("pollname") & "'", 0
        i = i + 1
     rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    
    EndSelectList
    
    
    CreateField "CATURL", "URL (laisser vide pour la création automatique)", ""

    InitSelectList "CATPARENT", "Catégorie parente", "", 1
    SQL = "select * from dbp_Categories order by CatId"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open SQL, My_Conn
    AddSelectList 0, "Root", 0
    Call cat_root(my_res, 0, 1)
    pr (my_res)
   
    EndSelectList
    
    pr ("<br><span class=""para"">Champs obligatoires pour les rubriques (notées 1), <br>éléments de menu gauche 1er niveau (notés 2)<br></span>")
    pr ("<br>")
    CreateField "CATREP", "Nom du Répertoire sur le serveur (1)", ""
    CreateField "CATX", "Coord X de l'image sur la home page (1) / hauteur de l'image (2)", 0
    CreateField "CATY", "Coord Y de l'image sur la home page (1)", 0
    
    EndForm2 ("ADD")

End Function
Public Function catSelect() As String

CheckConnexion
CheckAdminRights
pr ("<html>")
pr ("<head>")
pr ("<link rel=STYLESHEET href='PMainStyle1.asp' type='text/css'>")
pr ("<script src='../Portal_Java/Portal_Cat_tree.js'></script>")
pr ("<script>")
pr ("USETEXTLINKS = 0")
pr ("indexOfEntries = new Array")
pr ("nEntries = 0")
pr ("doc = document")
pr ("browserVersion = 0")
pr ("selectedFolder = 0")
pr ("</script>")
pr ("<script>")
pr ("USETEXTLINKS = 1")
pr ("</script>")
pr ("<body>")
pr ("<script>")
pr ("foldersTree = gFld('[Add Category]','../portal_ASP/Portal.ASP?mode=CategoryAddNew target=CatMain')")
Call arbre(0, 1)

My_Conn.Close
Set My_Conn = Nothing
pr ("</script>")
pr ("</head>")
pr ("<body background='Background.asp'>")
pr ("<center><table border=0 cellpadding=0 cellspacing=0 width=142 height=22>")
pr ("<tr><td align=center><font id=""Dheading""><b>&nbsp;Category&nbsp;</td></tr>")
pr ("</table></center>")
pr ("<br>")
pr ("<script>")
pr ("initializeDocument()")
pr ("</script>")
pr ("</body>")
pr ("</html>")
End Function

Sub arbre(my_parent, my_level)
Dim SQL1 As String
Dim rs1

        SQL1 = "SELECT dbp_CatUserList.CatId, dbp_Categories.CatCaption, dbp_Categories.CatPath,dbp_Categories.Parent, dbp_Categories.CatHasChilds, dbp_Categories.CatLevel, dbp_Categories.CatSort, dbp_Categories.CatExtURL "
        SQL1 = SQL1 & "FROM ((dbp_Clients INNER JOIN dbp_Users ON dbp_Clients.CId = dbp_Users.CId) INNER JOIN dbp_CatUserList ON dbp_Users.UserId = dbp_CatUserList.UserId) "
        SQL1 = SQL1 & "INNER JOIN dbp_Categories ON dbp_CatUserList.CatId = dbp_Categories.CatId "
        SQL1 = SQL1 & "WHERE (((dbp_Categories.Parent)=" & my_parent & ") AND ((dbp_Categories.CatLevel)=" & my_level & ") AND ((dbp_Clients.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_Users.UserId)=" & Session("candidat_USERId") & ")) "
        SQL1 = SQL1 & "ORDER BY dbp_Categories.CatSort;"
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open SQL1, My_Conn
        
        Do Until rs1.EOF

            If my_parent = 0 Then
                pr ("tv" & my_level & "   = insFld(foldersTree, gFld(""" & rs1("CatCaption") & """,""../Portal_ASP/Portal.asp?mode=catedition&Category=" & rs1("CatId") & " target=CatMain""))")
            Else
                pr ("tv" & my_level & "   = insFld(tv" & my_level - 1 & ", gFld(""" & rs1("CatCaption") & """,""../Portal_ASP/Portal.asp?mode=catedition&Category=" & rs1("CatId") & " target=CatMain""))")
            End If
            If rs1("cathaschilds") > 0 Then Call arbre(rs1("catid"), my_level + 1)
        
        rs1.MoveNext
        Loop
        rs1.Close

End Sub


Public Function catUpdate(ByVal catid As Integer) As Boolean
    CheckConnexion
    CheckAdminRights
    Dim typeurl, my_url
    Dim sortie As Boolean
    
    If Request("Category") = "0" Then
    
                '*************************************************************
                ' Calcul de CatId
                '*************************************************************
                SQL = "select * from dbp_Categories where CId=" & Session("candidat_APPLICATION") & " order by CatId "
                Set rs1 = Server.CreateObject("ADODB.Recordset")
                rs1.Open SQL, My_Conn
                Do Until rs1.EOF
                    newcatid = rs1("CatID")
                    rs1.MoveNext
                Loop
                newcatid = newcatid + 1
                rs1.Close
                Set rs1 = Nothing
                
                'recherche niveau ds arbo et/ou catégorie parente
                If Request.Form("CATPARENT") = 0 Then 'Root
                    catlevel = 1
                    CatPath = ChkString(Request.Form("CATCAPTION"))
                Else
                ' Recherche de la categorie parente
                    SQL = "SELECT dbp_Categories.Catlevel, dbp_Categories.CatCaption, dbp_Categories.CatId FROM dbp_Categories WHERE (CId=" & Session("candidat_APPLICATION") & " AND dbp_Categories.CatId=" & Request.Form("CATPARENT") & ");"
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    rs.Open SQL, My_Conn
                    CatPath = rs("CatCaption") + " : " + ChkString(Request.Form("CATCAPTION"))
                    catlevel = rs("Catlevel") + 1
                    rs.Close
                    My_Conn.Execute ("update dbp_categories set cathaschilds=1 where (CId=" & Session("candidat_APPLICATION") & " AND dbp_Categories.CatId=" & Request.Form("CATPARENT") & ");")
                    Set rs = Nothing
                End If
                'RECHERCHE du lien automatique
                typeurl = val(Request.Form("TYPEURL"))
                If typeurl > 0 Then
                
                        Select Case typeurl
                        
                            Case 1 'Page html
                                my_url = "../portal_asp/portal.asp?mode=display&amp;catid=" & newcatid
                            Case 2 'News
                            Case 3 'Recherche d'emploi
                                my_url = "../portal_job/main.asp"
                            Case 4 'Recherche d'emploi mode DRH
                                my_url = "../portal_job/employer.asp"
                            Case 5 'Recherche d'emploi mode admin
                                my_url = "../portal_job/admin.asp"
                            Case 6 'Contact/HelpDesk
                                my_url = "../portal_helpdesk/default.asp"
                            Case 7 'Sondage
                                my_url = "../portal_survey/collect.asp"
                            Case 8 ' résumé des events
                                my_url = "../portal_events/resume.asp"
                            Case Else
                               i = 9
                               sortie = False
                                'Recherche des events
                                SQL = "select event_id from dbe_events order by event_title"
                                Set rs1 = Server.CreateObject("ADODB.Recordset")
                                rs1.Open SQL, My_Conn
                                Do Until rs1.EOF
                                    If i = typeurl Then
                                        my_url = "../portal_events/default.asp?event_id=" & rs1("event_id")
                                        sortie = True
                                        Exit Do
                                    End If
                                    i = i + 1
                                 rs1.MoveNext
                                Loop
                                rs1.Close
                                Set rs1 = Nothing
                                If Not sortie Then
                                    'Recherche des faqs
                                    SQL = "select fldauto from dbfaq_faq order by name"
                                    Set rs1 = Server.CreateObject("ADODB.Recordset")
                                    rs1.Open SQL, My_Conn
                                    Do Until rs1.EOF
                                        If i = typeurl Then
                                            my_url = "../portal_faq/default.asp?faqid=" & rs1("fldauto")
                                            sortie = True
                                            Exit Do
                                        End If
                                        i = i + 1
                                     rs1.MoveNext
                                    Loop
                                    rs1.Close
                                    Set rs1 = Nothing
                                    If Not sortie Then
                                        'Recherche des sondages
                                        SQL = "select id from dbpoll_poll order by pollname"
                                        Set rs1 = Server.CreateObject("ADODB.Recordset")
                                        rs1.Open SQL, My_Conn
                                        Do Until rs1.EOF
                                            If i = typeurl Then
                                                my_url = "../portal_survey/poll.asp?id=" & rs1("id")
                                                sortie = True
                                                Exit Do
                                            End If
                                            i = i + 1
                                         rs1.MoveNext
                                        Loop
                                        rs1.Close
                                        Set rs1 = Nothing
                                    End If
                                End If
                        End Select
                End If
                
                
                
                ' INSERTION CATEGORY direct par ajout dans  dbp_Categories - pas de traitement particulier

                SQL = "Insert into dbp_Categories (CId, CatId, CatCaption, CatDescription, CatOwner, Parent, CatSort, CatPos, CatLevel,CatPath, CatExtUrl, CatHasChilds, CatAccessLog, CatApprove, CatStyle,catrep,catx,caty, cathtmlcontent)"
                If InStr(Request.Form("CATDESCRIPTION"), "<") > 0 Then
                    SQL = SQL & " Values (" & Session("candidat_APPLICATION") & ", " & newcatid & ",'" & Trim(ChkString(Server.HTMLEncode(Request.Form("CATCAPTION")))) & "','" & ChkString(Request.Form("CATDESCRIPTION")) & "', "
                Else
                    SQL = SQL & " Values (" & Session("candidat_APPLICATION") & ", " & newcatid & ",'" & Trim(ChkString(Server.HTMLEncode(Request.Form("CATCAPTION")))) & "','" & ChkString(Server.HTMLEncode(Request.Form("CATDESCRIPTION"))) & "', "
                End If
                SQL = SQL & 0 & "," & Request.Form("CATPARENT") & ",'" & Trim(Request.Form("CATSORT")) & "', 0 ," & catlevel & ",'" & CatPath & "','"
                SQL = SQL & IIf((typeurl > 0), my_url, Request.Form("CATURL")) & "',0,0,0," & Request.Form("CATSTYLE") & ",'" & Trim(ChkString(Request.Form("CATREP"))) & "'," & Request.Form("CATX") & "," & Request.Form("CATY") & ",'" & ChkString(Request.Form("cathtmlcontent")) & "')"
                My_Conn.Execute (SQL)
                My_Conn.Close
                Set My_Conn = Nothing
                Response.Redirect ("Portal.asp?mode=GroupPermUpdate&category=" & newcatid)
    Else
                If InStr(Request.Form("CATDESCRIPTION"), "<") > 0 Then
                    SQL = "UPDATE dbp_Categories  SET CatCAPTION='" & Trim(ChkString(Server.HTMLEncode(Request.Form("CATCAPTION")))) & "', CatDescription='" & Trim(ChkString(Request.Form("CATDESCRIPTION")))
                Else
                    SQL = "UPDATE dbp_Categories  SET CatCAPTION='" & Trim(ChkString(Server.HTMLEncode(Request.Form("CATCAPTION")))) & "', CatDescription='" & Trim(ChkString(Server.HTMLEncode(Request.Form("CATDESCRIPTION"))))
                End If
                If Request.Form("CATPARENT") = 0 Then 'Root
                    catlevel = 1
                    CatPath = Request.Form("CATCAPTION")
                Else
                ' Recherche de la categorie parente pour construire le CatPath
                    SQL2 = "SELECT dbp_Categories.Catlevel, dbp_Categories.CatCaption, dbp_Categories.CatId FROM dbp_Categories WHERE (CId=" & Session("candidat_APPLICATION") & " AND dbp_Categories.CatId=" & Request.Form("CATPARENT") & ");"
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    rs.Open SQL2, My_Conn
                    CatPath = rs("CatCaption") + " : " + Request.Form("CATCAPTION")
                    catlevel = rs("Catlevel") + 1
                    rs.Close
                    My_Conn.Execute ("update dbp_categories set cathaschilds=1 where (CId=" & Session("candidat_APPLICATION") & " AND dbp_Categories.CatId=" & Request.Form("CATPARENT") & ");")
                    Set rs = Nothing
                End If
                SQL = SQL & "',  Parent=" & Request.Form("CATPARENT") & ", CatSort='" & Trim(Request.Form("CATSORT")) & "', CatExtUrl ='" & Trim(ChkString(Server.HTMLEncode(Request.Form("CATURL"))))
                SQL = SQL & "', catlevel=" & catlevel & ", CatStyle=" & Request.Form("CATSTYLE") & ", catrep='" & Trim(ChkString(Request.Form("CATREP"))) & "', Catx=" & Request.Form("CATX") & ", Caty=" & Request.Form("CATY") & ", cathtmlcontent='" & ChkString(Request.Form("cathtmlcontent")) & "'  WHERE  ((dbp_Categories.CatId=" & Request.Form("CATId") & ") AND (dbp_Categories.CId=" & Session("candidat_APPLICATION") & "))"
                ' UPDATE CATEGORY dbp_Categories
                My_Conn.Execute (SQL)
                Response.Redirect ("Portal.asp?mode=GroupPermEdit&category=" & Request.Form("CATid"))
    End If

End Function
Public Function catDelete(ByVal catid As Integer) As Boolean
    
    CheckConnexion
    CheckAdminRights
    Category = Request.QueryString("category")
    
    SQL = "select parent from dbp_categories where ((dbp_Categories.CatId=" & Category & ") AND (dbp_Categories.CId=" & Session("candidat_APPLICATION") & "))"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, My_Conn
    If Not rs.EOF Then
        If rs("parent") > 0 Then
            Set RS2 = Server.CreateObject("ADODB.Recordset")
            SQL = "select count(catid) as nb from dbp_categories where ((dbp_Categories.parent=" & rs("parent") & ") AND (dbp_Categories.CId=" & Session("candidat_APPLICATION") & "))"
            RS2.Open SQL, My_Conn
            If Not RS2.EOF Then
                If RS2("nb") = 1 Then
                    My_Conn.Execute ("update dbp_categories set cathaschilds=0   WHERE (CId=" & Session("candidat_APPLICATION") & " AND dbp_Categories.CatId=" & rs("parent") & ");")
                End If
            End If
            RS2.Close
            Set RS2 = Nothing
        End If
    End If
    rs.Close
    'rs = Nothing
    ' Effacement dans Catégorie
    SQL = "Delete * From dbp_Categories    WHERE  ((dbp_Categories.CatId=" & Category & ") AND (dbp_Categories.CId=" & Session("candidat_APPLICATION") & "))"
    My_Conn.Execute (SQL)
    

    'Effacement dans GroupPerm
    SQL = "Delete * From dbp_GroupsPermission"
    SQL = SQL & " WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND  ((dbp_GroupsPermission.ObjTypeId)=100) AND ((dbp_GroupsPermission.ObjId)=" & Category & "));"
    My_Conn.Execute (SQL)
    
    'Effacement dans dbp_CatUserList
    SQL = "Delete * FROM dbp_CatUserList WHERE (((dbp_CatUserList.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_CatUserList.CatId)=" & Category & "));"
    My_Conn.Execute (SQL)
    
    Response.Redirect ("../Portal_asp/result.asp?message=category+successfully+deleted")
    

End Function
Public Function catEdition(ByVal catid As Integer) As String
      
    Dim s, i
    CheckConnexion
    CheckAdminRights
      
    Category = Request.QueryString("category")
    
    SQL = "select * from dbp_Categories  where ((CId=" & Session("candidat_APPLICATION") & ") AND  (CatId=" & Category & "));"
    
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open SQL, My_Conn
    
    FormCaption = rs1("CatCaption") + ": (edition)"
    
    InitForm "ADM_CAT_EDIT", "Portal.ASP?mode=catUpdate&category=" & rs1("catID")

    ' Variables Sessions
    
    CreateField "CATid", "ID de la Categorie", rs1("CatId")
    CreateField "CATCAPTION", "Titre", rs1("CatCaption")
    CreateTextArea "CATDESCRIPTION", "Description", rs1("CatDescription"), 100, 40, 5

    InitSelectList "CATStyle", "Type de catégorie", 1, 1
    AddSelectList 9, "Rubrique Principale, niv. arbre = 1, affichée sur home page+menu haut", rs1("CatStyle")
    AddSelectList 8, "Rubrique Principale, niv. arbre = 1, affichée sur menu haut", rs1("CatStyle")
    AddSelectList 1, "Rubrique type Outils, niv. arbre = 1", rs1("CatStyle")
    AddSelectList 6, "Elément de menu gauche, 1er niv", rs1("CatStyle")
    AddSelectList 3, "Elément de menu gauche, 1er niv, type News", rs1("CatStyle")
    AddSelectList 7, "Elément de menu gauche, 1er niv, raccourci sur menu haut", rs1("CatStyle")
    AddSelectList 5, "Elément de menu gauche, 2eme niv ou supérieur", rs1("CatStyle")
    'AddSelectList 2, "???????", rs1("CatStyle")
    'AddSelectList 4, "?????????", rs1("CatStyle")
    EndSelectList
    CreateField "CATSORT", "Ordre de tri", rs1("CatSort")
    CreateTextArea "cathtmlcontent", "Page HTML", rs1("cathtmlcontent"), 100, 80, 40

    CreateField "CATURL", "URL", rs1("CatExtUrl")

    selectedoption = rs1("Parent")
    rep = rs1("catrep")
    x = rs1("catx")
    Y = rs1("caty")
'    InitSelectList "CATPARENT", "Catégorie parente", SelectedOption, 1
'    my_catid = rs1("CatId")
'    rs1.Close
'
'    Set rs1 = Nothing
'    SQL = "select * from dbp_Categories  where (CId=" & session("candidat_APPLICATION") & " ) order by catlevel"
'    Set rs1 = Server.CreateObject("ADODB.Recordset")
'    rs1.Open SQL, My_Conn
'    AddSelectList 0, "Root", SelectedOption
'    Do Until rs1.EOF
'        If rs1("CId") = session("candidat_APPLICATION") And rs1("catid") <> my_catid Then
'            s = ""
'            For i = 1 To rs1("CatLevel")
'                s = s + "-"
'            Next i
'            s = s + ">"
'            AddSelectList rs1("CatId"), s + rs1("CatCaption"), SelectedOption
'        End If
'        rs1.MoveNext
'    Loop
'    rs1.Close
'    Set rs1 = Nothing
'
'    EndSelectList
    InitSelectList "CATPARENT", "Catégorie parente", "", 1
    SQL = "select * from dbp_Categories order by CatId"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open SQL, My_Conn
    AddSelectList 0, "Root", selectedoption
    Call cat_root2(my_res, 0, 1, selectedoption)
    pr (my_res)
    EndSelectList
    
    
    pr ("<br><span class=""para"">Champs obligatoires pour les rubriques (notées 1), <br>éléments de menu gauche 1er niveau (notés 2)<br></span>")
    pr ("<br>")
    
    CreateField "CATREP", "Nom du Répertoire sur le serveur (1)", rep
    CreateField "CATX", "Coord X de l'image sur la home page (1) / hauteur de l'image (2)", x
    CreateField "CATY", "Coord Y de l'image sur la home page (1)", Y

    Response.Write ("<br><hr>")
    CreateLink "Delete this category", "", "../Portal_ASP/Portal.asp?mode=catDelete&Category=" & Category
    Response.Write ("<br><hr>")
    CreateLink "Related Links", "", "../portal_asp/portal.asp?mode=topics&Category=" & Category & "&CatStyle=8"
    CreateLink "Add this category as a Related Link", "", "../portal_asp/portal.asp?mode=relatedlink&Category=" & Category & "&affiche=help"
    pr ("<br><hr>")
    EndForm2 ("update")

End Function

Public Function GroupPermEdit() As Boolean
    
    CheckConnexion
    CheckAdminRights
    Category = Request.QueryString("category")
    '=========================================
    ' Chargement de la liste des groupes issues de dbp_Groups
    '=========================================
    SQL = "SELECT * FROM dbp_Groups WHERE ((Cid=" & Session("candidat_APPLICATION") & ") AND (GroupParent=0)) ORDER BY GroupID;"
    Set Group = Server.CreateObject("ADODB.Recordset")
    Group.Open SQL, My_Conn
    '=========================================
    FormCaption = " Group Permission "
    InitForm "ADM_GRP_AUT_REP", "Portal.asp?mode=GroupPermEdit2&category=" & Category
    Do Until Group.EOF
            InitSelectList "G" & Group("GroupId"), Group("GroupDescr"), "", 1
            '=================================================
            ' Chargement des permissions d'un groupe pour la catégorie demandée
            '=================================================
            SQL2 = "SELECT * FROM dbp_GroupsPermission"
            SQL2 = SQL2 & " WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.GroupId)=" & Group("GroupId") & ") AND ((dbp_GroupsPermission.ObjTypeId)=100) AND ((dbp_GroupsPermission.ObjId)=" & Category & "));"
            Set permaccess = Server.CreateObject("ADODB.Recordset")
            permaccess.Open SQL2, My_Conn
            '===============================================
            ' Chargement de la table des permissions et Mise à jour des Combos
            '===============================================
            SQL1 = "SELECT * FROM dbp_PermsTypes WHERE (ObjTypeid=100) ORDER BY ObjPermType ;"
            Set PermId = Server.CreateObject("ADODB.Recordset")
            PermId.Open SQL1, My_Conn
            
            Do Until PermId.EOF
                    If permaccess.EOF Or permaccess.BOF Then                                                                                '=============
                            AddSelectList PermId("ObjPermType"), PermId("ObjPermDescr"), 0                       ' pas d'accès défini
                    Else
                            AddSelectList PermId("ObjPermType"), PermId("ObjPermDescr"), permaccess("ObjPermType")  ' On liste l'accès
                    End If
                    PermId.MoveNext
            Loop
            EndSelectList
            Set PermId = Nothing
        Group.MoveNext
    Loop

    EndForm2 ("Update")
    Set Group = Nothing
    Set PermId = Nothing
    Set permaccess = Nothing
    My_Conn.Close
    Set My_Conn = Nothing

End Function
Public Function GroupPermEdit2() As Boolean

    Category = Request.QueryString("category")

    CheckConnexion
    CheckAdminRights
    
    ' Pour l'Update , on efface tous les droits sur la catégorie et on réecrit par sécurité
    
    'Effacement dans GroupPerm   Ok
    SQL = "Delete * From dbp_GroupsPermission"
    SQL = SQL & " WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND  ((dbp_GroupsPermission.ObjTypeId)=100) AND ((dbp_GroupsPermission.ObjId)=" & Category & "));"
    My_Conn.Execute (SQL)
    
    'Effacement dans CatUserPref Ok
    SQL = "Delete * FROM dbp_CatUserList WHERE (((dbp_CatUserList.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_CatUserList.CatId)=" & Category & "));"
    My_Conn.Execute (SQL)

    ' On ouvre la Table dbp_Groups et pour chacun des groupes on écrit la permission issue de la request.form
    ' le champ est écrit dans request.form(group("GroupId"))

    SQL = "SELECT * FROM dbp_Groups WHERE ((Cid=1) AND (GroupParent=0)) ORDER BY GroupID;"
    
    Set Group = Server.CreateObject("ADODB.Recordset")
    
    Group.Open SQL, My_Conn
    
    Do Until Group.EOF
        
        ' Insertion des permissions groupes dans GroupPerm Ok
        
        SQL = "INSERT INTO dbp_GroupsPermission (Cid, GroupId, ObjTypeId, ObjPermId, ObjId, ObjPermType)"
        SQL = SQL & " Values (" & Session("candidat_APPLICATION") & "," & Group("GroupId") & ",100,1," & Category & "," & Request.Form("G" & Group("GroupId")) & ")"
          
        My_Conn.Execute (SQL)
            
        If Request.Form("G" & Group("GroupId")) > 0 Then
        
        ' Insertion des UsersPréferences pour chacun des users du groupe
        ' On liste les users du groupe par recherche de la valeur PermType 10 pour TypeId = 50
        
        Set UserGroup = Server.CreateObject("ADODB.Recordset")
        Set RsCheck = Server.CreateObject("ADODB.Recordset")
        
        SQL = "SELECT dbp_GroupsPermission.ObjId FROM dbp_GroupsPermission "
        SQL = SQL & "WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.GroupId)=" & Group("GroupId") & ") AND ((dbp_GroupsPermission.ObjTypeId)=50) AND ((dbp_GroupsPermission.ObjPermType)>0)  );"

        UserGroup.Open SQL, My_Conn
                
            Do Until UserGroup.EOF
                
                ' On vérifie que la catégorie n'est pas déjà existante dans CatUserPref pour le Group et le User assigné
                
    '            SQLCheck = "SELECT * from dbp_CatUserList WHERE ( (Cid=" & session("candidat_APPLICATION") & ") AND (CatId=" & Category & ") AND (UserId=" & UserGroup("ObjId") & "));"
   '             RsCheck.Open SQLCheck, My_Conn
                SQL1 = "delete from dbp_CatUserList "
                SQL1 = SQL1 & " where  Cid=" & Session("candidat_APPLICATION") & " and  CatId=" & Category & " and UserId=" & UserGroup("ObjId")
                My_Conn.Execute (SQL1)
                'If RsCheck.EOF Or RsCheck.BOF Then      ' si CatUserPref n'existe pas déjà
                        ' Insertion dans CatUserPref    1,CatId,User,0
                        SQL1 = "INSERT INTO dbp_CatUserList (Cid ,CatId, UserId, CatUserHide) "
                        SQL1 = SQL1 & "Values (" & Session("candidat_APPLICATION") & ", " & Category & "," & UserGroup("ObjId") & ",0)"
                        My_Conn.Execute (SQL1)
      '          End If
                
     '           RsCheck.Close
                UserGroup.MoveNext
            Loop
            Set UserGroup = Nothing
            End If
            Group.MoveNext
    Loop
    Set RsCheck = Nothing
    Set CheckPerm = Nothing
    Set Group = Nothing
    My_Conn.Close
    Set My_Conn = Nothing
    
    Response.Redirect ("../Portal_asp/result.asp?message=category+successfully+updated")
    
End Function

Public Function GroupPermUpdate()
    
    CheckConnexion
    CheckAdminRights
    
    Category = Request.QueryString("category")
    User = Request.QueryString("user")
    
    ' Chargement de la liste des groupes issues de dbp_Groups

    SQL = "SELECT * FROM dbp_Groups WHERE ((Cid=" & Session("candidat_APPLICATION") & ") AND (GroupParent=0)) ORDER BY GroupID;"
    If User = 0 Then SQL1 = "SELECT * FROM dbp_PermsTypes WHERE (ObjTypeid=100) ORDER BY ObjPermType ;"
    If Category = 0 Then SQL1 = "SELECT * FROM dbp_PermsTypes WHERE (ObjTypeid=50) ORDER BY ObjPermType ;"
    Set Group = Server.CreateObject("ADODB.Recordset")
     

    Group.Open SQL, My_Conn

    If Len(Trim(User)) > 0 Then InitForm "ADM_GRP_AUT_REP", "Portal.asp?mode=GroupPermUpdate2&user=" & User
    If Len(Trim(Category)) > 0 Then InitForm "ADM_GRP_AUT_REP", "Portal.asp?mode=GroupPermUpdate2&category=" & Category
    
    Do Until Group.EOF
        InitSelectList "G" & Group("GroupId"), Group("GroupDescr"), "", 1
            Set PermId = Server.CreateObject("ADODB.Recordset")
            PermId.Open SQL1, My_Conn
            Do Until PermId.EOF
                AddSelectList PermId("ObjPermType"), PermId("ObjPermDescr"), 0
                PermId.MoveNext
            Loop
            EndSelectList
            Set PermId = Nothing
            'If Group("GroupMailList")=1 then
                    'CreateCheckButton "CreateML","Mailing List :",group("GroupId"),""
            'end if
        Group.MoveNext
    Loop

    EndForm2 ("submit")
    Set Group = Nothing
    Set PermId = Nothing
    My_Conn.Close
    Set My_Conn = Nothing
End Function
Public Function GroupPermUpdate2()

    CheckConnexion
    CheckAdminRights

    Category = Request.QueryString("category")
    User = Request.QueryString("User")
    
    
    If Len(Trim(Category)) > 0 Then
    
        ' Pour l'Update , on efface tous les droits sur la catégorie et on réecrit par sécurité
    
        'Effacement dans GroupPerm
        
        SQL = "Delete * From dbp_GroupsPermission"
        SQL = SQL & " WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND  ((dbp_GroupsPermission.ObjTypeId)=100) AND ((dbp_GroupsPermission.ObjId)=" & Category & "));"
        My_Conn.Execute (SQL)
    
        'Effacement dans dbp_CatUserList
        
        SQL = "Delete * FROM dbp_CatUserList WHERE (((dbp_CatUserList.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_CatUserList.CatId)=" & Category & "));"
        My_Conn.Execute (SQL)
    
        ' On ouvre la Table dbp_Groups et pour chacun des groupes on écrit la permission issue de la request.form
        ' le champ est écrit dans request.form(group("GroupDescr"))

        SQL = "SELECT * FROM dbp_Groups WHERE ((Cid=" & Session("candidat_APPLICATION") & ") AND (GroupParent=0)) ORDER BY GroupID;"
    
        Set Group = Server.CreateObject("ADODB.Recordset")
    
        Group.Open SQL, My_Conn
    
        Do Until Group.EOF
        
            ' Insertion des permissions groupes dans GroupPerm
        
            SQL = "INSERT INTO dbp_GroupsPermission (Cid, GroupId, ObjTypeId, ObjPermId, ObjId, ObjPermType)"
            SQL = SQL & " Values (" & Session("candidat_APPLICATION") & "," & Group("GroupId") & ",100,1," & Category & "," & Request.Form("G" & Group("GroupId")) & ")"

            My_Conn.Execute (SQL)
        
            If Request.Form("G" & Group("GroupId")) > 0 Then
        
            ' Insertion des UsersPréferences pour chacun des users du groupe
            ' On liste les users du groupe par recherche de la valeur PermType 10 pour TypeId = 50
        
            Set UserGroup = Server.CreateObject("ADODB.Recordset")
            Set RsCheck = Server.CreateObject("ADODB.Recordset")
        
            SQL = "SELECT dbp_GroupsPermission.ObjId FROM dbp_GroupsPermission "
            SQL = SQL & "WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.GroupId)=" & Group("GroupId") & ") AND ((dbp_GroupsPermission.ObjTypeId)=50) AND ((dbp_GroupsPermission.ObjPermType)>0)  );"
        
            UserGroup.Open SQL, My_Conn
                Do Until UserGroup.EOF
                    
                    ' On vérifie que la catégorie n'est pas déjà existante dans CatUserPref pour le Group et le User assigné
                    
                    SQLCheck = "SELECT * from dbp_CatUserList WHERE ( (Cid=" & Session("candidat_APPLICATION") & ") AND (CatId=" & Category & ") AND (UserId=" & UserGroup("ObjId") & "));"
                    RsCheck.Open SQLCheck, My_Conn
                    If RsCheck.EOF Or RsCheck.BOF Then
                        
                        ' Insertion dans CatUserPref    1,CatId,User,0
                        SQL1 = "INSERT INTO dbp_CatUserList (Cid ,CatId, UserId, CatUserHide) "
                        SQL1 = SQL1 & "Values (" & Session("candidat_APPLICATION") & ", " & Category & "," & UserGroup("ObjId") & ",0)"
                        My_Conn.Execute (SQL1)
                    
                    End If
                    
                    RsCheck.Close
                    UserGroup.MoveNext
                Loop
                
                Set UserGroup = Nothing
                
                End If
            Group.MoveNext
        Loop
    
        Set Group = Nothing
        My_Conn.Close
        Set My_Conn = Nothing
    
        Response.Redirect ("../Portal_asp/result.asp?message=category+successfully+added")

    Else
            
        'On efface les droits du user sur les catégories pour plus de sécurité
        SQL = "delete from dbp_CatUserList WHERE ( (Cid=" & Session("candidat_APPLICATION") & ") AND (UserId=" & User & "));"
        My_Conn.Execute (SQL)
        
        ' On ouvre la Table dbp_Groups et pour chacun des groupes on écrit la permission issue de la request.form
        ' le champ est écrit dans request.form(group("GroupDescr"))

        SQL = "SELECT * FROM dbp_Groups WHERE (Cid=" & Session("candidat_APPLICATION") & ") ORDER BY GroupID;"
        Set Group = Server.CreateObject("ADODB.Recordset")
        Group.Open SQL, My_Conn
    
        Do Until Group.EOF
        
            ' Insertion des permissions user  dans GroupPerm
        
            SQL = "INSERT INTO dbp_GroupsPermission (Cid, GroupId, ObjTypeId, ObjPermId, ObjId, ObjPermType)"
            SQL = SQL & " Values (" & Session("candidat_APPLICATION") & "," & Group("GroupId") & ",50,1," & User & "," & Request.Form("G" & Group("GroupId")) & ")"
            
            ' Response.Write SQL & "<br>"
            
            My_Conn.Execute (SQL)
            
            If Request.Form("G" & Group("GroupId")) > 0 Then
                Set UserGroup = Server.CreateObject("ADODB.Recordset")
                'Set RsCheck = Server.CreateObject("ADODB.Recordset")
            
                ' Insertion des UsersPréferences pour chacun des users du groupe
                ' On liste les catégories visibles par le ou les groupes dont il fait partie
            
                SQL = "SELECT dbp_GroupsPermission.ObjId FROM dbp_GroupsPermission "
                SQL = SQL & "WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.GroupId)=" & Group("GroupId") & ") AND  ((dbp_GroupsPermission.ObjPermType)>0) AND ((dbp_GroupsPermission.ObjTypeId)=100));"
        
                UserGroup.Open SQL, My_Conn
                
                Do Until UserGroup.EOF
                    
                    ' On vérifie que la catégorie n'est pas déjà existante dans CatUserPref pour le Group et le User assigné
                    'SQLCheck = "SELECT * from dbp_CatUserList WHERE ( (Cid=" & session("candidat_APPLICATION") & ") AND (CatId=" & UserGroup("ObjId") & ") AND (UserId=" & User & "));"
                    'RsCheck.Open SQLCheck, My_Conn
                    '    If RsCheck.EOF Then
                    
                    ' Insertion dans CatUserPref    1,CatId,User,0
                    SQL1 = "INSERT INTO dbp_CatUserList (Cid ,CatId, UserId, CatUserHide) "
                    SQL1 = SQL1 & "Values (" & Session("candidat_APPLICATION") & ", " & UserGroup("ObjId") & "," & User & ",0)"
                    My_Conn.Execute (SQL1)
                    
                    '    End If
                    'RsCheck.Close
                    
                    UserGroup.MoveNext
                Loop
                Set UserGroup = Nothing
            End If
            Group.MoveNext
        Loop
        Set RsCheck = Nothing
        Set Group = Nothing
        My_Conn.Close
        Set My_Conn = Nothing
        Response.Redirect ("../Portal_asp/result.asp?message=Rights+for+user+are+set")
    
    
    End If

End Function
Public Function topicAddNew(ByVal catid As Integer) As String
    
    CheckConnexion
    categoryId = Request.QueryString("category")
    
'=============================================================
' Formulaire de saisie
'=============================================================

    FormCaption = "Add Topics"
    InitForm "ADM_TOP_ADD_REP", "Portal.ASP?mode=TopicsUpdate&topics=new&category=" & categoryId

    CId = Session("candidat_APPLICATION")


    CreateField "TOPid", "Topic ID", NewTopicId
    CreateField "TOPTITLE", "Title", ""
    CreateTextArea "TOPABSTRACT", "Contain", "", 2500, 80, 20
    'CreateHtmlEditor "TOPABSTRACT", "Contain", "", "ADM_TOP_ADD_REP"
    CreateField "START", "Appears from", Date
    CreateField "EXPIRE", "To (include)", DateAdd("m", 1, Date)
    InitSelectList "TOPTYPE", "Category", "", 1
    ' les récupérer à terme dans la base de données
    AddSelectList 1, "News", 0
    AddSelectList 2, "News for Home Page", 0
    AddSelectList 3, "Document Attach", 0
    AddSelectList 7, "Right column for Home Page", 0
    AddSelectList 4, "Application", 0
    AddSelectList 5, "w/HTML", 0
    AddSelectList 6, "Dynamic Url", 0
    AddSelectList 8, "Related Link for this category", 0
    EndSelectList
    
    
        InitSelectList "TYPEURL", "Type d'URL (création automatique)", 1, 1
        AddSelectList 0, "Autre (remplir le champ URL)", 0
        Dim i
        i = 1
        'Recherche des events
        SQL = "select event_title from dbe_events order by event_title"
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open SQL, My_Conn
        Do Until rs1.EOF
            AddSelectList i, "Evènement '" & rs1("event_title") & "'", 0
            i = i + 1
         rs1.MoveNext
        Loop
        rs1.Close
        Set rs1 = Nothing
        
        'Recherche des sondages
        SQL = "select pollname from dbpoll_poll order by pollname"
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open SQL, My_Conn
        Do Until rs1.EOF
            AddSelectList i, "Sondage '" & rs1("pollname") & "'", 0
            i = i + 1
         rs1.MoveNext
        Loop
        rs1.Close
        Set rs1 = Nothing
                            
    CreateField "TOPURL", "URL", ""
    CreateField "topicIcon", "Icon for right column", ""
    InitSelectList "TOPSTYLE", "Style set for this topic", PreviousStyle, 1
        AddSelectList 0, "Default", 0
        AddSelectList 3, "News", 0
    EndSelectList
    CreateCheckButton "NoTitle", "Remove Title", "Yes", ""
    CreateCheckButton "NoSignature", "Remove Signature", "No", ""
    CreateField "TITLEBG", "Title Background", ""
    CreateField "ABSTRACTBG", "Topic Background", ""

    ' Restriction a des sous-catégories
    Response.Write "<p class='smallerheader'>Default Permission :<br>"
    '===============
    SQL = "SELECT * FROM dbp_Groups WHERE ((Cid=" & Session("candidat_APPLICATION") & ") AND (GroupParent=0)) ORDER BY GroupID;"
    Set Group = Server.CreateObject("ADODB.Recordset")
    Group.Open SQL, My_Conn
    '=========================================
    Do Until Group.EOF
            '=================================================
            ' Chargement des permissions d'un groupe pour la catégorie demandée
            '=================================================
            SQL2 = "SELECT * FROM dbp_GroupsPermission"
            SQL2 = SQL2 & " WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.GroupId)=" & Group("GroupId") & ") AND ((dbp_GroupsPermission.ObjTypeId)=100) AND ((dbp_GroupsPermission.ObjId)=" & categoryId & "));"
            Set permaccess = Server.CreateObject("ADODB.Recordset")
            permaccess.Open SQL2, My_Conn
            
            '===============================================
            ' Chargement de la table des permissions et Mise à jour des Combos
            '===============================================
            If Not permaccess.EOF Then
                If permaccess("ObjPermType") > 0 Then
                    CreateCheckButton Group("GroupId"), Group("GroupDescr"), "Yes", "checked"
                Else
                    CreateCheckButton Group("GroupId"), Group("GroupDescr"), "No", ""
                End If
            Else
                    CreateCheckButton Group("GroupId"), Group("GroupDescr"), "No", ""
            End If
            Group.MoveNext
            Set permaccess = Nothing
    Loop
    
    Set Group = Nothing
    Set PermId = Nothing

    My_Conn.Close
    Set My_Conn = Nothing
    '===============
    EndForm2 ("ADD")

End Function
Public Function topicView(ByVal TopicID As Integer) As String

Response.Expires = -1440

CheckConnexion
'Filtrer un type de topic à afficher
Dim my_filter, my_id, my_father_rep, my_father_id, my_target, my_topicresume
my_filter = IIf(Len(Request.QueryString("filtre")) > 0, val(Request.QueryString("filtre")), 0)
pr ("<html>")
pr ("<head>")

If Request.QueryString("page") = "news" Then
    my_id = Request.QueryString("category")
    my_father my_id, my_father_rep, my_father_id
    If my_filter = 7 Then
        pr ("<link rel=STYLESHEET href=""PMainStyle1.asp?type=news"" type=""text/css"">")
        pr ("</head>")
        pr ("<script language=""JavaScript"">")
        pr ("if (navigator.appName == 'Netscape') {document.write(""<body  background='../portal_html" & IIf(my_father_rep <> "", my_father_rep, Request.QueryString("rep")) & "/bgnews7_ne.gif' topmargin='0' leftmargin='10'>"");}")
        pr ("else {document.write(""<body  background='../portal_html" & IIf(my_father_rep <> "", my_father_rep, Request.QueryString("rep")) & "/bgnews7.gif' topmargin='0' leftmargin='10'>"");}")
        pr ("</script>")
    Else
        pr ("<link rel=STYLESHEET href=""PMainStyle1.asp?type=news"" type=""text/css"">")
        pr ("</head>")
        pr ("<script language=""JavaScript"">")
        pr ("if (navigator.appName == 'Netscape') {document.write(""<body  background='../portal_html" & IIf(my_father_rep <> "", my_father_rep, Request.QueryString("rep")) & "/bgnews_ne.gif' topmargin='0' leftmargin='10'>"");}")
        pr ("else {document.write(""<body  background='../portal_html" & IIf(my_father_rep <> "", my_father_rep, Request.QueryString("rep")) & "/bgnews.gif' topmargin='0' leftmargin='10'>"");}")
        pr ("</script>")
    End If
Else
    pr ("<link rel=STYLESHEET href=""PMainStyle1.asp?type=topic"" type=""text/css"">")
    pr ("</head>")
    pr ("<body background=""Background.asp""  BOTTOMMARGIN=""100px"">")
End If

categoryId = IIf(Len(Request.QueryString("category")) = 0, -1, val(Request.QueryString("category")))
'CategoryPath = Request.QueryString("CatPath")
CATStyle = val(Request.QueryString("CatStyle"))
Order = Request.QueryString("Order")

If CATStyle = 0 Then CATStyle = 1 ' par sécurité
homepage = 0

' Si pas de CategoryID, c'est la page de une
If categoryId = 0 Then homepage = 1
If categoryId < 0 Then categoryId = 0

'============BANNIERE D'ACCUEIL==============================================
'Response.write "<p id=welcome> Welcome " & session("candidat_FULL_NAME") & ",<br> Your Personal News and main informations for the last " & datediff("d", session("candidat_LASTLOGIN"),now) & " day(s)"
'Response.Write "<p id=welcome> Your last Visit : " & session("candidat_LASTLOGIN")
'=============================================================================
If homepage = 0 Then
    '============VERFICATION DE L'APPARTENANCE AU GROUPE POUR ADD TOPICS======
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    Set RS2 = Server.CreateObject("ADODB.Recordset")
    
    SQL1 = "SELECT dbp_GroupsPermission.ObjId, dbp_GroupsPermission.GroupId "
    SQL1 = SQL1 + "FROM dbp_GroupsPermission INNER JOIN dbp_Groups ON dbp_GroupsPermission.GroupId = dbp_Groups.GroupId "
    SQL1 = SQL1 + "WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.ObjTypeId)=50) AND ((dbp_GroupsPermission.ObjId)=" & Session("candidat_USERId") & ") AND ((dbp_GroupsPermission.ObjPermType)=10)); "
        rs1.Open SQL1, My_Conn
        Do Until rs1.EOF
            Set RS2 = Server.CreateObject("ADODB.Recordset")
            SQL2 = "SELECT dbp_GroupsPermission.CId, dbp_GroupsPermission.GroupId, dbp_GroupsPermission.ObjTypeId, dbp_GroupsPermission.ObjId, dbp_GroupsPermission.ObjPermType "
            SQL2 = SQL2 + "FROM dbp_GroupsPermission INNER JOIN dbp_Categories ON dbp_GroupsPermission.ObjId = dbp_Categories.CatId "
            SQL2 = SQL2 + "WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.GroupId)=" & rs1("GroupId") & ") AND ((dbp_GroupsPermission.ObjTypeId)=100) AND ((dbp_GroupsPermission.ObjId)=" & categoryId & ") AND ((dbp_GroupsPermission.ObjPermType)>=20));"
            RS2.Open SQL2, My_Conn
            If Not RS2.EOF And YetTopic = 0 Then
                pr ("<welcome><center><a href='Portal.asp?mode=TopicsAddNew&category=" & categoryId & "'><img src='../Portal_Image/post-icon.gif'  border='0'></a><br>")
                pr ("<a href='Portal.asp?mode=TopicsAddNew&category=" & categoryId & "'>Post New</a></center>")
                YetTopic = 1
            End If
            RS2.Close
            Set RS2 = Nothing
        rs1.MoveNext
        Loop
    rs1.Close
    Set rs1 = Nothing

End If
If Len(Trim(Order)) = 0 Then Order = "ORDER BY type asc, topicstart desc, topicid"
If Request.QueryString("page") = "news" And my_filter = 0 And homepage = 0 Then

    Set rs1 = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT dbp_Objects.*, dbp_topics.size, dbp_Topics.cid, dbp_Topics.CatId, TopicId, dbp_Topics.Title, dbp_Topics.Type, dbp_Topics.URL, topicstyle, titlebackground, AbstractBackground,TopicExpires,TopicStart,NoTitle,NoSignature,TopicIcon "
    SQL = SQL + "FROM dbp_Categories INNER JOIN ((dbp_Topics INNER JOIN dbp_Objects ON dbp_Topics.TopicId = dbp_Objects.ObjId) INNER JOIN dbp_CatUserList ON dbp_Topics.CatId = dbp_CatUserList.CatId) ON (dbp_Categories.CatId = dbp_CatUserList.CatId) AND (dbp_Categories.CId = dbp_CatUserList.CId)"
    SQL = SQL + " WHERE (((dbp_Categories.Parent)=" & my_father_id & ") AND ((dbp_Topics.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_Objects.ObjTypeId)=101) AND ((dbp_CatUserList.UserId)=" & Session("candidat_USERId") & ") AND ((dbp_Categories.CatStyle)=3) AND ((dbp_Topics.Type)<>8))" & Order
    SQL = SQL + " UNION "
    SQL = SQL + "SELECT dbp_Objects.*, dbp_topics.size, dbp_Topics.cid, dbp_Topics.CatId, TopicId, dbp_Topics.Title, dbp_Topics.Type, dbp_Topics.URL, topicstyle, titlebackground, AbstractBackground,TopicExpires,TopicStart,NoTitle,NoSignature,TopicIcon "
    SQL = SQL + "FROM dbp_Categories INNER JOIN ((dbp_Topics INNER JOIN dbp_Objects ON dbp_Topics.TopicId = dbp_Objects.ObjId) INNER JOIN dbp_CatUserList ON dbp_Topics.CatId = dbp_CatUserList.CatId) ON (dbp_Categories.CatId = dbp_CatUserList.CatId) AND (dbp_Categories.CId = dbp_CatUserList.CId)"
    SQL = SQL + " WHERE (((dbp_Categories.catid)=" & categoryId & ") AND ((dbp_Topics.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_Objects.ObjTypeId)=101) AND ((dbp_CatUserList.UserId)=" & Session("candidat_USERId") & ") AND ((dbp_Topics.Type)=8))" & Order
    'Debug.Print SQL
Else

    'Recherche des topics à afficher en croisant avec les droits sur la catégorie
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT dbp_Topics.CatId, * "
    'SQL = SQL + "FROM ((dbp_Topics INNER JOIN dbp_Objects ON dbp_Topics.TopicId = dbp_Objects.ObjId) INNER JOIN dbp_Users ON dbp_Objects.ObjOwner = dbp_Users.UserId) INNER JOIN dbp_Categories ON dbp_Topics.CatId = dbp_Categories.CatId "
    SQL = SQL + "FROM (dbp_Topics INNER JOIN dbp_Objects ON dbp_Topics.TopicId = dbp_Objects.ObjId) INNER JOIN dbp_CatUserList ON dbp_Topics.CatId = dbp_CatUserList.CatId "
    
    
    If homepage = 0 Then
        'SQL = SQL + "WHERE (((dbp_Topics.CId)=" & session("candidat_APPLICATION") & ") AND ((dbp_Objects.ObjTypeId)=101) AND ((dbp_Topics.CatId)=" & categoryId & ")) ORDER BY TopicID " & Order
        SQL = SQL + "WHERE (((dbp_Topics.CatId)=" & categoryId & ") AND ((dbp_Topics.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_Objects.ObjTypeId)=101) AND ((dbp_CatUserList.UserId)=" & Session("candidat_USERId") & ")) " & Order
    Else
        'SQL = SQL + "WHERE (((dbp_Topics.CId)=" & session("candidat_APPLICATION") & ") AND ((dbp_Objects.ObjTypeId)=101)) ORDER BY TopicID " & Order
        SQL = SQL + "WHERE (((dbp_Topics.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_Objects.ObjTypeId)=101) AND ((dbp_CatUserList.UserId)=" & Session("candidat_USERId") & ")) " & Order
    End If
End If
'Debug.Print SQL
rs1.Open SQL, My_Conn
' ========== SI LA CATEGORIE EST VIDE ON AFFICHE LA BANNIERE SEULE =================================
If rs1.EOF Then
    pr ("<br><p id=desc>&nbsp;</p>")
Else
    pr ("<TABLE WIDTH='100%' BORDER='0'><tr>")
    If Request.QueryString("page") = "news" And rs1("type") <> 8 And my_filter = 0 Then
        pr ("<td><img src=""../portal_image/news.gif""></td></tr><tr>")
    End If
    prems = True
    Do While Not rs1.EOF
        If Request.QueryString("page") = "news" And my_filter = 0 And rs1("type") = 8 And prems And homepage = 0 Then
            pr ("</tr><tr><td><img src=""../portal_image/related.gif""></td></tr><tr>")
            prems = False
        End If
        zap = False
        If Session("candidat_ADMIN") <> 1 And Request.QueryString("page") = "news" Then
            If Len(Trim(rs1("topicstart"))) > 0 Then
                If rs1("topicstart") > Date Then
                    zap = True
                Else
                    If Len(Trim(rs1("topicexpires"))) > 0 Then
                        If rs1("topicexpires") < Date Then
                            zap = True
                        End If
                    End If
                End If
            End If
        End If
        If Not zap Then
            If homepage = 1 Then
                If my_filter > 0 Then
                    If rs1("type") <> my_filter Then zap = True
                Else
                    If rs1("Type") <> 2 Or DateDiff("s", Session("candidat_LASTLOGIN"), rs1("ObjLastChange")) = 0 Then zap = True 'Exit Do
                End If
            End If
            If rs1("type") = 8 And Session("candidat_ADMIN") <> 1 And Request.QueryString("page") <> "news" Then zap = True
            If Not zap Then
                If Request.QueryString("page") = "news" Then
                    
                    my_id = rs1("catid")
                    my_father my_id, my_father_rep, my_father_id
                    If Trim(rs1("URL")) <> "" Then
                        NewUrl = rs1("URL")
                    Else
                        NewUrl = "../portal_html/default.asp?mode=cat&catid=" & my_father_id & "&rep=" & my_father_rep & "&rep_haut=" & my_father_rep & "haut/&prev=home&current=cat&quickurl=" & rs1("catid")
                    End If
                    my_target = "target=""_top"""
                    'NewUrl = "./default.asp?mode=cat&catid=" & rs3("catid") & "&rep=" & rs3("catrep") & "&rep_haut=" & rs3("catrep") & "haut/&prev=home&current=cat&quickurl=" & rs1("catid")

                Else
                    NewUrl = rs1("URL")
                End If
                NewUrl = Replace(NewUrl, "%20", "+", 1, -1, 1)
                ItemId = ItemId + 1
                Select Case rs1("Type")
                    Case 1 'text
                    icone = "news.gif"
                    Case 2 'news
                    icone = "news.gif"
                    Case 3 'Document attached
                    icone = "document.gif"
                    Case 4 'Application
                    icone = "application.gif "
                    Case 5 'Url
                    icone = "link.gif"
                    Case 6 'Dynamic Url
                    icone = "go.gif"
                    Case 7 'Colonne droite de la page d'accueil
                    icone = IIf((Not IsNull(rs1("topicicon")) And rs1("topicicon") <> ""), rs1("topicicon") & ".gif", "")
                End Select
                BytesToTranfer = ""
                If rs1("Type") = 3 Then BytesToTranfer = " (" & rs1("Size") & " Bytes)"
                'Response.Write session("candidat_Admin") & rs1("ObjOwner") & session("candidat_USERId")
                If rs1("ObjOwner") = Session("candidat_USERId") Or Session("candidat_Admin") = 1 Then
                    Tedit = "<a href='../portal_asp/Portal.asp?Mode=TopicsEdit&Topics=" & rs1("TopicId") & "'><img src='../portal_image/update.gif' border=0></a>"
                    Tapprove = "<a href=../portal_asp/Portal.asp?Mode=TopicsEdit&Topics=" & rs1("TopicId") & ">Approve</a>"
                End If
                
                AllStyle = rs1("TopicStyle")
                If rs1("NoTitle") = True Then
                    Title = ""
                Else
                    Title = rs1("Title")
                End If
                If rs1("NoSignature") = True Then
                    MyAuthor = ""
                    MyEmail = ""
                Else
                    MyAuthor = rs1("UserDisplayName")
                    MyEmail = rs1("UserEmail")
                End If
                MyDate = rs1("topicstart")
                Session("candidat_BgPixTitle") = Trim(rs1("TitleBackground"))
                Session("candidat_BgPixAbstract") = rs1("AbstractBackground")
                ' on sépare la page en Catstyle colonnes égales
                
                col = col + 1
                MyTdWidth = Int(100)
                'MyTdWidth = Int(100 / catstyle)
                pr ("<td width='" & MyTdWidth & "%' valign='top'>")
                pr ("<table border='0'><tr><td>")
                If Request.QueryString("page") = "news" Then 'colonnes de droite
                    ViewNewsStyle Title & BytesToTranfer, NewUrl, "", "", "", "", icone, "", "", AllStyle, my_target
                Else
                    my_topicresume = rs1("Abstract")
                    ViewNewsStyle Title & BytesToTranfer, NewUrl, MyDate, my_topicresume, MyAuthor, MyEmail, icone, Tedit, Tapprove, AllStyle, ""
                End If
                pr ("</td></tr></table>")
                pr ("</td>")
                'CatStyle = CatStyle * 1
                'If CatStyle = col Then
                If 1 = col Then
                    col = 0
                    pr ("</tr><tr>")
                End If
                Session("candidat_BgPixTitle") = ""
                Session("candidat_BgPixAbstract") = ""
            End If
        End If
        If Not rs1.EOF Then rs1.MoveNext
    Loop
    pr ("</tr></table>")
End If
rs1.Close
Set rs1 = Nothing
My_Conn.Close

pr ("</body>")
pr ("</html>")
End Function
Public Function topicEdition(ByVal TopicID As Integer) As String
    
    CheckConnexion
    
    Topic = Request.QueryString("Topics")
    
    SQL = "select * from dbp_Topics  where (  (CId=" & Session("candidat_APPLICATION") & ") AND  (TopicId=" & Topic & ") );"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    
    rs1.Open SQL, My_Conn
    
    FormCaption = rs1("Title") + ": (edition)"
    
    InitForm "ADM_TOPIC_EDIT", "Portal.ASP?mode=TopicsUpdate&Topics=" & rs1("TopicID")

    ' Variables Sessions
    
    CreateField "TOPid", "Topic ID", Topic
    CreateField "TOPTITLE", "Title", rs1("Title")
    CreateTextArea "TOPABSTRACT", "Contain", rs1("Abstract"), 25000, 80, 20
    'CreateHtmlEditor "TOPABSTRACT", "Contain", rs1("Abstract"), "ADM_TOPIC_EDIT"
    CreateField "START", "Appears from", rs1("topicstart")
    CreateField "EXPIRE", "To (include)", rs1("topicexpires")
    InitSelectList "TOPTYPE", "Category", rs1("Type"), 1

    selectedoption = rs1("Type")


    
    AddSelectList 1, "News", selectedoption
    AddSelectList 2, "News for Home Page", selectedoption
    AddSelectList 3, "Document Attach", selectedoption
    AddSelectList 7, "Right column for Home Page", selectedoption
    AddSelectList 4, "Application", selectedoption
    AddSelectList 5, "w/HTML", selectedoption
    AddSelectList 6, "Dynamic Url", selectedoption
    AddSelectList 8, "Related link for this category", selectedoption
    EndSelectList
    CreateField "TOPURL", "URL", rs1("Url")
    CreateField "topicIcon", "Icon for right column", rs1("topicIcon")
    PreviousStyle = rs1("TopicStyle")
    If rs1("NoTitle") = True Then notitle = "checked"
    If rs1("NoSignature") = True Then nosignature = "checked"
    CreateField "TITLEBG", "Title Background :", rs1("TitleBackground")
    CreateField "ABSTRACTBG", "Topic Background :", rs1("AbstractBackground")
    rs1.Close
    Set rs1 = Nothing
    InitSelectList "TOPSTYLE", "Style set for this topic", PreviousStyle, 1
        AddSelectList 0, "Default", PreviousStyle
        AddSelectList 3, "News", PreviousStyle
    EndSelectList
    
    CreateCheckButton "NoTitle", "Remove Title", "yes", notitle
    CreateCheckButton "NoSignature", "Remove Signature", "yes", nosignature
    CreateLink "Delete this Topic", "", "../Portal_ASP/Portal.asp?mode=TopicsDelete&Topics=" & Topic
    Response.Write "<br>"
    EndForm2 ("UPDATE")
    
End Function
Public Function topicUpdate(ByVal TopicID As Integer) As Boolean
    Dim typeurl, my_url, my_start, my_expire
    CheckConnexion
    
    If Len(Trim(Request.Form("START"))) > 0 Then
        my_start = Format(Request.Form("START"), "dd/mm/yyyy")
    End If
    If Len(Trim(Request.Form("EXPIRE"))) > 0 Then
        my_expire = Format(Request.Form("EXPIRE"), "dd/mm/yyyy")
    End If
If Request("topics") = "new" Then
    
    categoryId = Request.QueryString("category")
    
    '=============================================================
    ' Calcul de TopicId
    '=============================================================

    SQL = "select * from dbp_Topics where CId=" & Session("candidat_APPLICATION") & " order by TopicId"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    

    rs1.Open SQL, My_Conn
    
    Do Until rs1.EOF
        NewTopicId = rs1("TopicID")
        rs1.MoveNext
    Loop
    
    NewTopicId = NewTopicId + 1
    rs1.Close
    Set rs1 = Nothing
    
    If InStr(Request.Form("TOPABSTRACT"), "<") > 0 And InStr(Request.Form("TOPABSTRACT"), ">") > 0 Then
            
            Title = ChkString(Request.Form("TOPTITLE"))
            abstract = ChkString(Request.Form("TOPABSTRACT"))
    Else
            Title = ChkString(Server.HTMLEncode(Request.Form("TOPTITLE")))
            abstract = ChkString(Server.HTMLEncode(Request.Form("TOPABSTRACT")))
    End If

    If Request.Form("NoTitle") = "" Then
        notitle = "no"
    Else
        notitle = "yes"
    End If
    If Request.Form("NoSignature") = "" Then
        nosignature = "no"
    Else
        nosignature = "yes"
    End If
    
    'RECHERCHE du lien automatique
    typeurl = val(Request.Form("TYPEURL"))
    If typeurl > 0 Then
        i = 1
        sortie = False
         'Recherche des events
         SQL = "select event_id from dbe_events order by event_title"
         Set rs1 = Server.CreateObject("ADODB.Recordset")
         rs1.Open SQL, My_Conn
         Do Until rs1.EOF
             If i = typeurl Then
                 my_url = Session("candidat_URLBASE") & "/portal_events/default.asp?event_id=" & rs1("event_id")
                 sortie = True
                 Exit Do
             End If
             i = i + 1
          rs1.MoveNext
         Loop
         rs1.Close
         Set rs1 = Nothing
         If Not sortie Then
             'Recherche des sondages
             SQL = "select id from dbpoll_poll order by pollname"
             Set rs1 = Server.CreateObject("ADODB.Recordset")
             rs1.Open SQL, My_Conn
             Do Until rs1.EOF
                 If i = typeurl Then
                     my_url = Session("candidat_URLBASE") & "/portal_survey/poll.asp?id=" & rs1("id")
                     sortie = True
                     Exit Do
                 End If
                 i = i + 1
              rs1.MoveNext
             Loop
             rs1.Close
             Set rs1 = Nothing
         End If
    End If

    ' INSERTION CATEGORY dbp_TOPIC

    SQL = "Insert into dbp_Topics (CId, TopicId, CatId, Title, Abstract, Type, Data, Url, ContentType, TopicStyle,NoTitle,NoSignature,AbstractBackground,TitleBackground,topicicon,topicstart,topicexpires) "
    SQL = SQL & " Values (" & Session("candidat_APPLICATION") & "," & NewTopicId & "," & categoryId & ",'" & Title & "','"
    SQL = SQL & abstract & "'," & Request.Form("TOPTYPE") & ", 'TEXTE'  ,'" & IIf(typeurl = 0, Request.Form("TOPURL"), my_url) & "','-'," & Request.Form("TOPSTYLE") & "," & notitle & "," & nosignature & ","
    SQL = SQL & "'" & Request.Form("ABSTRACTBG") & "','" & Request.Form("TITLEBG") & "','" & Request.Form("topicicon") & "',"
    SQL = SQL & Replace("'" & my_start & "', '" & my_expire & "' );", "''", "Null")

    My_Conn.Execute (SQL)
    
    SQL = "INSERT INTO dbp_Objects ( CId, ObjTypeId, ObjId, ObjCreated, ObjOwner, ObjLastChange )"
    SQL = SQL & " Values (" & Session("candidat_APPLICATION") & ",101," & NewTopicId & ",'" & Now() & "'," & Session("candidat_USERId") & ",'" & Now() & "');"

    My_Conn.Execute (SQL)

    If Request.Form("TOPTYPE") = 3 Then
        Response.Redirect ("Portal_file_upload.asp?topic=" & Request.Form("TOPid"))
    End If
    Response.Redirect ("../Portal_asp/result.asp?message=Topics+" & Request.Form("TOPTITLE") & "+added")

Else

    'if Request.Form("TOPTYPE")>4 and Request.Form("TOPTYPE")<8 then
    
    If InStr(Request.Form("TOPABSTRACT"), "<") > 0 And InStr(Request.Form("TOPABSTRACT"), ">") > 0 Then
            Title = ChkString(Request.Form("TOPTITLE"))
            abstract = ChkString(Request.Form("TOPABSTRACT"))
    Else
            Title = ChkString(Server.HTMLEncode(Request.Form("TOPTITLE")))
            abstract = ChkString(Server.HTMLEncode(Request.Form("TOPABSTRACT")))
    End If
    If Request.Form("NoTitle") = "" Then
        notitle = "no"
    Else
        notitle = "yes"
    End If
    If Request.Form("NoSignature") = "" Then
        nosignature = "no"
    Else
        nosignature = "yes"
    End If
        
    SQL = "UPDATE dbp_Topics  SET dbp_Topics.Title ='" & Title & "', dbp_Topics.Abstract ='" & abstract
    SQL = SQL & "', dbp_Topics.Type = " & Request.Form("TOPTYPE") & " , dbp_Topics.URL =' " & Request.Form("TOPURL")
    SQL = SQL & "', dbp_Topics.TopicStyle=" & Request.Form("TOPSTYLE") & ", dbp_Topics.Notitle=" & notitle & ", dbp_Topics.NoSignature=" & nosignature & ", "
    SQL = SQL & " dbp_Topics.TitleBackground='" & Request.Form("TITLEBG") & "', dbp_Topics.AbstractBackground='" & Request.Form("ABSTRACTBG") & "', dbp_topics.topicicon='" & Request.Form("topicicon") & "', "
    SQL = SQL & " dbp_Topics.topicstart=" & Replace("'" & my_start & "'", "''", "NULL") & ", dbp_Topics.topicexpires=" & Replace("'" & my_expire & "'", "''", "NULL") & ""
    SQL = SQL & " WHERE ((dbp_Topics.TopicId=" & Request.Form("TOPId") & ") AND (dbp_Topics.CId=" & Session("candidat_APPLICATION") & "))"
    
    ' MsgBox (SQL)
    ' UPDATE TOPICS  dbp_Topics
    
    My_Conn.Execute (SQL)

    SQL = "UPDATE  dbp_Objects  SET ObjLastChange='" & Now()
    SQL = SQL & "'  WHERE ( (CId =" & Session("candidat_APPLICATION") & ")  AND (ObjTypeId=101) AND (ObjId=" & Request.Form("TOPId") & "));"

    My_Conn.Execute (SQL)
    
    If Request.Form("TOPTYPE") = 3 Then
        Response.Redirect ("Portal_file_upload.asp?topic=" & Request.Form("TOPid"))
    End If

    Response.Redirect ("../Portal_asp/result.asp?message=Topics+successfully+updated")

End If
End Function
Public Function topicDelete(ByVal TopicID As Integer) As Boolean
    CheckConnexion
    
    Topics = Request.QueryString("Topics")
    
    SQL = "Delete * From dbp_Topics  WHERE  ((dbp_Topics.TopicId=" & Topics & ") AND (dbp_Topics.CId=" & Session("candidat_APPLICATION") & "))"

    My_Conn.Execute (SQL)
    
    SQL = "Delete * From dbp_Objects WHERE  ( ( dbp_Objects.ObjTypeId=101) AND ( dbp_Objects.ObjId=" & Topics & ") AND ( dbp_Objects.CId=" & Session("candidat_APPLICATION") & "))"

    My_Conn.Execute (SQL)

    Response.Redirect ("../Portal_asp/result.asp?message=Topics+successfully+deleted")

End Function
Public Function userAddNew() As String
    
    CheckConnexion
    CheckAdminRights
    
    SQL = "select * from dbp_Users where CId=" & Session("candidat_APPLICATION") & " order by UserId"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    
    rs1.Open SQL, My_Conn
    
    Do Until rs1.EOF
        NewUserId = rs1("UserId")
        rs1.MoveNext
    Loop
    NewUserId = NewUserId + 1
    rs1.Close
    Set rs1 = Nothing

'*************************************************************
' Formulaire de saisie
'*************************************************************
    FormCaption = "Add User  to Portal"
    FlagHTMLEditor = False
    pr ("<HTML>")
    pr ("<HEAD>")
    pr ("<link rel=STYLESHEET href='PMainStyle1.asp' type='text/css'>")
    pr ("</HEAD>")
    Dim my_directory
    Set my_directory = Server.CreateObject("CP8Portal.directory")
    my_directory.DirectoryMenu
    Set my_directory = Nothing
    pr ("<b><p id='dheading'>" & FormCaption & "</b></font></td>")
    pr ("</center>")
    pr ("<form action='Portal.ASP?mode=UserUpdate&User=0' method='post' name='ADM_USER_ADD_REP'>")

' Variables Session

    CId = Session("candidat_APPLICATION")

    CreateField "USERid", "User ID", NewUserId
    CreateField "LOGIN", "Login (required)", ""
    CreateCheckButton "ExternIdent", "External Identification (no control on password)", "extern", ""
    CreateField "helppw", "Suggestion for Password", generate_pw
    CreateMaskField "PASSW", "PassWord (8 char min)", ""
    CreateMaskField "PASSW1", "PassWord Again", ""
    CreateField "FIRST", "First Name (required)", ""
    CreateField "LAST", "Last Name (required)", ""
    CreateField "DISPNAME", "Display Name (required)", ""
    CreateField "EMAIL", "E-MAIL (required)", ""
    InitSelectList "LANG", "Language", 1, 1
    AddSelectList 1, "Francais", 1
    AddSelectList 2, "English", 0
    EndSelectList
    ' CreateCheckButton "CreateML","Mailing List :","1",""
    CreateCheckButton "contact", "Answer permission for Contact Us", "1", ""
    EndForm2 ("ADD")

End Function
Public Function userUpdate(ByVal UserID As Integer) As Boolean
Dim my_lang
User = Request.QueryString("User")
mode = Request.QueryString("Mode")
modeuser = Request.QueryString("ModeUser")

CheckConnexion
CheckAdminRights

' enregistrement d'un nouvel utilisateur

If User = 0 Then
    
    SQL = "select * from dbp_Users where CId=" & Session("candidat_APPLICATION") & " order by UserId"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    
    rs1.Open SQL, My_Conn
    
    Do Until rs1.EOF
        NewUserId = rs1("UserId")
        rs1.MoveNext
    Loop
    NewUserId = NewUserId + 1
    rs1.Close
    Set rs1 = Nothing
    
    errorpasw = 0
    
    If Request.Form("ExternIdent") = 0 Then ' si l'identification est externe on n'effectue pas de contrôle

        If Request.Form("PASSW") <> Request.Form("PASSW1") Then
                Response.Write "<p id=alarm>Error in password : Password and Password check are different !" & "<br><p>"
                errorpasw = 1
        End If
        
        If Len(Request.Form("PASSW")) < 8 Then
                Response.Write "<p id=alarm>Error in password : Password Must be 8 char Length !" & "<br><p>"
                errorpasw = 2
        End If
        
        If errorpasw > O Then
                Response.Write "<br> There are errors in password definitions. Please press BACK button  to correct"
                Response.End
        End If

        ' Hash Password
        HashIt (Request.Form("PASSW"))
    Else
        HashIt ("NONE") ' on hash n'importe quoi pour créer un password de sécurité
    End If

    Select Case Request.Form("LANG")
    
        Case 1
            my_lang = "FR"
        Case 2
            my_lang = "GB"
    End Select
    ' Insertion dans dbp_Users
    
    SQL = "Insert into dbp_Users (CId, UserID, UserLogin, UserPass, UserSecure, UserFirst, UserLast,UserDisplayName,UserEmail,UserServer, userlang)"
    SQL = SQL & " Values (" & Session("candidat_APPLICATION") & ", " & NewUserId & ",'" & Request.Form("LOGIN") & "','" & Session("candidat_HashData") & "','" & Session("candidat_HashSecure") & "', "
    SQL = SQL & "'" & Request.Form("FIRST") & "','" & Request.Form("LAST") & "','" & Request.Form("DISPNAME") & "','" & Request.Form("EMAIL") & "','" & Request.Form("ExternIdent") & "','" & my_lang & "');"
    'Response.Write SQL
    My_Conn.Execute (SQL)
    Session("candidat_HashData") = ""
    Session("candidat_HashSecure") = ""
    ' Insertion dans dbp_UserInfos
    
    SQL = "Insert into dbp_UserInfos (CID,UserId,FirstName,LastName,username) Values (" & Session("candidat_APPLICATION") & ","
    SQL = SQL & NewUserId & ",'" & Request.Form("FIRST") & "','" & Request.Form("LAST") & "','" & Request.Form("LOGIN") & "')"
    My_Conn.Execute (SQL)
    
    If InStr(1, UCase(Request.Form("LOGIN")), "GUEST") = 0 Then ' ===== Insertion dans contact
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "select users from db_keys", My_Conn
        If Not rs.EOF Then
            my_sid = rs("users")
        Else
            my_sid = 1
        End If
        rs.Close
        Set rs = Nothing
        My_Conn.Execute ("update db_keys set users=" & my_sid + 1)
        
        SQL = "insert into tblusers  (sid,uid,fname,email1,department,zone_id,isrep,dtcreated) values("
        SQL = SQL & my_sid & ", "
        SQL = SQL & "'" & Request.Form("LOGIN") & "', "
        SQL = SQL & "'" & Request.Form("FIRST") & " " & Request.Form("LAST") & "', "
        SQL = SQL & "'" & Request.Form("EMAIL") & "', "
        SQL = SQL & "0,0, "
        If Request.Form("contact") = "1" Then
            SQL = SQL & "1, "
        Else
            SQL = SQL & "0, "
        End If
        SQL = SQL & "'" & Now() & "');"
        My_Conn.Execute (SQL)
        
    End If
    
    My_Conn.Close
    Set My_Conn = Nothing
        
    Response.Redirect ("Portal.ASP?mode=GroupPermUpdate&user=" & Request.Form("USERid"))

End If

'==========================================================
' MODE EDITION SIMPLE
'==========================================================
If modeuser = "" Then
        Set cmdcheckuser = Server.CreateObject("ADODB.Recordset")
        ' ============== SQL pour la connexion =========================
        SQL = "SELECT * FROM dbp_Users "
        SQL = SQL & "WHERE  dbp_Users.UserId=" & User & " AND CId=" & Session("candidat_APPLICATION")
        cmdcheckuser.Open SQL, My_Conn
'*************************************************************
' Formulaire de saisie
'*************************************************************

        
        FormCaption = "Edit User " & cmdcheckuser("UserDisplayName")
        FlagHTMLEditor = False
        pr ("<HTML>")
        pr ("<HEAD>")
        pr ("<link rel=STYLESHEET href='PMainStyle1.asp' type='text/css'>")
        pr ("</HEAD>")
        Dim my_directory
        Set my_directory = Server.CreateObject("CP8Portal.directory")
        my_directory.DirectoryMenu
        Set my_directory = Nothing
        pr ("<b><p id='dheading'>" & FormCaption & "</b></font></td>")
        pr ("</center>")
        pr ("<form action='Portal.ASP?mode=UserUpdate&User=" & User & "&ModeUser=UPDATE' method='post' name='ADM_USER_EDIT'>")
        ' Variables Session
        CId = Session("candidat_APPLICATION")
        CreateField "USERid", "User ID", cmdcheckuser("UserId")
        CreateField "LOGIN", "Login", cmdcheckuser("UserLogin")
        If cmdcheckuser("UserServer") = "1" Then
            ischecked = "Checked"
        Else
            ischecked = ""
        End If
        CreateCheckButton "ExternIdent", "External Identification :", "1", ischecked
        CreateMaskField "PASSW", "PassWord", "********"
        CreateMaskField "PASSW1", "PassWord Again", "********"
        CreateField "FIRST", "First Name", cmdcheckuser("UserFirst")
        CreateField "LAST", "Last Name", cmdcheckuser("UserLast")
        CreateField "DISPNAME", "Display Name", cmdcheckuser("UserDisplayName")
        CreateField "EMAIL", "E-MAIL", cmdcheckuser("UserEMail")
       
        InitSelectList "LANG", "Language", 1, 1
            AddSelectList 1, "Francais", IIf(cmdcheckuser("USERLANG") = "FR", 1, 0)
            AddSelectList 2, "English", IIf(cmdcheckuser("USERLANG") = "GB", 2, 0)
        EndSelectList
        CreateLink "Other Informations", LinkIcon, "../portal_members/aspintranet.asp?mode=emp_frmemp&UserID=" & User
        CreateLink "Delete User", LinkIcon, "Portal.ASP?mode=UserUpdate&User=" & User & "&ModeUser=DELETE"
            
        Set rs = Server.CreateObject("ADODB.Recordset")
        SQL = "SELECT sid FROM tblusers "
        SQL = SQL & "WHERE  UId='" & cmdcheckuser("userlogin") & "' AND isrep=1"
        rs.Open SQL, My_Conn
        If Not rs.EOF Then
            CreateCheckButton "contact", "Répond aux questions du Contact Us :", "1", "Checked"
        Else
            CreateCheckButton "contact", "Répond aux questions du Contact Us :", "1", ""
        End If
        rs.Close
        Set rs = Nothing
        Response.Write ("<br><hr>")
        EndForm2 ("Update")
        Set cmdcheckuser = Nothing
End If

'==========================================================
' MODE UPDATE
'==========================================================

If modeuser = "UPDATE" Then
    
    If Request.Form("PASSW") <> "********" And Request.Form("ExternIdent") = 0 Then ' le password a été modifié
            errorpasw = 0
            If Request.Form("PASSW") <> Request.Form("PASSW1") Then
                Response.Write "<p id=alarm>Error in password : Password and Password check are different !" & "<br><p>"
                errorpasw = 1
            End If
            If Len(Request.Form("PASSW")) < 8 Then
                Response.Write "<p id=alarm>Error in password : Password Must be 8 char Length !" & "<br><p>"
                errorpasw = 2
            End If
            If errorpasw > O Then
                Response.Write "<br> There are errors in password definitions. Please press BACK button  to correct"
                Response.End
            End If
            ' Hash Password
            Session("candidat_HashData") = ""
            Session("candidat_HashSecure") = ""
            HashIt (Request.Form("PASSW"))
    
    End If
    
    
    '=====UPDATE DANS USER
    SQL = "UPDATE dbp_Users  SET "
    SQL = SQL & "UserLogin ='" & Server.HTMLEncode(Request.Form("LOGIN")) & "',"
    SQL = SQL & "UserFirst ='" & Server.HTMLEncode(Request.Form("FIRST")) & "',"
    SQL = SQL & "UserLast ='" & Server.HTMLEncode(Request.Form("LAST")) & "',"
    SQL = SQL & "UserDisplayName ='" & Server.HTMLEncode(Request.Form("DISPNAME")) & "',"
    SQL = SQL & "UserEMail ='" & Server.HTMLEncode(Request.Form("EMAIL")) & "', "
    SQL = SQL & "UserServer ='" & Request.Form("ExternIdent") & "', "
    Select Case Request.Form("LANG")
    Case 1
            SQL = SQL & "USERLANG ='FR'"
    Case 2
            SQL = SQL & "USERLANG ='GB'"
    End Select

    If Session("candidat_HashData") <> "" And Session("candidat_HashSecure") <> "" Then
            SQL = SQL & ", UserPass ='" & Server.HTMLEncode(Session("candidat_HashData")) & "'"
            SQL = SQL & ", UserSecure ='" & Server.HTMLEncode(Session("candidat_HashSecure")) & "' "
    End If
    'response.write SQL
    SQL = SQL & "WHERE  (UserId=" & User & ") AND CId=" & Session("candidat_APPLICATION")
    My_Conn.Execute (SQL)
    
    ' ===== update dans dbp_userinfo
    SQL = "UPDATE dbp_Userinfos  SET "
    SQL = SQL & "Username ='" & Server.HTMLEncode(Request.Form("LOGIN")) & "',"
    SQL = SQL & "Firstname ='" & Server.HTMLEncode(Request.Form("FIRST")) & "',"
    SQL = SQL & "Lastname ='" & Server.HTMLEncode(Request.Form("LAST")) & "',"
    SQL = SQL & "EMail ='" & Server.HTMLEncode(Request.Form("EMAIL")) & "' "
    SQL = SQL & "WHERE  (UserId=" & User & ") AND CId=" & Session("candidat_APPLICATION")
    My_Conn.Execute (SQL)
    
    If Not InStr(1, UCase(Request.Form("LOGIN")), "GUEST") > 0 Then
        ' ===== update dans tblusers
        SQL = "UPDATE tblusers  SET "
        SQL = SQL & "fname ='" & Request.Form("FIRST") & Request.Form("LAST") & "',"
        SQL = SQL & "EMail1 ='" & Request.Form("EMAIL") & "' "
        If Request.Form("contact") = "1" Then
            SQL = SQL & ", isrep =1 "
        Else
            SQL = SQL & ", isrep =0 "
        End If
        SQL = SQL & "WHERE  (uid ='" & Request.Form("LOGIN") & "')"
        My_Conn.Execute (SQL)
    End If
    '======CHECK DES GROUPES
    '=========================================
    ' Chargement de la liste des groupes issues de dbp_Groups
    '=========================================
    SQL = "SELECT * FROM dbp_Groups WHERE (Cid=" & Session("candidat_APPLICATION") & ") ORDER BY GroupID;"
    Set Group = Server.CreateObject("ADODB.Recordset")
    Group.Open SQL, My_Conn
    '=========================================
    InitForm "ADM_GRP_AUT_REP", "Portal.asp?mode=UserUpdate&modeUser=PERMISSION&user=" & User
    Do Until Group.EOF
            InitSelectList "G" & Group("GroupId"), Group("GroupDescr"), "", 1
            '=================================================
            ' Chargement des permissions d'un groupe pour la catégorie demandée
            '=================================================
            SQL2 = "SELECT * FROM dbp_GroupsPermission"
            SQL2 = SQL2 & " WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.GroupId)=" & Group("GroupId") & ") AND ((dbp_GroupsPermission.ObjTypeId)=50) AND ((dbp_GroupsPermission.ObjId)=" & User & "));"
            Set permaccess = Server.CreateObject("ADODB.Recordset")
            permaccess.Open SQL2, My_Conn
            ' Response.Write SQL2
            '===============================================
            ' Chargement de la table des permissions et Mise à jour des Combos
            '===============================================
            SQL1 = "SELECT * FROM dbp_PermsTypes WHERE (ObjTypeid=50) ORDER BY ObjPermType ;"
            Set PermId = Server.CreateObject("ADODB.Recordset")
            PermId.Open SQL1, My_Conn
            Do Until PermId.EOF
                    If permaccess.EOF Or permaccess.BOF Then                                                                                '=============
                            AddSelectList PermId("ObjPermType"), PermId("ObjPermDescr"), 0                       ' pas d'accès défini
                    Else
                            AddSelectList PermId("ObjPermType"), PermId("ObjPermDescr"), permaccess("ObjPermType")  ' On liste l'accès
                    End If
                    PermId.MoveNext
            Loop
            EndSelectList
            Set PermId = Nothing
        Group.MoveNext
    Loop
    EndForm2 ("Update")
    Set Group = Nothing
    Set PermId = Nothing
    Set permaccess = Nothing
    My_Conn.Close
    Set My_Conn = Nothing
End If

'==========================================================
' MODE UPDATE PERMISSION
'==========================================================

If modeuser = "PERMISSION" Then
    
    User = Request.QueryString("user")

    ' Pour l'Update , on efface tous les droits sur la catégorie et on réecrit par sécurité
    
    'Effacement dans GroupPerm
    SQL = "Delete * From dbp_GroupsPermission"
    SQL = SQL & " WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND  ((dbp_GroupsPermission.ObjTypeId)=50) AND ((dbp_GroupsPermission.ObjId)=" & User & "));"
    My_Conn.Execute (SQL)
    
    'Effacement dans CatUserPref
    SQL = "Delete * FROM dbp_CatUserList WHERE (((dbp_CatUserList.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_CatUserList.UserId)=" & User & "));"
    My_Conn.Execute (SQL)

    ' On ouvre la Table dbp_Groups et pour chacun des groupes on écrit la permission issue de la request.form
    ' le champ est écrit dans request.form(group("GroupId"))

    SQL = "SELECT * FROM dbp_Groups WHERE (Cid=1) ORDER BY GroupID;"
    
    Set Group = Server.CreateObject("ADODB.Recordset")
    
    Group.Open SQL, My_Conn
    
    Do Until Group.EOF
        
        ' Insertion des permissions groupes dans GroupPerm
        
        SQL = "INSERT INTO dbp_GroupsPermission (Cid, GroupId, ObjTypeId, ObjPermId, ObjId, ObjPermType)"
        SQL = SQL & " Values (" & Session("candidat_APPLICATION") & "," & Group("GroupId") & ",50,1," & User & "," & Request.Form("G" & Group("GroupId")) & ")"
        'Response.write SQL
        My_Conn.Execute (SQL)
            
        ' Insertion des UsersPréferences pour chacun des users du groupe
        ' On liste les users du groupe par recherche de la valeur PermType 10 pour TypeId = 50
        
        Set UserGroup = Server.CreateObject("ADODB.Recordset")
        Set RsCheck = Server.CreateObject("ADODB.Recordset")
        
        SQL = "SELECT dbp_GroupsPermission.ObjId FROM dbp_GroupsPermission "
        SQL = SQL & "WHERE (((dbp_GroupsPermission.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_GroupsPermission.GroupId)=" & Group("GroupId") & ") AND ((dbp_GroupsPermission.ObjTypeId)=100));"
        
        UserGroup.Open SQL, My_Conn
            Do Until UserGroup.EOF
                ' On vérifie que la catégorie n'est pas déjà existante dans CatUserPref pour le Group et le User assigné
                SQLCheck = "SELECT * from dbp_CatUserList WHERE ( (Cid=" & Session("candidat_APPLICATION") & ") AND (CatId=" & UserGroup("ObjId") & ") AND (UserId=" & User & "));"
                RsCheck.Open SQLCheck, My_Conn
                If RsCheck.EOF Or RsCheck.BOF Then
                    ' Insertion dans CatUserPref    1,CatId,User,0
                    SQL1 = "INSERT INTO dbp_CatUserList (Cid ,CatId, UserId, CatUserHide) "
                    SQL1 = SQL1 & "Values (" & Session("candidat_APPLICATION") & ", " & UserGroup("ObjId") & "," & User & ",0)"
                    My_Conn.Execute (SQL1)
                End If
                RsCheck.Close
                UserGroup.MoveNext
            Loop
            Set UserGroup = Nothing
            Group.MoveNext
    Loop
    
    Set Group = Nothing
    My_Conn.Close
    Set My_Conn = Nothing
    Response.Redirect ("../Portal_asp/result.asp?message=User+successfully+Updated")
End If

'==========================================================
' MODE DELETE
'==========================================================

If modeuser = "DELETE" Then
    SQL = "DELETE * From dbp_Users  WHERE  (UserId=" & User & ") AND CId=" & Session("candidat_APPLICATION")
    My_Conn.Execute (SQL)
    SQL = "DELETE * From dbp_UserInfos  WHERE  (UserId=" & User & ") AND CId=" & Session("candidat_APPLICATION")
    My_Conn.Execute (SQL)
    SQL = "DELETE * From dbp_CatUserList  WHERE  (UserId=" & User & ") AND CId=" & Session("candidat_APPLICATION")
    My_Conn.Execute (SQL)
    SQL = "DELETE * From dbp_GroupsPermission  WHERE ObjTypeId=50 AND ObjId=" & User & " AND CId=" & Session("candidat_APPLICATION")
    My_Conn.Execute (SQL)
    My_Conn.Close
    Set My_Conn = Nothing
    Response.Redirect ("../Portal_asp/result.asp?message=User+successfully+deleted")
End If

End Function
Public Function groupAddNew() As String
    
    CheckConnexion
    CheckAdminRights
    
    FormCaption = "Add a Group to Portal"
    Set Parent = Server.CreateObject("ADODB.Recordset")
    SQL2 = "SELECT * FROM dbp_Groups WHERE ( (Cid=" & Session("candidat_APPLICATION") & ") AND (GroupParent=0)) ORDER BY GroupID;"
    Parent.Open SQL2, My_Conn
        
    Do Until Parent.EOF
        NewGroupID = Parent("GroupId")
        Parent.MoveNext
    Loop
    NewGroupID = NewGroupID + 1
    Parent.MoveFirst
    InitForm "ADM_GRP_AUT_REP", "Portal.asp?mode=GroupUpdate&GroupId=" & NewGroupID
    CreateField "GROUP_ID", "Group Id", NewGroupID
    CreateField "GROUP_DESC", "Group Name", ""
    InitSelectList "GRPPARENT", "Parent Group", selectedoption, 1
    AddSelectList 0, "Root", selectedoption
    Do Until Parent.EOF
            AddSelectList Parent("GroupId"), Parent("GroupDescr"), selectedoption
            Parent.MoveNext
    Loop
    Parent.Close
    Set Parent = Nothing
    EndSelectList
    'CreateCheckButton "contactgp", "Groupe de type Contact", "1", ""
    pr ("<hr><br>")
    EndForm2 ("Add")
    
End Function
Public Function groupEdition() As String
    CheckConnexion
    CheckAdminRights
    '=========================================
    ' Chargement de la liste des groupes issues de dbp_Groups
    '=========================================
    SQL = "SELECT * FROM dbp_Groups WHERE (       (Cid=" & Session("candidat_APPLICATION") & ")      AND      (GroupId=" & Request.QueryString("Group") & " )) ORDER BY GroupID;"
    Set Group = Server.CreateObject("ADODB.Recordset")
    Set Parent = Server.CreateObject("ADODB.Recordset")
    Group.Open SQL, My_Conn
    '=========================================
    selectedoption = Group("GroupParent")
    FormCaption = "Edit User Group"
    InitForm "ADM_GRP_EDIT", "Portal.asp?mode=GroupUpdate&Group=" & Request.QueryString("Group")
        CreateField "GROUPNAME", "Groupe Name", Group("GroupDescr")
    InitSelectList "GRPPARENT", "Parent Group", selectedoption, 1
        SQL2 = "SELECT * FROM dbp_Groups WHERE ( (Cid=" & Session("candidat_APPLICATION") & ") AND (GroupParent=0)) ORDER BY GroupID;"
        Parent.Open SQL2, My_Conn
        AddSelectList 0, "Root", selectedoption
    Do Until Parent.EOF
            AddSelectList Parent("GroupId"), Parent("GroupDescr"), selectedoption
            Parent.MoveNext
    Loop
    Parent.Close
    Set Parent = Nothing
    EndSelectList
    If Group("GroupContactList") = 1 Then
        ischecked = "Checked"
    End If
    'CreateCheckButton "contactgp", "Groupe de type Contact", "1", ischecked
    CreateLink " Mailing list", LinkIcon, "Portal.asp?mode=GroupMail&step=1&Group=" & Request.QueryString("Group")
    'Response.Write "<input type='hidden' name='contactgpWas' value='" & ischecked & "'>"
    CreateLink " Delete Group", LinkIcon, "Portal.asp?mode=GroupDelete&Group=" & Request.QueryString("Group")
    Response.Write ("<hr><br>")
    EndForm2 ("Update")

    
End Function
Public Function groupUpdate() As Boolean
    CheckConnexion
    CheckAdminRights
    
    If Request("Group") = 0 Then
    
        '*************************************************************
        ' Recherche du prochain Group Id
        '*************************************************************
    
        SQL = "select * from dbp_Groups WHERE CId=" & Session("candidat_APPLICATION") & " order by GroupId;"
        Set rs1 = Server.CreateObject("ADODB.Recordset")

        rs1.Open SQL, My_Conn
        Do Until rs1.EOF
            NewGroupID = rs1("GroupId")
            rs1.MoveNext
        Loop
        NewGroupID = NewGroupID + 1
        rs1.Close
        Set rs1 = Nothing
        '=============================================================
        GroupContactList = Request.Form("contactgp") * 1
        SQL = "Insert into dbp_Groups (CId, GroupId,GroupDescr,GroupParent,GroupContactList)  Values (" & Session("candidat_APPLICATION") & ", " & NewGroupID & ",'" & Server.HTMLEncode(Request.Form("GROUP_DESC")) & "'," & Request.Form("GRPPARENT") & "," & GroupContactList & ");"
        My_Conn.Execute (SQL)
        My_Conn.Close
        Set My_Conn = Nothing
        Response.Redirect ("../Portal_asp/result.asp?message=Group+successfully+added")
    
    Else
    
        '=========================================
        ' Chargement de la liste des groupes issues de dbp_Groups
        '=========================================
    GroupID = Request.QueryString("Group")
    GroupName = Request.Form("GROUPNAME")
    GroupParent = Request.Form("GRPPARENT")
    GroupContactList = Request.Form("contactgp") * 1
    contactgpWas = Request.Form("contactgpWAS")
    SQL = "UPDATE dbp_Groups  SET GroupDescr ='" & Server.HTMLEncode(GroupName) & "',GroupParent=" & GroupParent & ",GroupContactList=" & GroupContactList & "  WHERE  CId=" & Session("candidat_APPLICATION") & " AND GroupId=" & GroupID & " ;"
    My_Conn.Execute (SQL)
    My_Conn.Close
    Response.Redirect ("../Portal_asp/result.asp?message=Group+successfully+updated")
    
    End If
End Function
Public Function groupDelete() As Boolean
    CheckConnexion
    CheckAdminRights
    '=========================================
    ' Chargement de la liste des groupes issues de dbp_Groups
    '=========================================
    GroupID = Request.QueryString("Group")
    SQL = "DELETE  * FROM dbp_Groups  WHERE  CId=" & Session("candidat_APPLICATION") & " AND GroupId=" & GroupID & " ;"
    My_Conn.Execute (SQL)
    SQL = "DELETE  * FROM dbp_GroupsPermission  WHERE  CId=" & Session("candidat_APPLICATION") & " AND GroupId=" & GroupID & " ;"
    My_Conn.Execute (SQL)
    My_Conn.Close
    Response.Redirect ("../Portal_asp/result.asp?message=Group+successfully+deleted")

End Function

Public Sub GroupMail()
    CheckConnexion
    
If Request.QueryString("step") = "1" Then
    
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT * FROM dbp_Groups WHERE ( (Cid=" & Session("candidat_APPLICATION") & ") AND (Groupid=" & Request.QueryString("group") & "));"
    rs.Open SQL, My_Conn
    If Not rs.EOF Then
        FormCaption = "Send email to Group " & rs("groupdescr")
    End If
    rs.Close
    Set rs = Nothing
    InitForm "ADM_GRP_MAIL", "Portal.asp?mode=GroupMail&step=2&group=" & Request.QueryString("group")
    CreateField "MAIL_FROM", "Email From", Session("candidat_username")
    CreateField "MAIL_FROM2", "From adress", Session("candidat_useremail")
    CreateField "MAIL_SUBJECT", "Subject", ""
    CreateTextArea "MAIL_BODY", "Message", "", 100, 30, 20
    EndForm2 ("Send")
    
ElseIf Request.QueryString("step") = "2" Then
    SQL = "SELECT dbp_GroupsPermission.CId, dbp_GroupsPermission.ObjId"
    SQL = SQL & " FROM dbp_GroupsPermission"
    SQL = SQL & " WHERE (((dbp_GroupsPermission.CId)=1) AND ( (dbp_GroupsPermission.ObjPermType)>0) AND ((dbp_GroupsPermission.GroupId)=" & Request.QueryString("group") & ") AND ((dbp_GroupsPermission.ObjTypeId)=50));"
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.Open SQL, My_Conn
    my_NameFrom = Request.Form("MAIL_FROM")
    my_EmailFrom = Request.Form("MAIL_FROM2")
    my_subject = Request.Form("MAIL_SUBJECT")
    my_Body = Request.Form("MAIL_BODY")
    Set rs = Server.CreateObject("ADODB.Recordset")
    Do While Not rs1.EOF
        SQL = "select * from dbp_users where cid=" & Session("candidat_application") & " and userid=" & rs1("ObjId") & ";"
        rs.Open SQL, My_Conn
        If Not rs.EOF Then
            If Trim(rs("useremail")) <> "" Then Call SendEmail(my_NameFrom, my_EmailFrom, Trim(rs("useremail")), my_subject, my_Body)
        End If
        rs.Close
        rs1.MoveNext
    Loop
    rs1.Close
    Set rs1 = Nothing
    Set rs = Nothing
    Response.Redirect ("../Portal_asp/result.asp?message=Mail+successfully+sent")
End If

    My_Conn.Close
End Sub

Public Function securityLogin(USERNAME As String, Password As String) As Boolean
        
        Set ConnPasswords = Server.CreateObject("ADODB.Connection")
        Set cmdcheckuser = Server.CreateObject("ADODB.Recordset")
        ConnPasswords.Open Session("candidat_DSN")
        ' ============== SQL pour la connexion =========================
        SQL = "SELECT * FROM dbp_Users "
        SQL = SQL & "WHERE (((dbp_Users.UserLogin)='" & USERNAME & "')  AND ((dbp_Users.CId)=" & Session("candidat_APPLICATION") & ")); "
        cmdcheckuser.Open SQL, ConnPasswords
        
        If cmdcheckuser.EOF And cmdcheckuser.BOF Then
                        Session("candidat_PASSWORDACCESS") = "No"
                        cmdcheckuser.Close
                        securityLogin = False
                        'Response.Write "Identification Inconnue<br>"
                        Exit Function
                Else
                        'Response.Write "Identification acceptée<br>"
                        Set CM = Server.CreateObject("AspCrypt.Crypt")
                        If CM.Crypt(cmdcheckuser("UserSecure"), Password) = cmdcheckuser("UserPass") Then  ' vérification du mot de passe
                                        Session("candidat_PASSWORDACCESS") = "Yes"
                                        Session("candidat_USERId") = cmdcheckuser("UserId")
                                        Session("candidat_FULL_NAME") = cmdcheckuser("UserDisplayName")
                                        Session("candidat_PASSWORD") = cmdcheckuser("UserPass")
                                        Session("candidat_USERNAME") = cmdcheckuser("UserLogin")
                                        Session("candidat_LASTLOGIN") = cmdcheckuser("LastLogin")
                                        Session("candidat_EMAIL") = cmdcheckuser("UserEMail")
                                        Session("candidat_language") = cmdcheckuser("userlang")
                                        'session("candidat_FILTER")=0
                                        
                                        ' Update last login date and time
                                        
                                        SQL = "UPDATE dbp_Users SET dbp_Users.LastLogin = Now(), dbp_Users.LoginDate = Now() "
                                        SQL = SQL + " WHERE (((dbp_Users.UserId)=" & cmdcheckuser("UserId") & "));"
                                        ConnPasswords.Execute (SQL)
                                        
                                        SQL = "UPDATE dbp_UserInfos SET dbp_UserInfos.DateLastAccess = Now(), dbp_UserInfos.logonCount=dbp_UserInfos.logonCount+1 "
                                        SQL = SQL + " WHERE (((dbp_UserInfos.UserId)=" & cmdcheckuser("UserId") & "));"
                                        ConnPasswords.Execute (SQL)
                                        cmdcheckuser.Close
                                        
                                        Session("candidat_ADMIN") = 0
                                        
                                        ' Check if user is an administrator
                                        SQL = "SELECT dbp_GroupsPermission.CId, dbp_GroupsPermission.GroupId, dbp_GroupsPermission.ObjTypeId, dbp_GroupsPermission.ObjId"
                                        SQL = SQL & " FROM dbp_GroupsPermission"
                                        SQL = SQL & " WHERE (((dbp_GroupsPermission.CId)=1) AND ( (dbp_GroupsPermission.ObjPermType)>0) AND ((dbp_GroupsPermission.GroupId)=1 or (dbp_GroupsPermission.GroupId)=2) AND ((dbp_GroupsPermission.ObjTypeId)=50) AND ((dbp_GroupsPermission.ObjId)=" & Session("candidat_USERID") & "));"
                                        cmdcheckuser.Open SQL, ConnPasswords
                                        If cmdcheckuser.EOF And cmdcheckuser.BOF Then
                                                Session("candidat_ADMIN") = 0
                                                Session("candidat_ACCESS_LEVEL") = "Low"
                                        Else
                                                Session("candidat_ADMIN") = 1
                                                Session("candidat_ACCESS_LEVEL") = "Full"
                                        End If
                                        cmdcheckuser.Close
                                        ' Load Mail Parameters
                                        SQL = "Select * from dbp_MailServer"
                                        cmdcheckuser.Open SQL, ConnPasswords
                                        If cmdcheckuser.EOF And cmdcheckuser.BOF Then
                                                                Session("candidat_MAILSERVER") = ""
                                        Else
                                                                Session("candidat_MAILSERVER") = cmdcheckuser("Portal_Mail_Server")
                                                                Session("candidat_MAILSERVERPATH") = cmdcheckuser("Portal_MailServer_Path")
                                                                Session("candidat_MAILMASTER") = cmdcheckuser("Portal_MailMaster")
                                                                Session("candidat_MAILDOMAIN") = cmdcheckuser("Portal_Mail_Domain")
                                        End If
                                        cmdcheckuser.Close

                                Else
                                        'Response.Write "Mauvais Mot de Passe<br>"
                                        Session("candidat_PASSWORDACCESS") = "No"
                                        cmdcheckuser.Close
                                End If
                End If
        
        Set CM = Nothing
        Set cmdcheckuser = Nothing
        ConnPasswords.Close
        Set ConnPasswords = Nothing
        'Response.Write ("Fin de la procédure Authentification<br>")
        securityLogin = True
 
End Function
Public Function portalInit(ByVal ZeDSN As String, ByVal ZeAppliID As Integer) As Boolean
    Session("candidat_DSN") = ZeDSN
    Session("candidat_Application") = ZeAppliID
    portalInit = True
    url_base = Session("candidat_URLBASE")
End Function
Sub ViewNewsHeader(CatPath, CatUrl)
    Response.Write BeforeTHead
    Response.Write "<p id=THead><a href=" & CatUrl & ">" & Trim(CatPath) & "</a></THead>" & vbCrLf
    Response.Write AfterTHead
End Sub

Sub ViewNewsStyle(Title, TopicUrl, TopicDate, TopicResume, TopicAuthor, AuthorMail, Icon, Tedit, Tapprove, AllStyle, Target)
    Dim my_topicresume
    
    BeforeTHead = "<table align='center' border='0'  width='100%'><tr><td>"
    AfterTHead = "</td></tr></table>"
    BeforeDheading = "<table align='center' border='0' width='100%'><tr>"
    AfterDheading = "</tr></table>"
    BeforeResume = "<table align='center' border='0' width='100%'><tr><td>"
    AfterResume = "</td></tr></table>"
    
    If Session("candidat_BgPixTitle") = "" Then
        BeforeDheading = vbCrLf & "<table align='left' border='0' width='100%'><tr>"
    Else
        BeforeDheading = vbCrLf & "<table background='" & Session("candidat_BgPixTitle") & "' align='left' border='0' width='100%'><tr>"
    End If
    pr (BeforeDheading)
    SQL = "select * from dbp_Styles where Style_Id=" & AllStyle
        Set RS2 = Server.CreateObject("ADODB.Recordset")
        RS2.Open SQL, My_Conn
        If Not RS2.EOF Then
            StyleHead = RS2("Style_Header")
            StyleDate = RS2("Style_Date")
            StyleBody = RS2("Style_Body")
            StyleAuthor = RS2("Style_Author")
        End If
    RS2.Close
    Set RS2 = Nothing
    
    If Title <> "" Then
        If Icon <> "news.gif" And Icon <> "" Then
            If Trim(TopicUrl) = "" Then
                pr ("<td width='100%' id=" & StyleHead & "><img src='../Portal_Image/" & Icon & "' border=0  align='absmiddle'>&nbsp;" & Trim(Title) & "</" & StyleHead & ">" & IIf(Len(Trim(TopicDate)) > 0, "<font id=" & StyleDate & "> - (" & TopicDate & ")</font>", "") & "</td>")
            Else
                pr ("<td width='100%' id=" & StyleHead & "><a href='" & TopicUrl & "' " & Target & "><img src='../Portal_Image/" & Icon & "' border=0  align='absmiddle'><font id='" & StyleHead & "'> " & Trim(Title) & "</a></" & StyleHead & ">" & IIf(Len(Trim(TopicDate)) > 0, "<font id=" & StyleDate & "> - (" & TopicDate & ")</font>", "") & "</td>")
            End If
        Else
            If Trim(TopicUrl) = "" Then
                pr ("<td width='100%'><font id='" & StyleHead & "'> " & Trim(Title) & "</font>" & IIf(Len(Trim(TopicDate)) > 0, "<font id=" & StyleDate & "> - (" & TopicDate & ")</font>", "") & "</td>")
            Else
                pr ("<td width='100%'><a href='" & TopicUrl & "' " & Target & "><font id='" & StyleHead & "'> " & Trim(Title) & "</font></a></" & StyleHead & ">" & IIf(Len(Trim(TopicDate)) > 0, "<font id=" & StyleDate & "> - (" & TopicDate & ")</font>", "") & "</td>")
            End If
        End If
    End If
    
    Response.Write AfterDheading & vbCrLf
    my_topicresume = TopicResume
    If Len(Trim(my_topicresume)) > 0 Then
        If Session("candidat_BgPixAbstract") = "" Then
            BeforeResume = vbCrLf & "</td></tr><tr><td width='100%' valign='top'><table align='center' border='0' width='100%' ><tr><td>"
        Else
            BeforeResume = vbCrLf & "</td></tr><tr><td width='100%' valign='top'><table background='" & Session("candidat_BgPixAbstract") & "' align='center' border='0' width='100%' ><tr><td>"
        End If
        
        pr (BeforeResume)
        pr ("<p id=" & StyleBody & ">" & my_topicresume & "</" & StyleBody & ">")
        If TopicAuthor <> "" Then
             If Tedit = "" Then
                 pr ("<p id=" & StyleDate & ">Last Update :" & TopicDate & "<br>by <a href='MAILTO:" & AuthorMail & "'>" & TopicAuthor & "</a>")
             Else
                 pr ("<p id=" & StyleDate & ">Last Update :" & TopicDate & "<br>by <a href='MAILTO:" & AuthorMail & "'>" & TopicAuthor & "</a>")
                 ' response.write "<td width='10%' <p id=boxhead>" & Tedit & "<br>"  & Tapprove & "</boxhead></td>"
                 pr (Tedit & "</boxhead>")
             End If
         Else
             If Tedit <> "" Then pr ("<br>" & Tedit)
         End If
        pr (AfterResume)
    Else
        If TopicAuthor <> "" Then
             If Tedit = "" Then
                 pr ("<p id=" & StyleDate & ">Last Update :" & TopicDate & "<br>by <a href='MAILTO:" & AuthorMail & "'>" & TopicAuthor & "</a>")
             Else
                 pr ("<p id=" & StyleDate & ">Last Update :" & TopicDate & "<br>by <a href='MAILTO:" & AuthorMail & "'>" & TopicAuthor & "</a>")
                 ' response.write "<td width='10%' <p id=boxhead>" & Tedit & "<br>"  & Tapprove & "</boxhead></td>"
                 pr (Tedit & "</boxhead>")
             End If
         Else
             If Tedit <> "" Then pr ("<tr><td>" & Tedit & "</td></tr>")
         End If
    End If
End Sub


Sub CheckAdminRights()
    If Session("candidat_ADMIN") <> 1 Then
                Session("candidat_USERId") = 0
                pr ("<p id=alarm>You have no permissions for accessing this area")
                pr ("You have been logged out </p>")
                pr ("<THead>Session is expired ...<Br>You must login now<Br><H2><a href='../Portal_ASP/Portal_login.asp' target=_top>Click Here</a>")
                Response.End
    End If
End Sub

Sub SendEmail(NameFrom, EmailFrom, MailTo, EmailSubject, EmailText)
Dim my_Mail

            '    Set MyEmail = Server.CreateObject("DynuEmail.Functions")
            '    MyEmail.Smtp = session("candidat_MAILSERVER")
            '    result = MyEmail.Send(EmailFrom, MailTo, EmailSubject, EmailText)
            '    Set MyEmail = Nothing
            'End Sub
    Set my_Mail = Server.CreateObject("CDONTS.NewMail")
    my_Mail.BodyFormat = 1 ' Text Only, 0 for HTML
    my_Mail.subject = EmailSubject
    my_Mail.To = MailTo
    my_Mail.Body = EmailText
    my_Mail.Send (NameFrom & "<" & EmailFrom & ">")
    Set my_Mail = Nothing
    
    
End Sub
Function doCode(str, oTag, cTag, roTag, rcTag)
    tx = Split(str, cTag)
    t = ""
    For i = 0 To UBound(tx)

      If LCase(oTag) = "[a]" Then
        p = InStr(1, tx(i), "[a]", 1)
        If p <> 0 Then
            tmp = Mid(tx(i), p)
            url = Mid(tmp, 4)
            If LCase(Left(url, 5)) = "http:" Then
                tmp1 = Replace(tmp, "[a]" & url, "<A href='" & url & "' Target=_Blank>Link</a>", 1, -1, 1)
            Else
                tmp1 = Replace(tmp, "[a]" & url, "<A href='http://" & url & "' Target=_Blank>Link</a>", 1, -1, 1)
            End If
            t = t & Replace(tx(i), tmp, tmp1)
        Else
            t = t & tx(i)
        End If
      Else
        cnt = InStr(1, tx(i), oTag, 1)
        Select Case cnt
            Case 0
                t = t & tx(i) & " "
            Case Else
                t = t & Replace(tx(i), oTag, roTag, 1, 1, 1)
                t = t & " " & rcTag & " "
        End Select
      End If
    Next
    doCode = t
End Function
Function ChkString(str)
     If str = "" Then
        str = " "
     Else
        If BadWordFiler = "true" Then
          bwords = Split(BadWords, "|")
          For i = 0 To UBound(bwords)
            str = Replace(str, bwords(i), String(Len(bwords(i)), "*"), 1, -1, 1)
          Next
        End If
     End If
     
     '  Do ASP Forum Code
     str = doCode(str, "[b]", "[/b]", "<b>", "</b>")
     str = doCode(str, "[i]", "[/i]", "<i>", "</i>")
     str = doCode(str, "[quote]", "[/quote]", "<BLOCKQUOTE><font size=1 face=arial>quote:<hr height=1 noshade>", "<hr height=1 noshade></BLOCKQUOTE></font><font face='" & DefaultFontFace & "' size=2>")
     str = doCode(str, "[a]", "[/a]", "<a>", "</a>")
     str = doCode(str, "[code]", "[/code]", "<pre>", "</pre>")
     
 
     str = Replace(str, "'", "''")
     str = Replace(str, "|", "/")
     
     ChkString = str
End Function

Sub CheckConnexion()

    If Session("candidat_DSN") = "" Then
        Response.Write ("<htm><body onload=""top.location='" & url_base & "';""></body></htm>")
    Else
        '============ On définit le DSN du fichier pour tout le Portail =========
        Set My_Conn = Server.CreateObject("ADODB.Connection")
        My_Conn.Open Session("candidat_DSN")
    End If
End Sub

Sub InitForm(NomForm, PostTo)
    FlagHTMLEditor = False
    pr ("<HTML>")
    pr ("<HEAD>")
    pr ("<link rel=STYLESHEET href='PMainStyle1.asp' type='text/css'>")
    pr ("</HEAD>")
    pr ("<BODY BACKGROUND='Background.asp'>")
    pr ("<p id='dheading'>" & FormCaption & "</p>")
    pr ("<form action='" & PostTo & "' method='post' name='" & NomForm & "'>")
End Sub

Sub InitFormNB(NomForm, PostTo)
    pr ("<HTML>")
    pr ("<HEAD>")
    pr ("<link rel=STYLESHEET href='PMainStyle1.asp' type='text/css'>")
    pr ("</HEAD>")
    pr ("<form action='" & PostTo & "' method='post' name='" & NomForm & "'>")
End Sub

Sub EndForm(ButtonValue)
    pr ("<br>")
    pr ("<table border='0' align='center' width='600px' >")
    If FlagHTMLEditor Then
        pr ("<tr><td align='left'><input name='submit' class='cmdflat' type='submit' value='" & ButtonValue & "' onclick=""SetVals();""></td></tr>")
    Else
        pr ("<tr><td align='left'><input name='submit' class='cmdflat' type='submit' value='" & ButtonValue & "'></td></tr>")
    End If
    pr ("</table>")
    pr ("</form>")
    pr ("</body>")
    pr ("</HTML>")
End Sub
    
Sub EndForm2(ButtonValue)
    pr ("<br>")
    pr ("<table border='0' width='600px'>")
    pr ("<tr><td align='left' valign='center' width='30%' height ='27'   background='fondmaj.gif'><input type='image' Name='submit'  src='" & ButtonValue & ".gif' border='0'></td></tr>")
    pr ("</table>")
    pr ("</form>")
    pr ("</BODY>")
    pr ("</HTML>")
End Sub
    
Sub CreateLink(LinkCaption, LinkIcon, LinkUrl)

    pr ("<table border='0' align='absmiddle' width='600px'>")
    pr ("<tr><td  align='left' width='30%' class='smallerheader'><img src='o4.gif' border=0>")
    pr ("<a href='" & LinkUrl & "'> " & LinkCaption & "</a></td></tr></table>")
End Sub


Sub CreateField(fieldNAME, fieldCAPTION, FieldVALUE)

    pr ("<table border='0' align='absmiddle' width='600px' >")
    pr ("<tr><td align='left' width='30%'  background='fondmaj.gif' class='smallerheader'> <img src='o5.gif'>" & fieldCAPTION & "  :  </td></tr>")
    pr ("<tr><td><input type='text' class='tbflat' name='" & fieldNAME & "' value='" & FieldVALUE & "' size='40' tabindex='1' ></td></tr>")
    pr ("</table>")

End Sub

Sub CreateCheckButton(fieldNAME, fieldCAPTION, FieldVALUE, FieldChecked)

    pr ("<table width='600Px' border='0'><tr><td>")
    pr ("<p class='smallerheader'><input type='checkbox' name='" & fieldNAME & "' value='" & FieldVALUE & "' " & FieldChecked & ">" & fieldCAPTION)
    pr ("</td></tr></table>")

End Sub

Sub CreateMaskField(fieldNAME, fieldCAPTION, FieldVALUE)

    pr ("<table border='0' align='absmiddle' width='600px' >")
    pr ("<tr><td align='left' width='30%'  background='fondmaj.gif' class='smallerheader'><img src='o5.gif'>" & fieldCAPTION & " :  </td></tr>")
    pr ("<tr><td><input type='password' class='tbflat' name='" & fieldNAME & "' value='" & FieldVALUE & "' size='10' tabindex='1' ></td></tr>")
    pr ("</table>")

End Sub

Sub CreateNonEditField(fieldNAME, fieldCAPTION, FieldVALUE)

    pr ("<table border='0' align='absmiddle' width='600px' >")
    pr ("<tr><td align='left' width='30%'  background='fondmaj.gif' class='smallerheader'><img src='o5.gif'>" & fieldCAPTION & " :  </td></tr>")
    pr ("</tr><td class='smallerheader'>" & FieldVALUE & "</td></tr>")
    pr ("</table>")

End Sub


Sub CreateTextArea(fieldNAME, fieldCAPTION, FieldVALUE, fieldSIZE, FieldCOLS, FieldROWS)

    pr ("<table border='0' align='absmiddle' width='600px' >")
    pr ("<tr><td align='left' class='smallerheader' valign='top' width='30%' height ='27' background='../portal_image/fondmaj.gif'>  <img src='../portal_image/o5.gif'> " & fieldCAPTION & ":</td></tr>")
    pr ("<tr><td colspan='3' ><textarea class='tbflat' name='" & fieldNAME & "' size='" & fieldSIZE & "' cols='" & FieldCOLS & "' rows='" & FieldROWS & "' tabindex='17'>" & FieldVALUE & "</textarea></td></tr>")
    pr ("</table>")

End Sub
Sub CreateHtmlEditor(fieldNAME, fieldCAPTION, FieldVALUE, NomForm)
 
  
pr (vbCrLf)
pr (vbCrLf)
pr ("<link rel=""stylesheet"" type=""text/css"" href=""../portal_HtmlEditor/ifedit.css"">")
pr ("<table border=0 cellspacing=0 cellpadding=3 width='600px'>")
pr ("<tr>")
pr ("<td colspan=""2"">")
pr ("<SCRIPT SRC=""../Portal_HTMLEditor/ifeditscript.js""></SCRIPT>")
pr ("<TEXTAREA NAME='" & fieldNAME & "' style='visibility:hidden;position:absolute;top:0px;left:0px'>" & FieldVALUE & "<P></TEXTAREA><BR>")
pr ("</td></tr><tr><td colspan=2 align=""center"">")
pr ("<BR></td>")
pr ("</tr></table></DIV>")
pr (vbCrLf)
pr (vbCrLf)
 
TargetObject = "document." & NomForm & "." & fieldNAME & ".value"
 

pr ("<SCRIPT>")
 
pr ("function SetVals() {")
pr (TargetObject & " = editor.GetHTML();")
pr ("    if (" & TargetObject & " == '<P></P>')")
pr ("       { " & TargetObject & " = ''; }")
pr ("}")
pr ("</SCRIPT>")
 
pr ("<BODY onload=""editor.SetHTML(" & TargetObject & ");""")
 
FlagHTMLEditor = True
 
End Sub

Sub InitSelectList(fieldNAME, fieldCAPTION, FieldVALUE, fieldSIZE)

    pr ("<table border='0' align='absmiddle' width='600px' >")
    pr ("<tr><td align='left'  width='30%'   height ='27px'  background='../portal_image/fondmaj.gif'  class='smallerheader'><img src='../portal_image/o5.gif'>" & fieldCAPTION & " : </td></tr>")
    If Fieldsvalue <> "" Then
        pr ("<tr><td align='left'><select class='tbflat' name='" & fieldNAME & "' size='" & fieldSIZE & "' Id =" & FieldVALUE & ">")
    Else
        pr ("<tr><td align='left'><select class='tbflat' name='" & fieldNAME & "' size='" & fieldSIZE & "' >")
    End If
End Sub

Sub AddSelectList(optionVALUE, optionCAPTION, FieldVALUE)
        
    If optionVALUE = FieldVALUE Then
            pr ("<option selected value='" & optionVALUE & "'>" & optionCAPTION & "</option>")
    Else
            pr ("<option value='" & optionVALUE & "'>" & optionCAPTION & "</option>")
    End If
        
End Sub

Sub EndSelectList()
    pr ("</select></td></tr></table>")
End Sub



Public Function pr(strPrint)
    Response.Write strPrint & vbCrLf
End Function
Public Function safeEntry(strField)
    strSafe = Trim(strField)
    strSafe = funReplace(strSafe, "'", "`")
    strSafe = funReplace(strSafe, "<", "&lt;")
    strSafe = funReplace(strSafe, ">", "&gt;")
    safeEntry = strSafe
End Function

Public Function CheckHash(DataToCheck, CryptedData, Salt)
            Set CM = Server.CreateObject("AspCrypt.Crypt")
            If CM.Crypt(Salt, DataToCheck) = CryptedData Then
                CheckHash = 1
            Else
                CheckHash = 0
            End If
End Function
Sub HashIt(DataToHash)
            Randomize
            Salt = ""
            For i = 1 To 10
                '65 is ASCII for "A"
                Salt = Salt & Chr(Int(Rnd * 26) + 65)
            Next
            ' Calculate Hash of (Password & Salt)
            Set CM = Server.CreateObject("AspCrypt.Crypt")
            Session("candidat_HashData") = CM.Crypt(Salt, DataToHash)
            Session("candidat_HashSecure") = Salt
End Sub
Function funY2K(d)
    strDate = Trim(d)
    If InStr(strDate, " ") Then
        strDate = Left(strDate, InStr(strDate, " "))
        trailer = Right(d, Len(d) - InStr(d, " "))
    End If
    If IsDate(strDate) Then
        dateY2K = strDate
        If InStr(strDate, "/") = 2 Then
            strMonth = Left(strDate, 1)
            If InStr(3, strDate, "/") = 4 Then
                strDay = Mid(strDate, 3, 1)
            Else
                strDay = Mid(strDate, 3, 2)
            End If
            strYear = Right(strDate, Len(strDate) - InStr(3, strDate, "/"))
        ElseIf InStr(strDate, "/") = 3 Then
            strMonth = Left(strDate, 2)
            If InStr(4, strDate, "/") = 5 Then
                strDay = Mid(strDate, 4, 1)
            Else
                strDay = Mid(strDate, 4, 2)
            End If
            strYear = Right(strDate, Len(strDate) - InStr(4, strDate, "/"))
        End If
        intYear = CInt(strYear)
        If intYear >= 0 And intYear < 51 Then
            strYear = "20" & strYear
        ElseIf intYear > 50 And intYear < 100 Then
            strYear = "19" & strYear
        End If
        
        funY2K = strMonth & "/" & strDay & "/" & strYear & " " & trailer
    Else
        funY2K = ""
    End If
End Function
Public Function funReplace(a, b, c)
    funReplace = Replace(a, b, c)
End Function
Function getLetter(num)
    If num = 1 Then
        getLetter = "a"
    ElseIf num = 2 Then
        getLetter = "b"
    ElseIf num = 3 Then
        getLetter = "c"
    ElseIf num = 4 Then
        getLetter = "d"
    ElseIf num = 5 Then
        getLetter = "e"
    ElseIf num = 6 Then
        getLetter = "f"
    ElseIf num = 7 Then
        getLetter = "g"
    ElseIf num = 8 Then
        getLetter = "h"
    ElseIf num = 9 Then
        getLetter = "i"
    ElseIf num = 10 Then
        getLetter = "j"
    ElseIf num = 11 Then
        getLetter = "k"
    ElseIf num = 12 Then
        getLetter = "l"
    ElseIf num = 13 Then
        getLetter = "m"
    ElseIf num = 14 Then
        getLetter = "n"
    ElseIf num = 15 Then
        getLetter = "o"
    ElseIf num = 16 Then
        getLetter = "p"
    ElseIf num = 17 Then
        getLetter = "q"
    ElseIf num = 18 Then
        getLetter = "r"
    ElseIf num = 19 Then
        getLetter = "s"
    ElseIf num = 20 Then
        getLetter = "t"
    ElseIf num = 21 Then
        getLetter = "u"
    ElseIf num = 22 Then
        getLetter = "v"
    ElseIf num = 23 Then
        getLetter = "w"
    ElseIf num = 24 Then
        getLetter = "x"
    ElseIf num = 25 Then
        getLetter = "y"
    ElseIf num = 26 Then
        getLetter = "z"
    Else
        getLetter = ""
    End If
End Function

Public Function translate(my_msg)

    Dim my_cid, my_language, rs, res

    my_cid = Session("candidat_Application")
    my_language = Session("candidat_language")
    strDSN = Session("candidat_DSN")
    Set cnx = Server.CreateObject("ADODB.Connection")
    cnx.Open strDSN
    Set rs = cnx.Execute("select txt_trans from dbp_trans where cid=" & my_cid & " and language='" & my_language & "' and txt ='" & my_msg & "'")
    If Not rs.EOF Then
        res = rs("txt_trans")
    Else
        res = my_msg
    End If
    rs.Close
    Set rs = Nothing
    cnx.Close
    Set cnx = Nothing
    translate = res

End Function

Public Function display_quick()

    CheckConnexion
    
    Dim rs, SQL, my_href

    SQL = "SELECT * FROM dbp_categories WHERE Cid=" & Session("candidat_APPLICATION") & " and catid =" & Request.QueryString("catid")
    Set rs = Server.CreateObject("ADODB.Recordset")
    
    rs.Open SQL, My_Conn
    
    If Not rs.EOF Then
        If rs("catexturl") = "" And rs("catstyle") = 3 Then
            my_href = Session("candidat_URLBASE") & "/portal_asp/portal.asp?mode=topics&Category=" & rs("CatId") & "&CatStyle=3"
        Else
            If InStr(1, rs("catexturl"), "http") = 0 And InStr(1, rs("catexturl"), "../") = 0 Then
                my_href = Session("candidat_URLBASE") & my_rep & rs("catexturl")
            Else
                my_href = rs("catexturl")
            End If
        End If
    End If
    rs.Close
    My_Conn.Close
    my_href = Replace(my_href, "&amp;", "&")
    Response.Redirect my_href

End Function

Public Function display_htm()

    CheckConnexion
    
    Dim rs, SQL

    SQL = "SELECT cathtmlcontent FROM dbp_categories WHERE Cid=" & Session("candidat_APPLICATION") & " and catid=" & Request.QueryString("catid") & " "
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, My_Conn
    If Not rs.EOF Then
        pr (rs("cathtmlcontent"))
    End If
    rs.Close
    My_Conn.Close
    
End Function
Public Sub my_father(id, f_rep, f_id)
    Dim rs, SQL
    
    SQL = "select catid, parent, catrep from dbp_categories where catid=" & id
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open SQL, My_Conn
    If Not rs.EOF Then
        If rs("parent") = 0 Then
            f_rep = rs("catrep")
            f_id = rs("catid")
        Else
            my_father rs("parent"), f_rep, f_id
        End If
    End If
    rs.Close
    'My_Conn.Close
    
End Sub

Public Function occurrence(str, mot)
    Dim s
    Dim sortie As Boolean
    Dim nb, pos, longueur As Integer
        
    sortie = False
    nb = 0
    s = str
    longueur = Len(mot)
    Do While Not sortie
        pos = InStr(1, s, mot)
        If pos > 0 Then
            nb = nb + 1
            s = Mid(s, pos + longueur)
        Else
            sortie = True
        End If
    Loop
    occurrence = nb
End Function
Public Function search()

Dim my_father_rep, my_html, my_desc As String, my_id, my_father_id, my_note As Long
Dim lst_note(100) As Integer
Dim lst_res(100) As String


    CheckConnexion
pr ("<html>")
pr ("<head>")
pr ("<link rel=STYLESHEET href=""PMainStyle1.asp?type=search"" type=""text/css"">")
pr ("</head>")
pr ("<body background=""Background.asp"">")

' ====================================ON LANCE LES RECHERCHES =====================================
If Request.QueryString("modesearch") = "doit" Then
    
    If Request.Form("search") <> "" Then
        
        SearchPattern = Request.Form("search")
        keywords = Split(SearchPattern, ",")
        keycnt = UBound(keywords)
        '=========================================================================
        ' ====================================ON LANCE LA RECHERCHE DANS LE PORTAIL======
        Set rs1 = Server.CreateObject("ADODB.Recordset")        ' Categories
        Set rs = Server.CreateObject("ADODB.Recordset")             ' Topics
        
        SQL = "SELECT dbp_Categories.catid, dbp_Categories.catcaption, dbp_Categories.CatDescription, dbp_categories.cathtmlcontent  FROM ( dbp_CatUserList INNER JOIN dbp_Categories ON dbp_CatUserList.CatId = dbp_Categories.CatId)"
        SQL = SQL & " WHERE  ((dbp_CatUserList.UserId)=" & Session("candidat_USERId") & ") AND ("
        
        cnt = 0
        
        For Each word In keywords
            SQL = SQL & " (dbp_Categories.CatDescription Like '%" & word & "%'"
            SQL = SQL & " OR dbp_Categories.cathtmlcontent Like '%" & word & "%')"
            If cnt < keycnt Then SQL = SQL & Request.Form("andor")
            cnt = cnt + 1
        Next
        
        SQL = SQL & ");"
        'Debug.Print SQL
        rs1.Open SQL, My_Conn
        
        flagfound = 0
        i = 0
        ' ========== AFFICHAGE DES RESULTATS =================================
        Do While Not rs1.EOF
            flagfound = 1
            
            ' recherche de la catégorie parente de niveau root pour info répertoire et cat id
            my_id = rs1("catid")
            my_html = Trim(rs1("cathtmlcontent"))
            my_desc = Trim(rs1("catdescription"))
            my_father my_id, my_father_rep, my_father_id
           
            cnt = 0
            keywords = Split(SearchPattern, ",")
            keycnt = UBound(keywords)
            res = "("
            For Each word In keywords
                my_note = 0
                If (Not IsNull(my_html)) And my_html <> "" Then
                    my_note = occurrence(UCase(my_html), UCase(word)) + my_note
                End If
                If (Not IsNull(my_desc)) And my_desc <> "" Then
                    my_note = occurrence(UCase(my_desc), UCase(word)) + my_note
                End If
                If cnt < keycnt Then
                    If cnt = 0 Then
                        res = res & word & ":" & my_note
                    Else
                        res = res & ", " & word & ":" & my_note
                    End If
                Else
                    If cnt > 0 Then
                        res = res & ", " & word & ":" & my_note & ")"
                    Else
                        res = "(" & word & ":" & my_note & ")"
                    End If
                End If
                lst_note(i) = lst_note(i) + my_note * (keycnt + 2 - cnt)
                cnt = cnt + 1
            Next
            res = "<td><a href=""../portal_html/default.asp?mode=cat&catid=" & my_father_id & "&rep=" & my_father_rep & "&rep_haut=" & my_father_rep & "haut/&prev=home&current=cat&quickurl=" & rs1("catid") & """ target=""_top"" >" & rs1("Catcaption") & "</a><font class=""xpetit_italique""> " & res & "</font></td>"
            res = "<tr>" & res & "</tr>"
            lst_res(i) = res
            i = i + 1
            If i = 60 Then Exit Do
            rs1.MoveNext
        Loop
        rs1.Close
        Set rs1 = Nothing
        
        
        ' On recommence le même traitement avec les topics...
        SearchPattern = Request.Form("search")
        keywords = Split(SearchPattern, ",")
        keycnt = UBound(keywords)
        
        SQL = "SELECT dbp_Topics.Abstract, dbp_Categories.catid, dbp_Topics.Title "
        SQL = SQL & "FROM (dbp_CatUserList INNER JOIN dbp_Categories ON dbp_CatUserList.CatId = dbp_Categories.CatId) INNER JOIN dbp_Topics ON dbp_Categories.CatId = dbp_Topics.CatId"
        SQL = SQL & " WHERE  (dbp_CatUserList.UserId=" & Session("candidat_USERId") & ") AND "
        cnt = 0
        For Each word In keywords
            SQL = SQL & " (dbp_Topics.Title Like '%" & word & "%'"
            SQL = SQL & " OR dbp_Topics.Abstract Like '%" & word & "%')"
            If cnt < keycnt Then SQL = SQL & Request.Form("andor")
            cnt = cnt + 1
        Next
        'Debug.Print SQL
        
        Select Case Request.Form("SearchDate")
            Case 1
                SQL = SQL & " AND (dbp_Topics.TopicStart=#" & DateAdd("d", Date, -1) & "#)"
            Case 5
                SQL = SQL & " AND (dbp_Topics.TopicStart>=#" & DateAdd("d", Date, -5) & "#)"
            Case 10
                SQL = SQL & " AND (dbp_Topics.TopicStart>=#" & DateAdd("d", Date, -10) & "#)"
            Case 30
                SQL = SQL & " AND (dbp_Topics.TopicStart>=#" & DateAdd("d", Date, -30) & "#)"
        End Select
        SQL = SQL & ";"
 '       Debug.Print SQL
        rs.Open SQL, My_Conn
        
        ' ========== AFFICHAGE DES RESULTATS =================================
        Do While Not rs.EOF
            flagfound = 1
            
            ' recherche de la catégorie parente de niveau root pour info répertoire et cat id
            my_id = rs("catid")
            my_html = Trim(rs("Abstract"))
            my_desc = Trim(rs("Title"))
            my_father my_id, my_father_rep, my_father_id
           
            cnt = 0
            keywords = Split(SearchPattern, ",")
            keycnt = UBound(keywords)
            res = "("
            For Each word In keywords
                my_note = 0
                If (Not IsNull(my_html)) And my_html <> "" Then
                    my_note = occurrence(UCase(my_html), UCase(word)) + my_note
                End If
                If (Not IsNull(my_desc)) And my_desc <> "" Then
                    my_note = occurrence(UCase(my_desc), UCase(word)) + my_note
                End If
                If cnt < keycnt Then
                    If cnt = 0 Then
                        res = res & word & ":" & my_note
                    Else
                        res = res & ", " & word & ":" & my_note
                    End If
                Else
                    If cnt > 0 Then
                        res = res & ", " & word & ":" & my_note & ")"
                    Else
                        res = "(" & word & ":" & my_note & ")"
                    End If
                End If
                lst_note(i) = lst_note(i) + my_note * (keycnt + 2 - cnt)
                cnt = cnt + 1
            Next
            res = "<td><a href=""../portal_html/default.asp?mode=cat&catid=" & my_father_id & "&rep=" & my_father_rep & "&rep_haut=" & my_father_rep & "haut/&prev=home&current=cat&quickurl=" & rs("catid") & """ target=""_top"" >" & my_desc & "</a><font class=""xpetit_italique""> " & res & "</font></td>"
            res = "<tr>" & res & "</tr>"
            lst_res(i) = res
            i = i + 1
            If i = 100 Then Exit Do
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        
        
        My_Conn.Close
        
        'tri suivant la note
        Tri_Fusion lst_note, 0, i, lst_res
        '============BANNIERE D'ACCUEIL==============================================
        pr ("<p id=dheading> " & i & " Search Results for  : " & Replace(SearchPattern, " ", Request.Form("andor")))
        pr ("<p id=body>")
        pr ("<table>")
        For j = i To 0 Step -1
            pr (lst_res(j))
        Next
        pr ("</table>")
        
        ' NO MATCH
        If flagfound <> 1 Then Response.Write "<font face='" & DefaultFontFace & "'>Sorry ! No matches found in portal</font>"
    
    Else ' Search = ""

        pr ("<font class=""styl2""><center>" & translate("Aucun mot à rechercher !") & "</font></center>")
        pr ("<p align=""center"">")
        If Session("candidat_language") = "FR" Then
            pr ("<a href=""javascript:history.go(-1)""><img src=""../portal_image/retour.gif"" border=""0""></a>")
        Else
            pr ("<a href=""javascript:history.go(-1)""><img src=""../portal_image/back.gif"" border=""0""></a>")
        End If

    End If ' Search = ""
Else
' ============================ON CONSTRUIT LE FORMULAIRE DE RECHERCHE ===========================
pr ("<div align=""center""><center>")

pr ("<form action=""../portal_asp/portal.asp?mode=search&modesearch=doit"" method=""post"" id=form1 name=form1>")
pr ("<br>")

pr ("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""238"">")
pr ("      <tr>")
pr ("            <td width=""19""><img src=""../portal_image/tab_hg.gif""></td>")
pr ("            <td width=""200"" background=""../portal_image/tab_bgh.gif"" align=""center"" class=""styldatesearch"">" & translate("Recherche par mot-clé") & " :</td>")
pr ("            <td width=""19""><img src=""../portal_image/tab_hd.gif"" width=""19"" height=""23""></td>")
pr ("     </tr>")
        
pr ("     <TR>")
pr ("            <td width=""19"" background=""../portal_image/tab_bgg.gif"">&nbsp;</td>")
pr ("            <td width=""200"" class=""fondtableau"" align=""center"">")
pr ("            <table>")
pr ("                <tr>")
pr ("                    <td colspan=""2"" align=""middle""><input type=""text"" name=""search"" size=""30"" class=""champ"" ></td>")
pr ("                </tr>")
pr ("                <tr>")
pr ("                    <td colspan=""2"" ><font class=""normalbleupetit"">")
pr ("                    <input type=""radio"" name=""andor"" value="" and "" checked>" & translate("Rechercher tous les mots") & "<br>")
pr ("                    <input type=""radio"" name=""andor"" value="" or "">" & translate("Trouver au moins un des mots"))
pr ("                    </font>")
pr ("                    </td>")
pr ("                </tr>")
pr ("                <tr>")
pr ("                    <td colspan=""2"" >&nbsp;</td>")
pr ("                </tr>")
pr ("                <tr>")
pr ("                    <td ><font class=""normalbleupetit"">" & translate("Recherche par date") & "&nbsp;:</td>")
pr ("                    <td >")
pr ("                    <SELECT NAME=""SearchDate"" class=""champ"">")
pr ("                        <OPTION value=""0"">" & translate("Indifférent") & "</option>")
pr ("                        <OPTION VALUE=""1"">" & translate("Depuis hier") & "</option>")
pr ("                        <OPTION VALUE=""5"">" & translate("Depuis 5 jours") & "</option>")
pr ("                        <OPTION VALUE=""10"">" & translate("Depuis 10 jours") & "</option>")
pr ("                        <OPTION VALUE=""30"">" & translate("Depuis 30 jours") & "</option>")
pr ("                    </SELECT>")
pr ("                    </td>")
pr ("                </tr>")
pr ("            </table>")
pr ("            </td>")
pr ("            <td width=""19"" background=""../portal_image/tab_bgd.gif"">&nbsp;</td>")
pr ("      </tr>")
pr ("      <tr>")
pr ("            <td width=""19"" height=""2""><img src=""../portal_image/tab_bg.gif"" width=""19"" height=""20""></td>")
pr ("            <td width=""200"" background=""../portal_image/tab_bgb.gif"" height=""2"">&nbsp;</td>")
pr ("            <td width=""19"" height=""2""><img src=""../portal_image/tab_bd.gif"" width=""19"" height=""20""></td>")
pr ("      </tr>")
pr ("    </table>")
pr ("<br>")
If Session("candidat_language") = "FR" Then
    pr ("<input type=""image"" value=""submit"" name=""search"" src=""../portal_image/rechercher.gif"" border=""0"">")
Else
    pr ("<input type=""image"" value=""submit"" name=""search"" src=""../portal_image/search-icon.gif"" border=""0"">")
End If
pr ("</form>")
pr ("</center>")
 End If
 
pr ("</body></html>")
On Error Resume Next
Set rs = Nothing
My_Conn.Close
Set My_Conn = Nothing

End Function

Public Sub Fusionner(a, p, q, r, Z)

 Dim b(100) As Integer
 Dim Y(100) As String
Dim i, j, k As Integer
Dim c As Integer
c = 0
For i = p To q
    b(i) = a(i)
    Y(i) = Z(i)
Next i
For j = q + 1 To r
    b(r + q + 1 - j) = a(j)
    Y(r + q + 1 - j) = Z(j)
Next j
i = p
j = r
For k = p To r
    c = c + 1
    If (b(i) < b(j)) Then
        a(k) = b(i)
        Z(k) = Y(i)
        i = i + 1
    Else
        a(k) = b(j)
        Z(k) = Y(j)
        j = j - 1
    End If
Next k

End Sub
Public Sub Tri_Fusion(a, p, r, Z)

Dim q, n1, n2, n3 As Integer

If (p < r) Then
     q = Int((p + r) / 2)
    Tri_Fusion a, p, q, Z
    Tri_Fusion a, q + 1, r, Z
    Fusionner a, p, q, r, Z
End If

End Sub

Public Sub rel_seek()
Dim my_catid
CheckConnexion
my_catid = Request.QueryString("Category")
my_mode = Request.QueryString("affiche")

Select Case my_mode
    
    Case "help"
        FormCaption = "Create new Related Links to display into other category"
        InitForm "ADM_TOP_ADD_REL", "Portal.ASP?mode=relatedlink&affiche=list&Category=" & my_catid
        CreateField "word", "Seek for this words into other categories", ""
        EndForm2 ("search-icon")
    Case "list"
        'idem que la recherche
                
        Dim my_father_rep, my_html, my_desc As String, my_id, my_father_id, my_note As Long
        Dim lst_note(100) As Integer
        Dim lst_res(100) As String
        CheckConnexion
        
            If Request.Form("word") <> "" Then
                SearchPattern = Request.Form("word")
                keywords = Split(SearchPattern, ",")
                keycnt = UBound(keywords)
                Set rs1 = Server.CreateObject("ADODB.Recordset")        ' Categories
                Set rs = Server.CreateObject("ADODB.Recordset")         ' Topics
                SQL = "SELECT dbp_Categories.catid, dbp_Categories.catcaption, dbp_Categories.CatDescription, dbp_categories.cathtmlcontent  FROM ( dbp_CatUserList INNER JOIN dbp_Categories ON dbp_CatUserList.CatId = dbp_Categories.CatId)"
                SQL = SQL & " WHERE  (dbp_CatUserList.UserId=" & Session("candidat_USERId") & ") AND ("
                cnt = 0
                For Each word In keywords
                    SQL = SQL & " (dbp_Categories.CatDescription Like '%" & word & "%'"
                    SQL = SQL & " OR dbp_Categories.cathtmlcontent Like '%" & word & "%')"
                    If cnt < keycnt Then SQL = SQL & " AND "
                    cnt = cnt + 1
                Next
                SQL = SQL & ");"
                rs1.Open SQL, My_Conn
                flagfound = 0
                i = 0
                Do While Not rs1.EOF
                    flagfound = 1
                    my_id = rs1("catid")
                    my_html = Trim(rs1("cathtmlcontent"))
                    my_desc = Trim(rs1("catdescription"))
                    cnt = 0
                    keywords = Split(SearchPattern, ",")
                    keycnt = UBound(keywords)
                    
                    res = "<table width='600Px' border='0'><tr><td>"
                    res = res & "<p class='smallerheader'><input type='checkbox' name='CAT' value=" & rs1("catid") & ">" & rs1("Catcaption")
                    res = res & "("
                    For Each word In keywords
                        my_note = 0
                        If (Not IsNull(my_html)) And my_html <> "" Then
                            my_note = occurrence(UCase(my_html), UCase(word)) + my_note
                        End If
                        If (Not IsNull(my_desc)) And my_desc <> "" Then
                            my_note = occurrence(UCase(my_desc), UCase(word)) + my_note
                        End If
                        If cnt < keycnt Then
                            If cnt = 0 Then
                                res = res & word & ":" & my_note
                            Else
                                res = res & ", " & word & ":" & my_note
                            End If
                        Else
                            If cnt > 0 Then
                                res = res & ", " & word & ":" & my_note & ")"
                            Else
                                res = res & word & ":" & my_note & ")"
                            End If
                        End If
                        lst_note(i) = lst_note(i) + my_note * (keycnt + 2 - cnt)
                        cnt = cnt + 1
                    Next
                    res = res & "</td></tr></table>"

                    lst_res(i) = res
                    i = i + 1
                    If i = 60 Then Exit Do
                    rs1.MoveNext
                Loop
                rs1.Close
                Set rs1 = Nothing
                SearchPattern = Request.Form("word")
                keywords = Split(SearchPattern, ",")
                keycnt = UBound(keywords)
                SQL = "SELECT dbp_Topics.Abstract, dbp_Categories.catid, dbp_Topics.Title "
                SQL = SQL & "FROM (dbp_CatUserList INNER JOIN dbp_Categories ON dbp_CatUserList.CatId = dbp_Categories.CatId) INNER JOIN dbp_Topics ON dbp_Categories.CatId = dbp_Topics.CatId"
                SQL = SQL & " WHERE  (dbp_CatUserList.UserId=" & Session("candidat_USERId") & ") AND "
                cnt = 0
                For Each word In keywords
                    SQL = SQL & " (dbp_Topics.Title Like '%" & word & "%'"
                    SQL = SQL & " OR dbp_Topics.Abstract Like '%" & word & "%')"
                    If cnt < keycnt Then SQL = SQL & " AND "
                    cnt = cnt + 1
                Next
                SQL = SQL & ";"
                rs.Open SQL, My_Conn
                Do While Not rs.EOF
                    flagfound = 1
                    my_id = rs("catid")
                    my_html = Trim(rs("Abstract"))
                    my_desc = Trim(rs("Title"))
                    my_father my_id, my_father_rep, my_father_id
                    cnt = 0
                    keywords = Split(SearchPattern, ",")
                    keycnt = UBound(keywords)
                    
                    res = "<table width='600Px' border='0'><tr><td>"
                    res = res & "<p class='smallerheader'><input type='checkbox' name='CAT' value=" & rs("catid") & ">" & my_desc
                    res = res & "("

                    For Each word In keywords
                        my_note = 0
                        If (Not IsNull(my_html)) And my_html <> "" Then
                            my_note = occurrence(UCase(my_html), UCase(word)) + my_note
                        End If
                        If (Not IsNull(my_desc)) And my_desc <> "" Then
                            my_note = occurrence(UCase(my_desc), UCase(word)) + my_note
                        End If
                        If cnt < keycnt Then
                            If cnt = 0 Then
                                res = res & word & ":" & my_note
                            Else
                                res = res & ", " & word & ":" & my_note
                            End If
                        Else
                            If cnt > 0 Then
                                res = res & ", " & word & ":" & my_note & ")"
                            Else
                                res = res & word & ":" & my_note & ")"
                            End If
                        End If
                        lst_note(i) = lst_note(i) + my_note * (keycnt + 2 - cnt)
                        cnt = cnt + 1
                    Next
                    res = res & "</td></tr></table>"
                    lst_res(i) = res
                    i = i + 1
                    If i = 100 Then Exit Do
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
                My_Conn.Close
                If flagfound <> 1 Then
                    Response.Write "<font face='" & DefaultFontFace & "'>Sorry ! No matches found in portal</font>"
                Else
                    Tri_Fusion lst_note, 0, i, lst_res
                    FormCaption = "Select all categories where you want the link appear"
                    InitForm "ADM_TOP_INS_REL", "Portal.ASP?mode=relatedlink&affiche=add&Category=" & my_catid
                    CreateNonEditField "ADM_TOP_INS_CAP", "Search Results for", Replace(SearchPattern, ",", " AND ")
                    For j = i To 0 Step -1
                        pr (lst_res(j))
                    Next
                    EndForm2 ("ADD")
                End If
                ' NO MATCH
            Else ' Search = ""
                pr ("<font class=""styl2""><center>" & translate("Aucun mot à rechercher !") & "</font></center>")
                pr ("<p align=""center"">")
                pr ("<a href=""javascript:history.go(-1)""><img src=""../portal_image/back.gif"" border=""0""></a>")
            End If '
            
        Case "add"
            For Each f In Split(Request.Form("CAT"), ",")
            
                Set rs1 = Server.CreateObject("ADODB.Recordset")
                
                SQL = "select * from dbp_categories where catid=" & my_catid & " and cid=" & Session("candidat_APPLICATION")
                rs1.Open SQL, My_Conn
                If Not rs1.EOF Then
                    Title = rs1("catcaption")
                    my_id = my_catid
                    my_father my_id, my_father_rep, my_father_id
                    my_url = "../portal_html/default.asp?mode=cat&catid=" & my_father_id & "&rep=" & my_father_rep & "&rep_haut=" & my_father_rep & "haut/&prev=home&current=cat&quickurl=" & my_catid
                End If
                rs1.Close
                
                ' AJOUT DANS dbp_topics
                '=============================================================
                ' Calcul de TopicId
                '=============================================================
                SQL = "select * from dbp_Topics where CId=" & Session("candidat_APPLICATION") & " order by TopicId"
                
                rs1.Open SQL, My_Conn
                Do Until rs1.EOF
                    NewTopicId = rs1("TopicID")
                    rs1.MoveNext
                Loop
                NewTopicId = NewTopicId + 1
                rs1.Close
                Set rs1 = Nothing
                
                notitle = "no"
                nosignature = "yes"
                
                SQL = "Insert into dbp_Topics (CId, TopicId, CatId, Title, Abstract, Type, Data, Url, ContentType, TopicStyle,NoTitle,NoSignature,AbstractBackground,TitleBackground,topicicon) "
                SQL = SQL & " Values (" & Session("candidat_APPLICATION") & "," & NewTopicId & "," & f & ",'" & Title & "',"
                SQL = SQL & "Null,8, 'TEXTE'  ,'" & my_url & "','-',0," & notitle & "," & nosignature & ","
                SQL = SQL & "Null,Null,Null)"
                My_Conn.Execute (SQL)
                
                SQL = "INSERT INTO dbp_Objects ( CId, ObjTypeId, ObjId, ObjCreated, ObjOwner, ObjLastChange )"
                SQL = SQL & " Values (" & Session("candidat_APPLICATION") & ",101," & NewTopicId & ",'" & Now() & "'," & Session("candidat_USERId") & ",'" & Now() & "');"
                My_Conn.Execute (SQL)
            Next
            Response.Redirect ("../Portal_asp/result.asp?message=Related+Link+successfully+added")
End Select

End Sub

Public Sub checkUser()

my_mode = Request.QueryString("affiche")
my_rep = "../portal_html" & Session("candidat_securerep")
isok = False
check = False

If my_mode = "securitycheck" Then
    old_lang = Session("candidat_language")
    my_dsn = Session("candidat_dsn")
    my_urlbase = Session("candidat_urlbase")
    'Response.End
    'Session.Abandon
    Session("candidat_dsn") = my_dsn
    Session("candidat_urlbase") = my_urlbase
    isok = securityLogin(Request.Form("login"), Request.Form("password"))
    check = True
    If isok Then Response.Write ("<htm><body onload=""top.location='" & Session("candidat_urlbase") & "/default.asp?mode=change';""></body></htm>")
End If

pr ("<html>")
pr ("<head>")
pr ("<link rel=STYLESHEET href=""PMainStyle1.asp?type=text"" type=""text/css"">")
pr ("</head>")
pr ("<body bgcolor=""#CBE7F2""><form action=""portal.Asp?mode=login&affiche=securitycheck"" method=""post"" >")
pr ("<center><table cellspacing=""0"" cellpadding=""0"" border=""0"">")
pr ("  <tr>")
pr ("    <td>")
pr ("      <table cellspacing=""0"" cellpadding=""0"" border=""0"">")
pr ("        <tr>")
pr ("          <td><img  src=""" & my_rep & "log_r1_c1.gif"" name=""log_r1_c1"" width=""219"" height=""190"" border=""0""></td>")
pr ("        </tr>")
pr ("        <tr>")
pr ("          <td>")
pr ("            <table cellspacing=""0"" cellpadding=""0"" border=""0"">")
pr ("              <tr>")
pr ("                <td><img  src=""" & my_rep & "log_r4_c1.jpg"" name=""log_r4_c1"" width=""58"" height=""130"" border=""0""></td>")
pr ("                <td>")
pr ("                  <table cellspacing=""0"" cellpadding=""0"" border=""0"">")
pr ("                    <tr bgcolor=""#F7C4CD"">")
pr ("                      <td width=""161"" height=""27"" border=""0"">")
pr ("                        <input type=""text"" name=""login"" size=""15"" maxlength=""15"">")
pr ("                      </td>")
pr ("                    </tr>")
pr ("                    <tr>")
pr ("                      <td><img  src=""" & my_rep & "log_r5_c2.gif"" name=""log_r5_c2"" width=""161"" height=""13"" border=""0""></td>")
pr ("                    </tr>")
pr ("                    <tr bgcolor=""#F7C4CD"">")
pr ("                      <td width=""161"" height=""31"" border=""0"">")
pr ("                        <input type=""password"" name=""password"" size=""15"" maxlength=""15"">")
pr ("                      </td>")
pr ("                    </tr>")
pr ("                    <tr>")
pr ("                      <td><img  src=""" & my_rep & "log_r7_c2.jpg"" name=""log_r7_c2"" width=""161"" height=""59"" border=""0""></td>")
pr ("                    </tr>")
pr ("                  </table>")
pr ("                </td>")
pr ("              </tr>")
pr ("            </table>")
pr ("          </td>")
pr ("        </tr>")
pr ("      </table>")
pr ("    </td>")
pr ("    <td>")
pr ("      <table cellspacing=""0"" cellpadding=""0"" border=""0"">")
pr ("        <tr>")
pr ("          <td>")
pr ("            <table cellspacing=""0"" cellpadding=""0"" border=""0"">")
pr ("              <tr>")
pr ("                <td><img  src=""" & my_rep & "log_r1_c3.jpg"" name=""log_r1_c3"" width=""178"" height=""132"" border=""0""></td>")
pr ("              </tr>")
pr ("              <tr>")
pr ("                <td>")
pr ("                  <table cellspacing=""0"" cellpadding=""0"" border=""0"">")
pr ("                    <tr>")
pr ("                      <td><input type=""image"" src=""" & my_rep & "log_r2_c3.gif"" width=""48"" height=""45"" border=""0""></td>")
pr ("                      <td><img  src=""" & my_rep & "log_r2_c4.jpg"" name=""log_r2_c4"" width=""130"" height=""45"" border=""0""></td>")
pr ("                    </tr>")
pr ("                  </table>")
pr ("                </td>")
pr ("              </tr>")
pr ("              <tr>")
pr ("                <td><img  src=""" & my_rep & "log_r3_c3.jpg"" name=""log_r3_c3"" width=""178"" height=""13"" border=""0""></td>")
pr ("              </tr>")
pr ("            </table>")
pr ("          </td>")
pr ("        </tr>")
pr ("        <tr>")
pr ("          <td><img  src=""" & my_rep & "log_r4_c3.jpg"" name=""log_r4_c3"" width=""178"" height=""130"" border=""0""></td>")
pr ("        </tr>")
pr ("      </table>")
pr ("    </td>")
pr ("  </tr>")
If check And Not isok Then pr ("<tr><td></td><td><font class=ecraserouge>" & translate("Identifiants erronés") & "</font></td></tr>")
pr ("</table></center></form>")
pr ("</body>")
pr ("</html>")


End Sub


Public Function generate_pw() As String

Y = 0
Z = 0
my_pass = ""
Do
Y = Y + 1
Z = Z + 1
    Randomize
    Ran = CInt(Rnd * 1)
        If Ran = 0 Then
            If CaseT = True Then
                Randomize
                Ran = CInt(Rnd * 25) + 97
                my_pass = my_pass & UCase(Chr(Ran))
            ElseIf CaseT = False Then
                Randomize
                Ran = CInt(Rnd * 25) + 97
                my_pass = my_pass & Chr(Ran)
            End If
        ElseIf Ran = 1 Then
            Randomize
            Ran = CInt(Rnd * 9)
            my_pass = my_pass & Ran
        End If
Loop Until Y = 8 And Z = 8
generate_pw = my_pass
End Function


Public Sub cat_root(res, my_parent, my_level)

Dim SQL1 As String
Dim rs1

        SQL1 = "SELECT dbp_CatUserList.CatId, dbp_Categories.CatCaption, dbp_Categories.CatPath,dbp_Categories.Parent, dbp_Categories.CatHasChilds, dbp_Categories.CatLevel, dbp_Categories.CatSort, dbp_Categories.CatExtURL "
        SQL1 = SQL1 & "FROM ((dbp_Clients INNER JOIN dbp_Users ON dbp_Clients.CId = dbp_Users.CId) INNER JOIN dbp_CatUserList ON dbp_Users.UserId = dbp_CatUserList.UserId) "
        SQL1 = SQL1 & "INNER JOIN dbp_Categories ON dbp_CatUserList.CatId = dbp_Categories.CatId "
        SQL1 = SQL1 & "WHERE (((dbp_Categories.Parent)=" & my_parent & ") AND ((dbp_Categories.CatLevel)=" & my_level & ") AND ((dbp_Clients.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_Users.UserId)=" & Session("candidat_USERId") & ")) "
        SQL1 = SQL1 & "ORDER BY dbp_Categories.CatSort;"
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open SQL1, My_Conn
        s = ""
        For j = 1 To (my_level - 1)
            s = s & "-"
        Next j
        Do Until rs1.EOF
            
            If my_parent = 0 Then
                res = res & "<option value='" & rs1("CatId") & "'>" & rs1("CatCaption") & "</option>" & vbCrLf
            Else
                res = res & "<option value='" & rs1("CatId") & "'>" & s & "> " & rs1("CatCaption") & "</option>" & vbCrLf
            End If
            If rs1("cathaschilds") > 0 Then Call cat_root(res, rs1("catid"), my_level + 1)
        
        rs1.MoveNext
        Loop
        rs1.Close

    
    Set rs1 = Nothing
    
End Sub

Public Sub cat_root2(res, my_parent, my_level, selectedoption)

Dim SQL1 As String
Dim rs1
        SQL1 = "SELECT dbp_CatUserList.CatId, dbp_Categories.CatCaption, dbp_Categories.CatPath,dbp_Categories.Parent, dbp_Categories.CatHasChilds, dbp_Categories.CatLevel, dbp_Categories.CatSort, dbp_Categories.CatExtURL "
        SQL1 = SQL1 & "FROM ((dbp_Clients INNER JOIN dbp_Users ON dbp_Clients.CId = dbp_Users.CId) INNER JOIN dbp_CatUserList ON dbp_Users.UserId = dbp_CatUserList.UserId) "
        SQL1 = SQL1 & "INNER JOIN dbp_Categories ON dbp_CatUserList.CatId = dbp_Categories.CatId "
        SQL1 = SQL1 & "WHERE (((dbp_Categories.Parent)=" & my_parent & ") AND ((dbp_Categories.CatLevel)=" & my_level & ") AND ((dbp_Clients.CId)=" & Session("candidat_APPLICATION") & ") AND ((dbp_Users.UserId)=" & Session("candidat_USERId") & ")) "
        SQL1 = SQL1 & "ORDER BY dbp_Categories.CatSort;"
        Set rs1 = Server.CreateObject("ADODB.Recordset")
        rs1.Open SQL1, My_Conn
        s = ""
        For j = 1 To (my_level - 1)
            s = s & "-"
        Next j
        Do Until rs1.EOF
            
            If my_parent = 0 Then
                If selectedoption = rs1("CatId") Then
                        res = res & "<option selected value='" & rs1("CatId") & "'>" & rs1("CatCaption") & "</option>" & vbCrLf
                Else
                        res = res & "<option value='" & rs1("CatId") & "'>" & rs1("CatCaption") & "</option>" & vbCrLf
                End If
            Else
                If selectedoption = rs1("CatId") Then
                        res = res & "<option selected value='" & rs1("CatId") & "'>" & s & "> " & rs1("CatCaption") & "</option>" & vbCrLf
                Else
                        res = res & "<option value='" & rs1("CatId") & "'>" & s & "> " & rs1("CatCaption") & "</option>" & vbCrLf
                End If
            End If
            If rs1("cathaschilds") > 0 Then Call cat_root2(res, rs1("catid"), my_level + 1, selectedoption)
        
        rs1.MoveNext
        Loop
        rs1.Close

    
    Set rs1 = Nothing
    
End Sub

Public Sub stand()

CheckConnexion

SQL = "select * from dbp_topics where title like 'Stand " & Request.QueryString("id") & "%';"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, My_Conn
If Not rs.EOF Then
pr ("<html>")
pr ("<head>")
pr ("<link rel=STYLESHEET href=""PMainStyle1.asp"" type=""text/css"">")
pr ("</head>")
pr ("<body background=""Background.asp""  BOTTOMMARGIN=""100px"">")
pr ("<welcome><center>")
pr ("</center>")
pr ("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""300px"" align=""left"" >")
pr ("    <tr>")
pr ("    <td align=""left""><img src=""../portal_image/tab_hg.gif"" width=""19px"" height=""17""></td>")
pr ("    <td ><img src=""../portal_image/tab_bgh.gif"" width=""100%"" height=""17""></td>")
pr ("    <td><img src=""../portal_image/tab_hd.gif"" width=""19px"" height=""17""></td>")
pr ("  </tr>")
pr ("    <tr>")
pr ("    <td background=""../portal_image/tab_bgg.gif"" height=""120"" width=""19px"">&nbsp;</td>")
pr ("    <td valign=""top"" class=""fondpage"" height=""120"">")
pr ("<table border='0' width=""100%"">")
pr ("  <tr>")
pr ("    <td>")
pr ("      <table align='left' border='0' width=""100%"">")
pr ("        <tr id=""Dheading"">")
pr ("          <td width=""100%"" id=""Dheading"">&nbsp;" & rs("title"))
pr ("          <br></td>")
pr ("        </tr>")
pr ("      </table>")
pr ("    </td>")
pr ("  </tr>")
pr ("  <tr>")
pr ("    <td width='100%' valign='top'>")
pr ("      <table align='center' border='0' width='100%' >")
pr ("       <tr>")
pr ("          <td>")
pr ("            <p id=""Dbody"">" & rs("abstract") & "<br>")
pr ("              <br><a href='javascript:window.close();'><img src='../portal_image/fermer.gif' border=0></a>")
pr ("          </td>")
pr ("        </tr>")
pr ("      </table>")
pr ("    </td>")
pr ("  </tr>")
pr ("</table>")
pr ("  </td>")
pr ("    <td  background=""../portal_image/tab_bgd.gif"" height=""120"" width=""19px"">&nbsp;</td>")
pr ("  </tr>")
pr ("  <tr>")
pr ("    <td align=""left""><img src=""../portal_image/tab_bg.gif"" width=""19px"" height=""17""></td>")
pr ("    <td ><img src=""../portal_image/tab_bgb.gif"" width=""100%"" height=""17""></td>")
pr ("    <td ><img src=""../portal_image/tab_bd.gif"" width=""19px"" height=""17""></td>")
pr ("  </tr>")
pr ("</table>")
pr ("</body>")
pr ("</html>")
End If
rs.Close
My_Conn.Close
End Sub
Public Sub reserve_stand()

CheckConnexion

SQL = "select * from dbp_topics where title like 'Stand " & Request.QueryString("id") & "%';"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, My_Conn
If Not rs.EOF Then
    my_abstract = rs("abstract")
    my_id = rs("topicid")
End If
rs.Close
SQL = "update dbp_topics set abstract='" & Mid(my_abstract, 1, InStr(1, my_abstract, "Disponible") - 1) & "Non Disponible' where topicid=" & my_id
My_Conn.Execute (SQL)
pr ("<html>")
pr ("<head>")
pr ("<link rel=STYLESHEET href=""PMainStyle1.asp"" type=""text/css"">")
pr ("</head>")
pr ("<body background=""Background.asp""  BOTTOMMARGIN=""100px"">")
pr ("<center><font class=texte><br><br>Votre réservation a bien été enregistrée et ajoutée à votre panier.</font>")
pr ("<br><br><a href='javascript:window.close();'><img src='../portal_image/fermer.gif' border=0></a>")
pr ("</center>")
pr ("</body>")
pr ("</html>")
My_Conn.Close
End Sub

Public Sub reserve_congres()

pr ("<html>")
pr ("<head>")
pr ("<link rel=STYLESHEET href=""PMainStyle1.asp"" type=""text/css"">")
pr ("</head>")
pr ("<body background=""Background.asp""  BOTTOMMARGIN=""100px"">")
pr ("<center><font class=texte><br><br>Votre réservation a bien été enregistrée.<br>Vous recevrez sous peu un email de confirmation</font>")
pr ("</center>")
pr ("</body>")
pr ("</html>")
End Sub

