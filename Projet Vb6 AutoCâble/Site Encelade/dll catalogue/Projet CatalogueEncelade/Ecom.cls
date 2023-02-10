VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ecom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub OnStartPage(PassedScriptingContext As ScriptingContext)
    Set MyScriptingContext = PassedScriptingContext
    Set Request = MyScriptingContext.Request
    Set Response = MyScriptingContext.Response
    Set Application = MyScriptingContext.Application
    Set Server = MyScriptingContext.Server
    Set Session = MyScriptingContext.Session

End Sub

Public Function Ecom_AjoutCaddie()

id_produit = Request("candidat_ItemWebId")
FieldRef = Request("candidat_fieldref")
ref_produit = Request("candidat_ref")
id_produit = Request("candidat_ref")
act = Request("act")




RacineImages = GetDefault("EcomPathImages", "../portal_ecom/img/", Session("candidat_ADOContact"))
    pr ("<html><head>")
    pr ("<link rel='stylesheet' href=""PMainStyle1.asp"">")
    pr ("<script language='javascript' src='../Portal_Java/Portal_Generic.js'></script>")
    pr ("<script language='javascript' src='../Portal_Ecom/ecom.js'></script>")
    pr ("</head><body>")
    pr ("<table border=""0"" width=""80%"" align=""center""><tr>")
    pr ("<td>" & ImgTransparent(1, 20) & "</td></tr>")
    pr ("<tr><td>")
    pr ("<IMG src=""" & GetDefault("EcomPathImages", "../portal_ecom/commander/img/", Session("candidat_ADOContact")) & "caddie.gif"" border=""0"" alt=""" & Text & """></td></tr>")
    pr ("<tr><td>")
    
    pr (ImgTitreEcommerce(translate("Votre devis")))

    pr ("</td></tr>")
    pr ("<tr><td>")
    
    Set Conn = Server.CreateObject("adodb.connection")
    Conn.Open Session("candidat_DSNcaddy")

    
    If CStr(act) = "" Then
        act = "aff"
    End If

    'If Len(CStr(Request("pognon"))) > 0 Then
    '    Session("candidat_pognon") = CStr(Request("pognon"))
    'End If
    
    'If Len(CStr(Session("candidat_pognon"))) < 1 Then
    '    Session("candidat_pognon") = "€"
    'End If
    
    'pognon = Session("candidat_pognon")
    
    'If Len(CStr(Request("last_page"))) > 0 Then
    '    Session("last_page") = Request("last_page")
    'End If
    
    
    id_caddie = Session("candidat_id_caddy")
    
    If Len(CStr(id_caddie)) < 1 Then
        Sql = "select max(id_caddie) from t_caddie"
        Set cur = Conn.Execute(Sql)
        id_caddie = inc(TOZS(cur(0)))
        Session("candidat_id_caddy") = id_caddie
        Set cur = Nothing
        Sql = "insert into t_caddie (id_caddie, id_produit, date_modif) values (" & id_caddie & ",-1, '" & Now() & " ')"
        Conn.Execute (Sql)
    End If
    
    
    
    If act = "ajout" Then
        qte_produit = Request("qte_produit")
        Sql = "select qte_produit from t_caddie where id_caddie = " & id_caddie & " and id_produit=" & id_produit
        Set cur = Conn.Execute(Sql)
        noprod = 1
        
        If Not cur.EOF Then
            nb = CInt(cur(0))
            noprod = 0
        End If
        Set cur = Nothing
        
        If noprod = 1 Then
            Sql = "insert into t_caddie (id_caddie, id_produit, qte_produit) values (" & id_caddie & "," & id_produit & "," & qte_produit & ")"
        Else
            nb = nb + qte_produit
            Sql = "update t_caddie set qte_produit = " & nb & " where id_caddie=" & id_caddie & " And id_produit = " & id_produit
        End If
        Conn.Execute (Sql)
        act = "aff"
    End If
    
    
    If act = "modif" Then
        qte_produit = Request("qte_produit")
        Sql = "update t_caddie set qte_produit = " & qte_produit & " where id_caddie = " & id_caddie & " and id_produit = " & id_produit
        Conn.Execute (Sql)
        act = "aff"
    End If
    
    
    If act = "suppr" Then
        Sql = "delete from t_caddie where id_caddie=" & id_caddie & " And id_produit = " & id_produit
        Conn.Execute (Sql)
        act = "aff"
    End If
    

    If act = "aff" Then
        Sql = "select count(*) from t_caddie where  id_caddie=" & id_caddie
        Set cur = Conn.Execute(Sql)
        id_cpt = CInt(cur(0))
        
        Sql = " SELECT t_caddie.id_caddie as ID_caddy, t_caddie.id_produit as ID_produit, t_caddie.qte_produit,"
        Sql = Sql & " t_tva.tva_valeur as Tva, t_tva.tva_id, dbp_Topics.Title as Title, dbp_Topics.prix as Prix, dbp_Topics.prixttc as TTC, dbp_Topics.prix * 6.55957 as prixf"
        Sql = Sql & " FROM t_tva INNER JOIN (t_caddie INNER JOIN dbp_Topics ON t_caddie.id_produit = dbp_Topics.TopicId) ON t_tva.tva_id = dbp_Topics.tva_id "
        Sql = Sql & " where t_caddie.id_caddie=" & id_caddie
        
        Set cur = Conn.Execute(Sql)
        rien = 0
        cpt = 0
        'total = 0#
        'totalf = 0#
        
        If cur.EOF Then
            rien = 1
            Session("candidat_caddyrempli") = 0
            pr (" <div align=""center"" class=""smallheader"">" & translate("<br><br>Votre caddie est vide.") & "</span></div>")
            pr ("<br><br><br><br><br><br>")
            pr ("<div align=""left"">" & BoutonRetour(1, "") & "</div>")
        Else
            Session("candidat_caddyrempli") = 1
            
              pr ("<center>")
              pr ("<table border=""0"" width='100%'><tr><td>")
                
                  pr ("<table border=0 width='100%' cellspacing='1' cellpadding='0'>")
                  pr ("<tr nowrap class=""TrEntete"" height=""30"">")
                  pr ("<form name='form_caddie' action='portal.asp?Mode=AjoutCommander' method='post'>")
                    pr ("<input type='hidden' name='id_produit' value=''>")
                    pr ("<input type='hidden' name='qte_produit' value=''>")
                    pr ("<input type='hidden' name='act' value=''>")

                  pr ("<td class='TextTrEntete' valign=""top"">" & ImgEnteteColonne(translate("Produit")) & "</td>")
                  'pr ("<td class='TextTrEntete' valign=""top"" align=""center"">" & ImgEnteteColonne(translate("Prix TTC")) & "</td>")
                  pr ("<td class='TextTrEntete' valign=""top"" align=""center"">" & ImgEnteteColonne(translate("Qté")) & "</td>")
                  'pr ("<td class='TextTrEntete' valign=""top"" align=""center"">" & ImgEnteteColonne(translate("Sous-total")) & "</td>")
                  pr ("<td class='TextTrEntete' valign=""top"" align=""center"">" & ImgEnteteColonne(translate("Action")) & "</td></tr>")
                    
        End If
        
        While Not cur.EOF
            id_produit = cur("ID_produit")
            titre = cur("Title")
            qte_produit = cur("qte_produit")
            cpt = cpt + qte_produit
            'If cur("TTC") = 1 Then
            '    Prix = fmt(cur("Prix"))
            '    prixf = fmt(cur("prixf"))
            '    sous_total_e = CDbl(cur("qte_produit")) * CDbl(cur("Prix"))
            '    sous_total = CDbl(cur("qte_produit")) * CDbl(cur("prixf"))
            'Else
            '    Tva = 1 + (cur("tva") / 100)
            '    Prix = fmt(cur("Prix") * Tva)
            '    prixf = fmt(cur("prixf") * Tva)
            '    sous_total_e = (CDbl(cur("qte_produit")) * CDbl(cur("Prix")) * Tva)
            '    sous_total = (CDbl(cur("qte_produit")) * CDbl(cur("prixf")) * Tva)
            'End If
            'total = (total + sous_total_e)
            'totalf = (totalf + sous_total)
            
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'A MODIFIER - INTEGRER VALEUR CANDIDAT
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            TDbg = TDbg + 1
            If (TDbg / 2) = Int(TDbg / 2) Then
                bgcolor = GetDefault("EcomTdColor1", "#E3Efff", Session("candidat_ADOContact"))
            Else
                bgcolor = GetDefault("EcomTdColor2", "#D3DFEF", Session("candidat_ADOContact"))
            End If

                  pr ("<tr class=""smalltext"" height=""30"">")
                    pr ("<td bgcolor=""" & bgcolor & """>" & titre & "</td>")
                    'pr ("<td align=""center"" bgcolor=""" & bgcolor & """>")
                    'If pognon = "FF" Then
                    '    Response.Write (prixf)
                    'Else
                    '    Response.Write (Prix)
                        pr ("<td align='center' bgcolor='" & bgcolor & "'>")
                        pr ("<input type='text' name='qte_" & id_produit & "' maxlength=3 size=2 value='" & qte_produit & "'>")
                        pr ("<img src='" & RacineImages & "/plus_moins.gif' alt='+/- 1' usemap='#mainmap" & id_produit & "' border=0 align='absmiddle'>")
                        pr ("<map name='mainmap" & id_produit & "'>")
                        pr ("<area shape=RECT coords='0,0,8,9' href='javascript:plus_un(" & id_produit & ") '>")
                        pr ("<area shape=RECT coords='0,10,8,19' href='javascript:moins_un(" & id_produit & ") '>")
                        pr ("</map> </td><td align='center' bgcolor='" & bgcolor & "'>")
                    'End If
                    
                    'If pognon = "FF" Then
                    '    Response.Write (fmt(sous_total))
                    'Else
                    '     Response.Write (fmt(sous_total_e))
                         pr ("<td valign='middle' align='center'  bgcolor='" & bgcolor & "'>&nbsp;<a href='javascript:modifie(" & id_produit & ") '><img src='" & RacineImages & "/mise_jour.gif' alt='Mise à jour' border=0></a>&nbsp;&nbsp;&nbsp;")
                        pr ("<a href='javascript:suppr(" & id_produit & ") '><img src='" & RacineImages & "/delete_small.gif' alt='Supprimer' border=0></a></td></tr>")
                    'End If
            cur.MoveNext
        Wend
        TDbg = 0
        
        
        If rien = 0 Then
        '          pr ("<tr class=""TrEntete"" height=""30"">")
        '            pr ("<td class='TextTrEntete' colspan='2' height='29'>" & translate("Total") & ":</td>")
        '            pr ("<td class='TextTrEntete' align='middle' height='29'>" & cpt & "</td>")
        '            pr ("<td class='TextTrEntete' align='middle' height='29'>")
        '             Response.Write (fmt(total))
        '            pr ("</td><td height='29' class='TextTrEntete'>&nbsp;</td></tr>")
                    ViewBasPageCaddy
                    pr ("</table>")
              pr ("</td></tr></table></center></form>")
        End If
    End If
    Conn.Close
    Set Conn = Nothing
  pr ("</td></tr></table>")
  
  pr ("</BODY></HTML>")


End Function


Public Function Ecom_ValiderCommande()

    Set Conn = Server.CreateObject("adodb.connection")
    Conn.Open Session("candidat_DSNcaddy")
    
    id_caddie = Session("candidat_id_caddy")
    
    Session("candidat_caddy_etape") = Request("etape")
    If Len(CStr(Session("candidat_caddy_etape"))) < 1 Then
        Session("candidat_caddy_etape") = 1
    End If
    
    'pognon = Session("candidat_pognon")
    
    pr ("<head>")
    pr ("<link rel=STYLESHEET href='PMainStyle1.asp' type='text/css'>")
    pr ("<script language='javascript' src='../Portal_Ecom/ecom.js'></script>")
    pr ("<script language='javascript' src='../Portal_Java/Portal_Generic.js'></script>")
'    pr ("<style>")
'    pr ("body         {")
 '   pr ("font-family: verdana;")
'    pr ("font-size:xx-small;")
'    pr ("font-weight: normal;")
'    pr ("background-color: White;")
'    pr ("margin: 0px;")
'    pr ("background : url(../flyway/img/bgciel.jpg);")
'    pr ("}")
'    pr ("</style>")
    pr ("</head>")


pr ("<BODY><br><br>")
ViewHautPage
pr ("<table width=""80%"" border=""0"" cellpadding='0' cellspacing='0' align='center'>")

pr ("<tr class='smallerheader'><td colspan='3'>")

    Select Case Session("candidat_caddy_etape")
        Case 1
            image = "etape1.gif"
            Text = translate("Frais de port")
        Case 2
            image = "etape2.gif"
            Text = translate("Commande finale")
        Case 3
            image = "etape3.gif"
            Text = translate("Identification")
        Case 4
            image = "etape4.gif"
            Text = translate("Vos coordonnées")
        Case 5
            image = "etape5.gif"
            Text = translate("Paiement")
    End Select
'    If Session("candidat_EcomPathImages") = "" Then Session("candidat_EcomPathImages") = GetDefault("EcomPathImages", "../portal_ecom/commander/img/")
'pr ("<IMG src=""" & Session("candidat_EcomPathImages") & image & """ border=""0"" alt=""" & Text & """ usemap=""#etapes""><br>")
'pr ("<map name=""etapes"">")
'    If Session("candidat_caddy_etape") > 4 Then pr ("<area shape=""rect"" coords=""560,143,659,160"" href=""portal.asp?mode=validercommande&etape=5"" target=""_self"">")
'    If Session("candidat_caddy_etape") > 3 Then pr ("<area shape=""rect"" coords=""476,143,558,160"" href=""portal.asp?mode=validercommande&etape=4"" target=""_self"">")
'    If Session("candidat_caddy_etape") > 2 Then pr ("<area shape=""rect"" coords=""387,143,475,160"" href=""portal.asp?mode=validercommande&etape=3"" target=""_self"">")
'    If Session("candidat_caddy_etape") > 1 Then pr ("<area shape=""rect"" coords=""294,143,385,160"" href=""portal.asp?mode=validercommande&etape=2"" target=""_self"">")
'    If Session("candidat_caddy_etape") > 0 Then pr ("<area shape=""rect"" coords=""201,144,292,159"" href=""portal.asp?mode=validercommande&etape=1"" target=""_self"">")
'pr ("</map>")
'pr ("<br><br>")
pr (ImgTitreEcommerce(Text))
                
'pr ("Etape " & Session("candidat_caddy_etape") & " : " & Text & " <br><br>")
'pr ("<br><br>" & Text & " <br><br>")
'pr ("------------------------------------------------------<br>")
pr ("</td></tr>")

pr ("<tr><td>")
  


    Select Case Session("candidat_caddy_etape")
        Case 1
            Call Ecom_ValiderCommande1
        Case 2
            Call Ecom_ValiderCommande2
        Case 3
            Call Ecom_ValiderCommande3
        Case 4
            Call Ecom_ValiderCommande4
        Case 5
            Call Ecom_ValiderCommande5
    End Select

    pr ("</td></tr></table>")
    pr (ViewBasPage(0, "80%", "center"))
pr ("<br><br></BODY></HTML>")

    Conn.Close
    Set Conn = Nothing

End Function


Public Function Ecom_ValiderCommande_TableauRecap()
        
        Set Conn = Server.CreateObject("adodb.connection")
        Conn.Open Session("candidat_DSNcaddy")
        
        If Session("candidat_id_caddy") = "" Then Session("candidat_id_caddy") = 0
        
        Sql = " SELECT t_caddie.id_caddie as ID_caddy, t_caddie.id_produit as ID_produit, t_caddie.qte_produit,"
        Sql = Sql & " t_tva.tva_valeur as Tva, t_tva.tva_id, dbp_Topics.Title as Title, dbp_Topics.prix as Prix, dbp_Topics.prixttc as TTC, dbp_Topics.prix * 6.55957 as prixf"
        Sql = Sql & " FROM t_tva INNER JOIN (t_caddie INNER JOIN dbp_Topics ON t_caddie.id_produit = dbp_Topics.TopicId) ON t_tva.tva_id = dbp_Topics.tva_id "
        Sql = Sql & " where t_caddie.id_caddie=" & Session("candidat_id_caddy")
        Set cur = Conn.Execute(Sql)
        rien = 0
        cpt = 0
        'total = 0#
        'totalf = 0#
        
        pr ("<center><table border=""0"" width=""100%"">")
            pr ("<tr><td class=""smallerheader"" colspan=""4"">")
            
            pr ("<table border=0 width=""100%"" cellpadding=""0"" cellspacing=""1""1><tr class=""TrEntete"" align=""top"">")
                pr ("<td class=""TextTrEntete"" valign=""top"">" & ImgEnteteColonne(translate("Désignation")) & "</td>")
        '        pr ("<td class=""TextTrEntete"" align='center' valign=""top"">" & ImgEnteteColonne(translate("Prix")) & "</td>")
                pr ("<td class=""TextTrEntete"" valign=""top"">" & ImgEnteteColonne(translate("Quantité")) & "</td>")
        '        pr ("<td class=""TextTrEntete"" align='center' valign=""top"">" & ImgEnteteColonne(translate("Sous-total")) & "</td></tr>")
                    
                
      While Not cur.EOF
      
            id_produit = cur("ID_produit")
            titre = cur("Title")
            qte_produit = cur("qte_produit")
            cpt = cpt + qte_produit
'            If cur("TTC") = 1 Then
'                Prix = fmt(cur("Prix"))
'                prixf = fmt(cur("prixf"))
'                sous_total_e = CDbl(cur("qte_produit")) * CDbl(cur("Prix"))
'                sous_total = CDbl(cur("qte_produit")) * CDbl(cur("prixf"))
'            Else
'                Tva = 1 + (cur("tva") / 100)
'                Prix = fmt(cur("Prix") * Tva)
'                prixf = fmt(cur("prixf") * Tva)
'                sous_total_e = (CDbl(cur("qte_produit")) * CDbl(cur("Prix")) * Tva)
'                sous_total = (CDbl(cur("qte_produit")) * CDbl(cur("prixf")) * Tva)
'            End If
'            total = (total + sous_total_e)
'            totalf = (totalf + sous_total)
            
            TDbg = TDbg + 1
            If (TDbg / 2) = Int(TDbg / 2) Then
                bgcolor = GetDefault("EcomTdColor1", "#E3Efff", Session("candidat_ADOContact"))
            Else
                bgcolor = GetDefault("EcomTdColor2", "#D3DFEF", Session("candidat_ADOContact"))
            End If

            
                pr ("<tr class='smalltext' Bgcolor=" & bgcolor & " height=""30"">")
                  pr ("<td>" & titre & "</td>")
'                  pr ("<td align='center'>" & Prix & "</td>")
                  pr ("<td align='center'>" & qte_produit & "</td>")
'                  pr ("<td align='center'>" & fmt(sous_total_e) & "</td>")
                  pr ("</tr>")
            cur.MoveNext
        Wend
          TDbg = 0
          
'          pr ("<tr class=""TrEntete"" height=""30"">")
'          pr ("<td class=""TextTrEntete"">" & translate("Total") & ":</td>")
'            pr ("<td align='center' class=""TExtTrEntete"">&nbsp;</td>")
'            pr ("<td class=""TextTrEntete"" align='center'>" & cpt & "</td>")
'            pr ("<td align='center' class=""TextTrEntete"">" & fmt(total) & "</td></tr>")
            
'            pr ("</table>")
        
    Conn.Close
    Set Conn = Nothing
        
End Function



Public Function Ecom_ValiderCommande_TableauRecap2()
        Set Conn = Server.CreateObject("adodb.connection")
        Conn.Open Session("candidat_DSNcaddy")

        Sql = " SELECT t_caddie.id_caddie as ID_caddy, t_caddie.id_produit as ID_produit, dbp_Topics.poids as poids, t_caddie.qte_produit,"
        Sql = Sql & " t_tva.tva_valeur as Tva, t_tva.tva_id, dbp_Topics.Title as Title, dbp_Topics.prix as Prix, dbp_Topics.prixttc as TTC, dbp_Topics.prix * 6.55957 as prixf"
        Sql = Sql & " FROM t_tva INNER JOIN (t_caddie INNER JOIN dbp_Topics ON t_caddie.id_produit = dbp_Topics.TopicId) ON t_tva.tva_id = dbp_Topics.tva_id "
        Sql = Sql & " where t_caddie.id_caddie=" & Session("candidat_id_caddy")
        Set cur = Conn.Execute(Sql)
        rien = 0
        cpt = 0
'        total = 0#
'        totalf = 0#
        
        leStr = "<center>"
        leStr = leStr & "<table border=""0"" width=""100%"" cellpadding=""0"">"
        leStr = leStr & "<tr class=""TrEntete"" height=""25"" valign=""top""> "
        leStr = leStr & "<td class=""TextTrEntete"" width=""40%"">" & ImgEnteteColonne(translate("Désignation")) & "</td>"
'        leStr = leStr & "<td class=""TextTrEntete"" width=""20%"" align=""center"">" & ImgEnteteColonne(translate("Prix TTC")) & " " & pognon & "</td>"
        leStr = leStr & "<td class=""TextTrEntete"" width=""20%"" align=""center"">" & ImgEnteteColonne(translate("Qté")) & "</td>"
'        leStr = leStr & "<td class=""TextTrEntete"" width=""20%"" align=""center"">" & ImgEnteteColonne(translate("Sous-total")) & " " & pognon & "</td></tr>"

        While Not cur.EOF
            id_produit = cur("ID_produit")
            titre = cur("Title")
            qte_produit = CInt(cur("qte_produit"))
            cpt = cpt + qte_produit
'            If cur("TTC") = 1 Then
'                Prix = fmt(cur("Prix"))
'                prixf = fmt(cur("prixf"))
'                sous_total_e = (CDbl(cur("qte_produit")) * CDbl(cur("Prix")))
'                total = total + sous_total_e
'            Else
'                Tva = 1 + (cur("tva") / 100)
'                Prix = fmt(cur("Prix") * Tva)
'                prixf = fmt(cur("prixf") * Tva)
'                sous_total_e = sous_total_e + (CDbl(cur("qte_produit")) * CDbl(cur("Prix")) * Tva)
'                total = total + sous_total_e
'            End If
'            Poids = CInt(cur("poids")) * qte_produit
'            i_poids = i_poids + Poids

            TDbg = TDbg + 1
            If (TDbg / 2) = Int(TDbg / 2) Then
                bgcolor = GetDefault("EcomTdColor1", "#E3Efff", Session("candidat_ADOContact"))
            Else
                bgcolor = GetDefault("EcomTdColor2", "#D3DFEF", Session("candidat_ADOContact"))
            End If
            leStr = leStr & "<tr class=""smalltext"" Bgcolor=" & bgcolor & " width=""30""> "
            leStr = leStr & "<td>" & titre & "</td>"
'            leStr = leStr & "<td align=""right"">" & Prix & "</td>"
            leStr = leStr & "<td align=""center"">" & qte_produit & "</td>"
'            leStr = leStr & "<td align=""right"">" & fmt(sous_total_e) & "</td>"
            leStr = leStr & "</tr>"
            cur.MoveNext
        Wend
        TDbg = 0
'                'Calcul frais de port
'                '====================
'                SQL = "select prix from  t_tarif_livraison  where poids_mini <= " & i_poids & " and poids_maxi > " & i_poids
'                Set cur = Conn.Execute(SQL)
'                If Not cur.EOF Then
'                    port_eu = CDbl(cur(0))
'                Else
'                    port_eu = 0
'                End If
'
'                SQL = "select label_pays, tarif_normal from t_pays where id_pays = " & Session("candidat_id_pays")
'                Set cur = Conn.Execute(SQL)
'                If Not cur.EOF Then
'                    coeff = CDbl(cur(1))
'                    label_pays = cur(0)
'                End If
'                Set cur = Nothing
'                If IsNull(coeff) Or (Len(CStr(coeff)) < 1) Then
'                    coeff = 1
'                End If
        
 '               port_eu = port_eu * coeff
 '               total = total + port_eu
 '               Session("total_eu") = CDbl(total)
        
'          leStr = leStr & "<tr class='smalltext' height=""25"">"
'          If port_eu = 0 Then
'            leStr = leStr & "<td>" & translate("Frais de port offerts") & "</td>"
'            leStr = leStr & "<td align=""right"" colspan=""3"">&nbsp;</td></tr>"
'          Else
'            leStr = leStr & "<td>" & translate("Frais de port") & " (<b>" & i_poids & "</b>g / " & label_pays & ") : </td>"
'            leStr = leStr & "<td colspan=""2"">&nbsp;</td>"
'            leStr = leStr & "<td align=""right"">" & fmt(port_eu) & "</td></tr>"
'          End If
          
'            leStr = leStr & "<tr class=""TrEntete"" height=""25"">"
'            leStr = leStr & "<td class=""TextTrEntete"">" & translate("Total TTC") & " (€)&nbsp;:</td>"
'            leStr = leStr & "<td width=""20%"" align=""right"" class=""TextTrEntete"">&nbsp;</td>"
'            leStr = leStr & "<td width=""20%"" align=""center"" class=""TextTrEntete"">" & cpt & "</td>"
'            leStr = leStr & "<td width=""20%"" align=""right"" class=""TextTrEntete"">" & fmt(total) & "</td></tr>"

    Conn.Close
    Set Conn = Nothing
    
    Ecom_ValiderCommande_TableauRecap2 = leStr
End Function


Public Function Ecom_ValiderCommande1()
        
  Set Conn = Server.CreateObject("adodb.connection")
  Conn.Open Session("candidat_DSNcaddy")
        
  pr (Ecom_ValiderCommande_TableauRecap())
    
  If Session("candidat_id_pays") = "" Then Session("candidat_id_pays") = 0

  pr ("<form name='form_etape" & Session("candidat_caddy_etape") & "' action='portal.asp?mode=validercommande' method='post'>")
        Session("candidat_caddy_etape") = 1
        etapesuivante = Session("candidat_caddy_etape") + 1
        pr ("<input type='hidden' name='etape' value='" & etapesuivante & "'>")
        pr ("<tr>")
        pr ("<td rowspan=""2"" align=""right""><img src=""" & Session("candidat_EcomPathImages") & "carteFT.gif"" border=""0""></td>")
        pr ("<td class=""paysliv"" colspan=""2"" valign=""middle"">" & translate("Pays de Livraison") & ":<br></td>")
        pr ("<td align=""center"" valign=""top""><a href=""javascript:valide_" & Session("candidat_caddy_etape") & "()"">")
        pr ("<img src=""" & Session("candidat_EcomPathImages") & "poursuivrelacommande.gif"" border=""0""></a></td>")
        pr ("<tr><td valign=""top"">")
                pr ("<select name='id_pays'>")
                    pr ("<option value='' selected>[" & translate("Choisissez") & "]")
                    pr ("<option value='999'>mon pays ne figure pas dans la liste")
                id_pays = Session("candidat_id_pays")
                Sql = "select id_pays, label_pays from t_pays where livrable = 1 order by 2"
                Set cur = Conn.Execute(Sql)
                While Not cur.EOF
                    pr ("<option value='" & cur(0) & "'")
                    If CInt(Session("candidat_id_pays")) = CInt(cur(0)) Then
                        Response.Write ("selected")
                        End If
                    pr (">" & cur(1))
                    cur.MoveNext
                Wend
                  pr ("</select></td>")
          pr ("</tr></table>")
          pr ("</form>")
    
    pr ("</td></tr></table>")
    Conn.Close
    Set Conn = Nothing
        
End Function



Public Function Ecom_ValiderCommande2()
        
        Set Conn = Server.CreateObject("adodb.connection")
        Conn.Open Session("candidat_DSNcaddy")
        
        If Session("candidat_id_pays") = "" Then Session("candidat_id_pays") = Request("id_pays")
        If IsEmpty(Session("candidat_id_pays")) And Session("candidat_id_pays") = "" Then Session("candidat_id_pays") = "0"
        
If Session("candidat_id_pays") = "999" Or Session("candidat_id_pays") = "0" Then

            pr ("<span class='alert'>" & translate("Votre pays ne figure pas encore dans notre liste") & ".</span><br><br>")
            pr ("<span class='alert'>" & translate("Merci de nous contacter afin de connaître les frais de port correspondant à votre commande") & ".</span><br><br>")
            pr (translate("E-mail") & ": <a href='mailto:" & GetDefault("EcomEmailCommandes", "commandes@euxia.net", Session("candidat_ADOContact")) & "'>" & GetDefault("EcomEmailCommandes", "commandes@euxia.net", Session("candidat_ADOContact")) & "</a></span><BR><br><center>")
        
            Exit Function
Else
       pr (Ecom_ValiderCommande_TableauRecap2())

    'test pour sauter l'étape 3 d'dentification si le user est déjà authentifié.
    '---------------------------------------------------------------------------
    If Session("USERNAME") = "Guest" Then
        pr ("<form name='form_etape" & Session("candidat_caddy_etape") & "' action='portal.asp?mode=validercommande' method='post'>")
        pr ("<span class='smallheader'>")
        etapesuivante = Session("candidat_caddy_etape") + 1
        pr ("<input type='hidden' name='etape' value='" & etapesuivante & "'>")
        pr ("<input type='hidden' name='is_client' value=''>")
        pr ("<input type='hidden' name='client' value='0'>")
        pr ("<input type='hidden' name='compte_client' value='" & Session("USERNAME") & "'>")
        pr ("</form>")

            pr ("<tr><td colspan=""3"">&nbsp;</td>")
            pr ("<td align=""right""><a href='javascript:valide_" & Session("candidat_caddy_etape") & "()'>")
            pr ("<img src=""" & Session("candidat_EcomPathImages") & "poursuivrelacommande.gif"" border=""0""></a></td>")
        pr ("</tr></table></center>")
        

    Else
        pr ("<form name='form_etape" & Session("candidat_caddy_etape") & "' action='portal.asp?mode=validercommande' method='post'>")
        pr ("<span class='smallheader'>")
        etapesuivante = Session("candidat_caddy_etape") + 1
        pr ("<input type='hidden' name='etape' value='" & etapesuivante & "'>")
        pr ("<input type='hidden' name='client' value='1'>")
        pr ("<input type='hidden' name='compte_client' value='" & Session("USERNAME") & "'>")
        pr ("</form>")

            pr ("<tr><td colspan=""3"">&nbsp;</td>")
            pr ("<td align=""right""><a href='javascript:valide_" & Session("candidat_caddy_etape") & "()'>")
            pr ("<img src=""" & Session("candidat_EcomPathImages") & "poursuivrelacommande.gif"" border=""0""></a></td>")
        pr ("</tr></table></center>")

    End If
    

    
End If

    Conn.Close
    Set Conn = Nothing
        
End Function



Public Function Ecom_ValiderCommande3()
        etapesuivante = Session("candidat_caddy_etape") + 1

If Session("Username") = "guest" Then
    If Request("err") = 1 Then
            Msg_Err = "<span class=warn>"
            Msg_Err = "<b>Echec de vote identication</b>.<br>"
            Msg_Err = Msg_Err & " <i>Nous vous invitons à vérifier la saisie de votre accès.<br>"
            Msg_Err = Msg_Err & " Si l'erreur periste merci de contacter notre <a href=mailto:contact@flywayonline.com>webmasteur</a><br></i>"
            Msg_Err = Msg_Err & " </span><br>"
    End If
    pr ("<center>")
    pr ("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">")
    pr ("<tr>")
        pr ("<td width=""1""><img src=""" & Session("candidat_EcomPathImages") & "tab1_hg.gif"" width=""17"" height=""37""></td>")
        pr ("<td background=""" & Session("candidat_EcomPathImages") & "tab1_bgh.gif"" height=""2"">")
            pr ("<div align=""left""><img src=""" & Session("candidat_EcomPathImages") & "dejauncompte.gif""></div></td>")
        pr ("<td height=""2""><img src=""" & Session("candidat_EcomPathImages") & "tab1_hd.gif"" width=""24"" height=""37""></td></tr>")
    pr ("<tr>")
        pr ("<td background=""" & Session("candidat_EcomPathImages") & "tab1_bgg.gif""width=""1""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bgg.gif"" width=""17"" height=""8""></td>")
        pr ("<td bgcolor=""#ECF1F8"">")
        
        pr ("<FORM method=""post"" action=""portal.asp"" name=""form_etape" & Session("candidat_caddy_etape") & """>")
            pr ("<input type='hidden' name='mode' value='VALIDERCOMMANDE'>")
            pr ("<input type='hidden' name='etape' value='" & etapesuivante & "'>")
            pr ("<input type='hidden' name='act' value='login'>")
            
          pr ("<TABLE border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">")
            'pr ("<TR><TD colspan=""3"" class=""smallerheader"" width=""50%"" bgcolor=""#FFFFF8"">" & translate("Accès client") & "</TD></TR>")
            pr ("<TR class=""smalltext""><TD rowspan=""3"">" & Msg_Err)
            pr (translate("Nous vous invitons à saisir votre login et mot de passe") & ".</TD></TR>")
            pr ("<TR class=""smalltext""><TD>" & translate("login") & "</TD>")
            pr ("<TD><INPUT type=""text"" name=""c_login""></TD></TR>")
            pr ("<TR class=""smalltext""><TD>" & translate("mot de passe") & "</TD>")
            pr ("<TD><INPUT type=""password"" name=""c_password""></TD></TR>")
            pr ("<TR class=""smalltext""><TD align=""right"" colspan=""3"">")
            pr ("<a href='javascript:valide_" & Session("candidat_caddy_etape") & "()'><img src =""" & Session("candidat_EcomPathImages") & "valider.gif"" border=""0""></a>")
            pr ("</TR>")
          pr ("</TABLE>")
        pr ("</form>")
          
          pr ("</td>")
          pr ("<td background=""" & Session("candidat_EcomPathImages") & "tab1_bgd.gif"" width=""25""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bgd.gif"" width=""24"" height=""6""></td></tr>")
        pr ("<tr>")
          pr ("<td width=""1""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bg.gif"" width=""17"" height=""20""></td>")
          pr ("<td background=""" & Session("candidat_EcomPathImages") & "tab1_bgb.gif""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bgb.gif"" width=""9"" height=""20""></td>")
          pr ("<td width=""25""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bd.gif"" width=""24"" height=""20""></td>")
        pr ("</tr></table>")
        
pr ("<br><br>")

    pr ("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">")
        pr ("<tr>")
        pr ("<td width=""1""><img src=""" & Session("candidat_EcomPathImages") & "tab1_hg.gif"" width=""17"" height=""37""></td>")
        pr ("<td background=""" & Session("candidat_EcomPathImages") & "tab1_bgh.gif"" height=""2"" colspan=""2"">")
            pr ("<div align=""left""><img src=""" & Session("candidat_EcomPathImages") & "pasdecompte.gif""></div></td>")
        pr ("<td height=""2""><img src=""" & Session("candidat_EcomPathImages") & "tab1_hd.gif"" width=""24"" height=""37""></td></tr>")
        pr ("<tr bgcolor=""#ECF1F8"">")
        pr ("<td background=""" & Session("candidat_EcomPathImages") & "tab1_bgg.gif""width=""1""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bgg.gif"" width=""17"" height=""8""></td>")
        pr ("<td colspan=""2"">")
        
        pr ("<table border=""0"" cellpadding=""0"" cellspacing=""5"" width=""100%"">")
        pr ("<tr><td align=""left"" align=""50%"">")
            
            pr ("<form action=""portal.asp"" name=""form_etape" & Session("candidat_caddy_etape") & "b"">")
                pr ("<input type='hidden' name='mode' value='VALIDERCOMMANDE'>")
                pr ("<input type='hidden' name='etape' value='" & etapesuivante & "'>")
                pr ("<input type='hidden' name='act' value='devenirclient'>")
                
                pr ("<TABLE border=""0"" cellspacing=""0"" cellpadding=""2"" width=""80%"">")
                  'pr ("<TR class=""smallerheader"" bgcolor=""#FFFFF8""><TD>" & translate("Devenir client") & "</TD></TR>")
                  pr ("<TR class=""smalltext""><TD>" & translate("Si vous n'avez aucun accès et souhaitez devenir client") & ".</td></TR>")
                  pr ("<TR><TD align=""right""><a href='javascript:valide_" & Session("candidat_caddy_etape") & "b()'><img src=""" & Session("candidat_EcomPathImages") & "devenirclient.gif"" alt=""" & translate("Devenir client") & """ border=""0""></A></TD></TR>")
                pr ("</TABLE>")
            pr ("</form>")

        pr ("</td><td align=""left"" align=""50%"">")

            pr ("<form action=""portal.asp"" name=""form_etape" & Session("candidat_caddy_etape") & "c"">")
                pr ("<input type='hidden' name='mode' value='VALIDERCOMMANDE'>")
                pr ("<input type='hidden' name='etape' value='" & etapesuivante & "'>")
                pr ("<input type='hidden' name='act' value='commanderapide'>")
                pr ("<TABLE border=""0"" cellspacing=""0"" cellpadding=""2"" width=""80%"">")
                    'pr ("<TR class=""smallerheader"" bgcolor=""#FFFFF8""><TD>" & translate("Commande Rapide") & "</TD></TR>")
                    pr ("<TR class=""smalltext""><TD>" & translate("Si vous souhaitez uniquement demander un devis") & ".</td></TR>")
                    pr ("<TR><td align=""right""><a href='javascript:valide_" & Session("candidat_caddy_etape") & "c()'><img src=""" & Session("candidat_EcomPathImages") & "commanderapide.gif"" alt=""" & translate("Commande rapide") & """ border=""0""></A></TD></TR>")
                pr ("</TABLE>")
            pr ("</form>")
        
        pr ("</td></tr></table>")
        
        pr ("</td>")
        pr ("<td background=""" & Session("candidat_EcomPathImages") & "tab1_bgd.gif"" width=""25""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bgd.gif"" width=""24"" height=""6""></td></tr>")
        pr ("<tr>")
        pr ("<td width=""1""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bg.gif"" width=""17"" height=""20""></td>")
        pr ("<td colspan=""2"" background=""" & Session("candidat_EcomPathImages") & "tab1_bgb.gif""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bgb.gif"" width=""9"" height=""20""></td>")
        pr ("<td width=""25""><img src=""" & Session("candidat_EcomPathImages") & "tab1_bd.gif"" width=""24"" height=""20""></td>")
    pr ("</tr></table>")
    pr ("</center>")
Else
    Response.Redirect ("portal.asp?mode=VALIDERCOMMANDE&etape=" & etapesuivante)
End If
End Function


Public Function Ecom_ValiderCommande4()
    Set Conn = Server.CreateObject("adodb.connection")
    Conn.Open Session("candidat_DSNcaddy")
    
    'Vérification du login & mot de passe
    If Request.Form("act") = "login" Then
        Sql = "SELECT UserId, UserLogin FROM dbp_users "
        Sql = Sql & " WHERE UserLogin='" & Request("c_login") & "'"
        Sql = Sql & " AND UserEmailPass='" & Request("c_password") & "'"
        Set Rs = Conn.Execute(Sql)
        If Not Rs.BOF Then
            Session("username") = Rs("UserLogin")
            Session("UserId") = Rs("UserId")
            is_client = 1
            Session("NewUser") = 0
        Else
            Session("username") = "guest"
            Session("candidat_caddy_etape") = Session("candidat_caddy_etape") - 1
            Response.Redirect ("portal.asp?mode=validercommande&err=1&etape=" & Session("candidat_caddy_etape"))
            Session("NewUser") = 1
        End If
   End If

   
            If Session("username") <> "guest" Then
                Sql = "SELECT UserId FROM dbp_users WHERE  UserId = " & Session("UserId")
                Set cur = Conn.Execute(Sql)
                If Not cur.EOF Then
                    id_client = cur(0)
                    Set cur = Nothing
                    
                    Sql = "SELECT c.LastName, "     '0
                    Sql = Sql & " c.FirstName, "    '1
                    Sql = Sql & " c.Address, "      '2
                    Sql = Sql & " c.Zip, "          '3
                    Sql = Sql & " c.City, "         '4
                    Sql = Sql & " c.id_pays, "      '5
                    Sql = Sql & " p.label_pays, "   '6
                    Sql = Sql & " c.Email, "        '7
                    Sql = Sql & " c.id_livraison, " '8
                    Sql = Sql & " c.fld2, "         '9 - numéro FNA"
                    Sql = Sql & " c.fld3 "          '10 - societe
                    Sql = Sql & " FROM  dbp_UserInfos c, t_pays p "
                    Sql = Sql & " WHERE p.id_pays = c.id_pays "
                    Sql = Sql & " And c.UserID = " & id_client
                    Set cur = Conn.Execute(Sql)
                        If cur.BOF Then
                        Session("NewUser") = 1
                        Msg_Err = "Votre compte n'a pas été reconnu.<br>"
                        Msg_Err = Msg_Err & " Nous vous invitons à ressaisir vos coordonnées pour les mettre à jour.<br>"
                        'Msg_Err = Msg_Err & "A l'issue vous recevrez un nouvel accès.""
                            c_nom = Session("UserLast")
                            c_prenom = Session("UserFirst")
                            c_societe = ""
                            c_adresse = ""
                            c_codepostal = ""
                            c_ville = ""
                            c_id_pays = 0
                            c_pays = ""
                            c_email = Session("EMAIL")
                            c_numfna = ""
                            id_livraison = 0
                        Else
                        Session("NewUser") = 0
                            c_nom = Trim(cur(0))
                            c_prenom = Trim(cur(1))
                            c_societe = Trim(cur(10))
                            c_adresse = Trim(cur(2))
                            c_codepostal = Trim(cur(3))
                            c_ville = Trim(cur(4))
                            c_id_pays = cur(5)
                            c_pays = cur(6)
                            c_email = Trim(cur(7))
                            c_numfna = Trim(cur(9))
                            If cur(8) = "" Or IsEmpty(cur(8)) Then
                                id_livraison = 0
                            Else
                                id_livraison = cur(8)
                            End If
                            
                        End If
                    Set cur = Nothing

                    Sql = "select l.beneficiare_liv, "
                    Sql = Sql & " l.adresse_liv, "
                    Sql = Sql & " l.cp_liv, "
                    Sql = Sql & " l.ville_liv, "
                    Sql = Sql & " l.id_pays, "
                    Sql = Sql & " p.label_pays "
                    Sql = Sql & " from  t_livraison l, t_pays p "
                    Sql = Sql & " where l.id_livraison = " & id_livraison
                    Sql = Sql & " and   p.id_pays = l.id_pays"

                    Set cur = Conn.Execute(Sql)
                        If cur.BOF Then
                            l_beneficiaire = ""
                            l_adresse = ""
                            l_cp = ""
                            l_ville = ""
                            l_id_pays = 0
                        Else
                            l_beneficiaire = cur(0)
                            l_adresse = cur(1)
                            l_cp = cur(2)
                            l_ville = cur(3)
                            l_id_pays = cur(4)
                        End If
                Else
                    is_client = -1
                    
                End If
                Set cur = Nothing
            Else
                Sql = "select max(UserId) from dbp_users"
                Set cur = Conn.Execute(Sql)
                id_client = inc(cur(0))
                Set cur = Nothing

                Sql = "select max(id_livraison) from t_livraison"
                Set cur = Conn.Execute(Sql)
                id_livraison = inc(cur(0))

                Set cur = Nothing
                
                c_nom = Session("c_nom")
                c_prenom = Session("c_prenom")
                c_adresse = Session("c_adresse")
                c_codepostal = Session("c_codepostal")
                c_ville = Session("c_ville")
                c_id_pays = Session("c_id_pays")
                c_email = Session("c_email")
                l_beneficiaire = Session("l_beneficiaire")
                l_adresse = Session("l_adresse")
                l_cp = Session("l_cp")
                l_ville = Session("l_ville")
                l_id_pays = Session("candidat_id_pays")
                compte_client = Session("username")
                'passwd_client = Session("passwd_client")
                
            End If
            
                        
'            If is_client > -1 Then
            If Session("username") <> "guest" Then
                is_client = 1
                Session("NewUser") = 0
            Else
                is_client = 0
                Session("NewUser") = 1
            End If
            
                pr ("<form name='form_etape" & Session("candidat_caddy_etape") & "' action='portal.asp?mode=validercommande' method='post'>")
                etapesuivante = Session("candidat_caddy_etape") + 1
                pr ("<input type='hidden' name='etape' value='" & etapesuivante & "'>")
                pr ("<input type='hidden' name='is_client' value='" & is_client & "'>")
                pr ("<input type='hidden' name='id_client' value='" & id_client & "'>")
                pr ("<input type='hidden' name='act' value='" & Request("act") & "'>")
                pr ("<input type='hidden' name='id_livraison' value='" & id_livraison & "'>")
                pr ("<input type='hidden' name='compte_client' value='" & Session("username") & "'>")
                pr ("<input type='hidden' name='modifie' value='0'>")
            
            'MESSAGE ERREUR EVENTUEL
            '=======================
        If Msg_Err <> "" Then pr ("<span class=""smalltext"">" & Msg_Err & "</span><br>")
        
        pr ("<table border=0 cellpadding=0 width='100%'><tr class='smallheader'>")
            
            'SAISIE DES COORDONNEES PERSONNELLES
            '===================================
            pr ("<tr><td colspan=2 class=""smallheader"" >")
                pr ("<table border=""0"" width=""100%"" cellpadding=""0"" cellspacing=""0"">")
                    pr ("<tr class=""smallheader"">")
                    pr ("<td width=""50%"" background=""" & Session("candidat_EcomPathImages") & "entete_bg.gif"">&nbsp;</td>")
                    pr ("<td width=""1%"" align=""right"" valign=""top""><img src=""" & Session("candidat_EcomPathImages") & "entete1.gif""></td>")
                    pr ("<td align=""right"" background=""" & Session("candidat_EcomPathImages") & "entete2.gif"">" & translate("Vos coordonnées personnelles") & "&nbsp; ")
                    pr ("<td width=""1%"" align=""left"" valign=""top""><img src=""" & Session("candidat_EcomPathImages") & "entete3.gif""></td></tr>")
                    pr ("<tr height=""10""><td colspan=""4"">")
                    pr ("<img src=""" & Session("candidat_EcomPathImages") & "transparent.gif"" border=""0"" width=""1"" height=""10""></td></tr>")
                pr ("</table>")
            pr ("</td></tr>")
            pr ("<tr class='smalltext'><td class='smalltext' width='26%'>" & translate("Nom") & ":</td>")
            pr ("<td width='74%'><input type='text' name='c_nom' value='" & c_nom & "' maxlength=40  size=20 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Prénom") & ":</td>")
            pr ("<td width='74%'><input type='text' name='c_prenom' value='" & c_prenom & "' maxlength=40  size=20 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Société / aéroclub") & ":</td>")
            pr ("<td width='74%'><input type='text' name='c_societe' value='" & c_societe & "' maxlength=40  size=20 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Adresse") & ":</td>")
            pr ("<td width='74%'><input type='text' name='c_adresse' value='" & c_adresse & "' maxlength=150  size=50 class='champ'></td></tr>")
            pr ("<tr><td class='smalltext' width='26%'>" & translate("CP") & ":</td>")
            pr ("<td width='74%'><input type='text' name='c_codepostal' value='" & c_codepostal & "' maxlength=10  size=10 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Ville") & ":</td>")
            pr ("<td width='74%'><input type='text' name='c_ville' value='" & c_ville & "' maxlength=60  size=20 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Pays") & ":</td>")
            pr ("<td width='74%'>")
              pr ("<select name='c_id_pays' class='champ'>")
                pr ("<option value='' selected>[" & translate("Choisissez") & "]")

                        Sql = "select id_pays, label_pays from t_pays where livrable = 1 order by 2"
                        Set cur = Conn.Execute(Sql)
                        While Not cur.EOF
                            pr ("<option value=" & cur(0))
                                If (CInt(l_id_pays) = CInt(cur(0))) Or (CInt(Session("candidat_id_pays")) = CInt(cur(0))) Then
                                    pr ("SELECTED")
                                End If
                                pr (" >" & cur(1))
                            cur.MoveNext
                        Wend
                        Set cur = Nothing
                        pr ("</select></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("E-Mail") & ":</td>")
            pr ("<td width='74%'><input type='text' name='c_email' value='" & c_email & "' maxlength=60 size=20 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Numéro FNA") & ":</td>")
            pr ("<td width='74%'><input type='text' name='c_NumFNA' value='" & c_numfna & "' maxlength=60 size=20 class='champ'></td></tr>")
            pr ("<tr><td colspan=2 &nbsp; </td></tr>")


                If Request("act") = "devenirclient" Then
                'DONNE LOGIN ET MOT DE PASSE POUR LE NOUVEAU CLIENT
                '==================================================
                    pr ("<tr><td colspan=2 class=""smallheader"" >")
                        pr ("<table border=""0"" width=""100%"" cellpadding=""0"" cellspacing=""0"">")
                            pr ("<tr class=""smallheader"">")
                            pr ("<td width=""50%"" background=""" & Session("candidat_EcomPathImages") & "entete_bg.gif"">&nbsp;</td>")
                            pr ("<td width=""1%"" align=""right"" valign=""top""><img src=""" & Session("candidat_EcomPathImages") & "entete1.gif""></td>")
                            pr ("<td align=""right"" background=""" & Session("candidat_EcomPathImages") & "entete2.gif"">" & translate("Informations de connexion au site") & "&nbsp; ")
                            pr ("<td width=""1%"" align=""left"" valign=""top""><img src=""" & Session("candidat_EcomPathImages") & "entete3.gif""></td></tr>")
                            pr ("<tr height=""10""><td colspan=""4"">")
                            pr ("<img src=""" & Session("candidat_EcomPathImages") & "transparent.gif"" border=""0"" width=""1"" height=""10""></td></tr>")
                        pr ("</table>")
                    pr ("</td></tr>")
                    pr ("<tr><td class=""smalltext"">" & translate("Votre login") & ":</td>")
                    pr ("<td  class=""smalltext"">")
                    pr ("<input type=""text"" name=""c_login"" value="""" maxlength=16 size=20 class=""champ""></tr>")
                    pr ("<tr><td class=""smalltext"">" & translate("Mot de passe") & ":</td>")
                    pr ("<td  class=""smalltext"">")
                    pr ("<input type=""password"" name=""c_passwd1"" value="""" maxlength=16 size=20 class=""champ""></tr>")
                    pr ("<tr><td class=""smalltext"">" & translate("Confirmez") & ":</td>")
                    pr ("<td  class=""smalltext"">")
                    pr ("<input type='password' name='c_passwd2' value='' maxlength=16 size=20 class=""champ""></tr>")
                    pr ("<input type='hidden' name='passwd_client' value=''>")
                    pr ("<tr><td colspan=2>&nbsp;</td></tr>")
                Else
                    pr ("<tr><td colspan=2>&nbsp;")
                    pr ("<input type='hidden' name='c_login' value='" & Session("username") & "'>")
                    pr ("<input type='hidden' name='c_passwd1' value='********'>")
                    pr ("<input type='hidden' name='c_passwd2' value='********'>")
                    pr ("</td></tr>")
                End If
                
                
            ' COORDONNES DE LIVRAISON
            '========================
            pr ("<tr><td colspan=2 class=""smallheader"" >")
                pr ("<table border=""0"" width=""100%"" cellpadding=""0"" cellspacing=""0"">")
                    pr ("<tr class=""smallheader"">")
                    pr ("<td width=""50%"" background=""" & Session("candidat_EcomPathImages") & "entete_bg.gif"">&nbsp;</td>")
                    pr ("<td width=""1%"" align=""right"" valign=""top""><img src=""" & Session("candidat_EcomPathImages") & "entete1.gif""></td>")
                    pr ("<td align=""right"" background=""" & Session("candidat_EcomPathImages") & "entete2.gif"">" & translate("Vos coordonnées de Livraison") & "&nbsp; ")
                    pr ("<td width=""1%"" align=""left"" valign=""top""><img src=""" & Session("candidat_EcomPathImages") & "entete3.gif""></td></tr>")
                    pr ("<tr>")
                    pr ("<td colspan=""3"" align=""left""><span class='smallertext'><a href='javascript:recopie()'>")
                    pr ("<img src=""" & Session("candidat_EcomPathImages") & "copiercoller.gif"" border=""0"" al=""" & translate("Recopier les informations personnelles") & """>&nbsp;&nbsp;" & translate("Recopier les informations personnelles") & "</a></span></td>")
                    pr ("<td>&nbsp;</td></tr>")
                pr ("</table>")
            pr ("</td></tr>")
'            pr ("<tr class='smalltext'>")
'            pr ("<td width='26%' align=""right""><img src=""" & Session("candidat_EcomPathImages") & "cadeau.gif"" alt=""cadeau"">&nbsp;&nbsp;</td>")
'            pr ("<td width='74%'><input type='checkbox' value='1' name='l_kdo' class='champ'>" & translate("Cochez la case si cette commande est un cadeau") & "</td></tr>")
            pr ("<tr class='smalltext'>")
            pr ("<td width='26%'>" & translate("Bénéficiaire") & ":</td>")
            pr ("<td width='74%'><input type='text' name='l_beneficiaire' value='" & c_nom & "' maxlength=40  size=30 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Adresse") & ":</td>")
            pr ("<td width='74%'><input type='text' name='l_adresse' value='" & l_adresse & "' maxlength=150  size=50 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("CP") & ":</td>")
            pr ("<td width='74%'><input type='text' name='l_cp' value='" & l_cp & "' maxlength=10  size=10 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Ville:") & "</td>")
            pr ("<td width='74%'><input type='text' name='l_ville' value='" & l_ville & "' maxlength=60  size=20 class='champ'></td></tr>")
            pr ("<tr class='smalltext'><td width='26%'>" & translate("Pays") & ":</td><td width='74%'>")
                        Sql = "select label_pays from t_pays where id_pays = " & Session("candidat_id_pays")
                        Set cur = Conn.Execute(Sql)
                            pr (cur("label_pays"))
                        Set cur = Nothing
            pr ("<input type='hidden' name='l_id_pays' value='" & Session("candidat_id_pays") & "'></td></tr>")
            pr ("<tr><td colspan=2 &nbsp; </td></tr>")
'            pr ("<tr class='smalltext'>")
'            pr ("<td width='26%' valign=""top"">" & translate("Message du cadeau") & ":</td>")
'            pr ("<td width='74%'><textarea name=""l_kdotxt"" cols=""40"" rows=""5"" wrap=""physical"" class=""champs"">&nbsp;</textarea></td></tr>")
            
            pr ("<tr><td colspan=""2"">" & ImgTransparent(1, 30) & "</td><tr>")
              
'            'CONDITIONS GENERALES DE VENTES
'            '==============================
'            pr ("<tr class=""smallertext"">")
'            pr ("<td>&nbsp;</td>")
'            pr ("<td>")
'            pr ("<input type='checkbox' name='accepte'>&nbsp;" & translate("J'accepte les") & "<a href='" & GetDefault("EcomPathCGV", "../portal_ecom/commander/conditions/conditions.html") & "' target='_blank'>")
'            pr (translate("conditions générales de vente") & "</a>&nbsp;")
'            pr ("</td></tr>")
            
            'pr ("<tr><td colspan=""2"">" & ImgTransparent(1, 30) & "</td><tr>")
            
            pr ("</table>")
        
            'VALIDATION & COMMANDER
            '======================
        pr ("<table border=0 width='100%' cellpadding='2' cellspacing='0' class='smallheader'>")
        pr ("<td align=""right""><a href='javascript:valide_" & Session("candidat_caddy_etape") & "()'><img src=""" & Session("candidat_EcomPathImages") & "poursuivrelacommande.gif"" border=""0""></a></td>")
        pr ("</tr></table></center></form>")
            
            'INFORMATIONS LEGALES
            '====================
        pr ("<table border=""0"" width=""100%"" cellpadding=""2"" cellspacing=""0"" class=""smallertext""><tr>")
        pr ("<td align=""left"">")
        pr ("Les informations qui vous concernent sont destinées à Encelade. Elles ne seront pas transmises à des tiers.")
        pr ("Vous disposez d'un droit d'accès, de modification, de rectification et de suppression des données qui vous concernent (art. 34 de la loi Informatique et Libertés)" & ".")
        pr ("Pour l'exercer, contactez-nous à" & " <a href=""mailto:" & GetDefault("EcomEmailContact", "contact@euxia.net", Session("candidat_ADOContact")) & """>" & Session("EcomEmailContact") & "</a></td>")
        pr ("</tr></table>")
        
        
    Conn.Close
    Set Conn = Nothing
End Function



Public Function Ecom_ValiderCommande5()
    Set Conn = Server.CreateObject("adodb.connection")
    Conn.Open Session("candidat_DSNcaddy")
'           liste_vars()
            is_client = ChkString(Request("is_client"))
            id_client = Request("id_client")
            id_livraison = Request("id_livraison")
            c_nom = Trim(ChkString(Request("c_nom")))
            c_prenom = Trim(ChkString(Request("c_prenom")))
            c_societe = Trim(ChkString(Request("c_societe")))
            c_adresse = Trim(ChkString(Request("c_adresse")))
            c_codepostal = Trim(ChkString(Request("c_codepostal")))
            c_ville = Trim(ChkString(Request("c_ville")))
            c_id_pays = Trim(ChkString(Request("c_id_pays")))
            c_email = Trim(ChkString(Request("c_email")))
            l_beneficiaire = Trim(ChkString(Request("l_beneficiaire")))
            l_adresse = Trim(ChkString(Request("l_adresse")))
            l_cp = Trim(ChkString(Request("l_cp")))
            l_ville = Trim(ChkString(Request("l_ville")))
            l_id_pays = Trim(ChkString(Request("l_id_pays")))
'            l_kdo = CInt(Request("l_kdo"))
                If l_kdo <> 1 Then l_kdo = 0
'            l_kdotxt = Trim(ChkString(Request("l_kdotxt")))
            c_numfna = Trim(ChkString(Request("c_Numfna")))
            passwd_client = Trim(ChkString(Request("c_passwd1")))
            LOGIN = Trim(ChkString(Request("c_login")))
            
        If Session("username") = "guest" And Request("act") <> "commanderapide" Then
            'Check Login
            Sql = "SELECT UserLogin from dbp_users "
            Sql = Sql & " WHERE UserLogin='" & LOGIN & "'"
            'SQL = SQL & " OR UserEMail='" & Request("c_email") & "'""
            Set Rs = Conn.Execute(Sql)
            
            'MsgErr si identifiant déjà utilisé.
            '-----------------------------------
            If Not Rs.EOF Or Not Rs.BOF Then
                pr ("<div class=""alert"" align=""center"">" & translate("Vous devez donnez un autre login") & ".<br>")
                pr ("<i>[" & LOGIN & "]</i> " & translate("est actuellement utilisé") & ".<br>")
                pr (translate("Choisissez un nouvau login en cliquant sur Retour") & ".<br>")
                'pr ("<br><br>" & BoutonRetour(0, "") & " <div>")
                Exit Function
            End If
        End If
                
        If Session("NewUser") = 1 Then 'nouveau client
                    Sql = "select max(UserId) from dbp_users"
                    Set cur = Conn.Execute(Sql)
                    id_client = cur(0)
                    If IsNull(id_client) Then
                        id_client = 1
                    End If
                    
                    Sql = "select max(id_livraison) from t_livraison"
                    Set cur = Conn.Execute(Sql)
                    id_livraison = cur(0)
                    If IsNull(id_livraison) Then
                        id_livraison = 1
                    End If
                    
                    id_client = id_client + 1
                    id_livraison = id_livraison + 1
                    
                    Session("NewUser") = 0
                    Session("id_client") = id_client
                    Session("id_livraison") = id_livraison

                    
                    'Mise à jour de Dbp_Users
                    UserLogin = c_nom & " " & c_prenom
                    HashIt (passwd_client)
                    
                    SQL0 = "insert into dbp_users (CId, UserId, UserLogin, UserPass, UserLast, UserFirst, UserDisplayName, UserEMail, UserSecure, UserEmailPass, UserDateCreation) values ("
                    SQL0 = SQL0 & Session("APPLICATION") & ", "
                    SQL0 = SQL0 & id_client & ", "
                    SQL0 = SQL0 & noquote2(LOGIN) & ", "
                    SQL0 = SQL0 & noquote2(Session("HashData")) & ", "
                    SQL0 = SQL0 & noquote2(c_nom) & ", "
                    SQL0 = SQL0 & noquote2(c_prenom) & ", "
                    SQL0 = SQL0 & noquote2(UserLogin) & ", "
                    SQL0 = SQL0 & noquote2(c_email) & ", "
                    SQL0 = SQL0 & noquote2(Session("HashSecure")) & ", "
                    SQL0 = SQL0 & noquote2(passwd_client) & ","
                    SQL0 = SQL0 & noquote2(Now()) & ")"
                    Conn.Execute (SQL0)
                    
                    'Mise à jour de Dbp_UserInfos
                    SQL1 = "insert into Dbp_UserInfos (CId, UserID, id_livraison, LastName, FirstName, Address, Zip, City, id_pays, fld2, fld3, Email) values ("
                    SQL1 = SQL1 & Session("APPLICATION") & ", "
                    SQL1 = SQL1 & id_client & ", "
                    SQL1 = SQL1 & id_livraison & ", "
                    SQL1 = SQL1 & noquote2(c_nom) & ", "
                    SQL1 = SQL1 & noquote2(c_prenom) & ", "
                    SQL1 = SQL1 & noquote2(c_adresse) & ", "
                    SQL1 = SQL1 & noquote2(c_codepostal) & ", "
                    SQL1 = SQL1 & noquote2(c_ville) & ", "
                    SQL1 = SQL1 & c_id_pays & ", "
                    SQL1 = SQL1 & noquote2(c_numfna) & ", "
                    SQL1 = SQL1 & noquote2(c_societe) & ", "
                    SQL1 = SQL1 & noquote2(c_email) & ")"
                    Conn.Execute (SQL1)
                  
                    'Mise à jour de Dbp_GroupsPermission : Groupe Visiteur
                    sql4 = "insert into Dbp_GroupsPermission (CId, GroupID, ObjTypeId, ObjPermId, ObjId, ObjPermType) values ("
                    sql4 = sql4 & Session("APPLICATION") & ", "
                    sql4 = sql4 & GetDefault("EcomGroupIdVisiteurs", "2", Session("candidat_ADOContact")) & ", "
                    sql4 = sql4 & "50, "
                    sql4 = sql4 & "1, "
                    sql4 = sql4 & id_client & ", "
                    sql4 = sql4 & "10)"
                    Conn.Execute (sql4)
                    
                    'Mise à jour de t_livraison
                    sql2 = "insert into t_livraison (id_livraison, id_pays, beneficiare_liv, adresse_liv, cp_liv, ville_liv, cadeau, cadeau_texte) values ("
                    sql2 = sql2 & id_livraison & ", "
                    sql2 = sql2 & l_id_pays & ", "
                    sql2 = sql2 & noquote2(l_beneficiaire) & ", "
                    sql2 = sql2 & noquote2(l_adresse) & ", "
                    sql2 = sql2 & noquote2(l_cp) & ", "
                    sql2 = sql2 & noquote2(l_ville) & ","
                    sql2 = sql2 & noquote2(l_kdo) & ","
                    sql2 = sql2 & noquote2(l_kdotxt) & ")"
                    Conn.Execute (sql2)
                
                Session("userid") = id_client
                Session("HashData") = ""
                Session("HashSecure") = ""
                
                'Rappel du login et mot de passe pour un nouveau client.
                '-------------------------------------------------------
             If Request("act") <> "commanderapide" Then
                pr ("<span class=""smalltext""><b>" & translate("Vos informations ont été correctement enregistrées") & ".</b><br><br>")
                    pr ("<span class=""alert""> " & translate("N'oubliez pas de noter vos") & " <b>" & translate("informations de connexion") & "</b> :</span><br>")
                    pr ("<span class=""smalltext"">" & translate("Identifiant") & " : " & LOGIN & "</span> <br>")
                    pr ("<span class=""smalltext"">" & translate("Mot de passe") & " : " & passwd_client & "</span>")
                    pr ("<br><br>")
              End If
                            
        Else                'déjà client
        UserDisplayName = Trim(c_nom) & " " & Trim(c_prenom)
                
                SQL0 = "update dbp_users set "
                SQL0 = SQL0 & " CId = " & Session("APPLICATION") & ", "
                'sql0 = sql0 &  " UserLogin = "" & noquote2(login) & ", "
                SQL0 = SQL0 & " UserLast = " & noquote2(c_nom) & ", "
                SQL0 = SQL0 & " UserFirst = " & noquote2(c_prenom) & ", "
                SQL0 = SQL0 & " UserDisplayName = " & noquote2(UserDisplayName) & ", "
                SQL0 = SQL0 & " UserEMail = " & noquote2(c_email)
                'sql0 = sql0 & " UserPass = "" & passwd_client & ", "
                'sql0 = sql0 & " UserEmailPas = "" & noquote2(passwd_client)
                SQL0 = SQL0 & " where  UserId=" & id_client
            
                    SQL1 = "update dbp_userinfos set "
                    SQL1 = SQL1 & " CId = " & Session("APPLICATION") & ", "
                    SQL1 = SQL1 & " LastName = " & noquote2(c_nom) & ", "
                    SQL1 = SQL1 & " FirstName = " & noquote2(c_prenom) & ", "
                    SQL1 = SQL1 & " fld3 = " & noquote2(c_societe) & ", "
                    SQL1 = SQL1 & " Address = " & noquote2(c_adresse) & ", "
                    SQL1 = SQL1 & " Zip = " & noquote2(c_codepostal) & ", "
                    SQL1 = SQL1 & " City = " & noquote2(c_ville) & ", "
                    SQL1 = SQL1 & " Email = " & noquote2(c_email) & ", "
                    SQL1 = SQL1 & " fld2 = " & noquote2(c_numfna) & ", "
                    SQL1 = SQL1 & " id_livraison = " & id_livraison & ", "
                    SQL1 = SQL1 & " id_pays = " & c_id_pays
                    SQL1 = SQL1 & " where  UserId=" & id_client
                    
                    'vérifie si client a un identifiant de livraison
                    sqltemp = "SELECT id_livraison FROM t_livraison where  id_livraison = " & id_livraison
                    Set test = Conn.Execute(sqltemp)
                    If test.BOF Then
                        'Insertion dans t_livraison
                        sql2 = "insert into t_livraison (id_livraison, id_pays, beneficiare_liv, adresse_liv, cp_liv, ville_liv, cadeau, cadeau_texte) values ("
                        sql2 = sql2 & id_livraison & ", "
                        sql2 = sql2 & l_id_pays & ", "
                        sql2 = sql2 & noquote2(l_beneficiaire) & ", "
                        sql2 = sql2 & noquote2(l_adresse) & ", "
                        sql2 = sql2 & noquote2(l_cp) & ", "
                        sql2 = sql2 & noquote2(l_ville) & ","
                        sql2 = sql2 & noquote2(l_kdo) & ","
                        sql2 = sql2 & noquote2(l_kdotxt) & ")"
                    Else
                        'Mise à jour de t_livraison
                        sql2 = "update t_livraison set "
                        sql2 = sql2 & " id_pays=" & l_id_pays & ", "
                        sql2 = sql2 & " beneficiare_liv=" & noquote2(l_beneficiaire) & ", "
                        sql2 = sql2 & " adresse_liv=" & noquote2(l_adresse) & ", "
                        sql2 = sql2 & " cp_liv=" & noquote2(l_cp) & ", "
                        sql2 = sql2 & " ville_liv=" & noquote2(l_ville) & ", "
                        sql2 = sql2 & " cadeau=" & l_kdo & ","
                        sql2 = sql2 & " cadeau_texte=" & noquote2(l_kdotxt) & ""
                        sql2 = sql2 & " where  id_livraison = " & id_livraison
                    End If
                    Set test = Nothing
                    
                Conn.Execute (sql2)
                Conn.Execute (SQL1)
                Conn.Execute (SQL0)
                    pr ("<span class=""smallheader"">" & translate("Vos informations ont été correctement mises à jour") & ".</span><br><p>")
                End If
            Sql = "select max(id_commande) from t_commande"
            
            Set cur = Conn.Execute(Sql)
            id_commande = 1
            If Not cur.EOF Then
                id_commande = cur(0) + 1
            End If
            Session("id_commande") = id_commande
            Sql = "insert into t_commande (id_commande, id_client, total, dcre_commande,etat_commande,currency_code) values ("
            Sql = Sql & id_commande & ", "
            Sql = Sql & id_client & ", "
            Sql = Sql & Replace(Session("total_eu"), ",", ".") & ", '" & Now() & "',1,978)"
            Conn.Execute (Sql)
            
    
            SQL1 = "select c.id_produit,c.qte_produit,p.prix from t_caddie c, dbp_topics p where c.id_caddie=" & Session("candidat_id_caddy") & " and c.id_produit=p.TopicId "
            Set cur = Conn.Execute(SQL1)
            While Not cur.EOF
            sql2 = "insert into t_ligne_commande(id_commande,id_article,quantity,prix) values("
            sql2 = sql2 & id_commande & ", "
            sql2 = sql2 & cur("id_produit") & ", "
            sql2 = sql2 & cur("qte_produit") & ", "
            sql2 = sql2 & fmt(cur("prix")) & ")"
            If cur("id_produit") <> "-1" Then

            Conn.Execute (sql2)
            End If
            cur.MoveNext
            Wend
        pr ("<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
        pr ("<tr>")
        pr ("<td class=""smallheader"" align=""center"" colspan=""2"">")
                pr (translate("Choisissez votre mode de règlement") & " :</td>")
           ' pr ("<td align=""left"">&nbsp;</td>")
            pr ("<td class=""smallertext"" width=""150"" align=""right"" colspan=""2"">")
                pr (translate("Sécurisé par") & "</td></tr>")
            pr ("<form name=""payer"" action=""" & GetDefault("EcomUrlPaiementBanque", "http://cyberplus.euxia.net/euxia/call_sips.asp", Session("candidat_ADOContact")) & """ method=""post"">")
            pr ("<input type=""hidden"" name=""transaction_id"" value=""" & id_commande & """>")
            
            montant_temp = FormatNumber(Session("total_eu"), 2)
            montant0 = Replace(montant_temp, ",", "")
            montant = Replace(montant0, ".", "")
            montant = Replace(montant, " ", "")
            montant = Format(montant)
            pr ("<input type=""hidden"" name=""f_amount"" value=" & montant & ">")
            pr ("<input type=""hidden"" name=""language"" value=""fr"">")
            
            pr ("<tr>")
                pr ("<td align=""center"">")
                    pr ("<a href=""../portal_ecom/cheque.asp"">")
                    pr ("<img src=""" & Session("candidat_EcomPathImages") & "cheque.gif"" alt=""" & translate("Paiement par chèque") & """ border=0></a>&nbsp;&nbsp;")
                    pr ("<a href=""javascript:document.payer.submit();"">")
                    pr ("<img src=""" & Session("candidat_EcomPathImages") & "cb.gif"" border=0 alt=""" & translate("Paiement sécurisé par CB") & """></a>")
                    pr ("</td>")
                pr ("<td align=""left"">&nbsp;</td>")
                pr ("<td align=""right"" valign=""top"">")
                    pr ("<a href=""http://www.cyberpluspaiement.com/"">")
                    pr ("<img src=""" & Session("candidat_EcomPathImages") & "logobp.gif"" border=""0"" alt=""Cyberplus""></a>&nbsp;&nbsp;</td>")
                    pr ("<td align=""left"" valign=""top"" width=""2""><a href=""http://www.banquepopulaire.fr/"" target=""_blank"">")
                    pr ("<img src=""" & Session("candidat_EcomPathImages") & "logobp2.gif"" border=""0"" alt=""Banque Populaire ""></a></td>")
            pr ("</tr>")
            pr ("<tr>")
                pr ("<td>&nbsp;</td>")
                pr ("<td align=""left"">&nbsp;</td>")
                pr ("<td align=""right"" valign=""top"" colspan=""2"">")
                    pr ("<img src=""" & Session("candidat_EcomPathImages") & "logoscartes.gif"" border=0 alt=""" & translate("Paiement accepté par CB") & """>")
            pr ("</tr>")
            
            pr ("</table>")
            pr ("</form></center>")
            
            Session("id_client") = id_client
            Session("id_livraison") = id_livraison
            Session("c_nom") = c_nom
            Session("c_prenom") = c_prenom
            Session("c_adresse") = c_adresse
            Session("c_codepostal") = c_codepostal
            Session("c_ville") = c_ville
            Session("c_id_pays") = c_id_pays
            Session("c_email") = c_email
            Session("l_beneficiaire") = l_beneficiaire
            Session("l_adresse") = l_adresse
            Session("l_cp") = l_cp
            Session("l_ville") = l_ville
            Session("l_id_pays") = l_id_pays
            Session("l_kdo") = l_kdo
            Session("l_kdotxt") = l_kdotxt
            Session("id_client") = id_client
            Session("id_livraison") = id_livraison
    Conn.Close
    Set Conn = Nothing
End Function

Public Function BoutonRetour(image, Class)
If Len(Class) > 1 Then
    css = "<span class=""" & Class & """>"
    fincss = "</span>"
Else
    css = ""
    fincss = ""
End If

If image = 1 Then
If Session("PathImagesSite") = "" Then Session("PathImagesSite") = GetDefault("PathImagesSite", "../euxia/", Session("candidat_ADOContact"))

    leStr = leStr & "<a onMouseOver=""message('" & translate("Retour") & "');return true;"" href=""javascript:history.go(-1);"">"
    leStr = leStr & "<img src=""" & Session("PathImagesSite") & "retour.gif"" alt=""" & translate("Retour") & """ border=""0""></a>"
Else
    leStr = leStr & "<a onMouseOver=""message('" & translate("Retour") & "');return true;"" href=""javascript:history.go(-1);"">"
    leStr = leStr & css
    leStr = leStr & translate("Retour") & fincss & "</a>"
End If

BoutonRetour = leStr
End Function

Public Function ImgTitreEcommerce(titre)
    pr ("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">")
    pr ("<tr><td>")

        pr ("<table border=""0"" background=""" & Session("EcomPathImages") & "bandebg.gif"" cellpadding=""0"" cellspacing=""0"" width=""100%"">")
        pr ("<tr>")
            pr ("<td width=""5%"" align=""left""><img src=""" & Session("EcomPathImages") & "bandeg.gif"" border=""0""></td>")
            pr ("<td class=""EcomTitreCaddie""> " & titre & "</td>")
            pr ("<td width=""5%"" align=""right""><img src=""" & Session("EcomPathImages") & "banded.gif"" border=""0""></td>")
            pr ("</tr></table>")
            
    pr ("</tr>")
    pr ("<tr><td>" & ImgTransparent(1, 5) & "</td>")
    pr ("</tr></table>")
End Function

Public Function ImgTransparent(width, height)
    ImgTransparent = "<img src=""" & Session("PortalPathImages") & "transparent.gif"" border=""0"" width=""" & width & """ height=""" & height & """>"
End Function

Public Function ImgEnteteColonne(titre)
leStr = "<table border=""0"" background=""" & Session("EcomPathImages") & "titre_bg.gif"" cellpadding=""0"" cellspacing=""0"">"
leStr = leStr & "<tr>"
leStr = leStr & "<td align=""left""><img src=""" & Session("EcomPathImages") & "titre_gauche.gif"" border=""0""></td>"
leStr = leStr & "<td class=""EcomTitreCaddie""> " & titre & "</td>"
leStr = leStr & "<td align=""right""><img src=""" & Session("EcomPathImages") & "titre_droite.gif"" border=""0""></td>"
leStr = leStr & "</tr></table>"
ImgEnteteColonne = leStr
End Function

Public Function ViewBasPage(category, twidth, align)

    leStr = "<div align=""" & align & """>"
    leStr = leStr & "<table border=""0"" width=""" & twidth & """ cellpadding=""0"" cellspacing=""0"">"
    leStr = leStr & "<tr>"
        leStr = leStr & "<td align=""left"">"
        If Session("CategoryAccueil_ID") = "" Then Session("CategoryAccueil_ID") = GetDefault("CategoryAccueil_ID", "3", Session("candidat_ADOContact"))
        If CInt(category) <> CInt(Session("CategoryAccueil_ID")) Then
            leStr = leStr & BoutonRetour(1, "")
        Else
            leStr = leStr & ImgTransparent(1, 1)
        End If
        leStr = leStr & "</td>"
        leStr = leStr & "<td align=""right"">"
        leStr = leStr & "<a href=""#haut""><img src=""" & Session("PathImagesSite") & "hautdepage.gif"" border=""0""></a>"
        leStr = leStr & "</td></tr>"
    leStr = leStr & "</table>"
    
    ViewBasPage = leStr
End Function


Public Sub ViewBasPageCaddy()


    pr ("<tr height=""30"" valign=""bottom""><td align=""left"">")
            pr ("<a href=""../" & Session("HomePage") & "?category=" & Session("category") & """ target=""_top"">")
            pr ("<img src=""" & Session("EcomPathImages") & "poursuivre.gif"" border=""0""></a>")
        pr ("</td>")
        pr ("<td colspan=""3"">" & ImgTransparent(1, 1) & "</td>")
        pr ("<td align=""right"">")
            pr ("<a href=""../" & Session("HomePage") & "?act=com "" target=""_top"">")
            pr ("<img src=""" & Session("EcomPathImages") & "confirmerlacommande.gif"" border=""0""></a>")
        pr ("</td></tr>")
    
    
    'If Session("totalqte") < 10 Then
            pr ("<tr height=""30"">")
            pr ("<td>" & ImgTransparent(1, 60) & "</td>")
            pr ("<td colspan=""4"" align=""right"" valign=""bottom"">")
            pr ("<a href=""#haut""><img src=""" & Session("PathImagesSite") & "hautdepage.gif"" border=""0""></a>")
            pr ("</td></tr>")
    'End If
End Sub


Public Sub ViewHautPage()
    pr ("<a name=""#haut""></a>")

End Sub


Function TOZS(v)
    v = v           ' AH, MICROSOFT... LE PIRE C'EST QUE CA MARCHE PLUS SANS CETTE LIGNE !
    If IsNull(v) Then
           TOZS = ""
    Else
           TOZS = CStr(v)
    End If
End Function

Function inc(i)
    If i <> "" Then
        inc = i + 1
    Else
        inc = 1
    End If
End Function
Function noquote2(s0)
    s0 = s0
    If (CStr(s0) = "NULL") Or IsNull(s0) Or (Len(CStr(s0)) = 0) Then
        noquote2 = "NULL"
    Else
        noquote2 = "'" & noquote(s0) & "'"
    End If
End Function
Function fmt(v)
    w = v
    If IsNull(w) Then
        fmt = "0.00"
    Else
        w = Round(CDbl(w), 2)
        fmt = Replace(CStr(w), ",", ".")
    End If
End Function
Function noquote(s0)
'   s = ""
'   c = ""
'
'   for i = 1 to len(s0)
'       c=mid(s0,i,1)
'       if c = "'" then
'           s = s & "''"
'       else
'           s = s & c
'       end if
'   next
'
'   noquote = s
    noquote = Replace(CStr(s0), "'", "''")
End Function
