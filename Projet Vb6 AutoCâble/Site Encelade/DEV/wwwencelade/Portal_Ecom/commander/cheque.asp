<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<%
set conn=server.createobject("adodb.connection")
conn.open session("dsn")

Set MyPortal = Server.CreateObject(Session("PortalComObject") & ".Portal")

if(Request("envoimail")=1) then
	strTo = Request("mail")
	id_caddie = Request("id_caddie")
	strSubject = "Bon de Commande - Flyway"
	strBody = "Bon de Commande n°" & session("id_commande") & vbCrLf & vbCrLf
	sql = 	"select c.id_produit, p.title, c.qte_produit,  p.prix, p.poids * 1000" &_
					" from	dbp_topics p, t_caddie c, t_tva t " &_
					" where	c.id_caddie=" & session("id_caddie") &_
					" and	c.id_produit is not null " &_
					" and	c.id_produit = p.TopicId " &_
					" and	t.tva_id = p.tva_id"
			set cur = conn.execute(sql)
			rien = 0
			cpt = 0
			total = cdbl(port_eu)
			totalf = cdbl(port_ff)
			i_poids = 0
			sous_total=0
			sous_total_e=0
			while not cur.eof
				id_produit = cur(0)
				titre = cur(1)
				qte_produit = cint(cur(2))
				cpt = cint(cpt) + qte_produit
				prix = MyPortal.fmt(cur(3))
				sous_total_e = cdbl(cur(2)) * cdbl(cur(3))
				total = total + sous_total_e
				poids = cint(cur(4)) * qte_produit
				i_poids = i_poids + poids
				strBody = strBody & qte_produit & "  " & titre & "  " &  prix & "TTC Euros " & vbCrLf 
				cur.movenext
			wend
 		    strBody=strBody & vbCrLf & "Sous Total : " & MyPortal.fmt(sous_total_e) & " Euros" & vbCrLf & vbCrLf
			
			sql = 	"select	prix " &_
					"from	t_tarif_livraison " &_
					"where	id_pays = " & session("id_pays") &_
					"and	poids_mini <= " & i_poids &_
					"and	poids_maxi > " & i_poids
			
			set cur = conn.execute(sql)
			if not cur.eof then
				port_eu = cdbl(cur(0))
			else
				port_eu = 0.00
			end if
	
			sql = "select label_pays, tarif_normal from t_pays where id_pays = " & session("id_pays")
			set cur = conn.execute (sql)
			if not cur.eof then
				coeff = cdbl(cur(1))
				label_pays = cur(0)
			end if
			set cur = nothing
			if isNull(coeff) or (len(cstr(coeff))<1) then
				coeff = 1.0
			end if
	
			port_eu = port_eu * coeff
			
			total = total + port_eu
			
			session("total_eu") = cstr(total)
				sql = "update t_commande set amount = " & MyPortal.fmt(total) & ", currency_code=978, etat_commande = 1, card_type = 'Chèque' where id_commande = " & session("id_commande")
			conn.execute(sql)
			
			strBody = strBody & "Frais de port : " & i_poids & "g / " & label_pays & " " & MyPortal.fmt(port_eu) & " Euros" & vbCrLf & vbCrLf &_
						"Quantite totale : " & cpt & vbCrLf & vbCrLf &_
						"Prix Total : " & MyPortal.fmt(total) & " Euros" & vbCrLf & vbCrLf &_
						"Merci et à bientôt sur " & session("UrlBase") &  vbCrLf & vbCrLf & vbCrLf
			
	
	Set objMail = Server.CreateObject("CDONTS.NewMail")
'	EmailCommercant = MyPortal.GetDefault("EcomEmailCommandes", "commandes@euxia.net")
	EmailCommercant = "amalin@euxia.net"
	'Mail pour le client
	objMail.To = strTo                    															'set 'To' address
	objMail.From = "serviceweb@euxia.net" 															'set 'From' address
	objMail.Value("Reply-To") =	EmailCommercant													     'set 'Reply to' address
	objMail.Subject = strSubject          															'set the subject line
	objMail.Body =  strBody               															'set the message content
	objMail.Send                          															'and send the message
	Set objMail = Nothing 'then destroy the component

	Set objMail = Server.CreateObject("CDONTS.NewMail")
	'Mail pour le commerçant
	objMail.To = 	EmailCommercant												          			'set 'To' address
	objMail.From = "serviceweb@euxia.net"           												'set 'From' address
	objMail.Value("Reply-To") =  MyPortal.GetDefault("EcomEmailCommandes", "commandes@euxia.net")	'set 'Reply to' address
	objMail.Subject = strSubject          															'set the subject line
	objMail.Body =  strBody               															'set the message content
	objMail.Send                          															'and send the message
		
	Set objMail = Nothing 'then destroy the component
	
%>
<HTML>
<HEAD>
    <link rel=STYLESHEET href='<%=Session("StyleSheet")%>' type='text/css'>
	<script language='javascript' src='../../Portal_Java/Portal_Generic.js'></script>
</HEAD>
<body>
<br><br><br>
<div align="center" class="smalltext">
<% =MyPortal.Translate("Un mail contenant le récaptitulatif de votre commande vous a été envoyé")%>.<br>
<% =MyPortal.Translate("Nous vous remercions pour votre intérêt envers Flyway") %>,&nbsp;
 
 <a href="<%=session("UrlBase")%>portal_ecom/kill_vars.asp" target="_top">
 <img src="" alt="Terminer la commande" border="0"></a>
</div>
</body>
</HTML>

<%else%>

<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel=STYLESHEET href='<%=Session("StyleSheet")%>' type='text/css'>
	<script language='javascript' src='../../Portal_Java/Portal_Generic.js'></script>
</head>
<body>
<br><br>
<table width="80%" align="center" >
	<tr class="smalltext">
      <td width="30%" align="left" >
    	 <a OnMouseOver="message('Imprimer'); return true;" href="#" OnClick="javascript:window.print();">
    	 <img src="../../flyway/images/imprimante.gif" alt"Imprimer la page" border="0">Imprimer la page</a>
      </td>
      <td align="right" class="smalltext">
	     Date de la commande: <%=now()%>
      </td>
	</tr>
</table>
<br>
<table width="75%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr class="smalltext"><td>
<% =MyPortal.Translate("Pour que votre commande numéro") & " " & session("id_commande")%>&nbsp;
<% =MyPortal.Translate("soit prise en compte") %>,&nbsp;
<% =MyPortal.Translate("veuillez nous adresser un chèque à l'ordre de Flyway à") %> :<br>
<blockquote>
<span class="smalltext">
 FLYWAY<br>
 1 avenue Pierre Salvi<br>
 95500 GONESSE - FRANCE <br>
</span>
</blockquote>
<span class="alert"><i><% =MyPortal.Translate("Attention") %> : </i></span><br>
<% =MyPortal.Translate("Veuillez toujours rappeler votre numéro de commande dans vos correspondances") %>: 
<span class="alert"><b><%=session("id_commande")%></b></span>
<br>
<br>
<%
'response.end
                    sql = "SELECT c.LastName, "
                    sql = sql & " c.FirstName, "
                    sql = sql & " c.Address, "
                    sql = sql & " c.Zip, "
                    sql = sql & " c.City, "
                    sql = sql & " c.id_pays, "
                    sql = sql & " p.label_pays, "
                    sql = sql & " c.Email, "
                    sql = sql & " c.id_livraison, "
                    sql = sql & " c.fld2 " 'numéro FLA
                    sql = sql & " FROM  dbp_UserInfos c, t_pays p "
                    sql = sql & " WHERE p.id_pays = c.id_pays "
                    sql = sql & " AND   c.UserID = " & session("userid")
set cur = conn.execute(sql)
c_nom = cur(0)
c_prenom = cur(1)
c_adresse = cur(2)
c_codepostal = cur(3)
c_ville = cur(4)
c_id_pays = cur(5)
c_pays = cur(6)
c_email = cur(7)
id_livraison=cur(8)
c_numfla=cur(9)
set cur = nothing
	sql = "select l.beneficiare_liv, " & VBCRLF
	sql = sql & " l.adresse_liv, " & VBCRLF
	sql = sql & " l.cp_liv, " & VBCRLF
	sql = sql & " l.ville_liv, " & VBCRLF
	sql = sql & " l.id_pays, " & VBCRLF
	sql = sql & " p.label_pays " & VBCRLF
	sql = sql & " from  t_livraison l, t_pays p " & VBCRLF
	sql = sql & " where l.id_livraison = " & id_livraison  & VBCRLF
	sql = sql & " and   p.id_pays = l.id_pays" & VBCRLF
	'response.write sql & "<br>"
	'response.end
	
set cur = conn.execute(sql)
l_beneficiaire = cur(0)
l_adresse = cur(1)
l_cp = cur(2)
l_ville = cur(3)
l_pays = cur(5)

set cur=nothing

       response.write MyPortal.Ecom_ValiderCommande_TableauRecap2()

%>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<%if session("l_kdo") = 1 then response.write("<tr><td colspan=""2""><img src="""" alt=""Cadeau""></td></tr>")%>
<tr class="smallheader">
<td width="50%"><i><% =MyPortal.Translate("Coordonnées de facturation") %>:</i></td>
<td width="50%"><i><% =MyPortal.Translate("Coordonnées de livraison") %>:</i></td></tr>
<tr class="smalltext"><td>
<%=c_nom&" "&c_prenom%><br>
<%=c_adresse%><br>
<%=c_codepostal%>&nbsp;
<%=c_ville%><br>
<%=c_pays%>
</td><td>
<%=l_beneficiaire%><br>
<%=l_adresse%><br>
<%=l_cp%>&nbsp;<%=l_ville%><br>
<%=l_pays%>
</td></tr></table>
<br>
<table border="0" cellspacing="0" cellpadding="0">
<form action="cheque.asp" method="post">
<tr>
<td class="smalltext"><% =MyPortal.Translate("Pour recevoir le bon de commande sur votre email") %> &nbsp; &nbsp; </td>
<td><input type="hidden" name="envoimail" value="1">
<input type="hidden" name="mail" value="<%= c_email%>">
<input type="hidden" name="id_commande" value="<%= session("id_commande")%>">
<input type="hidden" name="id_caddie" value="<%= session("id_caddie")%>">
<a href="javascript:form.submit();">
<Img src="<%=session("EcomPathImages")%>poursuivrelacommande.gif" alt="poursuivre la commande" border="0"></a></td>
</tr>
</form>
</table>

</td></tr></table>
</body>
</html>
<%
set conn = nothing
end if%>
