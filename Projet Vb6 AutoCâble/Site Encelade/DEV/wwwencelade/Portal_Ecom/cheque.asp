<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<%
set conn=server.createobject("adodb.connection")
conn.open session("dsn")

Set MyPortal = Server.CreateObject(Session("PortalComObject") & ".Portal")

if(Request("envoimail")=1) then
	strTo = Request("mail")
	id_caddie = Request("id_caddie")
	strSubject = "Bon de Commande - Flyway"
'	strBody= MyPortal.Ecom_ValiderCommande_TableauRecap2() & "</table>"
'	sql = 	"select c.id_produit, p.title, c.qte_produit,  p.prix, p.poids * 1000" &_
'					" from	dbp_topics p, t_caddie c, t_tva t " &_
'					" where	c.id_caddie=" & session("id_caddie") &_
'					" and	c.id_produit is not null " &_
'					" and	c.id_produit = p.TopicId " &_
'					" and	t.tva_id = p.tva_id"
'			set cur = conn.execute(sql)
'			rien = 0
'			cpt = 0
'			total = cdbl(port_eu)
'			totalf = cdbl(port_ff)
'			i_poids = 0
'			sous_total=0
'			sous_total_e=0
'			while not cur.eof
'				id_produit = cur(0)
'				titre = cur(1)
'				qte_produit = cint(cur(2))
'				cpt = cint(cpt) + qte_produit
'				prix = MyPortal.fmt(cur(3))
'				sous_total_e = cdbl(cur(2)) * cdbl(cur(3))
'				total = total + sous_total_e
'				poids = cint(cur(4)) * qte_produit
'				i_poids = i_poids + poids
'				strBody = strBody & qte_produit & "  " & titre & "  " &  prix & "TTC Euros " & vbCrLf 
'				cur.movenext
'			wend
' 		    
'			
'			sql = 	"select	prix " &_
'					"from	t_tarif_livraison " &_
'					"where	id_pays = " & session("id_pays") &_
'					"and	poids_mini <= " & i_poids &_
'					"and	poids_maxi > " & i_poids
'			
'			set cur = conn.execute(sql)
'			if not cur.eof then
'				port_eu = cdbl(cur(0))
'			else
'				port_eu = 0.00
'			end if
'	
'			sql = "select label_pays, tarif_normal from t_pays where id_pays = " & session("id_pays")
'			set cur = conn.execute (sql)
'			if not cur.eof then
'				coeff = cdbl(cur(1))
'				label_pays = cur(0)
	'		end if
'			set cur = nothing
'			if isNull(coeff) or (len(cstr(coeff))<1) then
'				coeff = 1.0
'			end if
'	
'			port_eu = port_eu * coeff
'			
'			total = total + port_eu
'			
'			session("total_eu") = cstr(total)
'				sql = "update t_commande set amount = " & MyPortal.fmt(total) & ", currency_code=978, etat_commande = 1, card_type = 'Chèque' where id_commande = " & session("id_commande")
'			conn.execute(sql)
			
			'strBody = strBody & "Frais de port : " & i_poids & "g / " & label_pays & " " & MyPortal.fmt(port_eu) & " Euros" & vbCrLf & vbCrLf &_
			'			"Quantite totale : " & cpt & vbCrLf & vbCrLf &_
			'			"Prix Total : " & MyPortal.fmt(total) & " Euros" & vbCrLf & vbCrLf &_
			'			"Merci et à bientôt sur " & session("UrlBase") &  vbCrLf & vbCrLf & vbCrLf
			
		strBody = ""
		strBody = strBody & "<html><head><link rel='stylesheet' href='http://flyway.euxia.net/portal_style/flyway.css'></head>"
		strBody = strBody & "<body background=" & session("EcomPathImages") & "/email/bgciel.jpg><br>"
		strBody = strBody & "<table border=0 cellpadding=0 cellspacing=0 width=500 align=center><tr>"
		strBody = strBody & "    <td width=10><img src=" & session("EcomPathImages") & "/email/hg.gif width=22></td>"
		strBody = strBody & "    <td background=bgh.gif width=455><img src=" & session("EcomPathImages") & "/email/haut.gif></td>"
		strBody = strBody & "    <td width=35><img src=" & session("EcomPathImages") & "/email/hd.gif width=25></td></tr>"
		strBody = strBody & "  <tr><td background=bgg.gif height=93 width=10><img src=" & session("EcomPathImages") & "/email/bgg.gif width=22 height=19></td>"
		strBody = strBody & "    <td height=93 width=455><table border=0 width=100% align=center><tr><td> "
		strBody = strBody & "            <table border=0 cellpadding=0 cellspacing=0 width=100% ><tr><td> "
		strBody = strBody & "                  <table border=0 background=bandebg.gif cellpadding=0 cellspacing=0 width=100% ><tr>" 
		strBody = strBody & "                      <td width=5% align=left><img src=" & session("EcomPathImages") & "/email/bandeg.gif border=0></td>"
		strBody = strBody & "                      <td class=EcomTitreCaddie> <b><font face=Verdana, Arial, Helvetica, sans-serif size=2 color=#00319C>Récapitulatif de votre commande</font></b></td>"
		strBody = strBody & "                      <td width=5% align=right><img src=" & session("EcomPathImages") & "/email/banded.gif border=0></td></tr></table></tr>"
		strBody = strBody & "              <tr><td><img src=" & session("EcomPathImages") & "/email/transparent.gif border=0 width=1 height=5></td></tr></table></td></tr>"
		strBody = strBody & "        <tr><td>" 
		strBody = strBody & MyPortal.Ecom_ValiderCommande_TableauRecap2 & "</table>"
		strBody = strBody & "        </td></tr></table></td>"
		strBody = strBody & "    <td background=bgd.gif height=93 width=35><img src=" & session("EcomPathImages") & "/email/bgd.gif width=25 height=15></td></tr>"
		strBody = strBody & "  <tr><td width=10><img src=" & session("EcomPathImages") & "/email/bg.gif width=22 height=26></td>"
		strBody = strBody & "    <td background=bgb.gif width=455><img src=" & session("EcomPathImages") & "/email/bgb.gif width=20 height=26></td>"
		strBody = strBody & "    <td width=35><img src=" & session("EcomPathImages") & "/email/bd.gif width=25 height=26></td></tr></table></BODY></HTML>"

	Set objMail = Server.CreateObject("CDONTS.NewMail")
'	EmailCommercant = MyPortal.GetDefault("EcomEmailCommandes", "commandes@euxia.net")
	EmailCommercant = "amalin@euxia.net"
	'Mail pour le client
	objMail.To = strTo                    															'set 'To' address
	objMail.BodyFormat = 0
	objMail.MailFormat = 0	
	objMail.From = "serviceweb@euxia.net" 															'set 'From' address
	objMail.Value("Reply-To") =	EmailCommercant													     'set 'Reply to' address
	objMail.Subject = strSubject          															'set the subject line
	objMail.Body =  strBody               															'set the message content
	objMail.Send                          															'and send the message
	Set objMail = Nothing 'then destroy the component

	Set objMail = Server.CreateObject("CDONTS.NewMail")
	'Mail pour le commerçant
	objMail.BodyFormat = 0
	objMail.MailFormat = 0	
	objMail.To = 	EmailCommercant												          			'set 'To' address
	objMail.From = "serviceweb@euxia.net"           												'set 'From' address
	objMail.Value("Reply-To") =  EmailCommercant													'set 'Reply to' address
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
<br><br><br><br><br><br><br>
<div align="center" class="smalltext">
<img src="<%=session("EcomPathImages")%>logo_petit.gif" alt="Flyway" border="0">
<br><br><br><br>
<% =MyPortal.Translate("Un mail contenant le récaptitulatif de votre commande vous a été envoyé")%>.<br><br>
<% =MyPortal.Translate("Nous vous remercions pour votre intérêt envers Flyway")%>.
 <br><br><br><br>
 <a href="<%=session("UrlBase")%>portal_ecom/kill_vars.asp" target="_top">
 <img src="<%=session("EcomPathImages")%>terminer.gif" alt="Terminer la commande" border="0"></a>
</div>
</body>
</HTML>

<%
response.end
else%>

<HTML>
<HEAD>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel=STYLESHEET href='<%=Session("StyleSheet")%>' type='text/css'>
	<script language='javascript' src='../../Portal_Java/Portal_Generic.js'></script>
</head>
<body>
<br><br>
<table width="80%" align="center" border="0" cellpadding="2" cellspacing="0">
	<tr class="smalltext">
      <td align="right" valign="top" colspan="2">
    	 <a OnMouseOver="message('Imprimer'); return true;" href="#" OnClick="javascript:window.print();">
    	 <img src="../../flyway/images/imprimante.gif" alt"Imprimer la page" border="0">
		 <br>
		 Imprimer la page</a>
      </td>
	</tr>
	<tr class="smalltext">
		<td width="50%">
			<% =MyPortal.Translate("Pour que votre commande soit prise en compte") %>,&nbsp;
			<% =MyPortal.Translate("veuillez nous adresser un chèque à l'ordre de Flyway suivi du bon de commande ci-dessous préalabrement imprimé") %>.
		</td>
		<td>
			<span class="alert"><i><% =MyPortal.Translate("Attention") %> : </i></span><br>
			<% =MyPortal.Translate("Veuillez toujours rappeler votre numéro de commande dans vos correspondances") %>: 
			<span class="alert"><b><%=session("id_commande")%></b></span>
		</td>
	</tr>
	<tr>
		<td colspan="2"><%=MyPortal.ImgTransparent(1,40)%></td>
	</tr>
	<tr>
		<td colspan="2"><%=MyPortal.ImgTitreEcommerce("Bon de commande")%></td>
	</tr>
	<tr class="smalltext">
      <td align="right" class="smalltext" valign="top" colspan="2">
	     <b>Date de la commande :</b> <%=now()%>
      </td>
	</tr>
	<tr>
		<td class="smalltext">
	 	<img src="<%=session("EcomPathImages")%>logo_petit.gif" border="0"><br>
	 	5 avenue Pierre Salvi<br>
	 	95500 GONESSE - FRANCE <br>
		</td>
		<td class="smalltext">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2">
<%
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
%>
	<tr>
		<td colspan="2" class="smalltext" align="right">
		<table width="70%" border="0" cellpadding="0" cellspacing="0">
			<%if session("l_kdo") = 1 then%>
			<tr class="smallheader">
				<td width="50%">&nbsp;</td>
				<td width="50%">		<% response.write("<img src=""" & session("EcomPathImages") & "cadeau.gif"" alt=""Cadeau"">")%></td></tr>
			<%end if%>
			<tr class="smallheader">
				<td width="50%"><i><% =MyPortal.Translate("Coordonnées de facturation") %>:</i></td>
				<td width="50%"><i><% =MyPortal.Translate("Coordonnées de livraison") %>:</i></td></tr>
			<tr class="smalltext">
				<td>
					<%=c_nom&" "&c_prenom%><br>
					<%=c_adresse%><br>
					<%=c_codepostal%>&nbsp;
					<%=c_ville%><br>
					<%=c_pays%>
				</td>
				<td>
					<%=l_beneficiaire%><br>
					<%=l_adresse%><br>
					<%=l_cp%>&nbsp;<%=l_ville%><br>
					<%=l_pays%>
				</td>
			</tr>
		</table>
		</td>	
	</tr>
	<tr>
		<td colspan="2"><%=MyPortal.ImgTransparent(1,40)%></td>
	</tr>
	<tr>
		<td colspan="2" class="smalltext">
<%
    response.write MyPortal.Ecom_ValiderCommande_TableauRecap2()
	response.write("</table>")
%>
		</td>	
	</tr>
	<tr>
		<td colspan="2"><%=MyPortal.ImgTransparent(1,20)%></td>
	</tr>
	<tr>
		<td colspan="2" class="smalltext">
		<table border="0" cellspacing="0" cellpadding="0" width="100%">
		<form action="cheque.asp" method="post" name="f_email">
			<tr>
				<td class="smalltext" align="right">
					<% =MyPortal.Translate("Pour recevoir le bon de commande sur votre email") %> &nbsp; &nbsp; </td>
				<td align="right">
					<input type="hidden" name="envoimail" value="1">
					<input type="hidden" name="mail" value="<%= c_email%>">
					<input type="hidden" name="id_commande" value="<%= session("id_commande")%>">
					<input type="hidden" name="id_caddie" value="<%= session("id_caddie")%>">
					<a href="javascript:f_email.submit();" >
					<Img src="<%=session("EcomPathImages")%>poursuivrelacommande.gif" alt="poursuivre la commande" border="0"></a></td>
			</tr>
		</form>
		</table>
		
	</td>
	</tr>
	<tr>
		<td colspan="2"><%=MyPortal.ImgTransparent(1,20)%></td>
	</tr>
</table>
</body>
</html>
<%
set conn = nothing
end if%>
