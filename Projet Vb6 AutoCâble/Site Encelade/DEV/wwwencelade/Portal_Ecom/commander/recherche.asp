<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus 1.2">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<!--#include file="../lib/admin-lib.asp"-->
</HEAD>

<BODY BGCOLOR="#FFFFFF">
<%
	set conn=server.createobject("adodb.connection")
	conn.open myDSN


	search = ucase(request("search"))
	searchType = request("searchType")
	demo = len(cstr(request("demo")))>0

%>
<br><br>
<font class="titre2">Résultats de votre Recherche:</font>
<hr noshade width="100%" size=1><p>
<%
	if demo then
		%>
<font class="erreur"><u>Vous êtes en mode recherche:</u> Vous verrez ici la liste des produits correspondants, 
mais vous ne pourrez pas les commander. Veuillez cliquer sur le bouton "Commander en ligne" de la page d'acceuil pour suivre</font>
le processus de commande.<hr noshade size=1 width="20%"><p>
		<%
	end if
	
		sql = "select p.ref_cas, p.titre, p.prix ,  p.prix  * 6.55957 as prixf, p.id_produit from t_produit p, t_tva t where p.tva_id = t.tva_id "
		if searchType="reference" then
			sql = sql & " and upper(ref_cas) like '%" & noquote(replace(search," ","%")) & "%' "
		else
			sql = sql & " and upper(convert(varchar(255),motscles)) like '%" & noquote(replace(search," ","%")) & "%' "
		end if
'		response.Write ("<br>" & sql & "<br>")
		set cur = conn.execute(sql)
		if cur.eof then
			%>
			<blockquote>
			<font class="erreur">Il n'y a pas de produit</font>
			</blockquote>
			<%
		end if
		while not cur.eof 
			ref_cas = toZS(cur(0))
			titre = toZS(cur(1))
			prix = fmt(cur(2))
			prixf = fmt(cur(3))
			id_produit = cur(4)
			%>
			<table border=0 width="<% if not demo then %>100%<% else %>400<% end if %>">
			<tr>
				<td>
					<ul>
					<li>
					<% if not demo then %>
					<a href="produit.asp?id_produit=<%=id_produit%>"><font class="titre1"><%=ref_cas%></font></a>
					<% else %>
					<font class="titre1"><%=ref_cas%></font>
					<% end if %>
					<br><%=titre%>
					</ul>
				</td>
				<td valign="top" width="20%" align="right">
				<font color="#0000FF">FF<%=prixf%><br><img src="../img/euro.gif" border=0><%=prix%></font>
				</td>
			</tr>
			</table>
			<%
			cur.movenext
		wend



	conn.close
	set conn = nothing

%>
</BODY>
</HTML>