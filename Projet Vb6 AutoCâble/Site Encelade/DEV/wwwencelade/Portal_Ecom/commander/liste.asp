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
<center>
<%
	set conn=server.createobject("adodb.connection")
	conn.open myDSN

	id_categorie = request("id_categorie")
	
	fourn_reference = request("fourn_reference")
	
	if len(cstr(fourn_reference))>0 then
		sql = "select fourn_photo, fourn_desc, fourn_design from t_reference_fourn where fourn_reference = " & noquote2(fourn_reference) 
	else
		sql = "select cat_photo, cat_desc, cat_label from t_categorie where id_categorie = " & id_categorie
	end if
	set cur=conn.execute(sql)
	cat_photo = toZS(cur(0))
	cat_desc = cur(1)
	design = cur(2)
	%>
	<font class="titre2"><%=design%></font><br>
	<%
	if len(cat_photo)>0 then
		%><img src="<%=cat_photo%>" border=0><p><%
	end if
	if len(cat_desc)>0 then
		%><div class="description"><%=cat_desc%></div><%
	end if
	
	response.write(barre & "</center>" & CRLF)
	
	if id_categorie>0 then
'		response.write ("fuck!<br>")
		sql = "select ref_multiple from t_categorie where id_categorie = " & id_categorie
		set cur = conn.execute(sql)
		ref_multiple = cur(0)
		set cur = nothing
	end if

	%>
	<!--
	id_categorie = <%=id_categorie%>
	fourn_reference = <%=fourn_reference%>
	ref_multiple = <%=ref_multiple%>
	len(cstr(fourn_reference)) = <%=len(cstr(fourn_reference))%>
	-->
	<%



	if (ref_multiple and (id_categorie>0)) then	
		' produit complexe
'		response.write ("Produit complexe<br>")
		
		sql = "select fourn_reference, fourn_photo, fourn_design, fourn_desc from t_reference_fourn where id_categorie = " & id_categorie
		set cur = conn.execute(sql)
		if cur.eof then
		%>
		<p><font class="erreur">Il n'y a aucun produit de cette catégorie</font><p>
		<%
		end if 
		while not cur.eof 
			fourn_reference = cur(0)
			fourn_photo = cur(1)
			fourn_design = cur(2)
			fourn_desc = cur(3)
			%>
			<table border=0 width="100%">
			<tr>
				<td colspan=2><a href="liste.asp?fourn_reference=<%=server.htmlencode(fourn_reference)%>"><font class="titre2"><%=fourn_design%></font></a></td>
			</tr>
			<tr>
				<td>
				<% if len(toZS(fourn_photo))>0 then %>
					<img src="<%=fourn_photo%>" border=0>
				<% else %>
					&nbsp;
				<% end if %>
				</td>
				<td align="right">
				<%=fourn_reference%>
			</tr>
			<tr>
				<td colspan=2>
				<div class="description"><%=fourn_desc%>
				</td>
			</tr>
			</table>
			<%
			cur.movenext
		wend
	
	else
'		response.write ("Produit simple<br>")
		' produit simple
		sql = "select p.ref_cas, p.titre, p.prix,  p.prix  * 6.55957 as prixf, p.id_produit from t_produit p, t_tva t where p.tva_id = t.tva_id "
		if (len(cstr(fourn_reference))=0) then
			sql = sql & "and id_categorie = " & id_categorie
		else
			sql = sql & "and fourn_reference = " & noquote2(fourn_reference)
		end if
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
			<table border=0 width="100%">
			<tr>
				<td>
					<ul>
					<li><a href="produit.asp?id_produit=<%=id_produit%>"><font class="titre1"><%=ref_cas%></font></a><br>
					<%=titre%>
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
	
	
	end if
	
	
	
	
	conn.close
	set conn = nothing

%>
</BODY>
</HTML>