<% response.addheader "Pragma","no-cache" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<!--#include file="lib/admin-lib.asp"-->
<script language="JavaScript">
function ajout()
{
	var f=document.prod_sel;

	f.act.value="ajout";
	f.action = "synth_card.asp";
	f.submit();
}

function modif()
{
	var f=document.prod_sel;

	if (f.id_produit.selectedIndex > -1)
	{
		f.act.value="modif";
		f.submit();
	}
	else
	{
		alert ("Vous devez sélectionner une catégorie");
	}
}

function suppr()
{
	var f=document.prod_sel;

	if (f.id_produit.selectedIndex > -1)
	{
		f.action = "synth_card.asp";
		f.act.value="suppr";
		f.submit();
	}
	else
	{
		alert ("Vous devez sélectionner une catégorie");
	}
}


function set_action()
{
	var f=document.prod_sel;
	f.action = f.destination.options[f.destination.selectedIndex].value;
	document.criteres.destination.value = f.action;
}
</script>
</HEAD>

<BODY BGCOLOR="#FFFFFF">
<%
	set conn=server.createobject("adodb.connection")
	conn.open myDSN

	multipart = 0


	act = request("act")
	if act = "" then
		act="aff"
	end if


	if act="aff" then
		id_categorie = request("id_categorie")
		if id_categorie = "" then
			id_categorie = "-1"
		end if
		fourn_reference = request("fourn_reference")
		if fourn_reference = "" then
			fourn_reference = "-1"
		end if
		destination = request("destination")
		if destination="" then
			destination = "synth_card.asp"
		end if
		%>
		<div align="center">
		<form name="criteres" action="produits.asp" method="post">
		<input type="hidden" name="destination" value="<%=destination%>">
		<input type="hidden" name="act" value="aff">
		<font class="titre1">Critères:</font><br>
		<table border=0 cellpadding=5>
		<tr>
			<td align="center">
				Catégorie:<br>
				<select name="id_categorie" OnChange="document.criteres.submit();">
				<option value="-1" <%=selected(id_categorie,"-1")%> >[Toutes]
				<%
				SQL = "select id_categorie, cat_label from t_categorie order by 2"
				set cur = conn.execute(SQL)
				while not cur.eof
					%>
					<option value="<%=cur(0)%>" <%=selected(cur(0),id_categorie)%> ><%=cur(1)%>
					<%
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
			<td align="center">
				Ref. Fournisseur:<br>
				<select name="fourn_reference" OnChange="document.criteres.submit();">
				<option value="-1" <%=selected(fourn_reference,"-1") %> >[Toutes]
				<%
				SQL = "select fourn_reference from t_reference_fourn order by 1"
'				response.write(SQL & "<br>")
				set cur = conn.execute(SQL)
				while not cur.eof
					%>
					<option value="<%=server.HTMLencode(cur(0))%>" <%=selected(cur(0),fourn_reference)%> ><%=cur(0)%>
					<%
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr>
		</table>
		</form>
		<% barre() %>
		<form name="prod_sel" action="synth_card.asp" method="post">
		<input type="hidden" name="act" value="">
		<%
		SQL="select id_produit, titre from t_produit "
		if cstr(id_categorie)<>"-1" then
			SQL = SQL & "where id_categorie = " & id_categorie & " "
		end if
		if cstr(fourn_reference)<>"-1" then
			if cstr(id_categorie)<>"-1" then
				SQL = SQL & " and "
			else
				SQL = SQL & " where "
			end if
			SQL = SQL & "fourn_reference = '" & noquote(fourn_reference) & "' "
		end if
		SQL=SQL & " order by 2"
		set cur = conn.execute(SQL)
		if not cur.eof then
			%>
			<font class="titre1">Choisissez:</font><br>
			<select name="id_produit" size=10>
				<%
				while not cur.eof
					response.write ("<option value=""" & cur(0) & """>" & cur(1) & CRLF)
					cur.movenext
				wend
				set cur=nothing
				%>
			</select>
			<%
		else
			response.write ("<i>Aucun produit ne correspond à ces critères...</i>")
		end if
		barre() %>
		<font class="titre1">Type de Fiche:</font><br>
		<select name="destination" size=2 onChange="set_action()">
		<option value="synth_card.asp" <%=selected(destination,"synth_card.asp")%>>Fiche de Synthèse
		<option value="exp_card.asp" <%=selected(destination,"exp_card.asp")%>>Fiche Expert (modification seulement)
		</select>
		<% barre() %>
		<input type="button" value="Nouvelle..." onClick="ajout();">&nbsp;
		<input type="button" value="Modifier" onClick="modif();">&nbsp;
		<input type="button" value="Supprimer" onClick="suppr();">
		</form>
		</div>
		<%
	end if


	if multipart = 1 then
		set upl=nothing
	end if
	conn.close()
	set conn = nothing
%>
</BODY>
</HTML>
