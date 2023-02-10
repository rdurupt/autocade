<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
	<link rel=STYLESHEET href='<%=Session("StyleSheet")%>' type='text/css'>
<!--#include file="lib/admin-lib.asp"-->
<script language="JavaScript">

function is_decimal(v)
{
	var f=document.prod_ajout;
	var pcpt=0;
	var i=0;
	var c="";
	var ret=true;

	if (v.length == 0)
	{
		return true;
	}
	else
	{
		for (i=0; i<v.length; i++)
		{
			c = v.charAt(i);
			if (c==".") { pcpt++ }
			else if ((c < "0") || (c > "9")) { ret = false; }
		}
	}
	return (ret && (pcpt < 2))
}


function ajout()
{
	var f=document.liste_tva;
	var err="";
	
	if (f.tva_valeur.value.length < 1)
	{
		err += "\nVous devez saisir une valeur pour la tva";
	}
	
	if (err == "")
	{
		f.act.value="ajout";
		f.submit();
	}
	else
	{
		alert ("Le(s) erreur(s) suivante(s) sont apparue(s):" + err);
	}
}

function modif(id)
{
	var f=document.liste_tva;
	var err="";

	eval ("f.tva_valeur.value = f.tva_valeur_"+id+".value;");
	eval ("f.livrable.checked = f.livrable_"+id+".checked;");
	eval ("f.coeff.value = f.coeff_"+id+".value;");

	if (f.tva_valeur.value.length < 1)
	{
		err += "\nVous devez saisir une valeur pour la tva";
	}

	if (err == "")
	{
		f.act.value="modif";
		f.tva_id.value = id;
		f.submit();
	}
	else
	{
		alert ("Le(s) erreur(s) suivante(s) sont apparue(s):" + err);
	}
}

function suppr(id)
{
	var f=document.liste_tva;
	var err="";

	if (confirm("Êtes-vous sûr de vouloir supprimer cette tva ?"))
	{
		f.tva_id.value = id;
		f.act.value = "suppr";
		f.submit();
	}
}

</script>
</HEAD>

<BODY>
&nbsp;&nbsp;&nbsp;<a href="javascript:history.go(-1)">Retour</a>
<br>
<div align="center" class="smallerheader">Gestion des TVA</div>


<%
	set conn=server.createobject("adodb.connection")
	conn.open session("DSN")

	act = request("act")
	if act = "" then
		act="aff"
	end if


	if act="ajout" then
		SQL = "select max(tva_id) from dbp_tva"	
		set cur=conn.execute(SQL)
		newtva_id = cur(0)+1
		set cur = nothing
		tva_valeur = request("tva_valeur")
		SQL = "insert into dbp_tva(tva_id, tva_valeur) values ("
		SQL=SQL & newtva_id & ", "
		SQL=SQL & noquote2(tva_valeur) & ") "
'		response.write (SQL)
		conn.execute(SQL)
		act="aff"
	end if

	if act="modif" then
		tva_id = request("tva_id")
		tva_valeur = request("tva_valeur")
		SQL = "update dbp_tva set "
		SQL=SQL & "tva_valeur = " & noquote2(tva_valeur) & ", "
		SQL=SQL & "where tva_id = " & tva_id
'		response.write (SQL)
		conn.execute(SQL)
		act="aff"
	end if

	if act="suppr" then
		erreurs = ""
		tva_id = request("tva_id")
		SQL = "select count(*) from dbp_topics where tva_id = " & tva_id
		set cur=conn.execute(SQL)
		if cur(0)<>0 then
			erreurs = "<br><li>Cette TVA est référencée par " & cur(0) & " produit(s)"
		end if
		set cur = nothing

		if erreurs = "" then
			SQL = "delete from dbp_tva where tva_id = " & tva_id
			conn.execute(SQL)
		else
			response.write ("<p><div align=""center"" class=""alert"">Cette TVA ne peut pas être supprimée:" & erreurs)
		end if
		act = "aff"
	end if

	if act="aff" then
		%>
		<div align="center">
		<form name="liste_tva" action="tva.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="tva_id" value="">
		<table border=0>
		<tr>
			<td>
			<table border=0 width="100%" cellpadding=2>
			<tr class="trEntete">
				<td class="texttrEntete" align="center">Id</td>
				<td class="texttrEntete" align="center">Valeur</td>
				<td class="texttrEntete" align="center">Action</td>
			</tr>
			<tr>
				<td colspan=3>&nbsp;</td>
			</tr>
			<tr align="center">
				<td>[Auto]</td>
				<td><input type="text" maxlength=5 name="tva_valeur" value=""></td>
				<td><a href="javascript:ajout();">Ajouter</a></td>
			</tr>
			<%
			old_l = ""
			l=""
			SQL = "select tva_id, tva_valeur "
			SQL=SQL & "from dbp_tva order by 1"
			set cur=conn.execute(SQL)
			while not cur.eof
				%>
				<tr bgcolor="#FFFFFF" align="center">
					<td><%=cur(0)%></td>
					<td><input type="text" maxlength=100 name="tva_valeur_<%=cur(0)%>" value="<%=cur(1)%>"></td>
					<td>
						<a href="javascript:suppr(<%=cur(0)%>)">Supprimer</a>
						<a href="javascript:modif(<%=cur(0)%>)">Mise à Jour</a>
					</td>
				</tr>
				<%
				cur.movenext
			wend
			%>
			</table>
			</td>
		</tr>
		</table>
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
