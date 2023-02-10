<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus 1.2">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
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
		ret = false;
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


function modif_tva(inp, tva_id)
{
	var f = document.f_tva;
	if (! is_decimal(inp.value))
	{
		alert ("La valeur saisie est incorrecte");
	}
	else
	{
		f.act.value = "modif_tva";
		f.id.value = tva_id;
		f.submit();
	}
}

function suppr_tva(tva_id)
{
	var f = document.f_tva;
	if (confirm ("Êtes-vous sûr de vouloir supprimer cet enregistrement ?"))
	{
		f.act.value = "suppr_tva";
		f.id.value = tva_id;
		f.submit();
	}
}

function ajout_tva()
{
	var f = document.f_tva;
	if (! is_decimal(f.tva_new.value))
	{
		alert ("La valeur saisie est incorrecte");
	}
	else
	{
		f.act.value = "ajout_tva";
		f.submit();
	}
}

</script>
</HEAD>

<BODY BGCOLOR="#FFFFFF">
<font class="titre">Taux de TVA</font>
<%

	set conn=server.createobject("adodb.connection")
	conn.open myDSN


	act = request("act")
	if act="" then
		act="aff"
	end if

	if act="ajout_tva" then
		tva_valeur = request("tva_new")
		SQL = "select count(*) from t_tva where tva_valeur = " & tva_valeur
		set cur = conn.execute (SQL)
		cpt = cur(0)
		set cur=nothing
		if cpt = 0 then
			SQL = "select max(tva_id) from t_tva"
			set cur=conn.execute(SQL)
			tva_id = inc(cur(0))
			set cur=nothing
			SQL = "insert into t_tva values (" & tva_id & ", " & tva_valeur & ")"
			conn.execute(SQL)
		else
			call jsalert("Ce taux de TVA existe déjà en base!\n\nVeuillez en saisir un nouveau")
		end if
	end if

	if act="modif_tva" then
		tva_id = request("id")
		tva_valeur= request("tva_" & tva_id)
		SQL = "update t_tva set tva_valeur = " & tva_valeur & " where tva_id=" & tva_id
		conn.execute (SQL)
	end if

	if act="suppr_tva" then
		tva_id = request("id")
		SQL = "select count(*) from t_produit where tva_id=" & tva_id
		set cur=conn.execute(SQL)
		cpt = cur(0)
		set cur=nothing
		if cpt=0 then
			SQL = "delete from t_tva where tva_id=" & tva_id
			conn.execute(SQL)
		else
			call jsalert("Ce taux de TVA est appliqué à " & cpt & " produit(s).\nVous ne pouvez pas le supprimer.")
		end if
	end if



	
	
	barre()

	%>
	<font class="titre1">Taux de TVA:</font>
	<form name="f_tva" action="reference.asp" method="post">
	<input type="hidden" name="id" value="">
	<input type="hidden" name="act" value="">
	<blockquote>
	<table border=0>
	<tr>
		<td bgcolor="#000000">
			<table border=0 cellpadding=5>
			<tr bgcolor="#808080">
				<td><font color="#FFFF00"><b>Num.</b></font></td>
				<td><font color="#FFFF00"><b>Valeur</b></font></td>
				<td><font color="#FFFF00"><b>Action</b></font></td>
			</tr>
			<%
			SQL = "select tva_id, tva_valeur from t_tva order by 1"
			set cur = conn.execute(SQL)
			while not cur.eof
				%>
				<tr bgcolor="#FFFFFF">
				<td><%=cur(0)%></td>
				<td><input type="text" name="tva_<%=cur(0)%>" value="<%=replace(cstr(cur(1)),",",".")%>" size=5></td>
				<td><a href="javascript:modif_tva(document.f_tva.tva_<%=cur(0)%>,<%=cur(0)%>)">M.A.J.</a> &nbsp;&nbsp;&nbsp;
					<a href="javascript:suppr_tva(<%=cur(0)%>)">Supprimer</a></td>
				</tr>
				<%
				cur.movenext
			wend
			%>
			<tr bgcolor="#FFFFFF">
				<td>*</td>
				<td><input type="text" name="tva_new" size=5></td>
				<td><a href="javascript:ajout_tva()">Ajouter</a></td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	</blockquote>
	</form>


<%
	conn.close()
	set conn = nothing
%>
</BODY>
</HTML>
