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
	var f=document.liste_pays;
	var err="";
	
	if (f.label_pays.value.length < 1)
	{
		err += "\nVous devez saisir le nom du pays";
	}
	if (! is_decimal(f.coeff.value))
	{
		err += "\nLe coefficient est invalide";
	}
	
	if (err == "")
	{
		f.act.value="ajout";
		f.submit();
	}
	else
	{
		alert ("Le(s) erreur(s) suivante(s) sont apparue(s):" + err);
		f.label_pays.focus();
	}
}

function modif(id)
{
	var f=document.liste_pays;
	var err="";

	eval ("f.label_pays.value = f.label_pays_"+id+".value;");
	eval ("f.livrable.checked = f.livrable_"+id+".checked;");
	eval ("f.coeff.value = f.coeff_"+id+".value;");

	if (f.label_pays.value.length < 1)
	{
		err += "\nVous devez saisir le nom du pays";
	}
	if (! is_decimal(f.coeff.value))
	{
		err += "\nLe coefficient est invalide";
	}

	if (err == "")
	{
		f.act.value="modif";
		f.id_pays.value = id;
		f.submit();
	}
	else
	{
		alert ("Le(s) erreur(s) suivante(s) sont apparue(s):" + err);
	}
}

function suppr(id)
{
	var f=document.liste_pays;
	var err="";

	if (confirm("Êtes-vous sûr de vouloir supprimer ce pays ?"))
	{
		f.id_pays.value = id;
		f.act.value = "suppr";
		f.submit();
	}
}

</script>
</HEAD>

<BODY>
<a name="#haut">
&nbsp;&nbsp;&nbsp;<a href="javascript:history.go(-1)">Retour</a>
<br>
<div align="center" class="smallerheader">Gestion des Pays</div>


<%
	set conn=server.createobject("adodb.connection")
	conn.open session("DSN")

	nb_pays = request("nb_pays")
	if cstr(nb_pays) = "" then
		SQL = "select count(*) from t_pays"
		set cur=conn.execute (SQL)
		nb_pays = cur(0)
		set cur = nothing
	end if


	act = request("act")
	if act = "" then
		act="aff"
	end if


	if act="ajout" then
		SQL = "select max(id_pays) from t_pays"
		set cur=conn.execute(SQL)
		id_pays = inc(cur(0))
		set cur = nothing
		label_pays = request("label_pays")
		label_pays = ucase(mid(label_pays, 1, 1)) & mid(label_pays, 2)
		livrable = request("livrable")
		if cstr(livrable) = "" then
			livrable = 0
		else
			livrable = 1
		end if
		coeff = toNull(request("coeff"))
		SQL = "insert into t_pays(id_pays, label_pays, livrable, tarif_normal, tarif_express) values ("
		SQL=SQL & id_pays & ", "
		SQL=SQL & noquote2(label_pays) & ", "
		SQL=SQL & livrable & ", "
		SQL=SQL & coeff & ", "
		SQL=SQL & "NULL) "
'		response.write (SQL)
		conn.execute(SQL)
		act="aff"
	end if

	if act="modif" then
		id_pays = request("id_pays")
		label_pays = request("label_pays")
		label_pays = ucase(mid(label_pays, 1, 1)) & mid(label_pays, 2)
		livrable = request("livrable")
		if cstr(livrable) = "" then
			livrable = 0
		else
			livrable = 1
		end if
		coeff = toNull(request("coeff"))
		tarif_express = toNull(request("tarif_express"))
		SQL = "update t_pays set "
		SQL=SQL & "label_pays = " & noquote2(label_pays) & ", "
		SQL=SQL & "livrable = " & livrable & ", "
		SQL=SQL & "tarif_normal = " & coeff & " "
		SQL=SQL & "where id_pays = " & id_pays
'		response.write (SQL)
		conn.execute(SQL)
		act="aff"
	end if

	if act="suppr" then
		erreurs = ""
		id_pays = request("id_pays")
		SQL = "select count(*) from dbp_userinfos where id_pays = " & id_pays
		set cur=conn.execute(SQL)
		if cur(0)<>0 then
			erreurs = "<br><li>Il est référencé par " & cur(0) & " client(s)"
		end if
		set cur = nothing

		SQL = "select count(*) from t_livraison where id_pays = " & id_pays
		set cur=conn.execute(SQL)
		if cur(0)<>0 then
			erreurs = "<br><li>Il est référencé par " & cur(0) & " adresse(s) de livraison"
		end if
		set cur = nothing

		if erreurs = "" then
			SQL = "delete from t_pays where id_pays = " & id_pays
			conn.execute(SQL)
		else
			response.write ("Ce produit ne peut pas être supprimé:" & erreurs)
			response.write ("<p>Cliquez <a href=""javascript:history.back()"">ici</a> ou cliquez sur le bouton 'Back' votre navigateur pour revenir à la liste")

		end if
		act = "aff"
	end if

	if act="aff" then
		%>
		<div align="center">
		<form name="liste_pays" action="pays.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="nb_pays" value="<%=nb_pays%>">
		<input type="hidden" name="id_pays" value="">
		<table border=0>
		<tr>
			<td>
			<table border=0 width="100%" cellpadding=2>
			<tr class="trEntete">
				<td class="texttrEntete" align="center">Id</td>
				<td class="texttrEntete" align="center">Nom</td>
				<td class="texttrEntete" align="center">Livrable</td>
				<td class="texttrEntete" align="center">Coefficient</td>
				<td class="texttrEntete" align="center">Action</td>
			</tr>
			<tr>
				<td colspan=6><font size="2"><b><i>Nouveau...</i></b></font></td>
			</tr>
			<tr align="center">
				<td>[Auto]</td>
				<td><%=etoile%><input type="text" maxlength=100 name="label_pays" value=""></td>
				<td><input type="checkbox" name="livrable"></td>
				<td><input type="text" maxlength=10 name="coeff" value="1"></td>
				<td><a href="javascript:ajout();">Ajouter</a></td>
			</tr>
			<%
			old_l = ""
			l=""
			SQL = "select id_pays, label_pays, livrable, tarif_normal, tarif_express "
			SQL=SQL & "from t_pays order by 2"
			set cur=conn.execute(SQL)
			while not cur.eof
				l=left (ucase(cur(1)),1)
				if l<>old_l then
					old_l = l
					%>
						<tr bgcolor="#FFFFFF">
							<td colspan=6><font size="3"><b><i><%=l%></i></b></font></td>
						</tr>
					<%
				end if
				%>
				<tr bgcolor="#FFFFFF" align="center">
					<td><%=cur(0)%></td>
					<td><input type="text" maxlength=100 name="label_pays_<%=cur(0)%>" value="<%=cur(1)%>"></td>
					<td><input type="checkbox" name="livrable_<%=cur(0)%>" <%=checked(cur(2))%> ></td>
					<td><input type="text" maxlength=10 name="coeff_<%=cur(0)%>" value="<%=replace(toZS(cur(3)),",",".")%>"></td>
					<td>
						<a href="javascript:suppr(<%=cur(0)%>)">Supprimer</a>&nbsp;-&nbsp;
						<a href="javascript:modif(<%=cur(0)%>)">Mise à Jour</a>&nbsp;-&nbsp;
						<a href="frais_de_port.asp?id_pays=<%=cur(0)%>">Frais de port</a>
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
&nbsp;&nbsp;&nbsp;<a href="javascript:history.go(-1)">Retour</a>&nbsp;&nbsp;&nbsp;<a href="#haut">Haut de page</a>
</BODY>
</HTML>
