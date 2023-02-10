<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
<link rel='stylesheet' href='<%=session("stylesheet")%>'>
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
	var f=document.liste_frais;
	var err="";
	
	if (! is_decimal(f.poids_mini.value))
	{
		err += "\nLe poids mini est invalide";
	}
	if (! is_decimal(f.poids_maxi.value))
	{
		err += "\nLe poids maxi est invalide";
	}
	if (! is_decimal(f.prix.value))
	{
		err += "\nLe prix est invalide";
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
	var f=document.liste_frais;
	var err="";

	eval ("f.poids_mini.value = f.poids_mini_"+id+".value;");
	eval ("f.poids_maxi.value = f.poids_maxi_"+id+".value;");
	eval ("f.prix.value = f.prix_"+id+".value;");

	if (! is_decimal(f.poids_mini.value))
	{
		err += "\nLe poids mini est invalide";
	}
	if (! is_decimal(f.poids_maxi.value))
	{
		err += "\nLe poids maxi est invalide";
	}
	if (! is_decimal(f.prix.value))
	{
		err += "\nLe prix est invalide";
	}

	if (err == "")
	{
		f.act.value="modif";
		f.id_tarif_livraison.value = id;
		f.submit();
	}
	else
	{
		alert ("Le(s) erreur(s) suivante(s) sont apparue(s):" + err);
	}
}

function suppr(id)
{
	var f=document.liste_frais;
	var err="";

	if (confirm("Êtes-vous sûr de vouloir supprimer ce pays ?"))
	{
		f.id_tarif_livraison.value = id;
		f.act.value = "suppr";
		f.submit();
	}
}

</script>
</HEAD>

<BODY>
<a name="#haut">
&nbsp;&nbsp;&nbsp;<a href="javascript:history.go(-1)">Retour</a>
<div align="center">
<font class="smallerheader">Frais de port des pays</font>
<%

	set conn=server.createobject("adodb.connection")
	conn.open session("dsn")

	id_pays = request("id_pays")
	id_tarif_livraison = request("id_tarif_livraison")
	
	nb_tarif = request("nb_tarif")
	if cstr(nb_tarif) = "" then
		SQL = "select count(*) from t_tarif_livraison"
		set cur=conn.execute (SQL)
		nb_tarif = cur(0)
		set cur = nothing
	end if
	
	act = request("act")
	if act = "" then
		act="aff"
	end if


	if act="ajout" then
		SQL = "select max(id_tarif_livraison) from t_tarif_livraison"
		set cur=conn.execute(SQL)
		id_tarif_livraison = inc(cur(0))
		set cur = nothing
		poids_mini = request("poids_mini")
		poids_maxi = request("poids_maxi")
		prix = request("prix")
		id_pays = request("id_pays")

		SQL = "insert into t_tarif_livraison(id_tarif_livraison, id_pays, poids_mini, poids_maxi, prix) values ("
		SQL=SQL & id_tarif_livraison & ", "
		SQL=SQL & id_pays & ", "
		SQL=SQL & poids_mini & ", "
		SQL=SQL & poids_maxi & ", "
		SQL=SQL & prix & ") "
'		response.write (SQL)
		conn.execute(SQL)
		act="aff"
	end if

	if act="modif" then
		poids_mini = request("poids_mini")
		poids_maxi = request("poids_maxi")
		prix = request("prix")
		id_tarif_livraison = request("id_tarif_livraison")
		
		SQL = "update t_tarif_livraison set "
		SQL=SQL & "poids_mini = " & poids_mini & ", "
		SQL=SQL & "poids_maxi = " & poids_maxi & ", "
		SQL=SQL & "prix = " & prix & " "
		SQL=SQL & "where id_tarif_livraison = " & id_tarif_livraison
'		response.write (SQL)
		conn.execute(SQL)
		act="aff"
	end if

	if act="suppr" then
		erreurs = ""
		id_tarif_livraison = request("id_tarif_livraison")

		if erreurs = "" then
			SQL = "delete from t_tarif_livraison where id_tarif_livraison = " & id_tarif_livraison
			conn.execute(SQL)
		else
			response.write ("Ce tarif ne peut pas être supprimé:" & erreurs)
			response.write ("<p>Cliquez <a href=""javascript:history.back()"">ici</a> ou cliquez sur le bouton 'Back' votre navigateur pour revenir à la liste")

		end if
		act = "aff"
	end if

	if act="aff" then
		%>
		<div align="center">
		<form name="liste_frais" action="frais_de_port.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="id_tarif_livraison" value="">
		<input type="hidden" name="nb_tarif" value="<%=nb_tarif%>">
		<input type="hidden" name="id_pays" value="<%=id_pays %>">
		<table border=0>
		<tr>
			<td>
			<table border=0 width="100%" cellpadding=2>
			<tr class="trEntete">
				<td class="textEntete" align="center"><b>Id_pays</b></td>
				<td class="textEntete" align="center"><b>Poids mini</b></td>
				<td class="textEntete" align="center"><b>Poids maxi</b></td>
				<td class="textEntete" align="center"><b>Prix</b></td>
				<td class="textEntete" align="center"><b>Action</b></td>
			</tr>
			<tr>
				<td colspan=6><font size="3"><b><i>Nouveau...</i></b></font></td>
			</tr>
			<tr align="center">
				<td>[Nouveau]</td>
				<td><%=etoile%><input type="text" maxlength=20 name="poids_mini" value="0"></td>
				<td><input type="text" maxlength=20 name="poids_maxi" value="0"></td>
				<td><input type="text" maxlength=20 name="prix" value="0"></td>
				<td><a href="javascript:ajout();">Ajouter</a></td>
			</tr>
			<tr align="center">
				<td colspan=5>&nbsp;</td>
			</tr>
			<%
			old_l = ""
			l=""
			SQL = "select id_tarif_livraison, poids_mini, poids_maxi, prix "
			SQL=SQL & "from t_tarif_livraison where id_pays=" & id_pays & " order by 2"
'			SQL=SQL & "from t_tarif_livraison order by 2"
			'response.write sql & "<br>"
			'response.write act & "<br>"  
			'response.end
			set cur=conn.execute(SQL)
			while not cur.eof
				l=left (ucase(cur(1)),1)
				%>
				<tr align="center">
					<td><%=cur(0)%></td>
					<td><input type="text" maxlength=20 name="poids_mini_<%=cur(0)%>" value="<%=replace(toZS(cur(1)),",",".")%>"></td>
					<td><input type="text" maxlength=20 name="poids_maxi_<%=cur(0)%>" value="<%=replace(toZS(cur(2)),",",".")%>"></td>
					<td><input type="text" maxlength=20 name="prix_<%=cur(0)%>" value="<%=replace(toZS(cur(3)),",",".")%>"></td>
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
<br>
&nbsp;&nbsp;&nbsp;<a href="pays.asp">Retour à la liste des pays</a>
&nbsp;&nbsp;&nbsp;<a href="#haut">Haut de page</a>
</div>
</BODY>
</HTML>
