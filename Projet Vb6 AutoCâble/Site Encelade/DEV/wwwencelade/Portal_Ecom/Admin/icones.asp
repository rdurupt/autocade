<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
	<link rel=STYLESHEET href='<%=Session("StyleSheet")%>' type='text/css'>
	<script language="javascript" src="../Portal_Java/Portal_Generic.js"></script>
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
	var f=document.f_icones;
	var err="";
	
	if (f.nom.value.length < 1)
	{
		err += "\nVous devez saisir le nom de l'icône";
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
	var f=document.f_icones;
	var err="";

	eval ("f.nom.value = f.nom_"+id+".value;");
	eval ("f.chemin.value = f.chemin_"+id+".value;");

	if (f.nom.value.length < 1)
	{
		err += "\nVous devez saisir le nom de votre icônes";
	}

	if (f.chem.value.length < 1)
	{
		err += "\nVous devez saisir le chemin de votre icônes";
	}
	
	if (err == "")
	{
		f.act.value="modif";
		f.IconId.value = id;
		f.submit();
	}
	else
	{
		alert ("Le(s) erreur(s) suivante(s) sont apparue(s):" + err);
	}
}

function suppr(id)
{
	var f=document.f_icones;
	var err="";

	if (confirm("Êtes-vous sûr de vouloir supprimer cet icône ?"))
	{
		f.IconId.value = id;
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
<div align="center" class="smallerheader">Gestion des Icônes</div>


<%
	set conn=server.createobject("adodb.connection")
	conn.open session("DSN")

	act = request("act")
	if act = "" then
		act="aff"
	end if


	if act="ajout" then
		SQL = "select max(IconId) from dbp_ecom_icones"
		set cur=conn.execute(SQL)
		IconId = inc(cur(0))
		set cur = nothing
		nom = request("nom")
		chemin = request("chemin")
		SQL = "insert into dbp_ecom_icones(IconId, nom, chemin) values ("
		SQL=SQL & IconId & ", "
		SQL=SQL & noquote2(nom) & ", "
		SQL=SQL & noquote2(chemin) & ") "
'		response.write (SQL)
		conn.execute(SQL)
		act="aff"
	end if

	if act="modif" then
		IconId = request("IconId")
		nom = request("nom")
		chemin = request("chemin")
		SQL = "update dbp_ecom_icones set "
		SQL=SQL & "nom = " & noquote2(nom) & ", "
		SQL=SQL & "chemin = " & noquote2(chemin)
		SQL=SQL & "where IconId = " & IconId
'		response.write (SQL)
		conn.execute(SQL)
		act="aff"
	end if

	if act="suppr" then
		erreurs = ""
		IconId = request("IconId")
		SQL = "select count(*) from dbp_topics where IconEcom = " & IconId
		set cur=conn.execute(SQL)
		if cur(0)<>0 then
			erreurs = "<br><li>Il est référencé par " & cur(0) & " produit(s)"
		end if
		set cur = nothing

		if erreurs = "" then
			SQL = "delete from dbp_ecom_icones where IconId = " & IconId
			conn.execute(SQL)
		else
			response.write ("<div align=center><br><br>Ce produit ne peut pas être supprimé:" & erreurs)
			response.write ("<p>Cliquez <a href=""javascript:history.back()"">ici</a> ou cliquez sur le bouton 'Back' votre navigateur pour revenir à la liste")
			response.end
		end if
		act = "aff"
	end if

	if act="aff" then
	Set Portal = Server.CreateObject(Session("PortalComObject") & ".Portal")
		%>
		<div align="center">
		<form name="f_icones" action="icones.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="IconId" value="">
		<table border=0>
		<tr>
			<td>
			<table border=0 width="100%" cellpadding=2>
			<tr class="trEntete">
				<td class="texttrEntete" align="center">Id</td>
				<td class="texttrEntete" align="center">Nom</td>
				<td class="texttrEntete" align="center"><%=Portal.GetDefault("EcomPathImagesIcones","../euxia/images/icones")%></td>
			</tr>
			<tr>
				<td colspan=3><font size="2"><b><i>Nouveau...</i></b></font></td>
			</tr>
			<tr align="center">
				<td>[Auto]</td>
				<td><%=etoile%><input type="text" maxlength=100 size="31" name="nom" value=""></td>
				<td><input type="text" name="chemin" size="32" value=""></td>
				<td><a href="javascript:ajout();">Ajouter</a></td>
			</tr>
			<tr>
				<td colspan=3>&nbsp;</td>
			</tr>

			<%
			SQL = "select IconId, nom, chemin "
			SQL = SQL & " from dbp_ecom_icones order by IconId"
			set cur=conn.execute(SQL)
			while not cur.eof
			%>
				<tr align="center">
					<td><%=cur("IconId")%></td>
					<td><input type="text" maxlength=100 size="32" name="nom_<%=cur("IconId")%>" value="<%=cur("Nom")%>"></td>
					<td><input type="text" name="chemin_<%=cur("IconId")%>" size="32" value="<%=cur("Chemin")%>"></td>
					<td>
						<img src="<%=Portal.GetDefault("EcomPathImagesIcones","../euxia/images/icones") & cur("Chemin")%>">&nbsp;-&nbsp;
						<a href="javascript:suppr(<%=cur(0)%>)">Supprimer</a>&nbsp;-&nbsp;
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

	Portal.Pr(Portal.ViewBasPage("80%","center"))
	
	Set Portal = nothing
%>
</BODY>
</HTML>
