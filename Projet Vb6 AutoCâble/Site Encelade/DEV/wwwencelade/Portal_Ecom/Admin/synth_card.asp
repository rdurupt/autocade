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

function is_date(d)
{
	var ret = true;
	if (d.length != 0)		// date vide
	{
		if (d.length != 10)
		{
			ret = false;
		}
		else
		{
			yy = d.substring(6,10);
			mm = d.substring(3,5);
			dd = d.substring(0,2);
			mm--;
			var mydate = new Date(yy,mm,dd);
			y2 = mydate.getYear();
			if (y2<1000) { y2 += 1900; }
			m2 = mydate.getMonth();
			d2 = mydate.getDate();
			if ( (yy != y2) || (mm != m2) || (dd != d2) )
			{
				ret = false;
			}
		}
	}
	return ret;
}



function ajout_ok()
{
	var f=document.prod_ajout;
	var err = "";
	var ret = true;

	if (f.ref_cas.value.length < 1)
	{
		err += "\nRéférence CAS vide";
		ret = false;
	}

	if (f.titre.value.length < 1)
	{
		err += "\nDésignation vide";
		ret = false;
	}

	if (f.motscles.value.length < 1)
	{
		err += "\nMots-Clés (fr) vide";
		ret = false;
	}

	if (f.motscles_gb.value.length < 1)
	{
		err += "\nMots-Clés (gb) vide";
		ret = false;
	}

	if (f.texte_principal.value.length < 1)
	{
		err += "\nDescriptif (fr) vide";
		ret = false;
	}

	if (f.texte_principal_gb.value.length < 1)
	{
		err += "\nDescriptif (gb) vide";
		ret = false;
	}

	if (! is_decimal(f.prix.value))
	{
		err += "\nPrix Invalide";
		ret = false;
	}

	if (f.tva_id.options.selectedIndex == -1)
	{
		err += "\nSélectionner un taux de TVA";
		ret = false;
	}

	if (! is_decimal(f.poids.value))
	{
		err += "\nPoids invalide";
		ret = false;
	}


	if (f.id_categorie.options.selectedIndex < 1)
	{
		err += "\nSélectionner une catégorie";
		ret = false;
	}


	if (! ret)
	{
		alert ("Les erreurs suivantes sont apparues:" + err);
	}
	else
	{
		f.act.value = "ajout_ok";
		f.submit();
	}
	
}

function modif_ok()
{
	var f=document.prod_modif;
	var err = "";
	var ret = true;

	if (f.titre.value.length < 1)
	{
		err += "\nLa désignation est vide";
		ret = false;
	}

	if (f.motscles.value.length < 1)
	{
		err += "\nMots-Clés (fr) vide";
		ret = false;
	}

	if (f.motscles_gb.value.length < 1)
	{
		err += "\nMots-Clés (gb) vide";
		ret = false;
	}

	if (f.texte_principal.value.length < 1)
	{
		err += "\nDescriptif (fr) vide";
		ret = false;
	}

	if (f.texte_principal_gb.value.length < 1)
	{
		err += "\nDescriptif (gb) vide";
		ret = false;
	}

	if (! is_decimal(f.prix.value))
	{
		err += "\nPrix Invalide";
		ret = false;
	}

	if (f.tva_id.options.selectedIndex == -1)
	{
		err += "\nSélectionner un taux de TVA";
		ret = false;
	}

	if (! is_decimal(f.poids.value))
	{
		err += "\nPoids invalide";
		ret = false;
	}

	if (f.id_categorie.options.selectedIndex < 1)
	{
		err += "\nSélectionner une catégorie";
		ret = false;
	}

	if (! ret)
	{
		alert ("Les erreurs suivantes sont apparues:" + err);
	}
	else
	{
		f.act.value = "modif_ok";
		f.submit();
	}
	
}

function suppr_ok()
{
	var f=document.prod_suppr;

	f.act.value = "suppr_ok";
	f.submit();
}

function cancel(f)
{
	f.action = "produits.asp";
	f.act.value="aff";
	f.submit();
}

</script>

</HEAD>
<BODY BGCOLOR="#FFFFFF">
<font class="titre">Produits</font>
<%
function retour()	' retourne à la liste de produits...
	response.write ("<script language=""JavaScript"">" & CRLF)
	response.write ("top.droite.location = ""produits.asp?act=aff"";" & CRLF)
	response.write ("</script>" & CRLF)
end function


	set conn=server.createobject("adodb.connection")
	conn.open myDSN


	act = request("act")
	if act="" then
		act="aff"
	end if

	id_produit=request("id_produit")


	if (act="modif") and (cstr(id_produit)<>"") then 
		barre()
		%>
		<div align="center">
		<font class="titre2">[&nbsp;<font color="#FF0000">Fiche de Synthèse</font>&nbsp;]
										&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
										[&nbsp;<a href="exp_card.asp?act=modif&id_produit=<%=id_produit%>" target="droite">Fiche Expert</a>&nbsp;]
		</font>
		</div>
		<%
	end if

	
	barre()
	

	if act="ajout" then
		id_categorie = request("id_categorie")
		fourn_reference = request("fourn_reference")
		%>
		<div align="center">
		<form name="prod_ajout" action="synth_card.asp" method="post">
		<input type="hidden" name="act" value="">
		<table border=0 cellpadding=3>
		<tr>
			<td align="left"><%=etoile%>Référence CAS:</td>
			<td><input type="text" maxlength=20 name="ref_cas" size=20></td>
		</tr><tr>
			<td align="left"><%=etoile%>Catégorie:</td>
			<td>
				<select name="id_categorie">
				<option value="">[Choisissez]
				<%
				SQL = "select id_categorie, cat_label from t_categorie order by 2"
				set cur = conn.execute(SQL)
				while not cur.eof
					%>	<option value="<%=cur(0)%>" <%=selected(cur(0),id_categorie)%> ><%=cur(1)%>	<%
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr><tr>
			<td align="left">Famille des<br>produits complexes</td>
			<td>
				<select name="fourn_reference">
				<option value="">[Choisissez]
				<%
				SQL = "select fourn_reference from t_reference_fourn order by 1"
				set cur = conn.execute(SQL)
				while not cur.eof
					%>	<option value="<%=server.HTMLencode(cur(0))%>" <%=selected(cur(0),fourn_reference)%> ><%=cur(0)%>	<%
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr><tr>
			<td align="left">Ref. JEP</td>
			<td><input type="text" maxlength=50 name="fourn_family" size=50></td>
		</tr><tr>
			<td align="left"><%=etoile%>Désignation:</td>
			<td><input type="text" maxlength=80 name="titre" size=80></td>
		</tr><tr>
			<td align="left"><%=etoile%>Mots-clés (fr):</td>
			<td><textarea name="motscles" cols=40 rows=5></textarea></td>
		</tr><tr>
			<td align="left"><%=etoile%>Mots-clés (gb):</td>
			<td><textarea name="motscles_gb" cols=40 rows=5></textarea></td>
		</tr><tr>
			<td align="left"><%=etoile%>Descriptif (fr):</td>
			<td><textarea name="texte_principal" cols=40 rows=5></textarea></td>
		</tr><tr>
			<td align="left"><%=etoile%>Descriptif (gb):</td>
			<td><textarea name="texte_principal_gb" cols=40 rows=5></textarea></td>
		</tr><tr>
			<td align="left"><%=etoile%>Photo:</td>
			<td><select name="photo">
				<option value="">Choisissez...
				<%
				Set fs = CreateObject("Scripting.FileSystemObject")
				dapath = LOCAL_PATH & "\img\produit\small"
				Set f = fs.GetFolder(dapath)
				    Set fc = f.Files
				    For Each f1 in fc
						response.write ("<option value=""" &  HTTP_PATH & "/img/produit/small/" & f1.name & """>" & f1.name & CRLF )
				    Next
				%>
				</select>
			</td>
		</tr><tr>
			<td align="left">Mise en ligne:</td>
			<td><select name="enligne" size=2>
				<option value="1">Oui
				<option value="0" selected>Non
				</select>
			</td>
		</tr><tr>
			<td align="left"><%=etoile%>Prix en Euros:</td>
			<td><input type="text" maxlength=50 name="prix" size=50></td>
		</tr><tr>
			<td align="left"><%=etoile%>Taux de TVA:</td>
			<td>
				<select name="tva_id">
				<%
				SQL = "select tva_id, convert(varchar,tva_valeur) from t_tva order by 1"
				set cur = conn.execute(SQL)
				while not cur.eof
					response.write ("<option value=""" & cur(0) & """>" & cur(1))
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr><tr>
			<td align="left"><%=etoile%>Poids:</td>
			<td><input type="text" maxlength=10 size=10 name="poids" value=""></td>
		</tr><tr>
			<td align="left">Produit lié:</td>
			<td>
				<select name="cross_selling">
				<option value="NULL" SELECTED>[Aucun]
				<%
				SQL = "select id_produit, titre from t_produit order by 2"
				set cur = conn.execute(SQL)
				while not cur.eof
					response.write ("<option value=""" & cur(0) & """>" & cur(1) & CRLF)
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr>
			<td colspan=2 align="left">
			Les champs précédés par <%=etoile%> sont obligatoires...
			</td>
			<td></td>
		</table>
		<% barre() %>		
		<input type="button" value="OK" onClick="ajout_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.prod_ajout);">
		</form>
		</div>
		<%
	end if

	if act="ajout_ok" then
		SQL = "select max(id_produit) from t_produit"
		set cur = conn.execute (SQL)
		id_produit = inc(cur(0))
		set cur=nothing
		id_categorie = toNull(request("id_categorie"))
		ref_cas = toNull(request("ref_cas"))
		titre = toNull(request("titre"))
		texte_principal = toNull(request("texte_principal"))
		texte_principal_gb = toNull(request("texte_principal_gb"))
		enligne = toNull(request("enligne"))
		prix = toNull(request("prix"))
		motscles = toNull(request("motscles"))
		motscles_gb = toNull(request("motscles_gb"))
		cross_selling = toNull(request("cross_selling"))
		fourn_family = toNull(request("fourn_family"))
		fourn_reference = toNull(request("fourn_reference"))
		photo = request("photo")
		tva_id = request("tva_id")
		poids = request("poids")
		page_sommaire = request("page_sommaire")
		
		SQL = "insert into t_produit (id_produit,id_categorie,ref_cas,titre,texte_principal,texte_principal_gb,enligne,prix,motscles,motscles_gb,photo,dcre_produit,dmod_produit,cross_selling,fourn_family,fourn_reference,tva_id,poids) values ("
		SQL=SQL & id_produit & ", "
		SQL=SQL & id_categorie & ", "
		SQL=SQL & "'" & noquote(ref_cas) & "', "
		SQL=SQL & "'" & noquote(titre) & "', "
		SQL=SQL & "'" & nl2br(noquote(texte_principal)) & "', "
		SQL=SQL & "'" & nl2br(noquote(texte_principal_gb)) & "', "
		SQL=SQL & enligne & ", "
		SQL=SQL & prix & ", "
		SQL=SQL & "'" & nl2br(noquote(motscles)) & "', "
		SQL=SQL & "'" & nl2br(noquote(motscles_gb)) & "', "
		SQL=SQL & noquote2(photo) & ", "
		SQL=SQL & "'" & now & "', "
		SQL=SQL & "'" & now & "', "
		SQL=SQL & cross_selling & ", "
		SQL=SQL & noquote2(fourn_family) & ", "
		SQL=SQL & noquote2(fourn_reference) & ", "
		SQL=SQL & tva_id & ", " 
		SQL=SQL & poids & ") "
		response.write(SQL)
		conn.execute (SQL)
		retour()
	end if

	if act="modif" then
		id_produit = request("id_produit")
		SQL = "select id_categorie, ref_cas, titre, texte_principal, texte_principal_gb, enligne, prix, motscles, motscles_gb, photo, dcre_produit, dmod_produit, cross_selling, fourn_family, fourn_reference, tva_id, poids, page_sommaire from t_produit where id_produit = " & id_produit
		set cur = conn.execute (SQL)
		id_categorie = toNull(cur(0))
		ref_cas = cur(1)
		titre = cur(2)
		texte_principal = cur(3)
		texte_principal_gb = cur(4)
		enligne = cur(5)
		prix = cur(6)
		motscles = cur(7)
		motscles_gb = cur(8)
		photo = cur(9)
		dcre_produit = cur(10)
		dmod_produit = cur(11)
		cross_selling = toNull(cur(12))
		fourn_family = toNull(cur(13))
		fourn_reference = toNull(cur(14))
		tva_id = cur(15)
		poids = cur(16)
		page_sommaire = cur(17)
		set cur=nothing
		%>
		<div align="center">
		<form name="prod_modif" action="synth_card.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="id_produit" value="<%=id_produit%>">
		<input type="hidden" name="photo_ori" value="<%=photo%>">
		<table border=0 cellpadding=3>
		<tr>
			<td align="left">Référence CAS:</td>
			<td><%=ref_cas%></td>
		</tr><tr>
			<td align="left">Désignation:</td>
			<td><input type="text" name="titre" value="<%=titre%>" size=80>
		</tr><tr>
			<td align="left"><%=etoile%>Catégorie:</td>
			<td>
				<select name="id_categorie">
				<option value="">[Choisissez]
				<%
				SQL = "select id_categorie, cat_label from t_categorie order by 2"
				set cur = conn.execute(SQL)
				while not cur.eof
					%>	<option value="<%=cur(0)%>" <%=selected(cur(0),id_categorie)%> ><%=cur(1)%>	<%
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr><tr>
			<td align="left">Famille des<br>produits complexes</td>
			<td>
				<select name="fourn_reference">
				<option value="">[Choisissez]
				<%
				SQL = "select fourn_reference from t_reference_fourn order by 1"
				set cur = conn.execute(SQL)
				while not cur.eof
					%>	<option value="<%=server.HTMLencode(cur(0))%>" <%=selected(cur(0),fourn_reference)%> ><%=cur(0)%>	<%
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr><tr>
			<td align="left">Ref. JEP</td>
			<td><input type="text" maxlength=50 name="fourn_family" size=50 value="<%=fourn_family%>"></td>
		</tr><tr>
			<td align="left"><%=etoile%>Mots-clés (fr):</td>
			<td><textarea name="motscles" cols=40 rows=5><%=br2nl(motscles)%></textarea></td>
		</tr><tr>
			<td align="left"><%=etoile%>Mots-clés (gb):</td>
			<td><textarea name="motscles_gb" cols=40 rows=5><%=br2nl(motscles)%></textarea></td>
		</tr><tr>
			<td align="left"><%=etoile%>Descriptif (fr):</td>
			<td><textarea name="texte_principal" cols=40 rows=5><%=br2nl(texte_principal)%></textarea></td>
		</tr><tr>
			<td align="left"><%=etoile%>Descriptif (gb):</td>
			<td><textarea name="texte_principal_gb" cols=40 rows=5><%=br2nl(texte_principal_gb)%></textarea></td>
		</tr><tr>
			<td align="left"><%=etoile%>Photo:</td>
			<td>
				<% if len(toZS(photo))>0 then %>
				<img src="<%=photo%>" alt="<%=photo%>" border=0><br>
				<% else %>
				<i>Aucune Photo</i><br>
				<% end if %>
				<select name="photo">
				<%
				Set fs = CreateObject("Scripting.FileSystemObject")
				dapath = LOCAL_PATH & "\img\produit\small"
				Set f = fs.GetFolder(dapath)
			    Set fc = f.Files
			    For Each f1 in fc
					rpath = HTTP_PATH & "/img/produit/small/" & f1.name
					s = "<option value=""" & rpath & """ "
					if rpath = photo then
						s = s & "SELECTED "
					end if
					s = s & ">" & f1.name & CRLF
					response.write (s)
			    Next
				%>
				</select>
			</td>
		</tr><tr>
			<td align="left">Mise en ligne:</td>
			<td><select name="enligne" size=2>
				<option value="1" <%=selected(enligne,1)%>>Oui
				<option value="0" <%=selected(enligne,0)%>>Non
				</select>
			</td>
		</tr><tr>
			<td align="left"><%=etoile%>Prix en Euros:</td>
			<td><input type="text" maxlength=50 size=50 name="prix" value="<%=replace(cstr(prix),",",".")%>"></td>
		</tr><tr>
			<td align="left"><%=etoile%>Taux de TVA:</td>
			<td>
				<select name="tva_id">
				<%
				SQL = "select tva_id, convert(varchar,tva_valeur) from t_tva order by 1"
				set cur = conn.execute(SQL)
				s = ""
				while not cur.eof
					s = "<option value=""" & cur(0) & """" 
					if cur(0) = tva_id then
						s = s & " SELECTED "
					end if
					s = s & ">" & cur(1)
					response.write (s)
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr><tr>
			<td align="left"><%=etoile%>Poids:</td>
			<td><input type="text" maxlength=10 size=10 name="poids" value="<%=replace(cstr(poids),",",".")%>"></td>
		</tr><tr>
			<td align="left">Produit lié:</td>
			<td>
				<select name="cross_selling">
				<option value="NULL" <%=selected(cross_selling,"NULL")%>>[Aucun]
				<%
				SQL = "select id_produit, titre from t_produit where id_produit<>" & id_produit & "order by 2"
				set cur = conn.execute(SQL)
				while not cur.eof
					response.write ("<option value=""" & cur(0) & """ " & selected(cur(0),cross_selling) & ">" & cur(1) & CRLF)
					cur.movenext
				wend
				set cur=nothing
				%>
				</select>
			</td>
		</tr>
			<td colspan=2 align="left">
			Les champs précédés par <%=etoile%> sont obligatoires...
			</td>
			<td></td>
		</table>
		<% barre() %>		
		<input type="button" value="OK" onClick="modif_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.prod_modif);">
		</form>
		</div>
		<%
	end if

	if act="modif_ok" then
		id_produit = toNull(request("id_produit"))
		id_categorie = request("id_categorie")
		titre = request("titre")
		texte_principal = toNull(request("texte_principal"))
		texte_principal_gb = toNull(request("texte_principal_gb"))
		enligne = toNull(request("enligne"))
		page_sommaire = toNull(request("page_sommaire"))
		prix = toNull(request("prix"))
		motscles = toNull(request("motscles"))
		motscles_gb = toNull(request("motscles_gb"))
		cross_selling = toNull(request("cross_selling"))
		fourn_family = toNull(request("fourn_family"))
		fourn_reference = request("fourn_reference")
		photo = request("photo")
		photo_ori = request("photo_ori")
		tva_id = request("tva_id")
		poids = request("poids")
		'on s'occupe de l'image...
'		if photo<>photo_ori then
'			randomize
'			set fs = server.createobject("Scripting.FileSystemObject")
			' on s'occupe de l'ancienne photo
'			SQL1="select photo from t_produit where id_produit=" & id_produit
'			set cur=conn.execute (SQL1)
'			if not isNull(cur(0)) then
'				kfile = LOCAL_PATH & "\img\produit\small\" & Mid(cur(0), InstrRev(cur(0), "/") + 1)
'				if fs.fileexists(kfile) then
'					fs.DeleteFile kfile
'				end if
'			end if
'			set cur = nothing
'		end if
		SQL = "update t_produit set "
		SQL=SQL & "id_categorie=" & id_categorie & ", "
		SQL=SQL & "titre=" & noquote2(titre) & ", "
		SQL=SQL & "texte_principal=" & noquote2(nl2br(texte_principal)) & ", "
		SQL=SQL & "texte_principal_gb=" & noquote2(nl2br(texte_principal_gb)) & ", "
		SQL=SQL & "enligne=" & enligne & ", "
		SQL=SQL & "page_sommaire=" & page_sommaire & ", "
		SQL=SQL & "prix=" & prix & ", "
		SQL=SQL & "motscles=" & noquote2(motscles) & ", "
		SQL=SQL & "motscles_gb=" & noquote2(motscles_gb) & ", "
		SQL=SQL & "cross_selling=" & cross_selling & ", "
		SQL=SQL & "fourn_family=" & noquote2(fourn_family) & ", "
		SQL=SQL & "fourn_reference=" & noquote2(fourn_reference) & ", "
		SQL=SQL & "photo='" & noquote(photo) & "', "
		SQL=SQL & "dmod_produit='" & now & "', "
		SQL=SQL & "tva_id = " & tva_id & ", "
		SQL=SQL & "poids = " & poids & " "
		SQL=SQL & "where id_produit = " & id_produit
'		response.write (SQL)
		conn.execute(SQL)
		retour()
	end if

	if act="suppr" then
		id_produit = request("id_produit")
		SQL="select titre, ref_cas from t_produit where id_produit = " & id_produit
		set cur=conn.execute (SQL)
		titre = cur(0)
		ref_cas = cur(1)
		set cur=nothing
		erreurs = ""
		' on vérifie si le produit n'est pas référencé en cross-selling
		SQL = "select count(*) from t_produit where cross_selling=" & id_produit
		set cur=conn.execute (SQL)
		if cur(0)<>0 then
			erreurs = "<br>Il est utilisé en cross-selling par " & cur(0) & " produits"
		end if
		set cur=nothing
		' et s'il est référencé dans un caddie actuel ...
		hier = dateadd("d",-1,now)
		SQL = "select count(*) from t_caddie where id_produit = " & id_produit & " and date_modif >= '" & hier & "'"
		set cur=conn.execute(SQL)
		if cur(0)<>0 then
			erreurs = erreurs & "<br>Il est utilisé par " & cur(0) & " paniers virtuels (commandes en cours)... Veuillez réessayer de le supprimer plus tard"
		end if
		set cur=nothing

		if erreurs<> "" then
			erreurs = "Problème: Ce produit ne peut être supprimé pour les raisons suivantes:" & erreurs
			response.write (erreurs)
			response.write ("<p>Cliquez <a href=""javascript:history.back()"">ici</a> ou cliquez sur le bouton 'Back' votre navigateur pour revenir à la liste")
		else
			%>
			<div align="center">
			<form name="prod_suppr" action="synth_card.asp" method="post">
			<input type="hidden" name="id_produit" value="<%=id_produit%>">
			<input type="hidden" name="act" value="">
			<font class="titre1">Êtes-vous sûr de vouloir supprimer le produit "<%=titre%>" (ref "<%=ref_cas%>") ?<p>
			<input type="button" value="OK" onClick="suppr_ok();">&nbsp;
			<input type="button" value="Annuler" onClick="cancel(document.prod_suppr);">
			</form>
			</div>
			<%
		end if
	end if

	if act="suppr_ok" then
		id_produit = request("id_produit")

		set fs = server.createobject("Scripting.FileSystemObject")
		' on s'occupe de l'ancienne photo
		SQL1="select photo from t_produit where id_produit=" & id_produit
		set cur=conn.execute (SQL1)
		if not isNull(cur(0)) then
			kfile = LOCAL_PATH & "\img\produit\small\" & Mid(cur(0), InstrRev(cur(0), "/") + 1)
			if fs.fileexists(kfile) then
				fs.DeleteFile kfile
			end if
		end if
		set cur = nothing

		SQL = "delete from t_produit where id_produit = " & id_produit
		conn.execute(SQL)
		retour()
	end if



	if multipart = 1 then
		set upl=nothing
	end if
	conn.close()
	set conn = nothing
%>
</BODY>
</HTML>
