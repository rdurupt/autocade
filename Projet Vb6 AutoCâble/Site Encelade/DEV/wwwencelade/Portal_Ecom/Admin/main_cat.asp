<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>

<!--#include file="lib/admin-lib.asp"-->
<script language="JavaScript">
function ajout()
{
	var f=document.cat_sel;

	f.act.value="ajout";
	f.submit();
}

function ajout_ok()
{
	var f=document.cat_ajout;
	var err = "Les erreurs suivantes sont apparues:";
	var ret = true;

	if (f.cat_label.value.length < 1)
	{
		err += "\nLe libellé est vide";
		ret = false;
	}

	if (f.cat_ref.value.length < 1)
	{
		err += "\nLa référence est vide";
		ret = false;
	}

	if (f.cat_desc.value.length < 1)
	{
		err += "\nLa description est vide";
		ret = false;
	}
/*
	if (f.cat_photo.options.selectedIndex<1)
	{
		err += "\nVous n'avez pas sélectionné de photo";
		ret = false;
	}
*/
	if (ret)
	{
		f.act.value = "ajout_ok";
		f.submit();
	}
	else
	{
		alert (err);
	}
}

function modif()
{
	var f=document.cat_sel;

	if (f.id_categorie.selectedIndex > -1)
	{
		f.act.value="modif";
		f.submit();
	}
	else
	{
		alert ("Vous devez sélectionner une catégorie");
	}
}

function modif_ok()
{
	var f=document.cat_modif;
	var err = "Les erreurs suivantes sont apparues:";
	var ret = true;

	if (f.cat_label.value.length < 1)
	{
		err += "\nLe libellé est vide";
		ret = false;
	}

	if (f.cat_ref.value.length < 1)
	{
		err += "\nLa référence est vide";
		ret = false;
	}

	if (f.cat_desc.value.length < 1)
	{
		err += "\nLa description est vide";
		ret = false;
	}
/*
	if (f.cat_photo.options.selectedIndex<1)
	{
		err += "\nVous n'avez pas sélectionné de photo";
		ret = false;
	}
*/
	if (ret)
	{
		f.act.value = "modif_ok";
		f.submit();
	}
	else
	{
		alert (err);
	}


}

function suppr()
{
	var f=document.cat_sel;

	if (f.id_categorie.selectedIndex > -1)
	{
		f.act.value="suppr";
		f.submit();
	}
	else
	{
		alert ("Vous devez sélectionner une catégorie");
	}
}

function suppr_ok()
{
	var f=document.cat_suppr;

	f.act.value = "suppr_ok";
	f.submit();
}

function cancel(f)
{
	f.action = "main_cat.asp";
	f.act.value="aff";
	f.submit();
}

</script>
</HEAD>
<BODY BGCOLOR="#FFFFFF">
<font class="titre">Catégories Principales</font>
<%	barre() 

	
	set conn=server.createobject("adodb.connection")
	conn.open myDSN


	act = request("act")
	if act="" then
		act="aff"
	end if

	if act="ajout" then
	
	 response.write LOCAL_PATH &"<br>"
	 'response.end
		%>
		<form name="cat_ajout" action="main_cat.asp" method="post">
		<input type="hidden" name="act" value="">
		<div align="center">
		<table border=0 cellpadding=5>
		<tr>
			<td><%=etoile%>Désignation:</td>
			<td><input type="text" maxlength=250 size="100" name="cat_label"></td>
		</tr><tr>
			<td><%=etoile%>Ref. Mulitple:</td>
			<td><select name="ref_multiple" size=2>
				<option value="1">Oui
				<option value="0" selected>Non
				</select>
			</td>
		</tr>
		<tr>
			<td><%=etoile%>Référence:</td>
			<td><input type="text" name="cat_ref" size=100 maxlength=50></td>
		</tr>
		<tr>
			<td><%=etoile%>Description:</td>
			<td><textarea name="cat_desc" cols=50 rows=4></textarea></td>
		</tr>
		<tr>
			<td>Photo Générique:</td>
			<td><select name="cat_photo">
				 <option value="">Choisissez... 
				<%
				Set fs = CreateObject("Scripting.FileSystemObject")
				dapath = LOCAL_PATH & "\img\famille\generique"
				'response.write dapath&"<br>" 
				'response.end
				Set f = fs.GetFolder(dapath)
			    Set fc = f.Files
			    For Each f1 in fc
					response.write ("<option value=""" &  HTTP_PATH & "/img/famille/generique/" & f1.name & """>" & f1.name & CRLF )
			    Next
				%>
			</td>
		</tr>
		</table>
		<% barre() %>		
		<input type="button" value="OK" onClick="ajout_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.cat_ajout);">
		</div>
		</form>
		<%
	end if

	if act="ajout_ok" then
		cat_label = request("cat_label")
		ref_multiple = request("ref_multiple")
		cat_photo = request("cat_photo")
		cat_ref = request("cat_ref")
		cat_desc = request("cat_desc")
		SQL = "select max(id_categorie) from t_categorie"
		set cur=conn.execute(SQL)
		maxi = inc(cur(0))
		SQL = "insert into t_categorie(id_categorie, cat_label, ref_multiple, cat_photo, cat_ref, cat_desc) values (" & maxi & ", '" & noquote(cat_label) & "', " & ref_multiple & ", " & noquote2(cat_photo) & ", " & noquote2(cat_ref) & ", " & noquote2(nl2br(cat_desc)) & ")"
		conn.execute (SQL)
		act = "aff"
	end if

	if act="modif" then
		id_categorie = request("id_categorie")
		SQL = "select cat_label, ref_multiple, cat_photo, cat_ref, cat_desc from t_categorie where id_categorie = " & id_categorie
		set cur = conn.execute(SQL)
		cat_label = cur(0)
		ref_multiple = cur(1)
		cat_photo = toZS(cur(2))
		cat_ref = toZS(cur(3))
		cat_desc = br2nl(toZS(cur(4)))
		%>
		<form name="cat_modif" action="main_cat.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="id_categorie" value="<%=id_categorie%>">
		<div align="center">
		<table border=0 cellpadding=5>
		<tr>
			<td><%=etoile%>Désignation:</td>
			<td><input type="text" maxlength=250 size="80" name="cat_label" value="<%=cat_label%>"></td>
		</tr>
		<tr>
			<td><%=etoile%>Référence:</td>
			<td><input type="text" name="cat_ref" size=100 maxlength=50 value="<%=cat_ref%>"></td>
		</tr>
		<tr>
			<td><%=etoile%>Description:</td>
			<td><textarea name="cat_desc" cols=50 rows=4><%=cat_desc%></textarea></td>
		</tr>
		<tr>
			<td><%=etoile%>Ref. Mulitple:</td>
			<td><select name="ref_multiple" size=2>
				<option value="1" <%=selected(ref_multiple,True)%>>Oui
				<option value="0" <%=selected(ref_multiple,False)%>>Non
				</select>
			</td>
		</tr>
		<input type="hidden" name="photo_ori" value="<%=cat_photo%>">
		<tr>
			<td>Photo Générique:</td>
			<td><img src="<%=cat_photo%>" alt="<%=cat_photo%>" border=0><br>
				<font size=1><i>Choisissez une nouvelle photo ou laissez inchangé pour ne pas la modifier...</i></font><br>
				<select name="cat_photo">
				<option value="">Choisissez...
				<%
				Set fs = CreateObject("Scripting.FileSystemObject")
				dapath = LOCAL_PATH & "\img\famille\generique"
				Set f = fs.GetFolder(dapath)
			    Set fc = f.Files
			    For Each f1 in fc
					s = "<option value=""" &  HTTP_PATH & "/img/famille/generique" & f1.name & """"
					if f1.name = Mid(cat_photo, InstrRev(cat_photo, "/") + 1) then
						s = s & " SELECTED "
					end if
					s = s & ">" & f1.name & CRLF 
					response.write (s)
			    Next
				%>
			</td>
		</tr>
		</table>
		<% barre() %>		
		<input type="button" value="OK" onClick="modif_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.cat_modif);">
		</div>
		</form>
		<%
	end if

	if act="modif_ok" then
		id_categorie = request("id_categorie")
		cat_label = request("cat_label")
		ref_multiple = request("ref_multiple")
		photo_ori = request("photo_ori")
		cat_photo = request("cat_photo")
		cat_ref = request("cat_ref")
		cat_desc = request("cat_desc")
		if photo_ori <> cat_photo then		' update photo
'			SQL1 = "select count(*) from t_categorie where cat_photo='" & photo_ori & "'"
'			set cur = conn.execute(SQL1)
'			if cur(0) = 1 then		' il n'y a que ce produit concerné par l'ancienne photo, alors on la vire
'				vfile = LOCAL_PATH & "\img\famille\generique\" & Mid(photo_ori, InstrRev(photo_ori, "/") + 1)
'				Set fs = CreateObject("Scripting.FileSystemObject")
'				if fs.fileexists(vfile) then
'					fs.DeleteFile (vfile)
'				end if
'				set fs=nothing
'			end if
'			set cur = nothing
		end if
		SQL = "update t_categorie set cat_label = '" & noquote(cat_label) & "', ref_multiple = " & ref_multiple & ", cat_photo=" & noquote2(cat_photo) & ", cat_ref=" & noquote2(cat_ref) & ", cat_desc=" & noquote2(nl2br(cat_desc)) & " where id_categorie = " & id_categorie
		conn.execute (SQL)
		act = "aff"
	end if

	if act="suppr" then
		erreurs = ""
		id_categorie = request("id_categorie")
		SQL = "select cat_label from t_categorie where id_categorie = " & id_categorie
		set cur = conn.execute (SQL)
		cat_label = cur(0)
		set cur=nothing
		SQL = "select count(*) from t_reference_fourn where id_categorie = " & id_categorie
		set cur=conn.execute(SQL)
		if cur(0) <> 0 then
			erreurs = "- " & cur(0) & " produit(s) complexe(s)<br>"
		end if
		set cur=nothing

		SQL = "select count(*) from t_produit where id_categorie = " & id_categorie
		set cur=conn.execute(SQL)
		if cur(0) <> 0 then
			erreurs = erreurs & "- " & cur(0) & " produit(s)<br>"
		end if
		set cur=nothing

		if erreurs="" then		'pas de problème:
			%>
			<div align="center">
			<font class="titre1">Êtes-vous sûr de vouloir supprimer la catégorie <%=cat_label&" ?"%></font>
			<p>
			<form name="cat_suppr" action="main_cat.asp" method="post">
			<input type="hidden" name="id_categorie" value="<%=id_categorie%>">
			<input type="hidden" name="act" value="">
			<input type="button" value="OK" onClick="suppr_ok();">&nbsp;
			<input type="button" value="Annuler" onClick="cancel(document.cat_suppr);">
			</div>
			</form>
			<%
		else
			erreurs = "Vous ne pouvez pas effacer cette catégorie car elle est utilisée par:<br>" & erreurs
			response.write (erreurs)
			response.write ("<p>Cliquez <a href=""javascript:history.back()"">ici</a> ou cliquez sur le bouton 'Back' votre navigateur pour revenir à la liste")

		end if
	end if

	if act="suppr_ok" then
		id_categorie = request("id_categorie")
'		SQL = "select cat_photo from t_categorie where id_categorie = " & id_categorie
'		set cur = conn.execute(SQL)
'		cat_photo = cur(0)
'		set cur=nothing
'		SQL = "select count(*) from t_categorie where cat_photo = " & noquote2(cat_photo)
'		set cur=conn.execute(SQL)
'		cpt = cur(0)
'		set cur = nothing
'		if cpt=1 then		' delete photo
'			vfile = LOCAL_PATH & "\img\famille\generique\" & Mid(cat_photo, InstrRev(cat_photo, "/") + 1)
'			Set fs = CreateObject("Scripting.FileSystemObject")
'			if fs.fileexists(vfile) then
'				fs.DeleteFile (vfile)
'			end if
'			set fs=nothing
'		end if
		SQL = "delete from t_categorie where id_categorie = " & id_categorie
		conn.execute (SQL)
		act = "aff"
	end if


	if act="aff" then
		%>	
		<form name="cat_sel" action="main_cat.asp" method="post">
		<input type="hidden" name="act" value="">
		<div align="center">
		<font class="titre1">Sélectionnez une catégorie:</font><br><br>
		<select name="id_categorie" size="5">	
		<%
		SQL = "select id_categorie, cat_label from t_categorie order by 2"
		set cur=conn.execute(SQL)

		while not cur.eof
			%>	<option value="<%=cur(0)%>"><%=cur(1)%>	<%
			cur.movenext
		wend
		set cur=nothing
		%>
		</select>
		<% barre() %>
		<input type="button" value="Nouvelle..." onClick="ajout();">&nbsp;
		<input type="button" value="Modifier" onClick="modif();">&nbsp;
		<input type="button" value="Supprimer" onClick="suppr();">
		</div>
		</form>
		<%
	end if

conn.close
set conn=nothing
%>
</BODY>
</HTML>
