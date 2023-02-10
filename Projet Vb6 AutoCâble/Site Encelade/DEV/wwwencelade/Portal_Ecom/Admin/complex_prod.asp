<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
<!--#include file="lib/admin-lib.asp"-->
<script language="JavaScript">
function ajout()
{
	var f=document.fourn_sel;

	f.act.value="ajout";
	f.submit();
}

function ajout_ok()
{
	var f=document.fourn_ajout;
	var err = "Les erreurs suivantes sont apparues:"
	var ret = true;

	if (f.fourn_reference.value.length < 1)
	{
		err += "\nVous devez entrer une référence";
		ret = false;
	}

	if (f.fourn_design.value.length < 1)
	{
		err += "\nVous devez entrer une désignation";
		ret = false;
	}

	if (f.fourn_desc.value.length < 1)
	{
		err += "\nVous devez entrer une description";
		ret = false;
	}
/*
	if (f.fourn_photo.selectedIndex < 1)
	{
		err += "\nVous devez entrer photo";
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
	var f=document.fourn_sel;

	if (f.fourn_reference.selectedIndex > -1)
	{
		f.act.value="modif";
		f.submit();
	}
	else
	{
		alert ("Vous devez sélectionner une référence");
	}
}

function modif_ok()
{
	var f=document.fourn_modif;

	var err = "Les erreurs suivantes sont apparues:"
	var ret = true;

	if (f.fourn_reference.value.length < 1)
	{
		err += "\nVous devez entrer une référence";
		ret = false;
	}

	if (f.fourn_design.value.length < 1)
	{
		err += "\nVous devez entrer une désignation";
		ret = false;
	}

	if (f.fourn_desc.value.length < 1)
	{
		err += "\nVous devez entrer une description";
		ret = false;
	}
/*
	if (f.fourn_photo.options.selectedIndex < 1)
	{
		err += "\nVous devez entrer photo";
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
	var f=document.fourn_sel;

	if (f.fourn_reference.selectedIndex > -1)
	{
		f.act.value="suppr";
		f.submit();
	}
	else
	{
		alert ("Vous devez sélectionner une référence");
	}
}

function suppr_ok()
{
	var f=document.fourn_suppr;

	f.act.value = "suppr_ok";
	f.submit();
}

function cancel(f)
{
	f.action = "complex_prod.asp";
	f.act.value="aff";
	f.submit();
}

</script>

</HEAD>

<BODY BGCOLOR="#FFFFFF">
<font class="titre">Produits Complexes</font>
<%	barre() 

	
	set conn=server.createobject("adodb.connection")
	conn.open myDSN

	act = request("act")
	if act="" then
		act="aff"
	end if

	if act="ajout" then
	
	LOCAL_PATH = "E:\webprod\wwwflywayonline\cas-aviation\public_html"
	response.write "[LOCAL_PATH =" & LOCAL_PATH & "]<br>[HTTP_PATH =" & HTTP_PATH & "]" 

		%>
		<div align="center">
		<form name="fourn_ajout" action="complex_prod.asp" method="post">
		<input type="hidden" name="act" value="">
		<table border=0 cellpadding=5>
		<tr>
			<td><%=etoile%>Référence</td>
			<td><input type="text" maxlength=20 size=20 name="fourn_reference">
		</tr><tr>
			<td>Catégorie</td>
			<td>
				<select name="id_categorie">
				<option value="NULL" SELECTED>[Aucune]
				<%
				SQL = "select id_categorie, cat_label from t_categorie where ref_multiple = 1 order by 2"
				set cur = conn.execute (SQL)
				while not cur.eof
					response.write ("<option value=""" & cur(0) & """>" & cur(1))
					cur.movenext
				wend
				%>
				</select>
			</td>
		</tr><tr>
			<td><%=etoile%>Désignation</td>
			<td><input type="text" name="fourn_design" size=250 maxlength=250></td>
		</tr><tr>
			<td><%=etoile%>Description</td>
			<td><textarea name="fourn_desc" cols=50 rows=4></textarea></td>
		</tr>
 		<tr>
			<td>Photo</td>
			<td><select name="fourn_photo">
				<option value="">Choisissez...
				<%
				Set fs = CreateObject("Scripting.FileSystemObject")
				dapath = LOCAL_PATH & "\img\famille\complexe"
				Set f = fs.GetFolder(dapath)
			    Set fc = f.Files
			    For Each f1 in fc
					response.write ("<option value=""" &  HTTP_PATH & "/img/famille/complexe/" & f1.name & """>" & f1.name & CRLF )
			    Next
				%>
				</select>
			</td>
		</tr>
		</table>
		<% barre() %>		
		<input type="button" value="OK" onClick="ajout_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.fourn_ajout);">
		</form>
		</div>
		<%
	end if

	if act="ajout_ok" then
		fourn_reference = request("fourn_reference")
		id_categorie = request("id_categorie")
		fourn_photo = request("fourn_photo")
		fourn_design = request("fourn_design")
		fourn_desc = request("fourn_desc")

		SQL = "select fourn_reference from t_reference_fourn where fourn_reference='" & noquote(fourn_reference) & "'"
		set cur = conn.execute(SQL)
		if not cur.eof then
			response.write ("La référence '" & fourn_reference & "' existe déjà! vous devez en saisir une autre...<br><br>")
			response.write ("Cliquez <a href=""javascript:history.back()"">ici</a> ou cliquez sur le bouton 'Back' votre navigateur pour saisir une nouvelle référence")
		else
'			response.write ("lfile = " & lfile & "<br>")
			SQL =       "insert into t_reference_fourn (fourn_reference, id_categorie, fourn_photo, fourn_design, fourn_desc) values ("
			SQL = SQL & "'" & noquote(fourn_reference) & "', "
			SQL = SQL & id_categorie & ","
			SQL = SQL & noquote2(fourn_photo) & ", "
			SQL = SQL & noquote2(fourn_design) & ", "
			SQL = SQL & noquote2(nl2br(fourn_desc)) & ")"
'			response.write(SQL)
			conn.execute (SQL)
			act = "aff"
		end if

	end if

	if act="modif" then
		fourn_reference = request("fourn_reference")
		SQL = "select id_categorie, fourn_photo, fourn_design, fourn_desc from t_reference_fourn where fourn_reference='" & noquote(fourn_reference) & "'"
		cur = conn.execute (SQL)
		id_categorie = cur(0)
		fourn_photo = cur(1)
		fourn_design = toZS(cur(2))
		fourn_desc = cur(3)
		fourn_desc = br2nl(toZS(fourn_desc))

		%>
		<div align="center">
		<form name="fourn_modif" action="complex_prod.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="fourn_reference" value="<%=server.HTMLencode(fourn_reference)%>">
		<input type="hidden" name="photo_ori" value="<%=fourn_photo%>">
		<table border=0 cellpadding=5>
		<tr>
			<td>Référence</td>
			<td><%=fourn_reference%></td>
		</tr><tr>
			<td>Catégorie</td>
			<td>
				<select name="id_categorie">
				<option value="NULL" <%=selected(id_categorie,"")%> >[Aucune]
				<%
				SQL = "select id_categorie, cat_label from t_categorie where ref_multiple = 1 order by 2"
				set cur = conn.execute (SQL)
				while not cur.eof
					response.write ("<option value=""" & cur(0) & """ " & selected(cur(0),id_categorie) & ">" & cur(1))
					cur.movenext
				wend
				%>
				</select>
			</td>
		</tr><tr>
			<td><%=etoile%>Désignation</td>
			<td><input type="text" name="fourn_design" size=250 maxlength=250 value="<%=fourn_design%>"></td>
		</tr><tr>
			<td><%=etoile%>Description</td>
			<td><textarea name="fourn_desc" cols=50 rows=4><%=fourn_desc%></textarea></td>
		</tr>
 		<tr>
			<td>Photo</td>
			<td><img src="<%=fourn_photo%>" border=0 alt="<%=fourn_photo%>"><br>
			<select name="fourn_photo">
				<%
				Set fs = CreateObject("Scripting.FileSystemObject")
				dapath = LOCAL_PATH & "\img\famille\complexe\"
				Set f = fs.GetFolder(dapath)
			    Set fc = f.Files
			    For Each f1 in fc
					rpath = HTTP_PATH & "/img/famille/complexe/" & f1.name
					s = "<option value=""" & rpath & """ "
					if rpath = fourn_photo then
						s = s & "SELECTED "
				end if
					s = s & ">" & f1.name & CRLF
					response.write (s)
			    Next
				%>
				</select>
			</td>
		</tr>
		</table>
		<% barre() %>		
		<input type="button" value="OK" onClick="modif_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.fourn_modif);">
		</form>
		</div>
		<%	
	end if

	if act="modif_ok" then
		fourn_reference = request("fourn_reference")
		id_categorie = request("id_categorie")
		photo_ori = request("photo_ori")
		fourn_photo = request("fourn_photo")
		fourn_design = request("fourn_design")
		fourn_desc = request("fourn_desc")

		SQL = "update t_reference_fourn "
		SQL = SQL & "set id_categorie = " & id_categorie
		SQL = SQL & ", fourn_design=" & noquote2(fourn_design)
		SQL = SQL & ", fourn_desc=" & noquote2(nl2br(fourn_desc))
		if photo_ori <> fourn_photo then		' update photo
			SQL1 = "select count(*) from t_reference_fourn where fourn_photo='" & photo_ori & "'"
			set cur = conn.execute(SQL1)
'			if cur(0) = 1 then		' il n'y a que ce produit concerné par l'ancienne photo, alors on la vire
'				vfile = LOCAL_PATH & "\img\famille\complexe\" & Mid(photo_ori, InstrRev(photo_ori, "/") + 1)
'				Set fs = CreateObject("Scripting.FileSystemObject")
'				if fs.fileexists(vfile) then
'					fs.DeleteFile (vfile)
'				end if
'				set fs=nothing
'			end if
			set cur = nothing
			SQL = SQL & ", fourn_photo = " & noquote2(fourn_photo) & " "
		end if
		SQL = SQL & " where fourn_reference = '" & noquote(fourn_reference) & "'"
		conn.execute (SQL)
		act = "aff"
	end if

	if act="suppr" then
		erreurs = ""
		fourn_reference = request("fourn_reference")
		SQL = "select count(*) from t_produit where fourn_reference = '" & noquote(fourn_reference) & "'"
		set cur = conn.execute(SQL)
		if cur(0) <> 0 then	' cette ref est utilisée par 1 ou des produits
			response.write ("Vous ne pouvez pas supprimer cette référence car " & cur(0) & " produit(s) l'utilisent")
			response.write ("<p>Cliquez <a href=""javascript:history.back()"">ici</a> ou cliquez sur le bouton 'Back' votre navigateur pour revenir à la liste")
		else
			%>
			<div align="center">
			<font class="titre1">Êtes-vous sûr de vouloir supprimer la catégorie <%=fourn_reference%> ?</font>
			<p>
			<form name="fourn_suppr" action="complex_prod.asp" method="post">
			<input type="hidden" name="fourn_reference" value="<%=server.HTMLencode(fourn_reference)%>">
			<input type="hidden" name="act" value="">
			<input type="button" value="OK" onClick="suppr_ok();">&nbsp;
			<input type="button" value="Annuler" onClick="cancel(document.fourn_suppr);">
			</form>
			</div>
			<%
		end if
		set cur = nothing
	end if

	if act="suppr_ok" then
		fourn_reference=request("fourn_reference")
'		SQL = "select fourn_photo from t_reference_fourn where fourn_reference = '" & noquote(fourn_reference) & "'"
'		set cur = conn.execute (SQL)
'		fourn_photo = cur(0)
'		if fourn_photo <> "" then	' ce produit a une photo
'			SQL1 = "select count(*) from t_reference_fourn where fourn_photo='" & fourn_photo & "'"
'			set cur = conn.execute(SQL1)
'			if cur(0) = 1 then		' il n'y a que ce produit concerné par l'ancienne photo, alors on la vire
'				vfile = LOCAL_PATH & "\img\famille\complexe\" & Mid(fourn_photo, InstrRev(fourn_photo, "/") + 1)
'				Set fs = CreateObject("Scripting.FileSystemObject")
'				if fs.fileexists(vfile) then
'					fs.DeleteFile (vfile)
'				end if
'				set fs=nothing
'			end if
'			set cur = nothing
'		end if
		SQL = "delete from t_reference_fourn where fourn_reference = '" & noquote(fourn_reference) & "'"
		conn.execute (SQL)
		act = "aff"
	end if




	if act="aff" then
		%>
		<div align="center">
		<form name="fourn_sel" action="complex_prod.asp" method="post">
		<input type="hidden" name="act" value="">
		<font class="titre1">Sélectionnez un produit:</font>
		<p>
		<select name="fourn_reference" size=5>
		<%
		SQL = "select fourn_reference from t_reference_fourn order by 1"
		set cur = conn.execute (SQL)
		while not cur.eof
			s0 = server.HTMLencode(cur(0))
			response.write ("<option value=""" & s0 & """>" & s0)
			cur.movenext
		wend
		%>
		</select>
		<% barre() %>
		<input type="button" value="Nouvelle..." onClick="ajout();">&nbsp;
		<input type="button" value="Modifier" onClick="modif();">&nbsp;
		<input type="button" value="Supprimer" onClick="suppr();">
		</form>
		</div>
		<%
	end if

	conn.close()
	set conn = nothing
%>
</BODY>
</HTML>
