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
function modif_ok()
{
	var f=document.prod_modif;
	var err = "";
	var ret = true;

	if (f.texte_detail.value.length < 1)
	{
		err += "\nTexte détaillé (fr) vide";
		ret = false;
	}

	if (f.texte_detail_gb.value.length < 1)
	{
		err += "\nTexte détaillé (gb) vide";
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

	barre()

	if (act="modif") and (cstr(id_produit)<>"") then 
		%>
		<div align="center">
		<font class="titre2">[&nbsp;<a href="synth_card.asp?act=modif&id_produit=<%=id_produit%>" target="droite">Fiche de Syntèse</a>&nbsp;]
										&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
										[&nbsp;<font color="#FF0000">Fiche Expert</font>&nbsp;]
		</font>
		</div>
		<%
	end if

	
	barre()



	if act="modif" then
		id_produit = request("id_produit")
		SQL = "select titre, ref_cas, texte_detail, texte_detail_gb, photo_expert from t_produit where id_produit=" & id_produit
		set cur=conn.execute(SQL)
		titre = cur(0)
		ref_cas = cur(1)
		texte_detail = cur(2)
		texte_detail_gb = cur(3)
		photo_expert = cur(4)
		set cur = nothing
		%>
		<div align="center">
		<form name="prod_modif" action="exp_card.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="id_produit" value="<%=id_produit%>">
		<input type="hidden" name="photo_expert_ori" value="<%=photo_expert%>">
		<table border=0 cellpadding=3>
		<tr>
			<td align="right">Référence CAS:</td>
			<td><%=ref_cas%></td>
		</tr><tr>
			<td align="right">Titre:</td>
			<td><%=titre%></td>
		</tr><tr>
			<td align="right"><%=etoile%>Texte détaillé (fr):</td>
			<td><textarea name="texte_detail" cols=40 rows=5><%=br2nl(texte_detail)%></textarea></td>
		</tr><tr>
			<td align="right"><%=etoile%>Texte détaillé (gb):</td>
			<td><textarea name="texte_detail_gb" cols=40 rows=5><%=br2nl(texte_detail_gb)%></textarea></td>
		</tr><tr>
			<td align="right" valign="top"><%=etoile%>Photo Détaillée:</td>
			<td><img src="<%=photo_expert%>" alt="<%=photo_expert%>" border=0><br>
				<select name="photo_expert">
				<%
				Set fs = CreateObject("Scripting.FileSystemObject")
				dapath = LOCAL_PATH & "\img\produit\large"
				Set f = fs.GetFolder(dapath)
			    Set fc = f.Files
			    For Each f1 in fc
					rpath = HTTP_PATH & "/img/produit/large/" & f1.name
					s = "<option value=""" & rpath & """ "
					if rpath = photo_expert then
						s = s & "SELECTED "
					end if
					s = s & ">" & f1.name & CRLF
					response.write (s)
			    Next
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
		id_produit = request("id_produit")
		texte_detail = request("texte_detail")
		texte_detail_gb = request("texte_detail_gb")
		photo_expert = request("photo_expert")
		photo_expert_ori = request("photo_expert_ori")

		'on s'occupe de l'image...
'		if photo_expert<>photo_expert_ori then
'			set fs = server.createobject("Scripting.FileSystemObject")
			' on s'occupe de l'ancienne photo
'			kfile = LOCAL_PATH & "\img\produit\large\" & Mid(photo_expert_ori, InstrRev(photo_expert_ori, "/") + 1)
'			if fs.fileexists(kfile) then
'				fs.DeleteFile kfile
'			end if
'		end if

		SQL = "update t_produit set "
		SQL=SQL & "texte_detail=" & noquote2(nl2br(texte_detail)) & ", "
		SQL=SQL & "texte_detail_gb=" & noquote2(nl2br(texte_detail_gb)) & ", "
		SQL=SQL & "photo_expert = " & noquote2(photo_expert) & ", "
		SQL=SQL & "dmod_produit = " & noquote2(now) & " "
		SQL=SQL & "where id_produit = " & id_produit

		conn.execute (SQL)
		retour()
	end if


	conn.close()
	set conn = nothing
%>
</BODY>
</HTML>
