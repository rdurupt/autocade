<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus 1.2">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<!--#include file="lib/admin-lib.asp"-->
</HEAD>

<BODY BGCOLOR="#FFFFFF">
<font class="titre">Purge du panier virtuel</font>
<%
	barre()

	set conn=server.createobject("adodb.connection")
	conn.open myDSN

	act = request("act")
	if act = "" then
		act="aff"
	end if

	if act="purge" then
		jours = cint(request("jours"))
		if jours>-30 then
			ladate = dateadd("d",cint(jours),now)
		else
			ladate = dateadd("m",-1,now)
		end if
		cpt = 0
		SQL = "select count(*) from t_caddie where date_modif < '" & ladate & "'"
		set cur=conn.execute(SQL)
		cpt = cur(0)
		if cpt=0 then
			cpt = "Aucun"
		end if
		set cur = nothing
		SQL = "delete from t_caddie where date_modif < '" & ladate & "'"
		conn.execute(SQL)
		response.write ("<font color=""#FF0000""><center>" & cpt & " enregistrement(s) supprimé(s)!</font><p>")
		act="aff"
	end if


	if act="aff" then
		%>
		<div align="center">
		<form name="caddie" action="caddie.asp" method="post">
		<input type="hidden" name="act" value="purge">
		Purger les paniers virtuels vieux de plus de&nbsp;
		<select name="jours">
		<option value="-1" SELECTED >1 jour
		<option value="-2">2 jours
		<option value="-7">7 jours
		<option value="-30">1 mois
		</select>
		<% barre() %>
		<input type="submit" value="Purger">
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
