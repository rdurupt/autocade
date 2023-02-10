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
function ajout()
{
	var f=document.adm_sel;

	f.act.value="ajout";
	f.submit();
}

function ajout_ok()
{
	var f=document.adm_ajout;
	var ret=true;
	var err="";

	if (f.login.value.length == 0)
	{
		err = "Vous devez entrer un login\n";
		ret = false;
	}
	if (f.password.value.length == 0)
	{
		err += "Vous devez entrer un mot de passe\n";
		ret = false;
	}
	if (f.password.value != f.password_conf.value)
	{
		err += "Mot de passe incorrect";
		f.password.value = "";
		f.password_conf.value = "";
		ret = false;
	}

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
	var f=document.adm_sel;

	if (f.id_adm_user.selectedIndex > -1)
	{
		f.act.value="modif";
		f.submit();
	}
	else
	{
		alert ("Vous devez sélectionner un Login");
	}

}

function modif_ok()
{
	var f=document.adm_modif;
	var ret=true;
	var err="";

	if (f.login.value.length == 0)
	{
		err = "Vous devez entrer un login\n";
		ret = false;
	}
	if (f.password.value.length == 0)
	{
		err += "Vous devez entrer un mot de passe\n";
		ret = false;
	}
	if (f.password.value != f.password_conf.value)
	{
		err += "Mot de passe incorrect";
		f.password.value = "";
		f.password_conf.value = "";
		ret = false;
	}

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
	var f=document.adm_sel;

	if (f.id_adm_user.selectedIndex > -1)
	{
		f.act.value="suppr";
		f.submit();
	}
	else
	{
		alert ("Vous devez sélectionner un Login");
	}
}

function suppr_ok()
{
	var f=document.adm_suppr;

	f.act.value = "suppr_ok";
	f.submit();
}

function cancel(f)
{
	f.action = "users.asp";
	f.act.value="aff";
	f.submit();
}

</script>
</HEAD>
<BODY BGCOLOR="#FFFFFF">
<font class="titre">Administrateurs</font>
<%	barre() 

	
	set conn=server.createobject("adodb.connection")
	conn.open myDSN


	act = request("act")
	if act="" then
		act="aff"
	end if

	if act="ajout" then
		%>
		<form name="adm_ajout" action="users.asp" method="post">
		<input type="hidden" name="act" value="">
		<div align="center">
		<table border=0 cellpadding=5>
		<tr>
			<td>Login</td>
			<td><input type="text" maxlength=20 size="20" name="login"></td>
		</tr><tr>
			<td>Password</td>
			<td><input type="password" name="password" maxlength=16 size=20></td>
		</tr><tr>
			<td>Confirmez</td>
			<td><input type="password" name="password_conf" maxlength=16 size=20></td>
		</tr>	
		</table>
		<% barre() %>		
		<input type="button" value="OK" onClick="ajout_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.adm_ajout);">
		</div>
		</form>
		<%
	end if

	if act="ajout_ok" then
		login = request("login")
		password = request("password")
		SQL = "select max(id_adm_user) from t_adm_users"
		set cur=conn.execute(SQL)
		maxi = inc(cur(0))
		
		SQL = "insert into t_adm_users(id_adm_user, login, password) values (" & maxi & ", '" & noquote(login) & "', '" & noquote(password) & "')"
		conn.execute (SQL)
		act = "aff"
	end if

	if act="modif" then
		id_adm_user = request("id_adm_user")
		SQL = "select login, password from t_adm_users where id_adm_user = " & id_adm_user
		set cur = conn.execute(SQL)
		login = cur(0)
		password = cur(1)
		%>
		<form name="adm_modif" action="users.asp" method="post">
		<input type="hidden" name="act" value="">
		<input type="hidden" name="id_adm_user" value="<%=id_adm_user%>">
		<div align="center">
		<table border=0 cellpadding=5>
		<tr>
			<td>Login</td>
			<td><input type="text" maxlength=20 size="20" name="login" value="<%=login%>"></td>
		</tr><tr>
			<td>Password</td>
			<td><input type="password" maxlength=16 size="16" name="password" value="<%=password%>"></td>
		</tr><tr>
			<td>Confirmez</td>
			<td><input type="password" maxlength=16 size="16" name="password_conf" value="<%=password%>"></td>
		</tr>

		</table>
		<% barre() %>		
		<input type="button" value="OK" onClick="modif_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.adm_modif);">
		</div>
		</form>
		<%
	end if

	if act="modif_ok" then
		id_adm_user = request("id_adm_user")
		login = request("login")
		password = request("password")
		SQL = "update t_adm_users set login = '" & noquote(login) & "', "
		SQL = SQL &	"password = '" & noquote(password) & "' where id_adm_user = " & id_adm_user
		conn.execute (SQL)
		act = "aff"
	end if

	if act="suppr" then
		id_adm_user = request("id_adm_user")
		SQL = "select login from t_adm_users where id_adm_user = " & id_adm_user
		set cur = conn.execute (SQL)
		login = cur(0)
		%>
		<div align="center">
		<font class="titre1">Êtes-vous sûr de vouloir supprimer l'administrateur <%=login&" ?"%></font>
		<p>
		<form name="adm_suppr" action="users.asp" method="post">
		<input type="hidden" name="id_adm_user" value="<%=id_adm_user%>">
		<input type="hidden" name="act" value="">
		<input type="button" value="OK" onClick="suppr_ok();">&nbsp;
		<input type="button" value="Annuler" onClick="cancel(document.adm_suppr);">
		</div>
		</form>
		<%
	end if

	if act="suppr_ok" then
		id_adm_user = request("id_adm_user")
		SQL = "delete from t_adm_users where id_adm_user = " & id_adm_user
		conn.execute (SQL)
		act = "aff"
	end if


	if act="aff" then
		%>
		<form name="adm_sel" action="users.asp" method="post">
		<input type="hidden" name="act" value="">
		<div align="center">
		<font class="titre1">Sélectionnez un Login:</font><br><br>
		<select name="id_adm_user" size="5">	
		<%
		SQL = "select id_adm_user, login from t_adm_users order by 2"
		set cur=conn.execute(SQL)

		while not cur.eof
			%>	<option value="<%=cur(0)%>"><%=cur(1)%>	<%
			cur.movenext
		wend
		%>
		</select>
		<% barre() %>
		<input type="button" value="Nouveau..." onClick="ajout();">&nbsp;
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
