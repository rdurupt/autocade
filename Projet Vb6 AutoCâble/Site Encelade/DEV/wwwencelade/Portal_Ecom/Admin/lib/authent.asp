<% 

function nquote(s0)
	s = ""
	c = ""

	for i = 1 to len(s0)
		c=mid(s0,i,1)
		if c = "'" then
			s = s & "''"
		else
			s = s & c
		end if
	next

	nquote = s
end function



	if request.cookies("CAS_LOGIN") <> "" then
		login = request.cookies("CAS_LOGIN")
		passwd = request.cookies("CAS_PASSWD")
	else
		login = request("login")
		passwd = request("passwd")
	end if


	cpt = request("cpt")
	if cpt="" then
		cpt=0
	else
		if cpt>3 then
			response.redirect "http://www.waltdisney.com"			
		else
			cpt = cpt + 1
		end if
	end if


	if login<>"" then
		set conn=server.createobject("adodb.connection")
		myDSN="DSN=casaviation;uid=casaviation;pwd=gyiodkbm"
		conn.open myDSN
		
		SQL = "select login, password from t_adm_users where login = '" & nquote(login) & "' and password = '" & nquote(passwd) & "'"
		set cur = conn.execute (SQL)

		if cur.eof then
			login = ""
			passwd = ""
			response.cookies("CAS_LOGIN").Expires = "31 Juillet 1992"
			response.cookies("CAS_PASSWD").Expires = "31 Juillet 1992"
		else
			response.cookies("CAS_LOGIN") = login & ":" & passwd
		end if

		conn.close
		set conn = nothing
	end if
	
	if login = "" then
		%>
		<html>
		<head>
		<title>CAS Aviation - Administration</title>
		</head>
		<body>
		<form name="auth" action="index.asp" method="post">
		<input type="hidden" name="cpt" value="<%=cpt+1%>">
		<table border=0 cellpadding=5>
		<tr>
			<td>Login:</td>
			<td><input type="text" name="login" maxlength="20" value="<%=login%>">
		</tr><tr>
			<td>Password:</td>
			<td><input type="password" name="passwd" maxlength="16">
		</tr>
		<br>
		</table>
		<input type="submit">
		</form>
		</body>
		</html>
		<%
		response.end
	end if
%>	
