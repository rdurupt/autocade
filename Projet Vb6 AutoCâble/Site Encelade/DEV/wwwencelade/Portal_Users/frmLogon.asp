<% 
Response.Expires = 0 

if Session("IsEvaluation") = "true" then
	Username = "Admin"
	Password = "new"
end if
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->

<html>
<head>
<title><%= GetDefault("emp_AppTitle","Employee Manager") %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "ASPIntranet.asp?mode=emp_lst&menuEmployee=EmployeeList"
}

</script>
<link rel="stylesheet" href="StyleSheet.css">
<body onLoad="document.frm.Username.focus()">

<% '****** Header ***** %>
<table width="100%" border="2" bgcolor="<%= GetDefault("web_bgcolor","#678DB8") %>" cellpadding="1" cellspacing="0">
<tr>
<td>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
<tr>
<form name="frm" action="subLogon.asp">
<td align="left">
<font size="3" color="<%= GetDefault("web_color","white") %>"><b><%= GetDefault("emp_AppTitle","Employee Manager") %></b></font>
</td>
</tr>
</table>
</td>
</tr>
</table>

<center>

<br>
<table border="0" cellpadding="0" cellspacing="0">

<tr>
<td>&nbsp;</td>
<td><font color="red"><b><%= Request("msg") %></b></font></td>
</tr>

<tr>
<td>Username&nbsp;</td>
<td><input type="text" name="Username" value="<%= Username %>"></td>
</tr>

<tr>
<td>Password&nbsp;</td>
<td><input type="password" name="Password" value="<%= Password %>"></td>
</tr>

<tr>
<td>&nbsp;</td>
<td align="left">
<input type="submit" value="OK">
</td>
</tr>
</table>

</form>
</center>
</body>
</html>
<% Conn.Close %>