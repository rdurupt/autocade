<% 
Response.Expires = 0 
Session("CurrentPage") = "frmPassword.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<!--#INCLUDE FILE="inc_Menus.asp"-->
<%
If Request("mode") = "execute" then
	Set RS1 = Conn.Execute("SELECT * FROM Users WHERE Password = '" & safeEntry(Request("Password1")) & "' AND UserID = " & Session("web_UserID"))
	If Not RS1.EOF Then
		if Session("IsEvaluation") <> "true" then
		    Conn.Execute("UPDATE Users SET Password = '" & safeEntry(Request("Password2")) & "' WHERE UserID = " & Session("web_UserID"))
		    msg = "Password changed successfully."
		end if
	Else
	    msg = "Unable to change password."
	End If
end if
%>
<html>
<head>
<title><%= GetDefault("web_AppTitle","ASP Intranet Suite") %></title>
</head>
<script language="javascript">
function javSubmit() {
	if (document.frm.Password2.value !== document.frm.Password3.value) {
		alert("New password and confirmation do not match.")
		return
	} 
	document.frm.submit()
}

</script>
<link rel="stylesheet" href="StyleSheet.css">
<body onLoad="document.frm.Password1.focus()">

<% Call GetMenu() %>

<% '****** Header ***** %>
<table width="100%" border="2" bgcolor="<%= GetUserDefault(Session("web_UserID"),"emp_bgcolor","#678DB8") %>" cellpadding="1" cellspacing="0">
<tr>
<td>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
<tr>
<form name="frm" action="frmPassword.asp">
<input type="hidden" name="mode" value="execute">

<td align="left">
<font size="3" color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy") %>"><b>Change Password</b></font>
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
<td><font color="red"><b><%= msg %></b></font></td>
</tr>

<tr>
<td align="right">Old Password&nbsp;</td>
<td align="left"><input type="password" name="Password1" value=""></td>
</tr>

<tr>
<td align="right">New Password&nbsp;</td>
<td align="left"><input type="password" name="Password2" value=""></td>
</tr>

<tr>
<td align="right">Confirm New Password&nbsp;</td>
<td align="left"><input type="password" name="Password3" value=""></td>
</tr>

<tr>
<td>&nbsp;</td>
<td align="left">
<input type="button" value="Submit" onClick="javSubmit()">
</td>
</tr>
</table>

</form>
</center>
</body>
</html>
<% Conn.Close %>