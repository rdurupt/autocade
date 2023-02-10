<% 
Response.Expires = 0 
Session("CurrentPage") = "emp_frmSetField.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<!--#INCLUDE FILE="inc_Menus.asp"-->
<%
if Request("mode") = "execute" then
	Set RS1 = Conn.Execute("SELECT FieldID FROM dbp_DirDefFields")
	do while not RS1.EOF 
		key = "key" & RS1("FieldID")
		strSQL = "UPDATE dbp_DirDefFields SET "
		strSQL = strSQL & "FieldAlias = '" & safeEntry(Request(key)) & "' "
		strSQL = strSQL & "WHERE FieldID = " & RS1("FieldID")
		Conn.Execute(strSQL)
		if len(Request(key)) = 0 then
			Conn.Execute("DELETE FROM dbp_DirFields WHERE FieldID = " & RS1("FieldID"))
		end if
		
		RS1.movenext
	loop
	msg = "Fields updated."
end if
%>
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
<body background="../Portal_Html/Images/Background.asp">

<% Call GetMenu() %>

<% '****** Header ***** %>
<table width="100%" border="2" bgcolor="<%= GetUserDefault(Session("web_UserID"),"emp_bgcolor","#678DB8") %>" cellpadding="0" cellspacing="0">
<tr>
<td>
<table width="100%" border="0">
<tr>
<form name="frm" action="emp_frmSetField.asp">
<input type="hidden" name="mode" value="execute">
<td align="left">
<font size="3" color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>"><b>Customize Fields</b></font>
</td>
<td align="right">
<input type="submit" value="Submit">&nbsp;
</td>

</tr>
</table>
</td>
</tr>
</table>

<center><font color="red"><b><%= msg %></b></font></center>

<br>
&nbsp;&nbsp;<b>Field definitions:</b><br>
&nbsp;&nbsp;To use customizable fields (fld1, fld2...fld10), just add an alias.
&nbsp;&nbsp;To disable customizable fields, set the alias to blank.
<br><br>
<center>

<TABLE border="0" cellpadding="1" cellspacing="4" width="95%">
<TR valign="top">
<TD>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td colspan="2" bgcolor="silver"><b>PreDefined Fields</b></td>
</tr>

<tr>
<td><b>Field</b></td>
<td><b>Alias</b></td>
</tr>
<% 
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldName ORDER BY FieldID") 
do while not RS0.EOF
	if left(RS0("FieldName"),3) <> "fld" then 
		pr ("<tr>")
		pr ("<td><font size='1'>" & RS0("FieldName") & "</font></td>")
		pr ("<td><input type='text' name='key" & RS0("FieldID") & "' value='" & RS0("FieldAlias") &"' size='30'></td>")
		pr ("</tr>")
	end if
	RS0.movenext
loop
%>

</table>
</TD>
<TD>
<table border="0" cellpadding="0" cellspacing="0" width="100%">

<tr>
<td colspan="2" bgcolor="silver"><b>Customizable Fields</b></td>
</tr>

<tr>
<td><b>Field</b></td>
<td><b>Alias</b></td>
</tr>
<% 
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldName ORDER BY FieldID") 
do while not RS0.EOF
	if left(RS0("FieldName"),3) = "fld" then 
		pr ("<tr>")
		pr ("<td><font size='1'>" & RS0("FieldName") & "</font></td>")
		pr ("<td><input type='text' name='key" & RS0("FieldID") & "' value='" & RS0("FieldAlias") &"' size='30'></td>")
		pr ("</tr>")
	end if
	RS0.movenext
loop
%>
</table>

</TD>
</TR>
</table>

</form>
</center>
</body>
</html>
<% Conn.Close %>
