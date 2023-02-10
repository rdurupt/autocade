<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->

<html>
<head>
<title><%= GetDefault("AppTitle","Employee Manager") %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "Employee.asp?mode=lst&menuEmployee=EmployeeList"
}
</script>
<link rel="stylesheet" href="StyleSheet.css">
<body>

<% Call GetMenu() %>

<% '****** Header ***** %>
<table width="100%" border="0" bgcolor="<%= Session("TitleBarColor") %>" cellpadding="0" cellspacing="0">
<tr>
<td>
<table width="100%" border="0">
<tr>
<form name="frm" action="subSetField.asp">
<td align="left">
<font size="3" color="<%= Session("TitleTextColor") %>"><b>Set Fields</b></font>
</td>
<td align="right">
<input type="submit" value="Submit">&nbsp;
</td>

</tr>
</table>
</td>
</tr>
</table>

<center><font color="red"><b><%= Request("msg") %></b></font></center>

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
Set RS0 = Conn.Execute("SELECT * FROM defFields WHERE FieldName ORDER BY FieldID") 
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
Set RS0 = Conn.Execute("SELECT * FROM defFields WHERE FieldName ORDER BY FieldID") 
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