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
<form name="frm" action="subField.asp">
<td align="left">
<font size="3" color="<%= Session("TitleTextColor") %>"><b>Select Fields</b></font>
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
&nbsp;&nbsp;Select the fields that you wish to view in the Employee List.
<br><br>
<table width="40%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>Field</td>
<td>Order</td>
</tr>
<% 
Set RS0 = Conn.Execute("SELECT * FROM dbp_dirFields WHERE UserID = " & Session("UserID") & " ORDER BY FieldOrder") 
do while not RS0.EOF
	Set RS1 = Conn.Execute("SELECT * FROM dbp_dirdefFields WHERE FieldID = " & RS0("FieldID")) 
	if not RS1.EOF then
		pr ("<tr>")
		pr ("<td>&nbsp;</td>")
		pr ("<td align='right'><input type='checkbox' name='lstFieldID' value='" & RS1("FieldID") &"' checked></td>")
		pr ("<td><font size='1'>" & RS1("FieldAlias") & "</font></td>")
		pr ("<td><input type='text' name='order" & RS1("FieldID") & "' value='" & RS0("FieldOrder") &"' size='4'></td>")
		pr ("</tr>")
	end if
	RS0.movenext
loop


Set RS0 = Conn.Execute("SELECT * FROM dbp_dirdefFields WHERE FieldAlias <> '' ORDER BY FieldID") 
do while not RS0.EOF
	Set RS1 = Conn.Execute("SELECT * FROM dbp_dirFields WHERE FieldID = " & RS0("FieldID")& " AND UserID = " & Session("UserID")) 
	if RS1.EOF then
		pr ("<tr>")
		pr ("<td>&nbsp;</td>")
		pr ("<td align='right'><input type='checkbox' name='lstFieldID' value='" & RS0("FieldID") &"'></td>")
		pr ("<td><font size='1'>" & RS0("FieldAlias") & "</font></td>")
		pr ("<td><input type='text' name='order" & RS0("FieldID") & "' value='' size='4'></td>")
		pr ("</tr>")
	end if
	RS0.movenext
loop
%>
</table>

</form>
</body>
</html>
<% Conn.Close %>