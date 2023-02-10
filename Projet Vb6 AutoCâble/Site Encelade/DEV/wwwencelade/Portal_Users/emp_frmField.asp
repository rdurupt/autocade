<% 
Response.Expires = 0 
Session("CurrentPage") = "emp_frmField.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<!--#INCLUDE FILE="inc_Menus.asp"-->

<%
if Request("mode") = "execute" then
	Conn.Execute("DELETE FROM dbp_DirFields WHERE UserID = " & Session("web_UserID"))
	strArray = split(Request("lstFieldID"), ",")
	
	for i = 0 to ubound(strArray)
		order = "order" & trim(strArray(i))
		FieldOrder = Request(order)
		if not isnumeric(FieldOrder) then
			FieldOrder = 0
		end if
		Conn.Execute("INSERT INTO dbp_DirFields(UserID,FieldID,FieldOrder) VALUES(" & Session("web_UserID") & "," & trim(strArray(i)) & "," & FieldOrder & ")")
	next
	
	'**** Reindex
	i = 0
	Set RS0 = Conn.Execute("SELECT * FROM dbp_DirFields WHERE UserID = " & Session("web_UserID") & " ORDER BY FieldOrder") 
	do while not RS0.EOF
		Set RS1 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldID = " & RS0("FieldID")) 
		if not RS1.EOF then
			i = i + 1
			Conn.Execute("UPDATE dbp_DirFields SET FieldOrder = " & i & " WHERE UserID = " & Session("web_UserID") & " AND FieldID = " & RS0("FieldID")) 
		end if
		RS0.movenext
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
<form name="frm" action="emp_frmField.asp">
<input type="hidden" name="mode" value="execute">
<td align="left">
<font size="3" color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>"><b>Select Fields</b></font>
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
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirFields WHERE UserID = " & Session("web_UserID") & " ORDER BY FieldOrder") 
do while not RS0.EOF
	Set RS1 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldID = " & RS0("FieldID")) 
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


Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldAlias <> '' ORDER BY FieldID") 
do while not RS0.EOF
	Set RS1 = Conn.Execute("SELECT * FROM dbp_DirFields WHERE FieldID = " & RS0("FieldID")& " AND UserID = " & Session("web_UserID")) 
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
