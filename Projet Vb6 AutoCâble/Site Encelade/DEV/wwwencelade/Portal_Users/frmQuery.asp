<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->

<html>
<head>
<title><%= GetDefault("AppTitle","Employee Manager") %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "Employee.asp?mode=lst"
}

</script>
<link rel="stylesheet" href="StyleSheet.css">
<body onLoad="document.frm.qryFirstName.focus()">

<% Call GetMenu() %>


<% '****** Header ***** %>
<table width="100%" border="3" bgcolor="<%= Session("TitleBarColor") %>">
<tr>
<td>
<table width="100%" border="0">
<tr>
<td align="left">
<font size="3" color="<%= Session("TitleTextColor") %>"><b>Query</b></font>
</td>
</tr>
</table>
</td>
</tr>
</table>

<form name="frm" action="Employee.asp">
<input type="hidden" name="mode" value="lst">
<input type="hidden" name="employeeQuery" value="true">

<table cellspacing="0" cellpadding="0" border="0">

<tr>
<td nowrap><b>First Name</b>&nbsp;</td>
<td><input type="text" size="30" name="qryFirstName" value="<%= Session("qryFirstName") %>"></td>
</tr>

<tr>
<td nowrap><b>Last Name</b>&nbsp;</td>
<td><input type="text" size="30" name="qryLastName" value="<%= Session("qryLastName") %>"></td>
</tr>

<tr>
<td nowrap><b>Title</b>&nbsp;</td>
<td><input type="text" size="30" name="qryTitle" value="<%= Session("qryTitle") %>"></td>
</tr>

<tr>
<td>&nbsp;</td>
<td>
<input type="submit" value=" OK ">
<input type="button" value="Cancel" onClick="javCancel()">
</td>
</tr>
</table>


</form>

</body>
</html>

<% Conn.Close %>