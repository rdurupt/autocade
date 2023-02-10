<% 
Response.Expires = 0 
Session("CurrentPage") = "emp_frmQuery.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<!--#INCLUDE FILE="inc_Menus.asp"-->
<!--#INCLUDE FILE="emp_TopMenu.asp"-->
<html>
<head>
<title><%= GetDefault("emp_AppTitle","Employee Manager") %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "ASPIntranet.asp?mode=emp_lst"
}

</script>
<link rel="stylesheet" href="StyleSheet.css">
<body background="../Portal_Html/Images/Background.asp" bgcolor="#F5F7FE" onLoad="document.frm.qryFirstName.focus()">

<% 
Call GetMenu() 
'**********  Get Field Labels
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields")
do while not RS0.EOF
	Session(RS0("FieldName")) = RS0("FieldAlias")
	RS0.movenext
loop 
%>


<% '****** Header ***** %>
<table width="100%" border="3" bgcolor="<%= GetUserDefault(Session("web_UserID"),"emp_bgcolor","#678DB8") %>">
<tr>
<td>
<table width="100%" border="0">
<tr>
<td align="left">
<font size="3" color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>"><b>Query</b></font>
</td>
</tr>
</table>
</td>
</tr>
</table>

<form name="frm" action="ASPIntranet.asp">
<input type="hidden" name="mode" value="emp_lst">
<input type="hidden" name="employeeQuery" value="true">

<table cellspacing="0" cellpadding="0" border="0">

<tr>
<td nowrap>&nbsp;&nbsp;<b><%= Session("FirstName") %></b>&nbsp;</td>
<td><input type="text" size="30" name="qryFirstName" value="<%= Session("qryFirstName") %>"></td>
</tr>

<tr>
<td nowrap>&nbsp;&nbsp;<b><%= Session("LastName") %></b>&nbsp;</td>
<td><input type="text" size="30" name="qryLastName" value="<%= Session("qryLastName") %>"></td>
</tr>

<tr>
<td nowrap>&nbsp;&nbsp;<b><%= Session("Title") %></b>&nbsp;</td>
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