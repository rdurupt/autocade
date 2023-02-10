<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<html>
<head>
<title><%= GetDefault("emp_AppTitle","Employee Manager") %></title>
</head>
<script src="inc_highlight.js" language="vbscript"></script>
<link rel=STYLESHEET href="../Portal_Styles/PNavStyle1.css" type="text/css">
<body background="../Portal_Html/Images/Background.asp" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table border="0" width="100%">
<tr>
	<td align="left" bgcolor="#004080"><font size="2" color="white"><b>&nbsp;Directory</b></font>&nbsp;</td>
		</tr>
</table>
<table width="100%" border="0" bgcolor="white" cellpadding="0" cellspacing="1">
<tr><td> </td></tr>
</table>
<table width="100%" border="0" bgcolor="gray" cellpadding="0" cellspacing="1">

<tr  bgcolor="white"><td nowrap align="center">
<a href="ASPIntranet.asp?mode=emp_lst&amp;menuEmployee=EmployeeList" target="emp2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;Members List&nbsp;</b></a>
</td></tr>

<tr  bgcolor="white"> <td nowrap align="center">
<a href="ASPIntranet.asp?mode=emp_chtOrganization&amp;menuEmployee=none" target="emp2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;Organization Chart&nbsp;</b></a>
</td></tr>

<tr  bgcolor="white" ><td nowrap align="center">
<a href="emp_frmQuery.asp" target="emp2"><span id="sp3"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;Query&nbsp;</b></a>
</td></tr>

<tr  bgcolor="white" ><td nowrap align="center">
<a href="emp_frmSetting.asp?menuEmployee=Administration" target="emp2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;Settings&nbsp;</b></a>
</td></tr>

<% if Session("Admin") = 1  then %>
<tr  bgcolor="white" ><td nowrap align="center">
<a href="modUser.asp?mode=web_lst" target="emp2"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">&nbsp;Manage Logons&nbsp;</b></span></a>
</td></tr>
<% end if %>
</table>
</center>
</body>
</html>
    
<% Conn.Close %>
