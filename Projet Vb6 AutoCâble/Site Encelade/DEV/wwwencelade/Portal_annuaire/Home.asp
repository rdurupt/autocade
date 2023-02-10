<% Response.Expires = 0 %>
<html>
<head>
<title>Resources Management</title>
</head>
<basefont FACE="Arial, Helvetica, sans-serif">
<link REL="stylesheet" HREF="style.css">
<body BGCOLOR="White" <%= LoadMe %>>
<img src="corner.gif" border="0" WIDTH="23" HEIGHT="20"><br clear="all">

<% '********************** HEADING ***************** %>
<table>
<tr>
<td valign="middle" align="left" NOWRAP>
&nbsp;<img src="contacts.gif" border="0" WIDTH="32" HEIGHT="32">
<font size="+3" face="Times New Roman" color="#004080"><strong>R</strong></font><font size="+3" color="#606060"><em>esource</em></font>&nbsp;
<font size="+3" face="Times New Roman" color="#004080"><strong>M</strong></font><font size="+3" color="#606060"><em>anagement</em></font>&nbsp;
</td>
</tr>
</table>
<% Session("fromPage") = "home.asp" %>

<% '********************** TITLEBAR ***************** %>
<table WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr>
<th NOWRAP BGCOLOR="#004080" ALIGN="Left">
<font SIZE="4" COLOR="WHITE">&nbsp;Main Menu</font>
</th>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->
<br>

<table CELLSPACING="1" CELLPADDING="3" BORDER="0">
<tr><td>&nbsp;</td><td bgcolor="#004080"><font color="white"><b>Records</b></font></td></tr>
<tr><td width="35">&nbsp;</td><td><a href="lstEmployee.asp"><font size="2">Employee List</font></a></td></tr>
<tr><td width="35">&nbsp;</td><td><a href="lstPosition.asp"><font size="2">Position List</font></a></td></tr>
<tr><td width="35">&nbsp;</td><td><a href="chtOrganization.asp"><font size="2">Organizational Chart</font></a></td></tr>
<tr><td>&nbsp;</td><td>&nbsp;</td></tr>

<tr><td>&nbsp;</td><td bgcolor="#004080"><font color="white"><b>Tasks</b></font></td></tr>
<tr><td width="35">&nbsp;</td><td><a href="Query.asp"><font size="2">Query</font></a></td></tr>
<% if Session("employeeAccess") > 1 then %>
	<tr><td width="35">&nbsp;</td><td><a href="frmNewEmp.asp"><font size="2">Add Employee</font></a></td></tr>
	<tr><td width="35">&nbsp;</td><td><a href="frmNewPos.asp"><font size="2">Add Position</font></a></td></tr>
<% end if %>

<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
	
<tr><td>&nbsp;</td><td bgcolor="#004080"><font color="white"><b>Administration</b></font></td></tr>
<tr><td width="35">&nbsp;</td><td><a href="frmPassword.asp"><font size="2">Change Password</font></a></td></tr>
<% if Session("employeeAccess") > 2 then %>
	<tr><td width="35">&nbsp;</td><td><a href="lstValidation.asp"><font size="2">Validation Tables</font></a></td></tr>
	<tr><td width="35">&nbsp;</td><td><a href="lstUser.asp"><font size="2">Users</font></a></td></tr>
<% end if %>

</table>
<br><br><br><br>
<br><br><br>
<br><br><br><font face="arial" size="1">Copyright © 1999 ASP Intranet Inc.&nbsp;&nbsp;All rights reserved.</font>


</body>
</html>
<% Conn.Close %>