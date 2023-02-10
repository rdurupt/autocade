<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<% Session("currEmployeeID") = Request("EmployeeID") %>
<html>
<head>
<title>Resource Management</title>
</head>

<basefont FACE="Arial, Helvetica, sans-serif">
<link REL="STYLESHEET" HREF="Style2.css">
<body BACKGROUND="Back2.jpg" BGCOLOR="White">

<!--#INCLUDE FILE="zHeading.asp"-->

<% '********************** TITLEBAR ***************** %>
<table BGCOLOR="#000080" WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr>
<th NOWRAP ALIGN="Left" BACKGROUND="NavBLUE.jpg">
<font SIZE="5" COLOR="WHITE">&nbsp;Delete Employee</font>
</th>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->


<% currKey = "curr" & Request("key") %>
<% Session(currKey) = Request("ID") %>
<form name="frm" action="subDeleteEmp.asp">
<input type="hidden" name="EmployeeID" value="<%= Request("EmployeeID") %>">


<table width="100%" border="0">
<tr>
<td width="25">&nbsp;</td>
<td width="50"><img src="warning.gif" border="0"></td>
<td>Delete employee?</td>
</tr>

<tr>
<td width="25">&nbsp;</td>
<td width="50">&nbsp;</td>
<td><input type="submit" name="subAction" value="OK">&nbsp;<input type="submit" name="subAction" value="Cancel"></td>
</tr>
</table>

<br><br><br><font face="arial" size="1">Copyright © 1998 S Hurdowar. All rights reserved.</font>
</body>
</html>



