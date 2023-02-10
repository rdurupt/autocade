<% Response.Expires = 0 %>
<!--#include file="ADOConnect.asp"-->
<!--#include file="inc_Utilities.asp"-->

<html>
<head>
<title><%= GetDefault("AppTitle","Employee Manager") %></title>
</head>

<script language="javascript">

</script>
<link rel="stylesheet" href="StyleSheet.css">
<body>

<% Call GetMenu() %>

<% '****** Header ***** %>
<table width="100%" border="2" bgcolor="<%= Session("TitleBarColor") %>" cellpadding="0" cellspacing="0">
<tr>
<td>
<table width="100%" border="0">
<tr>
<td align="left">
<font size="3" color="<%= Session("TitleTextColor") %>"><b>Settings</b></font>
</td>
</tr>
</table>
</td>
</tr>
</table>

<% '******  Message ****** %>
<center><font color="red"><b><%= Request("msg") %></b></font></center>

    
<form name="frm" action="Employee.asp">
<input type="hidden" name="mode" value="subSetting">

<table border="0">
<% if Session("employeeAccess") > 2 then %>
	<tr>
	<td align="right">Company Name</td>
	<td><input type="text" name="CompanyName" size="40" value="<%= GetDefault("CompanyName","DemoTech, Inc.") %>"></td>
	</tr>
	<tr>
	<td align="right">Application Title</td>
	<td><input type="text" name="AppTitle" size="40" value="<%= GetDefault("AppTitle","Employee Manager") %>"></td>
	</tr>
<% end if %>

<tr>
<td align="right">Width of menu frame</td>
<td>
<select name="MenuWidth">
<% 
MenuWidth = GetDefaultUser("MenuWidth","150")
strArray = split("100:125:150:160:170:180:190:200:210:220:230:240:250:260", ":")
for i = 0 to ubound(strArray)
	if cStr(strArray(i)) = MenuWidth then
		pr ("<option value='" & strArray(i) & "' selected>" & strArray(i))
	else
		pr ("<option value='" & strArray(i) & "'>" & strArray(i))
	end if
next
%>
</select>
</td>
</tr>

<tr>
<td align="right">Records per page</td>
<td>
<select name="PageNum">
<% 
PageNum = GetDefaultUser("PageNum","20")
strArray = split("10:15:20:25:30:35:40:45:50:55:60:65:70:80:90:100:200:300", ":")
for i = 0 to ubound(strArray)
	if cStr(strArray(i)) = PageNum then
		pr ("<option value='" & strArray(i) & "' selected>" & strArray(i))
	else
		pr ("<option value='" & strArray(i) & "'>" & strArray(i))
	end if
next
%>
</select>
</td>
</tr>

<tr>
<td align="right">Title bar background color</td>
<td><input type="text" name="TitleBarColor" value="<%= GetDefaultUser("TitleBarColor","#C5E9E7") %>"></td>
</tr>

<tr>
<td align="right">Title bar text color</td>
<td><input type="text" name="TitleTextColor" value="<%= GetDefaultUser("TitleTextColor","navy") %>"></td>
</tr>

<tr>
<td align="right"></td>
<td>
<input type="submit" name="userAction" value="Submit">&nbsp;
<input type="submit" name="userAction" value="Set to Default">
</td>
</tr>
    
</table>
</form>
</center>
	
</body>
</html>
<% Conn.Close %>
