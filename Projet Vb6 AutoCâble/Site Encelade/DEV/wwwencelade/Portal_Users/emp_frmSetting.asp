<% 
Response.Expires = 0 
Session("CurrentPage") = "emp_frmSetting.asp"
%>
<!--#include file="ADOConnect.asp"-->
<!--#include file="inc_Utilities.asp"-->
<!--#INCLUDE FILE="inc_Menus.asp"-->
<!--#INCLUDE FILE="emp_TopMenu.asp"-->
<%
if Request("mode") = "execute" then
	if Request("user_action") = "Set to Default" then
		if Session("Admin") = 1  then 
			Call SetDefault("emp_AppTitle","Employee Manager")
			Call SetDefault("emp_CompanyName", " ")
		end if
		Call SetUserDefault(Session("web_UserID"),"emp_MenuWidth","150")
		Call SetUserDefault(Session("web_UserID"),"emp_PageNum","15")
		Call SetUserDefault(Session("web_UserID"),"emp_bgcolor","#6CBFD0")
		Call SetUserDefault(Session("web_UserID"),"emp_color","navy")
		msg = "Settings set to default."
	else
		if Session("admin") = 1  then 
			Call SetDefault("emp_AppTitle",safeEntry(Request("emp_AppTitle")))
			Call SetDefault("emp_CompanyName",safeEntry(Request("emp_CompanyName")))
		end if
		Call SetUserDefault(Session("web_UserID"),"emp_MenuWidth",Request("emp_MenuWidth"))
		Call SetUserDefault(Session("web_UserID"),"emp_PageNum",Request("emp_PageNum"))
		Call SetUserDefault(Session("web_UserID"),"emp_bgcolor",safeEntry(Request("emp_bgcolor")))
		Call SetUserDefault(Session("web_UserID"),"emp_color",safeEntry(Request("emp_color")))
		msg = "Settings updated."
	end if
end if
%>
<html>
<head>
<title><%= GetDefault("emp_AppTitle","Employee Manager") %></title>
</head>

<script language="javascript">

</script>
<link rel="stylesheet" href="StyleSheet.css">
<body background="../Portal_Html/Images/Background.asp" bgcolor="#F5F7FE">

<% Call GetMenu() %>

<% '****** Header ***** %>
<table width="100%" border="2" bgcolor="<%= GetUserDefault(Session("web_UserID"),"emp_bgcolor","#678DB8") %>" cellpadding="0" cellspacing="0">
<tr>
<td>
<table width="100%" border="0">
<tr>
<td align="left">
<font size="3" color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy") %>"><b>Settings</b></font>
</td>
<form name="frm" action="emp_frmSetting.asp">
<td align="right">
<input type="submit" name="user_action" value="Submit">
<input type="submit" name="user_action" value="Set to Default">
</td>
</tr>
</table>
</td>
</tr>
</table>

<% '******  Message ****** %>
<center><font color="red"><b><%= msg %></b></font></center>

    

<input type="hidden" name="mode" value="execute">

<table border="0">
<%if Session("admin") = 1 then %>
	<tr>
	<td align="right">Company Name</td>
	<td><input type="text" name="emp_CompanyName" size="40" value="<%= GetDefault("emp_CompanyName","DemoTech, Inc.") %>"></td>
	</tr>
	<tr>
	<td align="right">Application Title</td>
	<td><input type="text" name="emp_AppTitle" size="40" value="<%= GetDefault("emp_AppTitle","Employee Manager") %>"></td>
	</tr>
<% end if %>

<tr>
<td align="right">Width of menu frame</td>
<td>
<select name="emp_MenuWidth">
<% 
emp_MenuWidth = GetUserDefault(Session("web_UserID"),"emp_MenuWidth","150")
strArray = split("100:110:120:125:130:135:140:145:150:160:170:180:190:200:210:220:230:240:250:260:300:325:350", ":")
for i = 0 to ubound(strArray)
	if cStr(strArray(i)) = emp_MenuWidth then
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
<select name="emp_PageNum">
<% 
PageNum = GetUserDefault(Session("web_UserID"),"emp_PageNum","20")
strArray = split("5:10:15:20:25:30:35:40:45:50:55:60:65:70:80:90:100:200:300:350:400", ":")
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
<td><input type="text" name="emp_bgcolor" value="<%= GetUserDefault(Session("web_UserID"),"emp_bgcolor","#C5E9E7") %>"></td>
</tr>

<tr>
<td align="right">Title bar text color</td>
<td><input type="text" name="emp_color" value="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy") %>"></td>
</tr>

    
</table>
</form>
</center>
	
</body>
</html>
<% Conn.Close %>
