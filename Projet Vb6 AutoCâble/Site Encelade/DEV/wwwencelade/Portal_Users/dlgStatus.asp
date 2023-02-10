<% Response.Expires= 0 %>
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<% 
strKey = "?x=1"
if len(Request("key")) > 0 then
	strKey = strKey & "&" & Request("key") & "=" & replace(Request("keyValue"), " ", "+") 
end if
if len(Request("key2")) > 0 then
	strKey = strKey & "&" & Request("key2") & "=" & replace(Request("keyValue2"), " ", "+") 
end if
if len(Request("key3")) > 0 then
	strKey = strKey & "&" & Request("key3") & "=" & replace(Request("keyValue3"), " ", "+")  
end if
if len(Request("key4")) > 0 then
	strKey = strKey & "&" & Request("key4") & "=" & replace(Request("keyValue4"), " ", "+")  
end if
if len(Request("key5")) > 0 then
	strKey = strKey & "&" & Request("key5") & "=" & replace(Request("keyValue5"), " ", "+")  
end if

strKey = replace(strKey, " ", "+")

start_function = "onLoad=""javClose()"""
cmd = ""
if Request("toPage") = "NoClose" then
	start_function = ""
	cmd = "<input type='button' value='Close' onClick='window.close()'>"
end if
%>
<html>
<head>
	<title>Status</title>
</head>
<script language="javascript">
function javClose() {
<% if len(Request("toPage")) > 0 and Request("toPage") <> "NoClose" then %>
	window.opener.location = "<%= Request("toPage") %><%= strKey %>";
<% end if %>
	setTimeout("window.close()",1500);
}
</script>

<body <%= start_function %>>
<link rel="stylesheet" href="StyleSheet.css">
<center>
<form name="frm">
<table width="100%">
<tr bgcolor="#003080"><td align="center"><font size="3" color="white"><b>Status</b></font></td></tr>
<tr><td align="center">&nbsp;</td></tr>
<tr><td align="center"><b><%= Request("msg") %></b></td></tr>
<tr><td align="center"><%= cmd %></td></tr>
</table>
</form>
</center>

</body>
</html>
