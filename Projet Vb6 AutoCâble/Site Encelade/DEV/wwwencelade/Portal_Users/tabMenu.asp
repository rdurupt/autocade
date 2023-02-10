<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<html>
<head>
<title><%= GetDefault("AppTitle","Employee Manager") %></title>
</head>
<script language="vbscript">
function vbMouseOver(i)
	if i = 1 then
		tr1.style.backgroundcolor = "navy"
		sp1.style.color = "white"
	elseif i = 2 then
		tr2.style.backgroundcolor = "navy"
		sp2.style.color = "white"
	elseif i = 3 then
		tr3.style.backgroundcolor = "navy"
		sp3.style.color = "white"
	end if
end function
function vbMouseOut(i)
	if i = 1 then
		tr1.style.backgroundcolor = "white"
		sp1.style.color = "navy"
	elseif i = 2 then
		tr2.style.backgroundcolor = "white"
		sp2.style.color = "navy"
	elseif i = 3 then
		tr3.style.backgroundcolor = "white"
		sp3.style.color = "navy"
	end if
end function
</script>
<link rel="stylesheet" href="TabSheet.css">
<body bgcolor="white">

<center>

<table border="0" bgcolor="white" cellpadding="0" cellspacing="1">
<tr><td> </td></tr>
</table>

<table border="0" bgcolor="gray" cellpadding="0" cellspacing="1">

<tr id="tr1" bgcolor="white" onMouseOver="vbMouseOver(1)" onMouseOut="vbMouseOut(1)"><td nowrap align="center">
<a href="Employee.asp?mode=lst&amp;menuEmployee=EmployeeList" target="emp2"><span id="sp1"><b>&nbsp;Employee List&nbsp;</b></span></a>
</td></tr>

<tr id="tr2" bgcolor="white" onMouseOver="vbMouseOver(2)" onMouseOut="vbMouseOut(2)"><td nowrap align="center">
<a href="Employee.asp?mode=chtOrganization&amp;menuEmployee=none" target="emp2"><span id="sp2"><b>&nbsp;Organization Chart&nbsp;</b></span></a>
</td></tr>

<tr id="tr3" bgcolor="white" onMouseOver="vbMouseOver(3)" onMouseOut="vbMouseOut(3)"><td nowrap align="center">
<a href="frmSetting.asp?menuEmployee=Administration" target="emp2"><span id="sp3"><b>&nbsp;Administration&nbsp;</b></span></a>
</td></tr>

</table>

</center>
</body>
</html>
    
<% Conn.Close %>
