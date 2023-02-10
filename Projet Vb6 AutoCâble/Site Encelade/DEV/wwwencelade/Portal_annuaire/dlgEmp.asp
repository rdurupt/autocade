<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<% 
if Request("mode") = "edit" then
	'*** UPDATE
	strSQL = "UPDATE Employees SET "
	strSQL = strSQL & "BossID = " & Request("BossID") & ","
	strSQL = strSQL & "FirstName = '" & safeEntry(Request("FirstName")) & "',"
	strSQL = strSQL & "MiddleName = '" & safeEntry(Request("MiddleName")) & "',"
	strSQL = strSQL & "LastName = '" & safeEntry(Request("LastName")) & "',"
	strSQL = strSQL & "Title = '" & safeEntry(Request("Title")) & "',"
	strSQL = strSQL & "WorkPhone = '" & safeEntry(Request("WorkPhone")) & "',"
	strSQL = strSQL & "WorkExt = '" & safeEntry(Request("WorkExt")) & "',"
	strSQL = strSQL & "Email = '" & safeEntry(Request("Email")) & "' "
	strSQL = strSQL & "WHERE UserID = " & Request("UserID")
	Conn.Execute(strSQL)
	response.redirect "dlgStatus.asp?toPage=Employee.asp&key=mode&keyValue=chtOrganization&key2=UserID&keyValue2=" & Request("UserID") & "&msg=Updating+employee..."
end if

if Request("mode") = "delete" then
	Set RS0 = Conn.Execute("SELECT * FROM Employees WHERE UserID = " & Request("UserID")) 
	if not RS0.EOF then
		BossID = RS0("BossID")
	end if
	'*** DELETE
	Conn.Execute("DELETE FROM Employees WHERE UserID = " & Request("UserID")) 
	response.redirect "dlgStatus.asp?toPage=Employee.asp&key=mode&keyValue=chtOrganization&key2=UserID&keyValue2=" & BossID & "&msg=Deleting+employee..."
end if

Set RS0 = Conn.Execute("SELECT * FROM Employees WHERE UserID = " & Request("UserID") ) 
if not RS0.EOF then
	strName = RS0("FirstName") & " " & RS0("LastName")
	BossID = RS0("BossID")
	FirstName = RS0("FirstName")
	MiddleName = RS0("MiddleName")
	LastName = RS0("LastName")
	Title = RS0("Title")
	WorkPhone = RS0("WorkPhone")
	WorkExt = RS0("WorkExt")
	Email = RS0("Email")
	Session("Subordinates") = ":" & Request("UserID") & ":"
	Call GetSubordinates(Request("UserID"))
	exclusive = replace(Session("Subordinates"), ":" & Request("UserID") & ":", "")
	
end if
%>
<html>
<head>
<title>Information for <%= strName %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "Employee.asp?mode=lst"
}
<% if len(exclusive) > 2 then %>
function javDelete() {
	alert("Cannot delete employee with subordinates.")
}	
<% else	 %>
function javDelete() {
	if (confirm("Delete Employee?")) {
		location.href = "dlgEmp.asp?mode=delete&UserID=<%= Request("UserID") %>"
	} else {
		return
	}
}	
<% end if %>
</script>
<link rel="stylesheet" href="StyleSheet.css">
<body onLoad="document.frm.Title.focus()">
<% '************* Titlebar %>
<table width="100%" border="2" bgcolor="#698DC5" cellpadding="1" cellspacing="0">
<tr>
<td>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
<tr>

<form name="frm" action="dlgEmp.asp">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="UserID" value="<%= Request("UserID") %>">

<td align="left">
<font size="3" color="white"><b><%= strName %></b></font>
</td>
<td align="right">
<input type="submit" value="Submit">&nbsp;
<input type="button" value="Delete" onClick="javDelete()" >&nbsp;
<input type="button" onClick="window.close()" value="Cancel">&nbsp;
</td>
</tr>
</table>
</td>
</tr>
</table>

<br>
<table border="0" width="100%">

<tr>
<td align="right"><font size="2">Title</font></td>
<td align="left"><input type="text" name="Title" size="30" value="<%= Title %>"></td>
</tr>

<tr>
<td align="right"><font size="2">First Name</font></td>
<td align="left"><input type="text" name="FirstName" size="30" value="<%= FirstName %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Middle Name</font></td>
<td align="left"><input type="text" name="MiddleName" size="30" value="<%= MiddleName %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Last Name</font></td>
<td align="left"><input type="text" name="LastName" size="30" value="<%= LastName %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Work Phone</font></td>
<td align="left"><input type="text" name="WorkPhone" size="30" value="<%= WorkPhone %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Work Extension</font></td>
<td align="left"><input type="text" name="WorkExt" size="30" value="<%= WorkExt %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Email</font></td>
<td align="left"><input type="text" name="Email" size="30" value="<%= Email %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Reports To:</font></td>
<td align="left">
<select name="BossID">
<option value='0'>Head of Company
<%
Set RS1 = Conn.Execute("SELECT * FROM Employees ORDER BY LastName")
do while not RS1.EOF
	strName = RS1("LastName") & ", " & RS1("FirstName") & " - "
	if len(strName) < 7 then
		strName = "OPEN - " & RS1("Title")
	else
		strName = strName & RS1("Title")
	end if

	if instr(Session("Subordinates"), ":" & RS1("UserID") & ":") = 0 then 
		if BossID = RS1("UserID") then
			pr ("<option value='" & RS1("UserID") & "' selected>" & strName)	
		else
			pr ("<option value='" & RS1("UserID") & "'>" & strName)	
		end if
	end if

	RS1.movenext
loop
%>
</select>
</td>
</tr>

</table>
</form>

</body>
</html>
<% Conn.Close %>
