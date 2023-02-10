<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<% 
if Request("UserID") <> "0" then
	Set RS0 = Conn.Execute("SELECT * FROM Employees WHERE UserID = " & Request("UserID") ) 
	Session("currUserID") = Request("UserID")
	if not RS0.EOF then
		strName = "Employee: " & RS0("FirstName") & " " & RS0("LastName")
		BossID = RS0("BossID")
		FirstName = RS0("FirstName")
		MiddleName = RS0("MiddleName")
		LastName = RS0("LastName")
		Address = RS0("Address")
		City = RS0("City")
		State = RS0("State")
		Zip = RS0("Zip")
		Country = RS0("Country")
		Title = RS0("Title")
		WorkPhone = RS0("WorkPhone")
		WorkExt = RS0("WorkExt")
		Fax = RS0("Fax")
		Email = RS0("Email")
		HomePhone = RS0("HomePhone")
		MobilePhone = RS0("MobilePhone")
		Notes = RS0("Notes")
	end if
	editMode = true
else
	strName = "New Employee"
	editMode = false
end if

%>
<html>
<head>
<title><%= GetDefault("AppTitle", "Employee Manager") %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "Employee.asp?mode=lst"
}	
function javDelete() {
	if (confirm("Delete Employee?")) {
		location.href = "Employee.asp?mode=subEmp&sub=delete&UserID=<%= Request("UserID") %>"
	} else {
		return
	}
}	
</script>
<link rel="stylesheet" href="StyleSheet.css">
<body onLoad="document.frm.FirstName.focus()">
<% Call GetMenu() %>
<% '************* Titlebar %>
<table width="100%" border="0" bgcolor="<%= Session("TitleBarColor") %>" cellpadding="1" cellspacing="0">
<tr>
<td>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
<tr>

<form name="frm" action="Employee.asp">
<input type="hidden" name="mode" value="subEmp">
<input type="hidden" name="UserID" value="<%= Request("UserID") %>">

<% if editMode = true then %>	
	<input type="hidden" name="sub" value="edit">
<% else %>
	<input type="hidden" name="sub" value="new">
<% end if %>

<td align="left">
<font size="3" color="<%= Session("TitleTextColor") %>"><b><%= strName %></b></font>
</td>
<td align="right">
<input type="submit" value="Submit">&nbsp;
<% if editMode = true then %>	
	<input type="button" value="Delete" onClick="javDelete()" >&nbsp;
<% end if %>
<input type="button" onClick="javCancel()" value="Cancel">&nbsp;
</td>
</tr>
</table>
</td>
</tr>
</table>

<table border="0" width="100%">
<tr>
<td align="right"><font size="2">Reports To:</font></td>
<td align="left" colspan="3">
<select name="BossID">
<option value='0'>Head of Company
<%
if editMode = true then
	Session("Subordinates") = ":" & Request("UserID") & ":"
	Call GetSubordinates(Request("UserID"))
end if
Set RS1 = Conn.Execute("SELECT * FROM Employees ORDER BY LastName")
do while not RS1.EOF
	strName = RS1("LastName") & ", " & RS1("FirstName") & " - "
	if len(strName) < 7 then
		strName = "OPEN - " & RS1("Title")
	else
		strName = strName & RS1("Title")
	end if
	if editMode = true then
		if instr(Session("Subordinates"), ":" & RS1("UserID") & ":") = 0 then 
			if BossID = RS1("UserID") then
				pr ("<option value='" & RS1("UserID") & "' selected>" & strName)	
			else
				pr ("<option value='" & RS1("UserID") & "'>" & strName)	
			end if
		end if
	else
		pr ("<option value='" & RS1("UserID") & "'>" & strName)	
	end if
	RS1.movenext
loop
%>
</select>
</td>
</tr>


<tr>
<td align="right"><font size="2">First Name</font></td>
<td><input type="text" name="FirstName" size="30" tabindex="1" value="<%= FirstName %>"></td>
<td><font size=2>Title</font></td>
<td><input type="text" name="Title" size="30" tabindex="9" value="<%= Title %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Middle Name</font></td>
<td><input type="text" name="MiddleName" size="30" tabindex="2" value="<%= MiddleName %>"></td>
<td><font size=2>Work Phone</font></td>
<td><input type="text" name="WorkPhone" size="30" tabindex="10" value="<%= WorkPhone %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Last Name</font></td>
<td><input type="text" name="LastName" size="30" tabindex="3" value="<%= LastName %>"></td>
<td><font size=2>Work Ext.</font></td>
<td><input type="text" name="WorkExt" size="30" tabindex="11" value="<%= WorkExt %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Address</font></td>
<td><textarea name="Address" cols="25" rows="2" tabindex="4"><%= Address %></textarea></td>
<td><font size=2>Fax</font></td>
<td><input type="text" name="Fax" size="30" tabindex="14" value="<%= Fax %>"></td>
</tr>

<tr>
<td align="right"><font size="2">City</font></td>
<td><input type="text" name="City" size="30" tabindex="5" value="<%= City %>"></td>
<td><font size=2>Email</font></td>
<td><input type="text" name="Email" size="30" tabindex="15" value="<%= Email %>"></td>
</tr>

<tr>
<td align="right"><font size="2">State</font></td>
<td><input type="text" name="State" size="30" tabindex="6" value="<%= State %>"></td>
<td><font size=2>Home Phone</font></td>
<td><input type="text" name="HomePhone" size="30" tabindex="12" value="<%= HomePhone %>"></td>
</tr>

<tr>
<td align="right"><font size="2">Postal Code</font></td>
<td><input type="text" name="Zip" size="30" tabindex="7" value="<%= Zip %>"></td>
<td><font size=2>Mobile Phone</font></td>
<td><input type="text" name="MobilePhone" size="30" tabindex="13" value="<%= MobilePhone %>"></td>
</tr>

<tr>
<td align="right"><font size=2>Country</font></td>
<td><input type="text" name="Country" size="30" tabindex="8" value="<%= Country %>"></td>
<td><font size=2>&nbsp;</font></td>
<td>&nbsp;</td>
</tr>

<tr>
<td align="right"><font size="2">Notes</font></td>
<td colspan="3"><textarea name="Notes" cols="50" rows="4" tabindex="17"><%= Notes %></textarea></td>
</tr>


<!-------------------------  Custom Fields ------------------- -->
<% 
Set RS1 = Conn.Execute("SELECT * FROM defFields WHERE FieldAlias <> '' ORDER BY FieldID") 
pr ("<tr><td>&nbsp;</td><td colspan='3'>&nbsp;</td></tr>")
do while not RS1.EOF
	if left(RS1("FieldName"),3) = "fld" and len(RS1("FieldAlias")) > 0 then
		strValue = ""
		if editMode = true then
			key = RS1("FieldName")
			strValue = RS0(key)
		end if
		pr ("<tr>")
		pr ("<td align='right'>" & RS1("FieldAlias") & "</td>")
		pr ("<td colspan='3'><input type='text' name='" & RS1("FieldName") & "' value='" & strValue & "' size='60'></td>")
		pr ("</tr>")
	end if
	RS1.movenext
loop
%>

</table>
</form>

</body>
</html>
<% Conn.Close %>

