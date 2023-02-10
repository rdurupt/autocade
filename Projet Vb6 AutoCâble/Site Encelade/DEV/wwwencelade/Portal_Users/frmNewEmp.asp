<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<% Session("fromPage") = "frmNewEmp.asp" %>
<% LoadMe = "onLoad=""document.frm.FirstName.focus()""" %>
<!--#INCLUDE FILE="zHeading.asp"-->
<script language="javascript">
function javDate() {
	window.open('dlgDate.asp','dialog','status=no,width=300,height=250,top=250,left=400')
}
function javCancel() {
	location.href = "Home.asp"
}	
function javAddPosition() {
	document.frm.action = "frmNewPos.asp"
	document.frm.submit()
}	
</script>

<% '********************** TITLEBAR ***************** %>
<table BGCOLOR="#004080" WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr>
<th NOWRAP ALIGN="Left">
<font SIZE="4" COLOR="WHITE">&nbsp;New Employee</font>
</th>
<form name="frm" action="subNewEmp.asp">
<td align="right">
<input TYPE="submit" NAME="subAction" VALUE="Submit">
<input type="button" onClick="javCancel()" value="Cancel">&nbsp;
</td>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->

<BR>
<% '********** Start fields ****** %>
<table border="0"> 

<tr>
<td width="150" align="right"><font size=2>First Name: </font></td>
<td><input type="text" name="FirstName" size="20" value="<%= Request("FirstName") %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Last Name: </font></td>
<td><input type="text" name="LastName" size="20" value="<%= Request("LastName") %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Extension: </font></td>
<td><input type="text" name="Extension" size="7" value="<%= Request("Extension") %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Email: </font></td>
<td><input type="text" name="Email" size="40" value="<%= Request("Email") %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>MobilePager: </font></td>
<td><input type="text" name="MobilePager" size="20" value="<%= Request("MobilePager") %>"></td>
</tr>
<tr>

<td width="150" align="right"><font size=2>Hire Date: </font></td>
<% if len(Request("msg")) > 0 then %>
	<td bgcolor="red">
<% else %>
	<td>
<% end if %>

<input type="text" name="HireDate" size="20" value="<%= Request("HireDate") %>">
<input type="button" value="..." onClick="javDate()"></td>
</tr>

<tr><td colspan="2">&nbsp;</td></tr>

<tr>
<td width="150" align="right">&nbsp;</td>
<td bgcolor="#004080"><font color="white"><b>Position Information</b></font></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Available Position:</font></td>
<td>
<select name="PositionID">
<option value="0">
<% Set RS1 = Conn.execute("SELECT * FROM Positions WHERE EmployeeID = null or EmployeeID = 0 ORDER BY Title") %>
<% Do While not RS1.EOF %>
	<% if Cstr(RS1("PositionID")) = Request("PositionID") then %>
		<option value="<%= RS1("PositionID") %>" selected><%= RS1("Title") %>
	<% else %>
		<option value="<%= RS1("PositionID") %>"><%= RS1("Title") %>
	<% end if %>
	<% RS1.MoveNext %>
<% Loop %>
</select>
<input type="button" value="Add Position" onClick="javAddPosition()">
</td>
</tr>

</table>
</form>

</body>
</html>
<% Conn.Close %>

