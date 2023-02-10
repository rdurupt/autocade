<% Response.Expires = 0 %>
<% LoadMe = "onLoad=""document.frm.FirstName.focus()""" %>
<!--#INCLUDE FILE="zHeading.asp"-->
<% Set RS0 = Conn.execute(sql0 & "AND Employees.EmployeeID = " & Request("EmployeeID")) 
if len(Request("FirstName")) > 0 or len(Request("LastName")) > 0 then
	FirstName = Request("FirstName")
	LastName = Request("LastName")
	Extension = Request("Extension")
	Email = Request("Email")
	MobilePager = Request("MobilePager")
	HireDate = Request("HireDate")
	PositionID = Request("PositionID")
else
	FirstName = RS0("FirstName")
	LastName = RS0("LastName")
	Extension = RS0("Extension")
	Email = RS0("Email")
	MobilePager = RS0("MobilePager")
	HireDate = RS0("HireDate")
	PositionID = RS0("PositionID")
end if
%>

<script language="javascript">
function javDate() {
	window.open('dlgDate.asp','dialog','status=no,width=300,height=250,top=250,left=400')
}
function javCancel() {
	location.href = "readEmp.asp?EmployeeID=<%= Request("EmployeeID") %>"
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
<font SIZE="4" COLOR="WHITE">&nbsp;Edit Employee</font>
</th>
<form name="frm" action="subEditEmp.asp">
<td align="right">
<input TYPE="submit" NAME="subAction" VALUE="Submit">
<input type="button" onClick="javCancel()" value="Cancel">&nbsp;
</td>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->
<input TYPE="hidden" NAME="fromPage" VALUE="frmEditEmp.asp">
<input TYPE="hidden" NAME="EmployeeID" VALUE="<%= Request("EmployeeID") %>">

<BR>
<% '********** Start fields ****** %>
<table border="0"> 

<tr>
<td width="150" align="right"><font size=2>First Name: </font></td>
<td><input type="text" name="FirstName" size="20" value="<%= FirstName %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Last Name: </font></td>
<td><input type="text" name="LastName" size="20" value="<%= LastName %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Extension: </font></td>
<td><input type="text" name="Extension" size="7" value="<%= Extension %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Email: </font></td>
<td><input type="text" name="Email" size="40" value="<%= Email %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Mobile Pager: </font></td>
<td><input type="text" name="MobilePager" size="20" value="<%= MobilePager %>"></td>
</tr>
<tr>

<td width="150" align="right"><font size=2>Hire Date: </font></td>
<% if len(Request("msg")) > 0 then %>
	<td bgcolor="red">
<% else %>
	<td>
<% end if %>

<input type="text" name="HireDate" size="20" value="<%= HireDate %>">
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
<option value="<%= RS0("PositionID") %>"><%= RS0("Title") %>
<% Set RS1 = Conn.execute("SELECT * FROM Positions WHERE EmployeeID = null or EmployeeID = 0 ORDER BY Title") %>
<% Do While not RS1.EOF %>
	<% if Cstr(RS1("PositionID")) = PositionID then %>
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

