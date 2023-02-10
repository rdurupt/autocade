<% Response.Expires = 0 %>
<% LoadMe = "onLoad=""document.frm.Title.focus()""" %>
<!--#INCLUDE FILE="zHeading.asp"-->
<%
public function GetDescendents(d,b)
	Set RS20 = Conn.Execute(sqlPOS & "AND BossID = " & d)
	do while not RS20.EOF
		if RS20("PositionID") = b then
			Session("IsSubordinate") = true
			exit do
		end if
		
		Call GetDescendents(RS20("PositionID"),b)
		RS20.movenext
	loop	
end function
%>
<% Set RS0 = Conn.execute(sqlPOS & "AND Positions.PositionID = " & Request("PositionID")) %>

<script language="javascript">

function javCancel() {
	document.frm.action = "readPos.asp?PositionID=<%= Request("PositionID") %>"
	document.frm.submit()
}	
function javSubmit() {
	if (document.frm.Title.value == "" || document.frm.Title.value == null) {
		alert("Position title required.")
		document.frm.Title.focus()
		return
	} else {
		document.frm.submit()
	}
}	
</script>

<% '********************** TITLEBAR ***************** %>
<table BGCOLOR="#004080" WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr>
<th NOWRAP ALIGN="Left">
<font SIZE="4" COLOR="WHITE">&nbsp;Edit Position</font>
</th>
<form name="frm" action="subEditPos.asp">
<td align="right">
<input type="button" onClick="javSubmit()" value="Submit">
<input type="button" onClick="javCancel()" value="Cancel">&nbsp;
</td>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->

<input TYPE="hidden" NAME="PositionID" VALUE="<%= Request("PositionID") %>">
<BR>
<% '********** Start fields ****** %>
<table border="0"> 
 
<% 
Incumbent = RS0("LastName") & ", " & RS0("FirstName")
if len(RS0("FirstName")) = 0 then
	Incumbent = RS0("LastName") 
end if
if len(RS0("LastName")) = 0 then
	Incumbent = RS0("FirstName") 
end if
%>

<tr>
<td width="150" align="right"><font size="2">Current Incumbent: </font></td>
<% if isnull(RS0("LastName")) and isnull(RS0("FirstName")) then %>
	<td>
	<select name="EmployeeID">
	<option value="0">OPEN
	<% Set RS1 = Conn.execute(sql0 & "AND Positions.Title = null ") %>
	<% Do While not RS1.EOF %>
		<option value="<%= RS1("EmployeeID") %>"><%= RS1("LastName") %>,&nbsp;<%= RS1("FirstName") %>
		<% RS1.MoveNext %>
	<% Loop %>
	</select>
	&nbsp;&nbsp;(You can assign employee here)
	</td>
<% else %>
	<td bgcolor="silver"><b><%= Incumbent %></b></td>
<% end if %>
</tr>
 
<tr>
<td width="150" align="right"><font size="2">Position Title: </font></td>
<td><input type="text" name="Title" size="20" value="<%= RS0("Title") %>"></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Reports to: </font></td>
<td>
<select name="BossID">
<% Set RS1 = Conn.execute(sqlPOS & "AND Positions.BossID <> " & Request("PositionID") & " AND Positions.PositionID <> " & Request("PositionID") & " ORDER BY Employees.LastName") %>
<% Do While not RS1.EOF %>
	<% if RS1("PositionID") = RS0("BossID") then %>
		<option value="<%= RS1("PositionID") %>" selected><%= RS1("LastName") %>,&nbsp;<%= RS1("FirstName") %>&nbsp;&nbsp;(<%= RS1("Title") %>)
	<% else %>
		<% Session("IsSubordinate") = False %>
		<% Call GetDescendents(Request("PositionID"),RS1("PositionID")) %>
		<% if Session("IsSubordinate") = false then %>
			<option value="<%= RS1("PositionID") %>"><%= RS1("LastName") %>,&nbsp;<%= RS1("FirstName") %>&nbsp;&nbsp;(<%= RS1("Title") %>)
		<% end if %>
	<% end if %>
	<% RS1.MoveNext %>
<% Loop %>
<% if RS0("BossID") = 0 then %>
	<option value="0" selected>Head of Company
<% else %>
	<option value="0">Head of Company
<% end if %>
</select>
</td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Department: </font></td>
<td>
<select name="Department">
<% Set RS1 = Conn.execute("SELECT * FROM v_Departments ORDER BY Department") %>
<% Do While not RS1.EOF %>
	<% if RS1("Department") = RS0("Department") then %>
		<option value="<%= RS1("Department") %>" selected><%= RS1("Department") %>
	<% else %>
		<option value="<%= RS1("Department") %>"><%= RS1("Department") %>
	<% end if %>
	<% RS1.MoveNext %>
<% Loop %>
</select>
</td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Division: </font></td>
<td>
<select name="Division">
<% Set RS1 = Conn.execute("SELECT * FROM v_Divisions ORDER BY Division") %>
<% Do While not RS1.EOF %>
	<% if RS1("Division") = RS0("Division") then %>
		<option value="<%= RS1("Division") %>" selected><%= RS1("Division") %>
	<% else %>
		<option value="<%= RS1("Division") %>"><%= RS1("Division") %>
	<% end if %>
	<% RS1.MoveNext %>
<% Loop %>
</select>
</td>
</tr>

</table>
</form>

</body>
</html>
<% Conn.Close %>
