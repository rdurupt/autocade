<% Response.Expires = "0" %>
<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="zHeading.asp"-->
<% 
fromPage = "home.asp"
if len(Session("fromPage")) > 0 then
	fromPage = Session("fromPage")
end if 
%>
<SCRIPT LANGUAGE="JavaScript">
function javCancel() {
    location.href = "<%= fromPage %>"
}
function javHome() {
    location.href = "Home.asp"
}
</SCRIPT>

<% '********************** TITLEBAR ***************** %>
<table WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr BGCOLOR="#004080">
<th NOWRAP ALIGN="Left">
<font SIZE="4" COLOR="WHITE">&nbsp;Query</font>
</th>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->

<br>
<form name="frm" ACTION="subQuery.Asp">
<table CELLSPACING="0" CELLPADDING="0" BORDER="0">

<tr>
<td width="25">&nbsp;</td>
<td BGCOLOR="#BFF9C4" NOWRAP><b>Employee Name</b>&nbsp;</td>
<td BGCOLOR="#BFF9C4">
<input type="text" size="30" name="EmployeeName" value="<%= Session("EmployeeName") %>">
</td>
<td BGCOLOR="#BFF9C4">&nbsp;&nbsp;</td>
</tr>

<tr><td colspan="4">&nbsp;</td></tr>

<tr>
<td width="25">&nbsp;</td>
<td BGCOLOR="#EB9B81" NOWRAP><b>Position Title</b>&nbsp;</td>
<td BGCOLOR="#EB9B81">
<input type="text" size="30" name="Title" value="<%= Session("Title") %>">
</td>
<td BGCOLOR="#EB9B81">&nbsp;&nbsp;</td>
</tr>

<tr><td colspan="4">&nbsp;</td></tr>

<tr>
<td width="25">&nbsp;</td>
<td BGCOLOR="#97E4F2" NOWRAP><b>Department</b>&nbsp;</td>
<td BGCOLOR="#97E4F2">
<select name="Department">
<option value="">
<% Set RS1 = Conn.execute("SELECT * FROM v_Departments ORDER BY Department") %>
<% Do While not RS1.EOF %>
	<% if RS1("Department") = Session("Department") then %>
		<option value="<%= RS1("Department") %>" selected><%= RS1("Department") %>
	<% else %>
		<option value="<%= RS1("Department") %>"><%= RS1("Department") %>
	<% end if %>
	<% RS1.MoveNext %>
<% Loop %>
</select>
</td>
<td BGCOLOR="#97E4F2">&nbsp;&nbsp;</td>
</tr>


<tr><td colspan="4">&nbsp;</td></tr>

<tr>
<td width="25">&nbsp;</td>
<td BGCOLOR="#DFE873" NOWRAP><b>Division</b>&nbsp;</td>
<td BGCOLOR="#DFE873">
<select name="Division">
<option value="">
<% Set RS1 = Conn.execute("SELECT * FROM v_Divisions ORDER BY Division") %>
<% Do While not RS1.EOF %>
	<% if RS1("Division") = Session("Division") then %>
		<option value="<%= RS1("Division") %>" selected><%= RS1("Division") %>
	<% else %>
		<option value="<%= RS1("Division") %>"><%= RS1("Division") %>
	<% end if %>
	<% RS1.MoveNext %>
<% Loop %>
</select>
</td>
<td BGCOLOR="#DFE873">&nbsp;&nbsp;</td>
</tr>

<tr><td colspan="4">&nbsp;</td></tr>

<tr>
<td width="25">&nbsp;</td>
<td>&nbsp;</td>
<td colspan="3"><input type="submit" name="subAction" value="Search">
&nbsp;<input type="button" value="Cancel" onClick="javCancel()"></td>
</tr>

</form>


</table>


<br>

</body>
</html>

<% Conn.Close %>