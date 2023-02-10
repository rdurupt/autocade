<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="zHeading.asp"-->
<% on error resume next %>
<% Session("currEmployeeID") = Request("EmployeeID") %>
<% Set RS0 = Conn.Execute(sql0 & "AND Employees.EmployeeID = " & Request("EmployeeID")) %>

<% 
fromPage = "Employee.asp?mode=lstEmployee" 
if len(Session("fromPage")) > 0 then
	fromPage = Session("fromPage")
end if
strButton = "Employee List"
if fromPage = "Employee.asp?mode=chtOrganization" then
	strButton = "Organization Chart"
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
<table BGCOLOR="#004080" WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr>
<th NOWRAP ALIGN="Left">
<font SIZE="4" COLOR="WHITE">&nbsp;Employee Information</font>
</th>
<form name="frm" action="frmEditEmp.asp">
<td align="right">
<% if Session("employeeAccess") > 1 then %>
	<input TYPE="SUBMIT" NAME="subAction" VALUE="Edit">
<% end if %>
<input type="button" value="<%= strButton %>" onClick="javCancel()">&nbsp;
<input type="button" value="Main Menu" onClick="javHome()">&nbsp;
</td>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->
<input type="hidden" name="EmployeeID" value="<%= Request("EmployeeID") %>">

<BR>
<% '********** Start fields ****** %>
<table border="0"> 

<tr>
<td width="150" align="right"><font size=2><b>Employee:</b> </font></td>
<td><%= RS0("LastName") %>,&nbsp;<%= RS0("FirstName") %></td>
</tr>

<% if len(RS0("Title")) > 0 then %>
	<% Title = RS0("Title") %>
<% else %>
	<% Title = "<font color='red'>NOT ASSIGNED (Edit to assign position)</font>" %>
<% end if %>
<tr>
<td width="150" align="right"><font size=2><b>Position Title:</b> </font></td>
<td><%= Title %></td>
</tr>

<tr>
<td width="150" align="right"><font size=2><b>Reports to:</b> </font></td>
<% Set RS1 = Conn.Execute(sql0 & "AND Positions.PositionID = " & RS0("BossID")) %>
<% 
Boss = "n/a"
if not RS1.EOF then
	Boss = RS1("LastName") & ", " & RS1("FirstName")
end if 
%> 
<td><%= Boss %></td>
</tr>

<tr>
<td width="150" align="right"><font size=2><b>Department:</b> </font></td>
<td><%= RS0("Department") %></td>
</tr>

<tr>
<td width="150" align="right"><font size=2><b>Division:</b> </font></td>
<td><%= RS0("Division") %></td>
</tr>

<tr>
<td width="150" align="right"><font size=2><b>Extension:</b> </font></td>
<td><%= RS0("Extension") %></td>
</tr>

<tr>
<td width="150" align="right"><font size=2><b>Email:</b> </font></td>
<td><%= RS0("Email") %></td>
</tr>

<tr>
<td width="150" align="right"><font size=2><b>MobilePager:</b> </font></td>
<td><%= RS0("MobilePager") %></td>
</tr>

</table>
</form>

</body>
</html>
<% Conn.Close %>
