<% Response.Expires = 0 %>
<% LoadMe = "onLoad=""document.frm.Title.focus()""" %>
<!--#INCLUDE FILE="zHeading.asp"-->
<script language="javascript">

<% 
fromPage = "Home.asp"
if len(Session("fromPage")) > 0 then
	fromPage = Session("fromPage")
end if 
%>

function javCancel() {
	document.frm.action = "<%= fromPage %>"
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
<font SIZE="4" COLOR="WHITE">&nbsp;New Position</font>
</th>
<form name="frm" action="subNewPos.asp">
<td align="right">
<input type="button" onClick="javSubmit()" value="Submit">
<input type="button" onClick="javCancel()" value="Cancel">&nbsp;
</td>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->
<input TYPE="hidden" NAME="fromPage" VALUE="frmNewEmp.asp">
<input TYPE="hidden" NAME="FirstName" VALUE="<%= Request("FirstName") %>">
<input TYPE="hidden" NAME="LastName" VALUE="<%= Request("LastName") %>">
<input TYPE="hidden" NAME="Extension" VALUE="<%= Request("Extension") %>">
<input TYPE="hidden" NAME="Email" VALUE="<%= Request("Email") %>">
<input TYPE="hidden" NAME="MobilePager" VALUE="<%= Request("MobilePager") %>">
<input TYPE="hidden" NAME="HireDate" VALUE="<%= Request("HireDate") %>">
<input TYPE="hidden" NAME="PositionID" VALUE="<%= Request("PositionID") %>">
<BR>
<% '********** Start fields ****** %>
<table border="0"> 

<tr>
<td width="150" align="right"><font size="2">Position Title: </font></td>
<td><input type="text" name="Title" size="20" value=""></td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Reports to: </font></td>
<td>
<select name="BossID">
<% Set RS1 = Conn.execute(sqlPOS) %>
<% Do While not RS1.EOF %>
	<option value="<%= RS1("BossID") %>"><%= RS1("LastName") %>,&nbsp;<%= RS1("FirstName") %>&nbsp;&nbsp;(<%= RS1("Title") %>)
	<% RS1.MoveNext %>
<% Loop %>
</select>
</td>
</tr>

<tr>
<td width="150" align="right"><font size=2>Department: </font></td>
<td>
<select name="Department">
<% Set RS1 = Conn.execute("SELECT * FROM v_Departments ORDER BY Department") %>
<% Do While not RS1.EOF %>
	<option value="<%= RS1("Department") %>"><%= RS1("Department") %>
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
	<option value="<%= RS1("Division") %>"><%= RS1("Division") %>
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

