<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<% LoadMe = "onLoad=""document.frm.addValue.focus()""" %>
<!--#INCLUDE FILE="zHeading.asp"-->
<script language="javascript">
function javDelete(ID) {
	if (confirm("Delete department?")) {
		location.href = "subDepartment.asp?subAction=Delete&DepartmentID=" + ID
	} else {
		return
	}
}
</script>

<% '********************** TITLEBAR ***************** %>
<table WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr BGCOLOR="#004080">
<th NOWRAP ALIGN="Left">
<font SIZE="4" COLOR="WHITE">&nbsp;Validation Table: Departments</font>
</th>
<form name="frm" action="subDepartment.asp">
<td align="right">
<input type="submit" name="subAction" value="Update">&nbsp;
<input type="submit" name="subAction" value="Cancel">&nbsp;
</td>

</tr>
</table>
<!--#INCLUDE FILE="zMessage.asp"-->

<br>
<table CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr><td width="30">&nbsp;</td><td BGCOLOR="SILVER"><b>Add New Record:</b></tr>

<tr>
<td width="30">&nbsp;</td>
<td>
<input type="text" name="addValue" SIZE="40" MAXLENGTH="40">
<input type="submit" name="subAction" VALUE=" Add ">
</td>
</tr>

</table>


<br><br>
<table CELLSPACING="1" BORDER="0">
<tr>

<td width="30" align="right">&nbsp;</td>
<td width="25" align="right" BGCOLOR="SILVER">&nbsp;&nbsp;</td>
<td ALIGN="Left" BGCOLOR="SILVER"><font SIZE="-1"><b>Department</b></font></td>

</tr>

<% Set RS1 = Conn.Execute("SELECT * FROM v_Departments ORDER BY Department") %>
<% do while not RS1.EOF %> 
	<tr>
	<td width="30" align="right">&nbsp;</td>

	<td width="25" align="center" BGCOLOR="White"><a href="javascript:javDelete(<%= RS1("DepartmentID") %>)"><img src="delete.gif" border="0" WIDTH="10" HEIGHT="8"></a>&nbsp;</td>
	<td bgcolor="White" nowrap><font SIZE="-1"><input type="text" name="lst" size="40" value="<%= RS1("Department") %>"></font></td>
	</tr>
	<% RS1.MoveNext %>
<% loop %>

</table>
</form>

</body>
</html>
<% Conn.Close %>