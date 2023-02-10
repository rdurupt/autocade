<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="zHeading.asp"-->

<% '********************** TITLEBAR ***************** %>
<table WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr BGCOLOR="#004080">
<th NOWRAP ALIGN="Left">
<font SIZE="4" COLOR="WHITE">&nbsp;Validation Table: <%= Request("fld") %></font>
</th>
<form name="frm" action="subValidation.asp">
<td align="right">
<input type="submit" name="subAction" value="Update">&nbsp;
<input type="submit" name="subAction" value="Cancel">&nbsp;
</td>

</tr>
</table>

<input type="hidden" name="tbl" value="<%= Request("tbl") %>">
<input type="hidden" name="fld" value="<%= Request("fld") %>">
<input type="hidden" name="key" value="<%= Request("key") %>">
<!--#INCLUDE FILE="zMessage.asp"-->
<br>

<table CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr><td width="30">&nbsp;</td><td BGCOLOR="SILVER"><b>Add New Record:</b></tr>

<tr>
<td width="30">&nbsp;</td>
<td>
<input TYPE="TEXT" NAME="addValue" SIZE="40" MAXLENGTH="40">
<input TYPE="SUBMIT" NAME="subAction" VALUE=" Add ">
</td>
</tr>

</table>


<br><br>
<table CELLSPACING="1" BORDER="0">
<tr>

<td width="30" align="right">&nbsp;</td>
<td width="25" align="right" BGCOLOR="SILVER">&nbsp;&nbsp;</td>
<td ALIGN="Left" BGCOLOR="SILVER"><font SIZE="-1"><b><%= Request("fld") %></b></font></td>

</tr>

<% Set RS1 = Conn.Execute("SELECT * FROM " & Request("tbl") & " ORDER BY " & Request("fld")) %>
<% thisID = Request("fld") & "ID" %>
<% do while not RS1.EOF %> 
	<tr>
	<td width="30" align="right">&nbsp;</td>

	<td width="25" align="center" BGCOLOR="White"><a href="subDeleteValidation.asp?tbl=<%= Request("tbl") %>&amp;key=<%= thisID %>&amp;ID=<%= RS1(thisID) %>&amp;fld=<%= Request("fld") %>"><img src="Delete.gif" border="0" WIDTH="10" HEIGHT="8"></a></td>
	<td BGCOLOR="White" NOWRAP><font SIZE="-1"><input type="text" name="lst" size="40" value="<%= RS1(Request("fld")) %>"></font></td>
	</tr>
	<% RS1.MoveNext %>
<% loop %>

</table>
</form>

</body>
</html>
<% Conn.Close %>