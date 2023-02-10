<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="zHeading.asp"-->
<%
public function GetDescendents(d)
	Set RS20 = Conn.Execute(sql0 & "AND BossID = " & d)
	do while not RS20.EOF
		response.write ("<tr><td width='150' align='right'>&nbsp;</td><td>" & vbcrlf)
		response.write (RS20("LastName") & ", " & RS20("FirstName") & "</td><td>(" & RS20("Title") & ")")
		response.write ("</td></tr>" & vbcrlf)
		Call GetDescendents(RS20("PositionID"))
		RS20.movenext
	loop
end function
%>

<% on error resume next %>
<% Session("currPositionID") = Request("PositionID") %>
<% Set RS0 = Conn.Execute(sqlPOS & "AND Positions.PositionID = " & Request("PositionID")) %>

<SCRIPT LANGUAGE="JavaScript">
function javList() {
    location.href = "lstPosition.asp"
}
function javHome() {
    location.href = "Home.asp"
}
</SCRIPT>

<% '********************** TITLEBAR ***************** %>
<table BGCOLOR="#004080" WIDTH="100%" CELLSPACING="0" CELLPADDING="0" BORDER="0">
<tr>
<th NOWRAP ALIGN="Left">
<font SIZE="4" COLOR="WHITE">&nbsp;Position Information</font>
</th>
<form name="frm" action="frmEditPos.asp">

<td align="right">
<% if Session("employeeAccess") > 1 then %>
	<input TYPE="SUBMIT" NAME="subAction" VALUE="Edit">
<% end if %>
<input type="button" value="Position List" onClick="javList()">&nbsp;
<input type="button" value="Main Menu" onClick="javHome()">&nbsp;
</td>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->
<input type="hidden" name="PositionID" value="<%= Request("PositionID") %>">

<BR>
<% '********** Start fields ****** %>
<table border="0"> 

<tr>
<td width="150" align="right"><font size=2><b>Position Title:</b> </font></td>
<td><%= RS0("Title") %></td>
</tr>

<% if len(RS0("LastName")) > 0 or len(RS0("FirstName")) > 0 then %>
	<% Incumbent = RS0("LastName") & ", " & RS0("FirstName") %>
<% else %>
	<% Incumbent = "<font color='red'>OPEN (Edit to assign employee)</font>" %>
<% end if %>

<tr>
<td width="150" align="right"><font size=2><b>Incumbent:</b> </font></td>
<td><%= Incumbent %></td>
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

<tr><td colspan="2">&nbsp;</td></tr>
<% Set RS1 = Conn.Execute("SELECT * FROM Positions WHERE BossID = " & Request("PositionID")) %>
	<% if not RS1.EOF then %>
	<tr>
	<td width="150" align="right">&nbsp;</td>
	<td bgcolor="#004080"><font color="white"><b>Subordinates</b></font></td>
	</tr>
	
	<% Call GetDescendents(Request("PositionID")) %>
<% end if %>
</table>
</form>

</body>
</html>
<% Conn.Close %>
