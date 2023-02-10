<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="zHeading.asp"-->

<script language="JavaScript">
function javDelete(UserID) {
    if (confirm("Delete user?")) {
        location.href = "subUser.asp?editMode=Delete&UserID=" + UserID
    }
}   
function javNew() {
    location.href = "frmNewUser.asp"
}
function javHome() {
    location.href = "Home.asp"
}
</script>

<% '********************** TITLEBAR ***************** %>
<% Set RS1 = Conn.Execute("SELECT Count(*) as [RecCount] FROM Users") %>
<table border="0" width="100%" bgcolor="navy">
<tr>
<td align="left">
<font color="white" size="4"><b>Users</b></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font color="white"><%= RS1("RecCount") %>&nbsp;Records</font>
</td>

<td align="right" width="200">&nbsp;</td>
<form name="frm">
<td align="right">
<input type="button" value="New User" onClick="javNew()">&nbsp;
<input type="button" value="Main Menu" onClick="javHome()">&nbsp;
</td>
</form>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->

<% Set RS0 = Conn.Execute("SELECT * FROM Users ORDER BY UserID") %>
<table bgcolor="#C5C2DC" cellspacing="1" width="100%">
<% If Not RS0.EOF Then %>
    <tr bgcolor="#0080C0">
    <td width="15">&nbsp;</td>
    <td><b><font color="white">ID</font></b></td>
    <td><b><font color="white">First Name</font></b></td>
    <td><b><font color="white">Last Name</font></b></td>
    <td><b><font color="white">Access Type</font></b></td>
    <td><b><font color="white">Username</font></b></td>
    <td><b><font color="white">Create Date</font></b></td>
    </tr>
<% End If %>
    
<% Do While Not RS0.EOF %>
    <% If CStr(RS0("UserID")) = CStr(Session("CurrUserID")) Then %>
    	<tr bgcolor="#C4F8FB">
	<% else %>
    	<tr bgcolor="#FBF9FF">
    <% End If %>
	<% 

    If RS0("AccessLevel") = 1 Then
        AccessLevel = "Read Only"
    ElseIf RS0("AccessLevel") = 2 Then
        AccessLevel = "Read/Write"
    ElseIf RS0("AccessLevel") = 3 Then
        AccessLevel = "Admin"
    Else
        AccessLevel = "Undetermined"
    End If

    %>
	<td width="15"><a href="javascript:javDelete(<%= RS0("UserID") %>)"><font color="Red"><b>x</b></font></td>
	<td><a href="frmEditUser.asp?UserID=<%= RS0("UserID") %>"><%= RS0("UserID") %></a></td>
	<td><a href="frmEditUser.asp?UserID=<%= RS0("UserID") %>"><%= RS0("FirstName") %></a></td>
	<td><a href="frmEditUser.asp?UserID=<%= RS0("UserID") %>"><%= RS0("LastName") %></a></td>
	<td><a href="frmEditUser.asp?UserID=<%= RS0("UserID") %>"><%= AccessLevel %></a></td>
	<td><a href="frmEditUser.asp?UserID=<%= RS0("UserID") %>"><%= RS0("UserName") %></a></td>
	<td><a href="frmEditUser.asp?UserID=<%= RS0("UserID") %>"><%= RS0("CreateDate") %></a></td>
    
	</tr>
    <% RS0.movenext %>
<% Loop %>
    
</table>
<br>
<font size="1">Copyright © 1998 <a href="http://www.mayanetics.com">S. Hurdowar</a>. All rights reserved.
</body>
</html>
    
<% Conn.Close %>
