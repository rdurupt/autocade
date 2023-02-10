<% Response.Expires = 0 %>
<% LoadMe = "onLoad=""document.frm.FirstName.focus()""" %>
<!--#INCLUDE FILE="zHeading.asp"-->
<% Set RS0 = Conn.Execute("SELECT * FROM Users WHERE UserID = " & Request("UserID")) %>
<% Session("currUserID") = Request("UserID") %>

<script language="JavaScript">
function javSubmit() {
    if (document.frm.FirstName.value == "") {
        alert ("FirstName required.")
        document.frm.FirstName.focus()
        return
    }
    if (document.frm.UserName.value == "") {
        alert ("Username required.")
        document.frm.UserName.focus()
        return
    }
    document.frm.submit()
}
    
function javCancel() {
    location.href = "lstUser.asp"
}
</script>

<% '********************** TITLEBAR ***************** %>

<table border="0" width="100%" bgcolor="#004080">
<tr>
<td align="left"><font color="white" size="4"><b>Edit User</b></font></td>
<form name="frm" action="subUser.asp">
<td align="right">
<input type="button" value="Submit" onClick="javSubmit()">&nbsp;
<input type="button" value="Cancel" onClick="javCancel()">&nbsp;
</td>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->
<input type="hidden" name="editMode" value="Edit">
<input type="hidden" name="UserID" value="<%= Request("UserID") %>">

<table border="0">
    
<tr>
<td align="right"><font size=2>FirstName</font></td>
<td><input type="text" name="FirstName" size="20" value="<%= RS0("FirstName") %>"></td>
</tr>
    
<tr>
<td align="right"><font size=2>LastName</font></td>
<td><input type="text" name="LastName" size="20" value="<%= RS0("LastName") %>"></td>
</tr>
    
<tr>
<td align="right"><font size=2>Username</font></td>
<td><input type="text" name="UserName" size="20" value="<%= RS0("Username") %>"></td>
</tr>

<% 
If RS0("AccessLevel") = 3 Then
    sel3 = "selected"
ElseIf RS0("AccessLevel") = 2 Then
    sel2 = "selected"
Else
    sel1 = "selected"
End If 
%>
    
<tr>
<td align="right"><font size=2>Access Level</font></td>
<td>
<select name="AccessLevel">
<option value="1" <%= sel1 %>>Read Only
<option value="2" <%= sel2 %>>Read/Write
<option value="3" <%= sel3 %>>Admin
</select>
</td>
</tr>
   
</table>
    
    
</form>
    
</body>
</html>
<% Conn.Close %>