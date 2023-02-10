<% Response.Expires = 0 %>
<% LoadMe = "onLoad=""document.frm.FirstName.focus()""" %>
<!--#INCLUDE FILE="zHeading.asp"-->

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
<td align="left"><font color="white" size="4"><b>New User</b></font></td>
<form name="frm" action="subUser.asp">
<td align="right">
<input type="button" value="Submit" onClick="javSubmit()">&nbsp;
<input type="button" value="Cancel" onClick="javCancel()">&nbsp;
</td>
</tr>
</table>

<!--#INCLUDE FILE="zMessage.asp"-->
<input type="hidden" name="editMode" value="New">

<table border="0">
    
<tr>
<td align="right"><font size=2>FirstName</font></td>
<td><input type="text" name="FirstName" size="20"></td>
</tr>
    
<tr>
<td align="right"><font size=2>LastName</font></td>
<td><input type="text" name="LastName" size="20"></td>
</tr>
    
<tr>
<td align="right"><font size=2>Username</font></td>
<td><input type="text" name="UserName" size="20"></td>
</tr>

<tr>
<td align="right"><font size=2>Password</font></td>
<td><input type="Password" name="Password" size="20"></td>
</tr>

<tr>
<td align="right"><font size=2>Access Level</font></td>
<td>
<select name="AccessLevel">
<option value="1">Read Only
<option value="2">Read/Write
<option value="3">Admin
</select>
</td>
</tr>
   
</table>
    
</form>
    
</body>
</html>
<% Conn.Close %>