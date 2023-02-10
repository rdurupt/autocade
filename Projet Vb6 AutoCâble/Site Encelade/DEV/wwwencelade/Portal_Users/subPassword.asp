<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="ADOConnect.asp"-->

<% 
p1 = replace(Request("p1"), "'", "`")
p2 = replace(Request("p2"), "'", "`")
Set RS1 = Conn.Execute("SELECT * FROM Users WHERE Password = '" & p1 & "' AND UserID = " & Session("employeeUserID"))
If Not RS1.EOF Then
	'Conn.Execute("UPDATE Users SET Password = '" & p2 & "' WHERE UserID = " & Session("employeeUserID"))
    pg = "frmPassword.asp?msg=Demo:+Password+not+changed."
Else
    pg = "frmPassword.asp?msg=Status:+Unable+to+change+password."
End If
%>
<html>
<script language="JavaScript">
function GoThere() {
    location.href = "<%= pg %>"
}
</script>
<body onload="GoThere()">
</body>
</html>
<% Conn.Close %>