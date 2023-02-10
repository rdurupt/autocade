<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="ADOConnect.asp"-->

<%
Conn.Execute("DELETE FROM Employees WHERE EmployeeID = " & Request("EmployeeID"))
Conn.Execute("UPDATE Positions SET EmployeeID = 0 WHERE EmployeeID = " & Request("EmployeeID"))
Call SetDefault("PositionUpdated","true") 
pg = "lstEmployee.asp?msg=Status:+Employee+deleted."

Conn.Close
%>

<html>
<script language="JavaScript">
function GoThere() {
	location.href = "<%= pg %>"
}
</script>
<body onLoad="GoThere()">
</body>
</html>