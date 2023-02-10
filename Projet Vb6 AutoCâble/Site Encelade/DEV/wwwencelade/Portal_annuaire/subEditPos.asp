<!-- Copyright 1999 (c) by S. Hurdowar -->
<!--#INCLUDE FILE="ADOConnect.asp"-->
<% 
if Request("EmployeeID") <> 0 then	
	Conn.Execute("UPDATE Positions SET EmployeeID = " & Request("EmployeeID") & " WHERE PositionID = " & Request("PositionID"))
end if

sql = "UPDATE Positions SET "
sql = sql & "Title = '" & replace(Request("Title"),"'","`") & "',"
sql = sql & "Department = '" & Request("Department") & "',"
sql = sql & "Division = '" & Request("Division") & "',"
sql = sql & "BossID = " & Request("BossID") & " "
sql = sql & "WHERE PositionID = " & Request("PositionID")
Conn.Execute(sql)

pg = "readPos.asp?PositionID=" & Request("PositionID") & "&msg=Status:+Position+updated."

Call SetDefault("PositionUpdated","true") 
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
