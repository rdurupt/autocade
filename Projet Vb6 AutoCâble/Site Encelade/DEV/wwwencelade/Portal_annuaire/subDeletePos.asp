<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="ADOConnect.asp"-->

<%
Session("currPositionID") = Request("PositionID")
Set RS1 = Conn.Execute("SELECT * FROM Positions WHERE BossID = " & Request("PositionID"))
if not RS1.EOF then	
	Set RS0 = Conn.Execute("SELECT * FROM Positions WHERE PositionID = " & Request("PositionID"))
	pg = "lstPosition.asp?msg=Status:+Cannot+delete '" & RS0("Title") & "' because+position+has+subordinates."
else
	Conn.Execute("DELETE FROM Positions WHERE PositionID = " & Request("PositionID"))
	pg = "lstPosition.asp?msg=Status:+Position+deleted."
end if
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