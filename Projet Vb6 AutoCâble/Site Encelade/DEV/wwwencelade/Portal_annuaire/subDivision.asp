<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<%
function funSafeEntry(strField)
	strSafe = replace(strField,"'","`")
	strSafe = replace(strSafe,"<","&lt;")
	strSafe = replace(strSafe,">","&gt;")
	funSafeEntry = strSafe
end function

subAction = trim(Request("subAction"))
If subAction = "Cancel" Then
    pg = "lstValidation.asp"
ElseIf subAction = "Add" and len(Request("addValue")) > 0 Then
	sql = "INSERT INTO v_Divisions (Division) VALUES ('"
	sql = sql & funSafeEntry(Request("addValue")) & "')"
	Conn.Execute(sql)
    pg = "valDivision.asp?msg=Status:+Division+added."
ElseIf subAction = "Update" and len(Request("lst")) > 0 Then
	Set RS1 = Conn.Execute("SELECT * FROM v_Divisions ORDER BY Division")
	i = 0
	strArray = split(Request("lst"),",")
	do while not RS1.EOF
		newValue = trim(strArray(i))
		newValue = funSafeEntry(newValue)
		sql = "UPDATE v_Divisions SET Division = '" & newValue & "' "
		sql = sql & "WHERE DivisionID = " & RS1("DivisionID") 
		Conn.Execute(sql)
		i = i + 1	
		RS1.Movenext
	loop
	pg = "valDivision.asp?msg=Status:+Division+updated."
ElseIf subAction = "Delete" then
	Conn.Execute("DELETE FROM v_Divisions WHERE DivisionID = " & Request("DivisionID"))
	pg = "valDivision.asp?msg=Status:+Division+deleted."
else
	pg = "valDivision.asp?msg=Status:+Nothing+to+update."
End If

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