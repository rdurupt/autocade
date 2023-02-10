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
	sql = "INSERT INTO v_Departments (Department) VALUES ('"
	sql = sql & funSafeEntry(Request("addValue")) & "')"
	Conn.Execute(sql)
    pg = "valDepartment.asp?msg=Status:+Department+added."
ElseIf subAction = "Update" and len(Request("lst")) > 0 Then
	Set RS1 = Conn.Execute("SELECT * FROM v_Departments ORDER BY Department")
	i = 0
	strArray = split(Request("lst"),",")
	do while not RS1.EOF
		newValue = trim(strArray(i))
		newValue = funSafeEntry(newValue)
		sql = "UPDATE v_Departments SET Department = '" & newValue & "' "
		sql = sql & "WHERE DepartmentID = " & RS1("DepartmentID") 
		Conn.Execute(sql)
		i = i + 1	
		RS1.Movenext
	loop
	pg = "valDepartment.asp?msg=Status:+Department+updated."
ElseIf subAction = "Delete" then
	Conn.Execute("DELETE FROM v_Departments WHERE DepartmentID = " & Request("DepartmentID"))
	pg = "valDepartment.asp?msg=Status:+Department+deleted."
else
	pg = "valDepartment.asp?msg=Status:+Nothing+to+update."
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