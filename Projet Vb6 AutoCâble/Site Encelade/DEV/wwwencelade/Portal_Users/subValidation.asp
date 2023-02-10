<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<%
subAction = trim(Request("subAction"))
If subAction = "Cancel" Then
    pg = "lstValidation.asp"
ElseIf subAction = "Add" and len(Request("addValue")) > 0 Then
	sql = "INSERT INTO " & Request("tbl") & " (" & Request("fld") & ") VALUES ('"
	sql = sql & Request("addValue") & "')"
	Conn.Execute(sql)
    pg = "frmValidation.asp?msg=Status:+Record+added.&tbl=" & Request("tbl") & "&fld=" & Request("fld")
ElseIf subAction = "Update" and len(Request("lst")) > 0 Then
	Set RS1 = Conn.Execute("SELECT * FROM " & Request("tbl") & " ORDER BY " & Request("fld"))
	i = 0
	strArray = split(Request("lst"),",")
	strID = Request("fld") & "ID"
	do while not RS1.EOF
		sql = "UPDATE " & Request("tbl") & " SET " & Request("fld") & " = '" & trim(strArray(i)) & "' "
		sql = sql & "WHERE " & strID & " = " & RS1(strID) 
		Conn.Execute(sql)
		i = i + 1	
		RS1.Movenext
	loop
	pg = "frmValidation.asp?msg=Status:+Validation+table+updated.&tbl=" & Request("tbl") & "&fld=" & Request("fld")
else
	pg = "frmValidation.asp?msg=Status:+Nothing+to+update.&tbl=" & Request("tbl") & "&fld=" & Request("fld")
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