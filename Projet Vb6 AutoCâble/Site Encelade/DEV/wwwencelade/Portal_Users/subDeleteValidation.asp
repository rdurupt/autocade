<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="ADOConnect.asp"-->

<%
Conn.Execute("DELETE FROM " & Request("tbl") & " WHERE " & Request("key") & " = " & Request("ID"))
pg = "frmValidation.asp?msg=Status:+Validation+record+deleted.&tbl=" & Request("tbl") & "&fld=" & Request("fld")

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