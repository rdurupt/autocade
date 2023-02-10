<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<%
if Request("qryReset") = "true" then
	Session("EmployeeName") = ""
	Session("Title") = ""
	Session("Department") = ""
	Session("Division") = ""
else
	Session("EmployeeName") = Request("EmployeeName")
	Session("Title") = Request("Title")
	Session("Department") = Request("Department")
	Session("Division") = Request("Division")
end if
fromPage = "lstEmployee.asp"
if len(Session("fromPage")) > 0 then
	fromPage = Session("fromPage")
end if
%>


<html>
<script language="JavaScript">
function GoThere() {
	location.href = "<%= fromPage %>"
}
</script>
<body onLoad="GoThere()">
</body>
</html>