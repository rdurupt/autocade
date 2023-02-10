
<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!-- Copyright 1999 (c) by S. Hurdowar -->
<!--#INCLUDE FILE="ADOConnect.asp"-->
<% 
function funSafeEntry(strField)
	strSafe = replace(strField,"'","`")
	strSafe = replace(strSafe,"<","&lt;")
	strSafe = replace(strSafe,">","&gt;")
	funSafeEntry = strSafe
end function

HireDate = trim(Request("HireDate"))

if not isdate(HireDate) AND len(HireDate) <> 0 then
	FirstName = replace(Request("FirstName"), " ","+")
	LastName = replace(Request("LastName"), " ","+")
	Extension = replace(Request("Extension"), " ","+")
	EMail = replace(Request("EMail"), " ","+")
	MobilePager = replace(Request("MobilePager"), " ","+")

	url = "frmNewEmp.asp?FirstName=" & FirstName
	url = url & "&LastName=" & LastName
	url = url & "&Extension=" & Extension
	url = url & "&Email=" & Email
	url = url & "&MobilePager=" & MobilePager
	url = url & "&PositionID=" & Request("PositionID")
	url = url & "&HireDate=" & HireDate
	url = url & "&msg=Status:+Invalid+hire+date."
	pg = url
else
	FirstName = funSafeEntry(Request("FirstName"))
	LastName = funSafeEntry(Request("LastName"))
	Extension = funSafeEntry(Request("Extension"))
	EMail = funSafeEntry(Request("EMail"))
	MobilePager = funSafeEntry(Request("MobilePager"))

	sql = "INSERT INTO Employees (FirstName,LastName,Extension,EMail,MobilePager) "
	if isDate(HireDate) then
		sql = "INSERT INTO Employees (FirstName,LastName,Extension,EMail,HireDate,MobilePager) "
	end if
	
	sql = sql & "VALUES ("
	sql = sql & "'" & FirstName & "',"
	sql = sql & "'" & LastName & "',"
	sql = sql & "'" & Extension & "',"
	sql = sql & "'" & EMail & "',"
	if isDate(HireDate) then
		sql = sql & "'" & HireDate & "',"
	end if
	sql = sql & "'" & MobilePager & "')"

	Conn.Execute(sql)
	
	Set RS1 = Conn.Execute("SELECT Max(EmployeeID) as NewID FROM Employees")
	Session("currEmployeeID") = RS1("NewID")
	
	if Request("PositionID") > 0 then
		Conn.Execute("UPDATE Positions SET EmployeeID = " & RS1("NewID")  & " WHERE PositionID = " & Request("PositionID") )
		Call SetDefault("PositionUpdated","true") 
	end if
	
	pg = "Home.asp?msg=Status:+Employee+added."
end if

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

