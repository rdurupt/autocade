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

	url = "frmEditEmp.asp?EmployeeID=" & Request("EmployeeID")
	url = url & "&FirstName=" & FirstName
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
	
	sql = "UPDATE Employees SET "
	sql = sql & "FirstName = '" & FirstName & "',"
	sql = sql & "LastName = '" & LastName & "',"
	sql = sql & "Extension = '" & Extension & "',"
	sql = sql & "EMail = '" & EMail & "',"
	if isDate(HireDate) then
		sql = sql & "HireDate = '" & HireDate & "',"
	end if
	sql = sql & "MobilePager = '" & MobilePager & "' "
	sql = sql & "WHERE EmployeeID = " & Request("EmployeeID") 
	Conn.Execute(sql)
	
	if len(Request("PositionID")) > 0 then
		if cInt(Request("PositionID")) > 0 then
			Conn.Execute("UPDATE Positions SET EmployeeID = 0 WHERE EmployeeID = " & Request("EmployeeID") )
			Conn.Execute("UPDATE Positions SET EmployeeID = " & Request("EmployeeID")  & " WHERE PositionID = " & Request("PositionID") )
			Call SetDefault("PositionUpdated","true") 
		end if
	end if
	pg = "readEmp.asp?EmployeeID=" & Request("EmployeeID") & "&msg=Status:+Employee+information+updated."
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

