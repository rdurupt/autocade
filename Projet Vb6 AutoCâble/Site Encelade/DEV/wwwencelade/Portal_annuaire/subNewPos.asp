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

sql = "INSERT INTO Positions (Title,Department,Division,BossID) "
sql = sql & "VALUES ("
sql = sql & "'" & funSafeEntry(Request("Title")) & "',"
sql = sql & "'" & Request("Department") & "',"
sql = sql & "'" & Request("Division") & "',"
sql = sql & "" & Request("BossID") & ")"
Conn.Execute(sql)

Set RS1 = Conn.Execute("SELECT Max(PositionID) as NewID FROM Positions")
Session("currPositionID") = RS1("NewID")

if Session("fromPage") = "frmNewEmp.asp" then
	FirstName = replace(Request("FirstName"), " ","+")
	LastName = replace(Request("LastName"), " ","+")
	Extension = replace(Request("Extension"), " ","+")
	EMail = replace(Request("EMail"), " ","+")
	MobilePager = replace(Request("MobilePager"), " ","+")
	HireDate = replace(Request("HireDate"), " ","+")
	
	url = "frmNewEmp.asp?FirstName=" & FirstName
	url = url & "&LastName=" & LastName
	url = url & "&Extension=" & Extension
	url = url & "&Email=" & Email
	url = url & "&MobilePager=" & MobilePager
	url = url & "&PositionID=" & Request("PositionID")
	url = url & "&HireDate=" & HireDate
	url = url & "&msg=Status:+Position+added." 
	pg = url
else
	pg = "Home.asp?msg=Status:+Position+added."
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

