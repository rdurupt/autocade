<% 
Response.Expires = "0" 
Session("fromPage") = "chtOrganization.asp"

Set Conn = Server.CreateObject("ADODB.Connection")  
Conn.Mode = 3 
Conn.Open Session("ADOEmployee") 

public function GetDefault(fld)
	Set RSDef = Conn.Execute("SELECT " & fld & " FROM dbp_defaultSettings")
	GetDefault = RSDef(fld)
end function

public function SetDefault(fld,val)
	Conn.Execute("UPDATE dbp_defaultSettings SET  " & fld & " = '" & val & "'")
end function

public function pr(strPrint)
	response.write (strPrint & vbcrlf)
end function

public function HideDescendents(d,l)
	Set RS20 = Conn.Execute(sql0 & "AND BossID = " & d & " ORDER BY Employees.LastName")

	do while not RS20.EOF
		Conn.Execute("INSERT INTO Sys_HideItems VALUES('x" & RS20("PositionID") & l & "x')")
		Call HideDescendents(RS20("PositionID"), l)
		RS20.movenext
	loop
end function

public function HideChildren(ID,lv)
	Set RS11 = Conn.Execute(sql0 & "AND BossID = " & ID & " ORDER BY Employees.LastName")
	Set RS13 = Conn.Execute("SELECT Count(*) as RecCount FROM Positions WHERE EmployeeID <> 0 and EmployeeID <> null AND BossID = " & ID )
	z = 0
	do while not RS11.EOF
		z = z + 1	

		if z = RS13("RecCount") then
			Conn.Execute("INSERT INTO Sys_LastItems VALUES(" & RS11("PositionID") & ")")
			Call HideDescendents(RS11("PositionID"), lv + 1)
		end if
	
		Call HideChildren(RS11("PositionID"), lv + 1)
		RS11.movenext
	loop
end function

public function GetChildren(PosID,lvl)
	Set RS1 = Conn.Execute(sql0 & "AND BossID = " & PosID & " ORDER BY Employees.LastName")

	do while not RS1.EOF
		Set RS5 = Conn.Execute("SELECT * FROM Sys_LastItems WHERE PositionID = " & RS1("PositionID"))
		if not RS5.EOF then	
			strEnd = "<img src='line1.gif' border='0'>"
		else
			strEnd = "<img src='line2.gif' border='0'>"
		end if
		
		for j = 0 to lvl
			Set RS6 = Conn.Execute("SELECT * FROM Sys_HideItems WHERE HideItem = 'x" & RS1("PositionID") & j & "x'")
			if j = 0 or not RS6.EOF then
				response.write ("<img src='line0.gif' border='0'>")
			else
				response.write ("<img src='line3.gif' border='0'>")
			end if

		next 
		
		strAnchor = "<a href='readEmp.asp?EmployeeID=" & RS1("EmployeeID") & "'>"
		response.write (strEnd & "<img src='line4.gif' border='0'>" & strAnchor & RS1("LastName") & ", " & RS1("FirstName") & ", " & RS1("Title") & "</a><br>")
		Call GetChildren(RS1("PositionID"), lvl+1)
		RS1.MoveNext
	loop 
end function

sql0 = "SELECT Employees.*, Positions.PositionID, Positions.BossID, Positions.Title, Positions.Department, Positions.Division FROM Employees LEFT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Employees.EmployeeID <> null " 
sql00 = "SELECT Count(*) as RecCount FROM Employees LEFT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Employees.EmployeeID <> null " 

sqlPOS = "SELECT Employees.*, Employees.EmployeeID , Positions.* FROM Employees RIGHT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Positions.PositionID <> null " 
sqlPOSCount = "SELECT Count(*) as RecCount FROM Employees RIGHT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Positions.PositionID <> null " 


pr ("<html>")
pr ("<head>")
pr ("<title>Human Resource Management</title>")
pr ("</head>")
pr ("<basefont FACE='Arial, Helvetica, sans-serif'>")
pr ("<link REL='stylesheet' HREF='Style.css'>")
pr ("<SCRIPT LANGUAGE='JavaScript'>")
pr ("function javHome() {")
pr ("    location.href = 'Home.asp'")
pr ("}")
pr ("</SCRIPT>")
pr ("<body BGCOLOR='White'>")
pr ("<img src='corner.gif' border='0'><br clear=all>")
'********************** HEADING ***************** 
pr ("<TABLE>")
pr ("<TR>")
pr ("<TD valign=middle align='left' NOWRAP>")
pr ("&nbsp;")
pr ("<font size=+3 face='Times New Roman' color='#004080'><STRONG>R</STRONG></font><font size=+3 color='#606060'><em>esource</em></font>&nbsp;")
pr ("<font size=+3 face='Times New Roman' color='#004080'><STRONG>M</STRONG></font><font size=+3 color='#606060'><em>anagement</em></font>&nbsp;")
pr ("</TD>")
pr ("</TR>")
pr ("</TABLE>")

'********************** TITLEBAR ***************** 
pr ("<table BGCOLOR='#004080' WIDTH='100%' BORDER='0'>")
pr ("<tr>")
pr ("<th NOWRAP ALIGN='Left'>")
pr ("<font SIZE='4' COLOR='WHITE'>&nbsp;Organizational Chart</font>")
pr ("</th>")
pr ("<form name='frm' action='frmEditEmp.asp'>")
pr ("<td align='right'><input type='button' value='Main Menu' onClick='javHome()'>&nbsp;</td>")
pr ("</tr>")
pr ("</table>")
pr ("</form>")

'********************** MESSAGE ***************** 
if len(Request("msg")) > 0 then 
	pr ("<table bgcolor='#F3F4BD' width='100%' border='0'><tr><td><font size='-1' color='maroon'>&nbsp;&nbsp;" & Request("msg") & "</font></td></tr></table>")
end if  

'**********************  Start Main Program here ***************

if GetDefault("PositionUpdated") = "true" or GetDefault("PositionUpdated") = "" then
	Conn.Execute("DELETE FROM Sys_HideItems")
	Conn.Execute("DELETE FROM Sys_LastItems")
	Set RS2 = Conn.Execute(sql0 & "AND BossID = 0")
	do while not RS2.EOF 
		Call HideChildren(RS2("PositionID"),0)
		RS2.MoveNext
	loop 
	Call SetDefault("PositionUpdated","false") 
end if

Set RS2 = Conn.Execute(sql0 & "AND BossID = 0 ORDER BY Employees.LastName")
do while not RS2.EOF 
	strAnchor = "<a href='readEmp.asp?EmployeeID=" & RS2("EmployeeID") & "'>"
	response.write ("<img src='line0.gif' border='0'>" & strAnchor & RS2("LastName") & ", " & RS2("FirstName") & ", " & RS2("Title") & "</a><br>")
	Call GetChildren(RS2("PositionID"),0)
	RS2.MoveNext
loop 
 
pr ("</body>")
pr ("</html>")
Conn.Close 
%>