<%
Response.Expires = "0" 
Set Conn = Server.CreateObject("ADODB.Connection")  
Conn.Mode = 3 
Conn.Open Session("ADOEmployee") 


public function pr(strPrint)
	response.write (strPrint & vbcrlf)
end function


Session("fromPage") = "lstPosition.asp"
qrySet = false

sql0 = "SELECT Employees.*, Positions.PositionID, Positions.BossID, Positions.Title, Positions.Department, Positions.Division FROM Employees LEFT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Employees.EmployeeID <> null " 
sql00 = "SELECT Count(*) as RecCount FROM Employees LEFT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Employees.EmployeeID <> null " 

sqlPOS = "SELECT Employees.*, Employees.EmployeeID , Positions.* FROM Employees RIGHT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Positions.PositionID <> null " 
sqlPOSCount = "SELECT Count(*) as RecCount FROM Employees RIGHT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Positions.PositionID <> null " 

strSQL= ""
criter = "Record Set:&nbsp;"
if len(Session("EmployeeName")) > 0 then
	criter = criter & "Employee name like '" & Session("EmployeeName") & "';&nbsp;&nbsp;"
	strSQL= strSQL & "AND Employees.FirstName LIKE '%" & Session("EmployeeName") & "%' OR Employees.LastName  LIKE '%" & Session("EmployeeName") & "%' "
	qrySet = true
end if

if len(Session("Title")) > 0 then
	criter = criter & "Title of " & Session("Title") & ";&nbsp;&nbsp;"
	strSQL= strSQL & "AND Positions.Title LIKE '%" & Session("Title") & "%' "
	qrySet = true
end if

if len(Session("Department")) > 0 then
	criter = criter & "In department " & Session("Department") & ";&nbsp;&nbsp;"
	strSQL= strSQL & "AND Positions.Department = '" & Session("Department") & "' "
	qrySet = true
end if

if len(Session("Division")) > 0 then
	criter = criter & "In division " & Session("Division") & ";&nbsp;&nbsp;"
	strSQL = strSQL & "AND Positions.Division = '" & Session("Division") & "'"
	qrySet = true
end if

sqlCount = sqlPOSCount & strSQL

'********* Order by
If Request("positionSortBy") <> "" Then
    Session("positionSortBy") = Request("positionSortBy")
    strSQL = strSQL & "ORDER BY " & Request("positionSortBy")
ElseIf Session("positionSortBy") <> "" Then
    strSQL = strSQL & "ORDER BY " & Session("positionSortBy")
Else
    Session("positionSortBy") = "Title"
    strSQL = strSQL & "ORDER BY Positions.Title"
End If

strSQL = sqlPOS & strSQL

Set RS0 = Conn.Execute(strSQL)

'*********** Get record count 
Set RS11 = Conn.Execute(sqlCount) 
strStatus = criter & "&nbsp;&nbsp;&nbsp;&nbsp;" & RS11("RecCount") & " Records" 

pr ("<html>")
pr ("<head>")
pr ("<title>Human Resource Management</title>")
pr ("</head>")

pr ("<SCRIPT LANGUAGE='JavaScript'>")
pr ("function javHome() {")
pr ("    location.href = 'Home.asp'")
pr ("}")

pr ("function javQuery() {")
pr ("    location.href = 'Query.asp'")
pr ("}")

pr ("function javReset() {")
pr ("    location.href = 'subQuery.asp?qryReset=true'")
pr ("}")
pr ("function javDelete(ID) {")
pr ("    if (confirm('Delete position?')) {")
pr ("       location.href = 'subDeletePos.asp?PositionID=' + ID")
pr ("    }")
pr ("} ")

pr ("</SCRIPT>")

pr ("<link REL='stylesheet' HREF='Style.css'>")

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
pr ("<font SIZE='4' COLOR='WHITE'>&nbsp;Position List</font>")
pr ("</th>")
pr ("<form name='frm' action='frmEditEmp.asp'>")
pr ("<td align='right'><input type='button' value='Query' onClick='javQuery()'>&nbsp;")
if qrySet = true then 
	pr ("<input type='button' value='Reset Query' onClick='javReset()'>&nbsp;")
end if 
pr ("<input type='button' value='Main Menu' onClick='javHome()'>&nbsp;</td>")
pr ("</tr>")
pr ("</table>")


'********************** STATUS ***************** 
pr ("<table bgcolor='#F3F4BD' width='100%' border='0'><tr><td><font size='-1' color='maroon'>&nbsp;&nbsp;" & strStatus & "</font></td></tr></table>")

'********************** MESSAGE ***************** 
if len(Request("msg")) > 0 then 
	pr ("<table bgcolor='#F3F4BD' width='100%' border='0'><tr><td><font size='-1' color='maroon'>&nbsp;&nbsp;" & Request("msg") & "</font></td></tr></table>")
end if  
 

'********************** SORT HEADERS ***************** 
pr ("<table bgcolor='#C0BCCB' width='100%' cellspacing='1'>")
If Not RS0.EOF Then 
    pr ("<tr bgcolor='#FC8D7A'>")
    pr ("<td width='15'><b>&nbsp;</b></td>")
    
    If InStr(Session("positionSortBy"), "Title") Then 
        If InStr(Session("positionSortBy"), "Desc") Then 
            pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Title'><img src='SortDesc.gif' border='0'>Title</a></b></td>")
        Else 
            pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Title+Desc'><img src='SortAss.gif' border='0'>Title</a></b></td>")
        End If 
    Else 
        pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Title'>Title</a></b></td>")
    End If 
	
    If InStr(Session("positionSortBy"), "LastName") Then 
        If InStr(Session("positionSortBy"), "Desc") Then 
            pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Employees.LastName'><img src='SortDesc.gif' border='0'>Incumbent</a></b></td>")
        Else 
            pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Employees.LastName+Desc'><img src='SortAss.gif' border='0'>Incumbent</a></b></td>")
        End If 
    Else 
        pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Employees.LastName'>Incumbent</a></b></td>")
    End If 
	
    pr ("<td nowrap><b>Reports To</b></td>")
	
    If InStr(Session("positionSortBy"), "Department") Then 
        If InStr(Session("positionSortBy"), "Desc") Then 
            pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Department'><img src='SortDesc.gif' border='0'>Department</a></b></td>")
        Else 
            pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Department+Desc'><img src='SortAss.gif' border='0'>Department</a></b></td>")
        End If 
    Else 
        pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Department'>Department</a></b></td>")
    End If  
	
    If InStr(Session("positionSortBy"), "Division") Then 
        If InStr(Session("positionSortBy"), "Desc") Then 
            pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Division'><img src='SortDesc.gif' border='0'>Division</a></b></td>")
        Else 
            pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Division+Desc'><img src='SortAss.gif' border='0'>Division</a></b></td>")
        End If 
    Else 
        pr ("<td nowrap><b><a href='lstPosition.asp?positionSortBy=Positions.Division'>Division</a></b></td>")
    End If  
 
    pr ("</tr>")
End If 

Do While Not RS0.EOF 
on error resume next 
	if len(RS0("LastName")) > 0 then 
	    If CStr(RS0("PositionID")) = Session("currPositionID") Then 
	        pr ("<tr bgcolor='#C1E8F7'>")
	    Else 
	        pr ("<tr bgcolor='#F5F7FE'>")
	    End If 
		Incumbent = RS0("LastName") & ", " &  RS0("FirstName") 
	else 
		Incumbent = "<font color='red'>OPEN</font>" 
		pr ("<tr bgcolor='#F5F7FE'>")
	end if
	
	ReportsTo = ""
	strSQL1 = "SELECT Employees.FirstName, Employees.LastName FROM Employees, Positions WHERE Positions.PositionID = " & RS0("BossID") & " AND Employees.EmployeeID = Positions.EmployeeID"
	Set RS21 = Conn.Execute(strSQL1) 
	if not RS21.EOF then
		ReportsTo = RS21("LastName") & ", " & RS21("FirstName")
	end if
	
	If Session("employeeAccess") > 1 then 
	    pr ("<td width='20' align='center'><a href='javascript:javDelete(" & RS0("PositionID") & ")'><img src='delete.gif' border='0'></a>&nbsp;</td>")
	else 
		pr ("<td width='20' align='center'>&nbsp;</td>")
	end if 
    pr ("<td nowrap><a href='readPos.asp?PositionID=" & RS0("PositionID") & "'>" & RS0("Title") & "</a>&nbsp;</td>")
	pr ("<td nowrap><a href='readPos.asp?PositionID=" & RS0("PositionID") & "'>" & Incumbent & "</a>&nbsp;</td>")
    pr ("<td nowrap><a href='readPos.asp?PositionID=" & RS0("PositionID") & "'>" & ReportsTo & "</a>&nbsp;</td>")
	pr ("<td nowrap><a href='readPos.asp?PositionID=" & RS0("PositionID") & "'>" & RS0("Department") & "</a>&nbsp;</td>")
	pr ("<td nowrap><a href='readPos.asp?PositionID=" & RS0("PositionID") & "'>" & RS0("Division") & "</a>&nbsp;</td>")
	pr ("</tr>")
    RS0.Movenext 
Loop 

pr ("</table>")
pr ("</form>")
pr ("</body>")
pr ("</html>")

Conn.Close 

%>
