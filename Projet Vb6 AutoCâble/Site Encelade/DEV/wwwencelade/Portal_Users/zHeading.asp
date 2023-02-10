<!--#INCLUDE FILE="ADOConnect.asp"-->
<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<html>
<head>
<title>Resources Management</title>
</head>
<basefont FACE="Arial, Helvetica, sans-serif">
<link REL="stylesheet" HREF="style.css">
<body BGCOLOR="White" <%= LoadMe %>>
<img src="corner.gif" border="0" WIDTH="23" HEIGHT="20"><br clear="all">

<% '********************** HEADING ***************** %>
<table>
<tr>
<td valign="middle" align="left" NOWRAP>
&nbsp;
<font size="+3" face="Times New Roman" color="#004080"><strong>R</strong></font><font size="+3" color="#606060"><em>esource</em></font>&nbsp;
<font size="+3" face="Times New Roman" color="#004080"><strong>M</strong></font><font size="+3" color="#606060"><em>anagement</em></font>&nbsp;
</td>
</tr>
</table>

<% sql0 = "SELECT Employees.*, Positions.PositionID, Positions.BossID, Positions.Title, Positions.Department, Positions.Division FROM Employees LEFT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Employees.EmployeeID <> null " %>
<% sql00 = "SELECT Count(*) as RecCount FROM Employees LEFT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Employees.EmployeeID <> null " %>


<% sqlPOS = "SELECT Employees.*, Employees.EmployeeID , Positions.* FROM Employees RIGHT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Positions.PositionID <> null " %>
<% sqlPOSCount = "SELECT Count(*) as RecCount FROM Employees RIGHT JOIN Positions ON Employees.EmployeeID = Positions.EmployeeID WHERE Positions.PositionID <> null " %>