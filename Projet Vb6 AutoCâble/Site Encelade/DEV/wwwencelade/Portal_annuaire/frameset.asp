<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<html>
<head>
<title><%= GetDefault("AppTitle", "Employee Manager") %></title>
</head>
    
<frameset cols="<%= GetDefaultUser("MenuWidth", "145") %>,*">
    <frame name="emp1" src="tabMenu.asp" marginheight="0" marginwidth="0" scrolling="auto" frameborder="0">
    <frame name="emp2" src="Employee.asp?menuEmployee=EmployeeList&mode=lst" marginheight="0" marginwidth="0" frameborder="0">
</frameset>

</html>
<% Conn.Close %>
