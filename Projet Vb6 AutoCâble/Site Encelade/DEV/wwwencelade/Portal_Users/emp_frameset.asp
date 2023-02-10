<% 
        Session("employeePage") = ""
        Session("employeeSearch") = ""
        Session("employeeQuery") = ""
        Session("employeeLetter") = ""
        Session("FieldName") = ""
        Session("FieldValue") = "" 
	Response.redirect "aspIntranet.asp?mode=emp_lst&reset=true&nomenu=1" 
%>
