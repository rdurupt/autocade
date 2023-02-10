<% 
Session("IsEvaluation") = "true"
if len(Session("web_UserID")) = 0 then
	pg = "frmLogon.asp"
else
	pg = "emp_frameset.asp"
end if
Response.redirect pg
%>
