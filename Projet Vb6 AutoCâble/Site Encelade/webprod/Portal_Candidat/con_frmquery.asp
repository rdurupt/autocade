<%
Session("candidat_UserType") = "Administrator"
'session("candidat_UserType") = ""
session("portal_candidat_Recherhe")=1
    'session("strTmplMain")=""
	response.redirect "default.asp"
	%>