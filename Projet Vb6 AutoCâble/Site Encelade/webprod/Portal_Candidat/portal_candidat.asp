<%

'session("candidat_UserType") = ""
session("candidat_UserType") = "Administrator"
'session("candidat_UserType") = ""
'session("portal_candidat_Recherhe")=1
session("portal_candidat_Recherhe")=0
'    session("strTmplMain")=""
	Response.Redirect "default.asp"
%>