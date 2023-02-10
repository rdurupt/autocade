<%
url = session("urlBase")
If session("username") <> "guest" then


end if
 
	session.abandon
	response.redirect url
	
%>
