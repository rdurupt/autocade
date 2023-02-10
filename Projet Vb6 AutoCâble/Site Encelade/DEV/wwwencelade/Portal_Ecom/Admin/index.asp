<%
if session("ADMIN") <> 1 then
	response.redirect(session("UrlBase"))
end if
%>
<html>
<head>
<title>CAS Aviation - Administration</title>
</head>
<FRAMESET COLS="250,*">
	<FRAME NAME="gauche" SRC="menu.asp" MARGINWIDTH=2 MARGINHEIGHT=2 SCROLLING=Auto>
	<FRAME NAME="droite" SRC="welcome.asp" MARGINWIDTH=8 MARGINHEIGHT=4 SCROLLING=Auto>
</FRAMESET>
</html>
