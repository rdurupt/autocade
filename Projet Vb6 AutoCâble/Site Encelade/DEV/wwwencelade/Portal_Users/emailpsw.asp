<% 
Response.Expires = 0 
%>
<!--#INCLUDE FILE="../portal_Asp/portal_common_db.asp"-->
<html>
<link rel="stylesheet" href="../Portal_styles/PMainStyle1.asp">
<body background="../Portal_Html/Images/Background.asp">
<br><br><br><br><br><br><br><br><left><p id=Dheading>
<% 
    USERNAME = Replace(Request("namelogin"),"'","")
    SQL = "select * from dbp_Users where UserLogin='" & USERNAME & "'"
    set rs1 = Server.CreateObject("ADODB.Recordset")
    rs1.open SQL, My_conn
	if rs1.eof or rs1.bof then
		%>
		 Erreur:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[<%=USERNAME%>] nom d'utilisateur est inconnu
		<%
		Response.end
	end if	
    ' Envoi de l'e-mail
    %>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Le mot de passe vient d'être envoyé par e-mail
    <%
    MyEmail=rs1("UserEmail")
    EmailFrom="Webmaster I-Graal<cgoaziou@i-graal.com>"
    EmailSubject="Votre compte i-graal : Mot de passe "
    EmailText=" Votre mot de passe est : "  & rs1("UserEmailPass")
    set mymail = server.createobject("DynuEmail.Functions")
    mymail.Smtp = "Mail.1ngk.com" ' session("MAILSERVER")
    result = mymail.Send(EmailFrom,MyEmail,EmailSubject, EmailText)
    set mymail = nothing     

%>
</left>
</html>
</body>    
	
