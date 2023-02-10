<%response.expires=-1440%>
<html>
<!--#INCLUDE FILE="../Portal_Asp/Portal_Common_db.asp"-->
<head>
<link rel=STYLESHEET href="../Portal_Styles/PMainStyle1.asp" type="text/css">
</head>
<body  bgcolor="#000000" background="../Portal_Html/Images/Background.asp">
<%
Response.write "<br><br><br><br><br><br><table align='left' width='80%'><tr><td><p id=welcome><font size=2>"
Response.write "Bienvenue " & Session("FULL_NAME") & ", <br>Votre compte a été ouvert sur I-graal.<br>"
Response.write "Votre mot de passe vient de vous être envoyé par e-mail<br> Il vous sera nécessaire pour accéder à notre service<br>"
Response.write "Vous devez vous connecter sur <dheading><a href='http://members.i-graal.com'>http://members.i-graal.com</a>" & "<br>"
Response.write "Votre login : " & request("email") & "<br>"
Response.write "Votre compte a été crédité de 10 Graals en cadeau de bienvenue<br>"
%>
