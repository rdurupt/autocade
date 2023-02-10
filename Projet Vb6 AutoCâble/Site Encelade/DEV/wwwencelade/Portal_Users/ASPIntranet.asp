<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<% 
Conn.Close
Dim obj
Session("userStyleSheet")="usereuxia.css"
Set obj = Server.CreateObject("userportal.Directory")
obj.Main
Set obj = Nothing 
%>
