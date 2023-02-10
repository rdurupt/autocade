<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<% 
Conn.Close
Dim obj
Set obj = Server.CreateObject("annuaire.Directory")
obj.Main
Set obj = Nothing 
%>
