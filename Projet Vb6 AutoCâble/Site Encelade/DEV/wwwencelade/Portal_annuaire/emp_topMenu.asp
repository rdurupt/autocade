<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<% 

Dim obj
Set obj = Server.CreateObject("annuaire.Directory")
obj.DirectoryMenu
Set obj = Nothing 
%>
