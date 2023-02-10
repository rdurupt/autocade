<HTML>
<HEAD><TITLE>Variables de serveur HTTP</TITLE></HEAD>
<BODY BGCOLOR=#FFFFFF>
<H1>Variables de serveur HTTP</H1>

<TABLE BORDER=1>
<TR><TD VALIGN=TOP><B>Variable</B></TD><TD VALIGN=TOP><B>Valeur</B></TD></TR>
<% For Each key In Request.ServerVariables %>
<TR>
<TD><% = key %></TD>
<TD>
<%
if Request.ServerVariables(key) = "" Then
Response.Write "&nbsp" ' Pour faire apparaître une bordure
' autour de la cellule d'un tableau
else 
Response.Write Request.ServerVariables(key)
end if
Response.Write "</TD>"
%>
</TR>
<% Next %>
</TABLE>
<BR>
<BR>

</BODY>
</HTML>
