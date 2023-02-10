
<!--#INCLUDE FILE="ADOConnect.asp"-->
 <!--#INCLUDE FILE="con_topmenu.asp"-->

<%
session("contact_devis")="http://gnv.euxia.net/demoencelade/Portal_Candidat/Devis/"
Dim candidat
Set candidat = Server.CreateObject("CatalogueEncelade.Candidat")
conn.close
set conn=Nothing
candidat.Main

Set candidat = Nothing 
%>



