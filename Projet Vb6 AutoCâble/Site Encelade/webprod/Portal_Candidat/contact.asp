
<!--#INCLUDE FILE="ADOConnect.asp"-->
 <!--#INCLUDE FILE="con_topmenu.asp"-->

<%
session("contact_devis")="http://localhost/demoencelade/Portal_Candidat/Devis/"
conn.close
set conn=Nothing
Set candidat  = Nothing 
Dim candidat
Set candidat = Server.CreateObject("CatalogueEncelade.Candidat")

candidat.Main

Set candidat  = Nothing 

%>



