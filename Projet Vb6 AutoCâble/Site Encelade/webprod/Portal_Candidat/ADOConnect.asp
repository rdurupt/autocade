
<%

if trim(session("candidat_ADOContact"))="" then
	session("candidat_ADOContact")= "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=C:\webprod\dbsportail\wwwencelade\Encelade_Menu.mdb"  
end if

Set Conn = Server.CreateObject("ADODB.Connection") 
Conn.Mode = 16


 Conn.Open session("candidat_ADOContact")
%>
