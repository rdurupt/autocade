
<%
set fso= Server.CreateObject("Scripting.FileSystemObject")
'aa= Server.MapPath("/demoencelade/Portal_Candidat")
'response.write aa
'Response.End
If FSO.FileExists("C:\webprod\dbsportail\wwwencelade\Maintenance.txt") = True Then

	response.redirect "../Portal_Candidat/Mantenace.htm"

else
	if trim(session("candidat_ADOContact"))="" then
		session("candidat_ADOContact")= "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=C:\webprod\dbsportail\wwwencelade\Encelade_Menu.mdb"  
	end if

	Set Conn = Server.CreateObject("ADODB.Connection") 
	Conn.Mode = 16
 	Conn.Open session("candidat_ADOContact")
end if
set fso=nothing
%>
