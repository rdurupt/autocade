
<%
'Session("ADOIntranet") = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("ASPIntranet.mdb") 

'******************  Sample DSN-less connections
'Session("ADOIntranet") = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & "c:\Inetpub\wwwvirtual\mayanetics\data\ASPIntranet.mdb"  
'Session("ADOIntranet") = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & "c:\Inetpub\wwwroot\nggug\members\employee2.mdb"  
Session("ADOIntranet")= session("DSN")
Set Conn = Server.CreateObject("ADODB.Connection")  
Conn.Open Session("ADOIntranet") 
%>
