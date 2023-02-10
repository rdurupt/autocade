
<!--#INCLUDE FILE="ADOConnect.asp"-->
<%
 
  set Rs =server.CreateObject("ADODB.Recordset")
set rs=Conn.Execute("SELECT BaseDefault.Path FROM BaseDefault;")
Set MyconnM = Server.CreateObject("ADODB.Connection") 
if rs.EOF=false then

DSN	="DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & rs("Path")
end if
MyconnM.Open DSN
set Rs=Nothing

Session.Timeout = 600
session("candidat_IsEvaluation") = "false"
session("candidat_web_UserID")=session("USERId")
if len(session("candidat_web_UserID")) = 0 then
	pg = "frmLogon.asp"
else
	pg = "Contact.asp?mode=CON_lst"
	Set RS1 = MyconnM.Execute("SELECT * FROM Users WHERE UserID= " & session("candidat_web_UserID"))
		If Not RS1.EOF Then
    		session("candidat_web_UserID") = RS1("UserID")
		
			session("candidat_conFirstName") = RS1("Firstname")
			session("candidat_conLastName")=RS1("LastName")
			session("candidat_conLocation")=RS1("Location")
		end if
end if

MyconnM.close
set RS1=Nothing
Session("candidat_con_lstPage") = ""
				Session("candidat_contactSearch") = ""
				Session("candidat_contactQuery") = ""
				Session("candidat_contactLetter") = ""
				Session("candidat_SubCategory") = ""
				Session("CloseWereSearch") = ""
				Session("candidat_contactSearch") = ""
				Session("candidat_contactSortBy")=""
				Session("CloseWherNbLiges")=""
				Session("Encelade_contactSearch")=""
Response.redirect pg
%>
