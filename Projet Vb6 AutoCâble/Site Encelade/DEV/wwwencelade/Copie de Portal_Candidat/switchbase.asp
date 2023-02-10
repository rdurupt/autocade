<%
	 If Trim("" & Request.QueryString("CatId")) <> "" Then 
	 	 Session("Encelade_CatId") = Trim("" & Request.QueryString("CatId")) 
		 if  Session("SaveEncelade_CatId")<> Session("Encelade_CatId") then
				Session("candidat_con_lstPage") = ""
				Session("candidat_contactSearch") = ""
				Session("candidat_contactQuery") = ""
				Session("candidat_contactLetter") = ""
				Session("candidat_SubCategory") = ""
				Session("CloseWereSearch") = ""
				Session("candidat_contactSearch") = ""
				Session("candidat_contactSortBy")=""
				Session("CloseWherNbLiges")=""
		end if
		 	
	end if
	Session("SaveEncelade_CatId") = Session("Encelade_CatId")		
	Session("candidat_MainTable")="con_Contacts"
	
	Session("candidat_CurrentPage")="contact.asp?mode=con_lst"
	Response.redirect("Contact.asp?mode=con_lst&Category=All")
	
	
		
	
	
%>
