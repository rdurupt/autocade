<%

' Gestion des LoadTemplate
If "งง" & Session("PortalModuleTmpl") = "งง" And CInt(Session("PortalFrames")) = 1 Then
	Session("LinkAfterModule") = Session("PortalPath") & "Portal_Cats/Portal_Cats.asp?h_mode=" & Request("h_mode")
	Session("PortalModuleTmpl") = "*"
	If Request("s_groupid").Count > 0 Then
		Session("CatGrp") = Request("s_groupid")
	Else
		If Request("h_getcatgrp").Count > 0 Then
			Session("CatGrp") = Request("h_getcatgrp")
		Else
			Session("CatGrp") = 0
		End If
	End If
	Response.Redirect Session("PortalPath") & "Portal_Cats/Portal_Cats.asp?h_mode=LoadTemplate&h_tmplid=" & Session("PortalCatsTmpl")
	Response.End
End If

Dim Portal
Set Portal = Server.CreateObject(Application("PortalComObject") & ".PortalCats")
Portal.Main
Set Portal = Nothing
Response.End
%>