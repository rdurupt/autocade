<%

' Gestion des LoadTemplate
If "งง" & Session("PortalModuleTmpl") = "งง" And CInt(Session("PortalFrames")) = 1 Then
	If Request("h_groupid").Count > 0 Then
		sGrp = "&h_groupid=" & Request("h_groupid")
	Else
		sGrp = ""
	End If
	Session("LinkAfterModule") = Session("PortalPath") & "Portal_Groups/Portal_Groups.asp?h_mode=" & Request("h_mode") & sGrp
	Session("PortalModuleTmpl") = "*"
	Response.Redirect Session("PortalPath") & "Portal_Groups/Portal_Groups.asp?h_mode=LoadTemplate&h_tmplid=" & Session("PortalGroupsTmpl") & sGrp
	Response.End
End If

Dim Portal
Set Portal = Server.CreateObject(Application("PortalComObject") & ".PortalGroupes")
Portal.Main
Set Portal = Nothing
Response.End
%>