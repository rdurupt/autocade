<%	

	' definition du repertoire virtuel
	Application("PortalPath") = "/demoencelade/"
	Application("SessionMsg") = "<br><br><br><center><div style=""font-family:Tahoma;font-size:x-small;text-align:center;border:1px solid #000000;padding:5px;"">Votre session a expir�e.<br><br>Vous allez etre redirig� vers la page d'accueil.</div></center>"
	
	
	' d�finition du DSN Base de donn�es
	Session("DSN")= "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=C:\Dev\wwwencelade\wwwencelade\_Database\Euxiaportal.mdb"
	' zone application dans la base de donn�es
	Session("Application")=1
	' Langue
	Session("Language")="##"

        	Set Portal = Server.CreateObject(Application("PortalComObject") & ".Portal")
	
	Portal.PortalInit Session("DSN"), Session("Application")
	Portal.SecurityLogin "guest", "12345678"


	Set Portal = Nothing

	If CInt(Session("PortalFrames")) = 1 Then

		Response.Redirect Session("PortalPath") & "Portal_Asp/Portal.asp?h_mode=LoadTemplate&h_tmplid=" & Session("PortalTemplate")

	Else

		Response.Redirect (Session("LinkAfterLogin"))

	End If
	
%>
