<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<!--#INCLUDE FILE="../portal_Java/MenuTools.asp"-->
<link rel=STYLESHEET href="../Portal_Styles/PMainStyle1.asp" type="text/css">
	<%
	
	CreateNewMenu 600,MenuColor,MenuFonts,BorderSize,ImageSet
	CreateNewBar 200,300
		AddMenuItem "Annuaire des Membres",""
		CloseMenuBar
	CreateNewBar 75,300
		AddMenuItem  "Liste",""
		AddMenuItem  "Listes des Membres","javascript:parent.frames['main'].location='ASPIntranet.asp?mode=emp_lst&menuEmployee=EmployeeList'"
		AddMenuItem  "Synoptique Hiérarchique ","javascript:parent.frames['main'].location='ASPIntranet.asp?mode=emp_chtOrganization&menuEmployee=none'"
		CloseMenuBar
	CreateNewBar 75,300
		AddMenuItem  "Outils ",""
		AddMenuItem  "Recherche ","javascript:parent.frames['main'].location='emp_frmQuery.asp'"
		AddMenuItem  "Configuration ","javascript:parent.frames['main'].location='emp_frmSetting.asp?menuEmployee=Administration'"
		CloseMenuBar

if Session("Admin") = 1  then
	CreateNewBar 75,300
		AddMenuItem  "Login ","javascript:parent.frames['main'].location='modUser.asp?mode=web_lst'"
		CloseMenuBar
End if
	PositionneMenu 1,1
	EndMenu
%>
<body background="../../Portal_Html/Images/Background.asp" onload="init()"
<br>
<br>
<br>




    

