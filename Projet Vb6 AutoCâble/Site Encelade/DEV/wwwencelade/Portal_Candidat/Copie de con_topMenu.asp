<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="MenuTools.asp"-->
<%

set Rs =server.CreateObject("ADODB.Recordset")
set Myconn = Server.CreateObject("ADODB.Connection")
Myconn.Open session("candidat_ADOContact")
set rs=Myconn.Execute("SELECT BaseDefault.Path FROM BaseDefault;")

if Rs.EOF=false then

	DSN	="DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Rs("Path")

end if
Myconn.close
set Myconn = Nothing
set Myconn = Server.CreateObject("ADODB.Connection")

Myconn.Open DSN
set Rs= Nothing


Function GetDefault(fld,def)

    Set ConnSP = Server.CreateObject("ADODB.Connection")
	ConnSP.Open session("candidat_ADOContact")

    Set RS100 = ConnSP.Execute("SELECT * FROM Defaults WHERE defName = '" & fld & "'")
    If Not RS100.EOF Then
        GetDefault = trim(RS100("defValue"))
	else
		ConnSP.Execute("INSERT INTO Defaults(defName,defValue) VALUES('" & fld & "','" & def & "')")
		GetDefault = def
    End If
    ConnSP.Close
        Set ConnSP = Nothing
      Set RS100 = Nothing
End Function

	Response.buffer=true
		'             border,   headerFG,  headerBG,  headrHiFg, hdrHiBg ,itmFgColor, itmBgColor, itmHiFgColor, itmHiBgColor
	MenuColor=  "'#FFFFFF', '#FFFFFF', '#9B9FC3', '#FFFFFF', '#7779A9', '#FFFFFF', '#9B9FC3', '#7779A9', '#C7CBE4'" 
	'            <---------- HEADER ----------------><---------- ITEMS-----------------------> 
	MenuFonts=" 'Verdana', 'plain', 'bold', 'xx-small', 'Verdana', 'plain', 'normal', '11' "
	'			BorderSize,Height,SepSize
	BorderSize="1,2,1"
	ImageSet= "'','',''"
	

	CreateNewMenu 600,MenuColor,MenuFonts,BorderSize,ImageSet
	CreateNewBar 200,200
	
		AddMenuItem "Types de pièces",""
		
		if session("candidat_Category")="" then session("candidat_Category")="All"
		
			Set RS1 = Myconn.Execute("SELECT * FROM menu  WHERE menu.PasVisible=false ORDER BY libelle")
		do while not RS1.EOF 
		
			AddMenuItem RS1("libelle"),"switchbase.asp?CatId=" & rs1("CatId")
			RS1.movenext
		loop
		
		CloseMenuBar
		
	
   	CreateNewBar 200,200
		AddMenuItem  "Outils",""
		AddMenuItem  "Recherche","Contact.asp?mode=search"
		CloseMenuBar
	if session("candidat_UserType") = "Administrator" then	
		CreateNewBar 200,200
			AddMenuItem  "Administration",""
			'AddMenuItem  "Gestion des comptes","javascript:parent.frames['main'].location='modUser.asp?mode=web_lst'"
    		AddMenuItem  getdefault("tables","Liste des tables"),"javascript:parent.frames['main'].location='con_lstCategory.asp'"
    		AddMenuItem  "Paramétrages","javascript:parent.frames['main'].location='Contact.asp?mode=con_frmSetting'"
    		AddMenuItem  "Configuration","javascript:parent.frames['main'].location='Contact.asp?mode=Config'"
		CloseMenuBar
	End if
	

	PositionneMenu 1,1
	EndMenu
		
	If Request("Mode")<>"con_frmcandidat" then
		response.write ("<body background=""background.asp"" onload=""init();"">")
	End if
	
	Response.Write("<br><br><br>")
	myconn.close


set Myconn = Nothing
%>


    

