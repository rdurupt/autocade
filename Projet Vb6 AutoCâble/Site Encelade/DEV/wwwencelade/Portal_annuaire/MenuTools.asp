<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%

if session("javapuce")="" then
	puce="<img src='puce_carre.gif'>&nbsp;"
else
	puce=session("javapuce")
end if
	pplus="<img src='pplus.gif'>&nbsp;"

Sub CreateNewMenu(MenuSize,MenuColor,MenuFonts,BorderSize,ImgSet)

	if session("javapuce")="" then
		puce="<img src='puce_carre.gif'>&nbsp;"
	else
		puce=session("javapuce")
	end if
	pplus="<img src='pplus.gif'>&nbsp;"
		if Session("MenuColor")="" then
		'             border,   headerFG,  headerBG,  headrHiFg, hdrHiBg ,itmFgColor, itmBgColor, itmHiFgColor, itmHiBgColor
		MenuColor=  "'#000000', '#000000', '#347AB6', '#F8F8F8', '#000080', '#000080', '#C0C0FF','#FFFFFF', '#000080' " 
	else
		MenuColor= Session("MenuColor")
	end if	

	'            <---------- HEADER ----------------><---------- ITEMS-----------------------> 
	MenuFonts=" 'Verdana', 'plain', 'bold', 'xx-small', 'Verdana', 'plain', 'bold', 'xx-small' "
	'			BorderSize,Height,SepSize
	BorderSize="1, 1, 1 "
	
	if Session("ImageSet")="" then
		ImageSet= "'','" & Session("sMenu_B_On") & "','" & Session("sMenu_B_Off") & "'"
	else
		ImageSet = Session("ImageSet")
	end if	

	Response.Write vbcrlf
	Response.Write "<script language=""JavaScript"" src=""dhtmllib.js""></script>" & vbcrlf
	Response.Write "<script language=""JavaScript"" src=""navbar.js""></script>" & vbcrlf

	Response.Write "<script language=""JavaScript"">" & vbcrlf
	Response.Write "function init()" & vbcrlf 
	Response.Write "{" & vbcrLf
  	Response.Write "bar.create();" & vbcrlf
	Response.Write "}" & vbcrlf
	Response.Write "var bar = new NavBar(" & MenuSize & ");" & vbcrlf
	Response.Write "var menu;" & vbcrlf
	Response.Write "bar.setColors(" & MenuColor & ");" & vbcrlf
	Response.Write "bar.setFonts(" & MenuFonts  & ");" & vbcrlf
	Response.Write "bar.setSizes(" & BorderSize & ");" & vbcrlf
	Response.Write "bar.setImages(" & ImgSet & ");" & vbcrlf
	Response.Write "bar.setAlign('left');" & vbcrlf
	
end sub

Sub CreateNewBar(TopSize,ItemSize)
	Response.Write "menu = new NavBarMenu(" & TopSize & "," & ItemSize & ");" & vbcrlf
end sub

Sub CloseMenuBar
	Response.Write "bar.addMenu(menu);" & vbcrlf
end sub

Sub AddMenuItem(MyCaption,MyLink)
	Response.Write "menu.addItem(new NavBarMenuItem(" & chr(34) &  puce & MyCaption & chr(34) & ", " & chr(34) & Mylink & chr(34) & "));" & vbcrlf
End Sub

Sub EndMenu
	Response.Write "</script>" & vbcrlf
End Sub

Sub PositionneMenu(xx,yy)
	Response.Write "bar.moveBy(" & xx & ", " & yy & ");" & vbcrlf
End sub
 


%>
