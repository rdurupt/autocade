<html>
<head>
<link rel="stylesheet" href="PMainStyle1.asp">
</head>

<%

Sub CreateNewMenu(MenuSize,MenuColor,MenuFonts,BorderSize,ImgSet)
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
	
end sub

Sub CreateNewBar(TopSize,ItemSize)
	Response.Write "menu = new NavBarMenu(" & TopSize & "," & ItemSize & ");" & vbcrlf
end sub

Sub CloseMenuBar
	Response.Write "bar.addMenu(menu);" & vbcrlf
end sub

Sub AddMenuItem(MyCaption,MyLink)
	Response.Write "menu.addItem(new NavBarMenuItem(" & chr(34) & chr(34) & chr(34) & MyCaption & chr(34) & ", " & chr(34) & Mylink & chr(34) & "));" & vbcrlf
	'Response.Write "menu.addItem(new NavBarMenuItem(" & chr(34) & MyCaption & chr(34) & ", " & chr(34) & Mylink & chr(34) & "));" & vbcrlf
End Sub

Sub EndMenu
	Response.Write "</script>" & vbcrlf
End Sub

Sub PositionneMenu(xx,yy)
	Response.Write "bar.moveBy(" & xx & ", " & yy & ");" & vbcrlf
End sub


	'             border,   headerFG,  headerBG,  headrHiFg, hdrHiBg ,itmFgColor, itmBgColor, itmHiFgColor, itmHiBgColor
	MenuColor=  "'#FFFFFF', '#000080', '#C0F0C0', '#FF0000', '#C0C0F0', '#000080', '#C0C0FF','#FFFFFF', '#000080' " 
	'            <---------- HEADER ----------------><---------- ITEMS-----------------------> 
	MenuFonts=" 'Verdana', 'plain', 'normal', 'xx-small', 'Verdana', 'plain', 'normal', 'xx-small' "
	'			BorderSize,Height,SepSize
	BorderSize="1, 3, 1 "
	ImageSet= "'','bouton_on.jpg','bouton_off.jpg'"
%>
