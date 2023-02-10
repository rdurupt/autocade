<%
Sub CreateNewMenu()
	Response.Write "<script language=""JavaScript"" src=""navcond.js""></script>" & vbcrlf
	Response.Write "<script language=""JavaScript"">" & vbcrlf
	Response.Write "function init()" & vbcrlf 
	Response.Write "{" & vbcrLf
  	Response.Write "bar.create();" & vbcrlf
	Response.Write "}" & vbcrlf
	Response.Write "var bar = new NavBar(600);" & vbcrlf
	Response.Write "var menu;" & vbcrlf
	                   'setColors(  bdColor,    hdrFgColor,  hdrBgColor,  hdrHiFgColor, hdrHiBgColor,itmFgColor, itmBgColor, itmHiFgColor, itmHiBgColor
	Response.Write "bar.setColors(""#D5D5F8"", ""#C8C8C8"", ""#000080"", ""#F8F8F8"", ""#000080"", ""#000090"", ""#A4A4F4"", ""#F8F8F8"", ""#000080"");" & vbcrlf
	Response.Write "bar.setFonts(""Arial"", ""plain"", ""bold"", ""xx-small"", ""Verdana"", ""plain"", """", ""xx-small"");" & vbcrlf
	Response.Write "bar.setSizes(1, 4, 2);" & vbcrlf
end sub

Sub CreateNewBar()
	Response.Write "menu = new NavBarMenu('',200);" & vbcrlf
end sub

Sub CloseMenuBar
	Response.Write "bar.addMenu(menu);" & vbcrlf
end sub

Sub AddMenuItem(MyCaption,MyLink)
	Response.Write "menu.addItem(new NavBarMenuItem(" & chr(34) & MyCaption & chr(34) & ", " & chr(34) & Mylink & chr(34) & "));" & vbcrlf
End Sub

Sub EndMenu
	Response.Write "</script>" & vbcrlf
End Sub

Sub PositionneMenu(xx,yy)
	Response.Write "bar.moveBy(" & xx & ", " & yy & ");" & vbcrlf
End sub
%>
