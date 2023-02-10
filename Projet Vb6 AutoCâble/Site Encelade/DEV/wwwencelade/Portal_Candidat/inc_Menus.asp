<%

Public Function GetMenu()

	pr ("<script language='javascript'>")
	pr ("function javAbout() {")
	pr ("	window.open('Contact.asp?mode=dlgAbout','dlgabout','resizable=yes,status=no,top=150,left=150,width=400,height=150')")
	pr ("}")
	pr ("</script>")
	
	pr ("<script language='vbscript'>")
	pr ("function mnMouseOver(i)")
	pr ("	if i = 1 then")
	pr ("		td1.style.backgroundcolor = ""white""")
	pr ("		sp1.style.color = ""#9B9FC3""")
	for k = 2 to 10
	pr ("	elseif i = " & k & " then")
	pr ("		td" & k & ".style.backgroundcolor = ""white""")
	pr ("		sp" & k & ".style.color = ""#9B9FC3""")
	next
	pr ("	end if")
	pr ("end function")
	
	pr ("function mnMouseOut(i)")
	pr ("	if i = 1 then")
	pr ("		td1.style.backgroundcolor = ""#9B9FC3""")
	pr ("		sp1.style.color = ""white""")
	for k = 2 to 10
	pr ("	elseif i = " & k & " then")
	pr ("		td" & k & ".style.backgroundcolor = ""#9B9FC3""")
	pr ("		sp" & k & ".style.color = ""white""")
	next
	pr ("	end if")
	pr ("end function")
	pr ("</script>")
	
	
	'****************  CONTACT MANAGER
	
    If UCase(Session("candidat_CurrentPage")) = UCase("con_frmSetSubField.asp") Or UCase(Session("candidat_CurrentPage")) = UCase("con_frmSetting.asp") Or UCase(Session("candidat_CurrentPage")) = UCase("con_frmField.asp") Or UCase(Session("candidat_CurrentPage")) = UCase("con_frmSetField.asp") Then
		pr ("<table border='0' bgcolor='white' cellspacing='1'><tr>")
		pr ("<td id='td1' nowrap bgcolor='#9B9FC3' onMouseOver='mnMouseOver(1)' onMouseOut='mnMouseOut(1)'><a href='Contact.asp?mode=CON_FRMSETTING'><span id='sp1'><b>&nbsp;Paramètres&nbsp;</b></span></a></td>")
		pr ("<td id='td2' nowrap bgcolor='#9B9FC3' onMouseOver='mnMouseOver(2)' onMouseOut='mnMouseOut(2)'><a href='con_frmField.asp'><span id='sp2'><b>&nbsp;Champs affichés&nbsp;</b></span></a></td>")
		if session("candidat_UserType") = "Administrator" then 
			pr ("<td id='td5' nowrap bgcolor='#9B9FC3' onMouseOver='mnMouseOver(5)' onMouseOut='mnMouseOut(5)'><a href='con_frmSetField.asp'><span id='sp5'><b>&nbsp;Personnalisation&nbsp;</b></span></a></td>")
		end if
		pr ("</tr></table>")
	end if
end function
%>
