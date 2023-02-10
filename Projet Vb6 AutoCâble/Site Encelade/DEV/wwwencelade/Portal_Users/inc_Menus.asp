<%

Public Function GetMenu()

	pr ("<script language='javascript'>")
	pr ("function javAbout() {")
	pr ("	window.open('ASPIntranet.asp?mode=dlgAbout','dlgabout','resizable=yes,status=no,top=150,left=150,width=400,height=150')")
	pr ("}")
	pr ("</script>")
	
	pr ("<script language='vbscript'>")
	pr ("function mnMouseOver(i)")
	pr ("	if i = 1 then")
	pr ("		td1.style.backgroundcolor = ""navy""")
	pr ("		sp1.style.color = ""white""")
	for k = 2 to 10
	pr ("	elseif i = " & k & " then")
	pr ("		td" & k & ".style.backgroundcolor = ""navy""")
	pr ("		sp" & k & ".style.color = ""white""")
	next
	pr ("	end if")
	pr ("end function")
	
	pr ("function mnMouseOut(i)")
	pr ("	if i = 1 then")
	pr ("		td1.style.backgroundcolor = ""white""")
	pr ("		sp1.style.color = ""navy""")
	for k = 2 to 10
	pr ("	elseif i = " & k & " then")
	pr ("		td" & k & ".style.backgroundcolor = ""white""")
	pr ("		sp" & k & ".style.color = ""navy""")
	next
	pr ("	end if")
	pr ("end function")
	pr ("</script>")

	if Session("CurrentPage") = "emp_frmSetting.asp" or Session("CurrentPage") = "emp_frmSetField.asp"  or Session("CurrentPage") = "emp_frmField.asp"  then
		pr ("<table border='0' bgcolor='gray' cellspacing='1'><tr>")
		pr ("<td id='td1' nowrap bgcolor='white' onMouseOver='mnMouseOver(1)' onMouseOut='mnMouseOut(1)'><a href='emp_frmSetting.asp'><span id='sp1'><b>&nbsp;Settings&nbsp;</b></span></a></td>")
		pr ("<td id='td2' nowrap bgcolor='white' onMouseOver='mnMouseOver(2)' onMouseOut='mnMouseOut(2)'><a href='emp_frmField.asp'><span id='sp2'><b>&nbsp;Select Fields&nbsp;</b></span></a></td>")
		if Session("Admin") = 1  then 
			pr ("<td id='td5' nowrap bgcolor='white' onMouseOver='mnMouseOver(5)' onMouseOut='mnMouseOut(5)'><a href='emp_frmSetField.asp'><span id='sp5'><b>&nbsp;Customize Fields&nbsp;</b></span></a></td>")
		end if
		pr ("</tr></table>")
	end if

end function
%>
