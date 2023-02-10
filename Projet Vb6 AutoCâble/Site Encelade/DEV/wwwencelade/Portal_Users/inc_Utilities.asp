<%
Public Function GetMenu()
	if len(Request("menuEmployee")) > 0 then
		Session("menuEmployee") = Request("menuEmployee")
	end if
	pr ("<script language='javascript'>")
	pr ("function javAbout() {")
	pr ("	window.open('Employee.asp?mode=dlgAbout','dlgabout','resizable=yes,status=no,top=150,left=150,width=400,height=150')")
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

	if Session("menuEmployee") = "EmployeeList" then
		pr ("<table border='0' bgcolor='gray' cellspacing='1'><tr>")
		pr ("<td id='td1' nowrap bgcolor='white' onMouseOver='mnMouseOver(1)' onMouseOut='mnMouseOut(1)'><a href='Employee.asp?mode=lst'><span id='sp1'><b>&nbsp;Employee List&nbsp;</b></span></a></td>")
		pr ("<td id='td2' nowrap bgcolor='white' onMouseOver='mnMouseOver(2)' onMouseOut='mnMouseOut(2)'><a href='frmEmp.asp?UserID=0'><span id='sp2'><b>&nbsp;New Employee&nbsp;</b></span></a></td>")
		pr ("<td id='td3' nowrap bgcolor='white' onMouseOver='mnMouseOver(3)' onMouseOut='mnMouseOut(3)'><a href='frmQuery.asp'><span id='sp3'><b>&nbsp;Query&nbsp;</b></span></a></td>")
		if len(Session("employeeSearch")) > 0  or len(Session("employeeQuery")) > 0 then
			if len(Session("employeeSearch")) > 0 then
				pr ("<td id='td4' nowrap bgcolor='white' onMouseOver='mnMouseOver(4)' onMouseOut='mnMouseOut(4)'><a href='Employee.asp?mode=lst&reset=true'><span id='sp4'><b>&nbsp;Reset Search&nbsp;</b></span></a></td>")
			else
				pr ("<td id='td4' nowrap bgcolor='white' onMouseOver='mnMouseOver(4)' onMouseOut='mnMouseOut(4)'><a href='Employee.asp?mode=lst&reset=true'><span id='sp4'><b>&nbsp;Reset Query&nbsp;</b></span></a></td>")
			end if
		end if
		pr ("</tr></table>")
	end if
	if Session("menuEmployee") = "Administration" then
		pr ("<table border='0' bgcolor='gray' cellspacing='1'><tr>")
		pr ("<td id='td1' nowrap bgcolor='white' onMouseOver='mnMouseOver(1)' onMouseOut='mnMouseOut(1)'><a href='frmSetting.asp'><span id='sp1'><b>&nbsp;Settings&nbsp;</b></span></a></td>")
		pr ("<td id='td2' nowrap bgcolor='white' onMouseOver='mnMouseOver(2)' onMouseOut='mnMouseOut(2)'><a href='frmField.asp'><span id='sp2'><b>&nbsp;Select Fields&nbsp;</b></span></a></td>")
		pr ("<td id='td3' nowrap bgcolor='white' onMouseOver='mnMouseOver(3)' onMouseOut='mnMouseOut(3)'><a href='frmPassword.asp'><span id='sp3'><b>&nbsp;Change Password&nbsp;</b></span></a></td>")
		if len(Session("employeeAccess")) > 0 then
			pr ("<td id='td4' nowrap bgcolor='white' onMouseOver='mnMouseOver(4)' onMouseOut='mnMouseOut(4)'><a href='Employee.asp?mode=modUser&sub=lst'><span id='sp4'><b>&nbsp;Manage Users&nbsp;</b></span></a></td>")
			pr ("<td id='td5' nowrap bgcolor='white' onMouseOver='mnMouseOver(5)' onMouseOut='mnMouseOut(5)'><a href='frmSetField.asp'><span id='sp5'><b>&nbsp;Customize Fields&nbsp;</b></span></a></td>")
		end if

		pr ("<td id='td6' nowrap bgcolor='white' onMouseOver='mnMouseOver(6)' onMouseOut='mnMouseOut(6)'><a href='javascript:javAbout()'><span id='sp6'><b>&nbsp;About&nbsp;</b></span></a></td>")
		pr ("</tr></table>")
	end if
end function


Function GetSubordinates(Boss)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOIntranet")
	Set RS000 = ConnSP.Execute("SELECT * FROM dbp_UserInfos WHERE BossID = " & Boss)
	do while not RS000.EOF
		Session("Subordinates") = Session("Subordinates") & ":" & RS000("UserID") & ":"
		Call GetSubordinates(RS000("UserID"))
		RS000.movenext
	loop
    ConnSP.Close
End Function
'************************************ defaults

Function SetDefault(fld, val)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOIntranet")
    ConnSP.Execute ("UPDATE dbp_defaultSettings SET defValue = '" & val & "' WHERE defName = '" & fld & "'")
    ConnSP.Close
End Function

Function GetDefault(fld,def)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOIntranet")
    Set RS100 = ConnSP.Execute("SELECT * FROM dbp_defaultSettings WHERE defName = '" & fld & "'")
    If Not RS100.EOF Then
        GetDefault = trim(RS100("defValue"))
	else
		ConnSP.Execute("INSERT INTO dbp_defaultSettings(defName,defValue) VALUES('" & fld & "','" & def & "')")
		GetDefault = def
    End If
    ConnSP.Close
End Function

Function SetUserDefault(usr, defName, defValue)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOIntranet")
	ConnSP.Execute("DELETE FROM dbp_defaultUserSettings WHERE UserID = " & session("UserID") & " AND defName = '" & defName & "'")
	ConnSP.Execute("INSERT INTO dbp_defaultUserSettings(UserID,defName,defValue) VALUES(" & session("UserID") & ",'" & defName & "','" & defValue & "')")
    ConnSP.Close
End Function

Function GetUserDefault(usr, defName, defValue)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOIntranet")
    Set RSSpec = ConnSP.Execute("SELECT defValue FROM dbp_defaultUserSettings WHERE defName = '" & defName & "' AND UserID = " & session("UserID"))
    If Not RSSpec.EOF Then
        GetUserDefault = trim(RSSpec("defValue"))
    Else
        GetUserDefault = defValue
    End If
    ConnSP.Close
End Function


function funY2K(d)
	strDate = trim(d)
	if instr(strDate, " ") then
		strDate = left(strDate,instr(strDate, " "))
		trailer = right(d,len(d) - instr(d, " ")) 
	end if
	if isdate(strDate) then
		dateY2K = strDate
		if instr(strDate,"/") = 2 then
			strMonth = left(strDate,1)
			if instr(3,strDate,"/") = 4 then
				strDay = mid(strDate,3,1)
			else
				strDay = mid(strDate,3,2)
			end if
			strYear = right(strDate,len(strDate) - instr(3,strDate,"/"))
		elseif instr(strDate,"/") = 3 then
			strMonth = left(strDate,2)
			if instr(4,strDate,"/") = 5 then
				strDay = mid(strDate,4,1)
			else
				strDay = mid(strDate,4,2)
			end if
			strYear = right(strDate,len(strDate) - instr(4,strDate,"/"))
		end if
		intYear = cInt(strYear)
		if intYear >= 0 and intYear < 51 then
			strYear = "20" & strYear
		elseif intYear > 50 and intYear < 100 then
			strYear = "19" & strYear
		end if
		strFinal = strMonth & "/" & strDay & "/" & strYear & " " & trailer
		funY2K = trim(strFinal)
	else
		funY2K = ""
	end if
end function


Public Function getSmallMonth(mn)
    If mn = 1 Then
        getSmallMonth = "Jan"
    ElseIf mn = 2 Then
        getSmallMonth = "Feb"
    ElseIf mn = 3 Then
        getSmallMonth = "Mar"
    ElseIf mn = 4 Then
        getSmallMonth = "Apr"
    ElseIf mn = 5 Then
        getSmallMonth = "May"
    ElseIf mn = 6 Then
        getSmallMonth = "Jun"
    ElseIf mn = 7 Then
        getSmallMonth = "Jul"
    ElseIf mn = 8 Then
        getSmallMonth = "Aug"
    ElseIf mn = 9 Then
        getSmallMonth = "Sep"
    ElseIf mn = 10 Then
        getSmallMonth = "Oct"
    ElseIf mn = 11 Then
        getSmallMonth = "Nov"
    ElseIf mn = 12 Then
        getSmallMonth = "Dec"
    Else
        getSmallMonth = "Unk"
    End If
End Function

Public Function MonthName(mn)
    If mn = 1 Then
        MonthName = "January"
    ElseIf mn = 2 Then
        MonthName = "February"
    ElseIf mn = 3 Then
        MonthName = "March"
    ElseIf mn = 4 Then
        MonthName = "April"
    ElseIf mn = 5 Then
        MonthName = "May"
    ElseIf mn = 6 Then
        MonthName = "June"
    ElseIf mn = 7 Then
        MonthName = "July"
    ElseIf mn = 8 Then
        MonthName = "August"
    ElseIf mn = 9 Then
        MonthName = "September"
    ElseIf mn = 10 Then
        MonthName = "October"
    ElseIf mn = 11 Then
        MonthName = "November"
    ElseIf mn = 12 Then
        MonthName = "December"
    Else
        MonthName = "Unknown Month"
    End If
End Function

Public Function GetLastDay(mn)
    If mn = 2 Then
        GetLastDay = 28
    ElseIf mn = 4 Or mn = 6 Or mn = 9 Or mn = 11 Then
        GetLastDay = 30
    Else
        GetLastDay = 31
    End If
End Function

Public Function pr(str)
    Response.Write (str & vbCrLf)
End Function

Public function safeEntry(strField)
	strSafe = trim(strField)
	strSafe = funReplace(strSafe,"'","`")
	strSafe = funReplace(strSafe,"<","&lt;")
	strSafe = funReplace(strSafe,">","&gt;")
	safeEntry = strSafe
end function


function getLetter(num)
	if num = 1 then
		getLetter = "a"
	elseif num = 2 then
		getLetter = "b"
	elseif num = 3 then
		getLetter = "c"
	elseif num = 4 then
		getLetter = "d"
	elseif num = 5 then
		getLetter = "e"
	elseif num = 6 then
		getLetter = "f"
	elseif num = 7 then
		getLetter = "g"
	elseif num = 8 then
		getLetter = "h"
	elseif num = 9 then
		getLetter = "i"
	elseif num = 10 then
		getLetter = "j"
	elseif num = 11 then
		getLetter = "k"
	elseif num = 12 then
		getLetter = "l"
	elseif num = 13 then
		getLetter = "m"
	elseif num = 14 then
		getLetter = "n"
	elseif num = 15 then
		getLetter = "o"
	elseif num = 16 then
		getLetter = "p"
	elseif num = 17 then
		getLetter = "q"
	elseif num = 18 then
		getLetter = "r"
	elseif num = 19 then
		getLetter = "s"
	elseif num = 20 then
		getLetter = "t"
	elseif num = 21 then
		getLetter = "u"
	elseif num = 22 then
		getLetter = "v"
	elseif num = 23 then
		getLetter = "w"
	elseif num = 24 then
		getLetter = "x"
	elseif num = 25 then
		getLetter = "y"
	elseif num = 26 then
		getLetter = "z"
	else
		getLetter = ""
	end if
end function

function htmlString(strField)
	strHTML = trim(strField)
	strHTML = replace(strHTML," ","+")
	htmlString = strHTML
end function

public Function funReplace(a,b,c)
	funReplace = replace(a,b,c)
end function

%>
