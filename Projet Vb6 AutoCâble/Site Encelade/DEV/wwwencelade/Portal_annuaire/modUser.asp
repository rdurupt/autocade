<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<!--#INCLUDE FILE="emp_TopMenu.asp"-->
<%  
	'----  Unauthorized entry
	if session("Admin") <>1 then
		Response.Redirect "frmIllegal.asp"
	end if
	
	Session("CurrentPage") = "modUser.asp"
    If Request("mode") = "web_lst" Then
        pr ("<html>")
        pr ("<head>")
        pr ("<title>Manage Users</title>")
            
        pr ("<script language='JavaScript'>")
        pr ("function javNewUser() {")
        pr ("    location.href = '../portal_asp/portal.asp?mode=UserAddNew'")
        pr ("}")
        pr ("</script>")
        pr ("</head>")
		
	    pr ("<script language='VBscript'>")
	    pr ("function vbMouseOver(a)")
	    pr ("   a.style.backgroundcolor = ""#C1E8F7""")
	    pr ("end function")
		
	    pr ("function vbMouseOut(a)")
	    pr ("   a.style.backgroundcolor = ""#F5F7FE""")
	    pr ("end function")
		
	    pr ("function vbEditUsr(ID)")
	    pr ("   location.href = ""../Portal_Asp/portal.asp?mode=UserUpdate&user="" & ID")
	    pr ("end function")
	    pr ("</script>")
		
		
        pr ("<link rel='stylesheet' href='" & session("StyleSheet") &"'>")
        pr ("<body background='" & session("Background") & "'>")

        Set RS1 = Conn.Execute("SELECT Count(*) as RecCount FROM  dbp_UserInfos ")
        
        '***************  TitleBar
        pr ("<table border='0' width='100%' bgcolor='" & GetUserDefault(Session("web_UserID"),"usr_bgcolor","#678DB8") & "' cellpadding='1' cellspacing='0'><tr><td>")
        pr ("<table border='0' width='100%' cellpadding='1' cellspacing='0'>")
        pr ("<tr valign='bottom'>")
		pr ("<form name='frm'>")
        pr ("<td align='left' nowrap>")
        pr ("<img src='user.gif' border='0'>&nbsp;<font size='3' color='" & GetUserDefault(Session("web_UserID"),"usr_color","navy")  & "'><b>User Management</b></font>")
        pr ("</td>")
        pr ("<td align='right' nowrap>")
        pr ("<input type='button' value='New Logon' onClick='javNewUser()'>")
        pr ("</td>")
        pr ("</tr>")
        pr ("</table>")
        pr ("</td></tr></table>")
        
        If Len(Request("msg")) > 0 Then
            pr ("<table border='0'><tr><td><font color='red'><b>" & Request("msg") & "</b></font></td></tr></table>")
        End If
        
        If Len(Request("userLetter")) > 0 Then
            Session("userLetter") = Request("userLetter")
        End If
        If Request("userLetter") = "All" Then
            Session("userLetter") = ""
        End If
    
        If Len(Session("userLetter")) > 0 Then
            strSQL = "SELECT * FROM dbp_UserInfos WHERE LastName LIKE '" & Session("userLetter") & "%' "
            strCOUNT = "SELECT Count(*) AS RecCount FROM dbp_UserInfos  WHERE LastName LIKE '" & Session("userLetter") & "%'"
        Else
            strSQL = "SELECT * FROM  dbp_UserInfos  "
            strCOUNT = "SELECT Count(*) AS RecCount FROM  dbp_UserInfos  "
        End If
        
        '********* Order by
        If Len(Request("UserSortBy")) > 0 Then
            Session("UserSortBy") = Request("UserSortBy")
            strSQL = strSQL & " ORDER BY " & Request("UserSortBy")
        ElseIf Len(Session("UserSortBy")) > 0 Then
            strSQL = strSQL & " ORDER BY " & Session("UserSortBy")
        Else
            Session("UserSortBy") = "UserID"
            strSQL = strSQL & " ORDER BY UserID"
        End If
    
        Set RS0 = Conn.Execute(strSQL)
        Set RSCount = Conn.Execute(strCOUNT)
        RecCount = RSCount("RecCount")
        
        pr ("<table bgcolor='#CFCEDB' cellspacing='1' width='100%'>")
        pr ("<tr>")
        pr ("<td colspan='100%' nowrap align='left'>")
        If Len(Session("userLetter")) > 0 Then
            pr ("<a href='modUser.asp?mode=web_lst&userPage=1&userLetter=All'>[All]</a>&nbsp;&nbsp;")
        Else
            pr ("<a href='modUser.asp?mode=web_lst&userPage=1&userLetter=All'><font size='3' color='red'><b>[All]</b></font></a>&nbsp;&nbsp;")
        End If
        
        For t = 1 To 26
            ltr = getLetter(t)
            ltr = UCase(ltr)
            If Session("userLetter") = ltr Then
                pr ("<font size='3' color='red'><b>" & ltr & "</b></font>&nbsp;")
            Else
                pr ("<a href='modUser.asp?mode=web_lst&userPage=1&userLetter=" & ltr & "'>" & ltr & "</a>&nbsp;")
            End If
        Next
        
        pr ("</td></tr>")
        pr ("</table>")
        
        pr ("<table bgcolor='#BBCCEC' cellspacing='1' width='100%'>")
        
        pr ("<tr bgcolor='#004080'>")
        
        If InStr(Session("UserSortBy"), "UserID") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserID'><img src='SortDesc.gif' border='0'><font color='white'><b>ID</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserID+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>ID</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserID'><font color='white'><b>ID</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "FirstName") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=FirstName'><img src='SortDesc.gif' border='0'><font color='white'><b>First Name</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=FirstName+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>First Name</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=FirstName'><font color='white'><b>First Name</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "LastName") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LastName'><img src='SortDesc.gif' border='0'><font color='white'><b>Last Name</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LastName+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Last Name</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LastName'><font color='white'><b>Last Name</b></font></a></td>")
        End If
        
        
        If InStr(Session("UserSortBy"), "Username") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=Username'><img src='SortDesc.gif' border='0'><font color='white'><b>Username</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=Username+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Username</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=Username'><font color='white'><b>Username</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "DateCreate") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateCreate'><img src='SortDesc.gif' border='0'><font color='white'><b>Create Date</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateCreate+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Create Date</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateCreate'><font color='white'><b>Create Date</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "DateLastAccess") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateLastAccess'><img src='SortDesc.gif' border='0'><font color='white'><b>Last Logon</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateLastAccess+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Last Logon</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateLastAccess'><font color='white'><b>Last Logon</b></font></a></td>")
        End If
		
        If InStr(Session("UserSortBy"), "LogonCount") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LogonCount'><img src='SortDesc.gif' border='0'><font color='white'><b>Log Count</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LogonCount+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Log Count</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LogonCount'><font color='white'><b>Log Count</b></font></a></td>")
        End If
		
        pr ("</tr>")
         
        '************  Initiate page counter
        userPage = 1
        endItem = 20
        beginItem = 1
        i = 0
        iCounter = 0
        
        intPage = Trim(Request("userPage"))
        If Len(intPage) > 0 And IsNumeric(intPage) Then
            Session("userPage") = intPage
        End If
        
        If Len(Session("userPage")) > 0 Then
            userPage = CInt(Session("userPage"))
            endItem = userPage * 20
            beginItem = endItem - (20 - 1)
        End If

        Do While Not RS0.EOF
            i = i + 1
            If i >= beginItem And i <= endItem Then
                col = ""
	            If CStr(RS0("UserID")) = CStr(Session("currUserID")) Then
	                pr ("<tr id='tr" & RS0("UserID") & "' style='background:#C1E8F7;' valign='top' OnMouseOver='vbMouseOver(tr" & RS0("UserID") & ")' OnMouseOut='vbMouseOut(tr" & RS0("UserID") & ")' OnClick='vbEditUsr(" & RS0("UserID") & ")'>")
	            Else
	                pr ("<tr id='tr" & RS0("UserID") & "' style='background:#F5F7FE;' valign='top' OnMouseOver='vbMouseOver(tr" & RS0("UserID") & ")' OnMouseOut='vbMouseOut(tr" & RS0("UserID") & ")' OnClick='vbEditUsr(" & RS0("UserID") & ")'>")
	            End If
                pr ("<td><a href='../Portal_Asp/portal.asp?mode=UserUpdate&user=" & RS0("UserID") & "'><font size='1'>" & RS0("UserID") & "</font></a></td>")
                pr ("<td><a href='../Portal_Asp/portal.asp?mode=UserUpdate&user=" & RS0("UserID") & "'><font size='1'>" & RS0("FirstName") & "</font></a></td>")
                pr ("<td><a href='../Portal_Asp/portal.asp?mode=UserUpdate&user=" & RS0("UserID") & "'><font size='1'>" & RS0("LastName") & "</font></a></td>")
                pr ("<td><a href='../Portal_Asp/portal.asp?mode=UserUpdate&user=" & RS0("UserID") & "'><font size='1'>" & RS0("UserName") & "</font></a></td>")
                pr ("<td><a href='../Portal_Asp/portal.asp?mode=UserUpdate&user=" & RS0("UserID") & "'><font size='1'>" & funY2K(RS0("DateCreate")) & "</font></a></td>")
                pr ("<td><a href='../Portal_Asp/portal.asp?mode=UserUpdate&user=" & RS0("UserID") & "'><font size='1'>" & funY2K(RS0("DateLastAccess")) & "</font></a></td>")
                pr ("<td align='center'><a href='../Portal_Asp/portal.asp?mode=UserUpdate&user=" & RS0("UserID") & "'><font size='1'>" & RS0("LogonCount") & "</font></a></td>")          
                pr ("</tr>")
                iCounter = iCounter + 1
            End If
            If i = endItem Then
                Exit Do
            End If
            RS0.movenext
        Loop
            
        pr ("</table>")
        pr ("</form>")
        '***********************************  Page Count  *****************
        pageCount = RecCount / 20
        
        If RecCount = 0 Then
            pageCount = 0
        ElseIf pageCount < 1 Then
            pageCount = 1
        ElseIf InStr(pageCount, ".") Then
            intLeft = Left(pageCount, InStr(pageCount, "."))
            pageCount = intLeft + 1
        End If
        
        If iCounter > 0 Then
            pr ("<form name='frm' action='modUser'>")
            pr ("<input type='hidden' name='mode' value='web_lst'>")
            brCount = 20 - iCounter
            For j = 0 To brCount
                pr ("<br>")
            Next
            prevPage = userPage - 1
            nextPage = userPage + 1
            
            pr ("<hr>")
            pr ("<table border='0' cellspacing='0'>")
            pr ("<tr><td>&nbsp;</td><td nowrap>Page:&nbsp;")
            
            '******************* Previous Page ***************           
            If prevPage < 1 Then
                pr ("<img src='leftEnd_.gif' border='0'><img src='leftOne_.gif' border='0'>")
            Else
                pr ("<a href='modUser.asp?mode=web_lst&userPage=1'><img src='leftEnd.gif' border='0'></a>")
                pr ("<a href='modUser.asp?mode=web_lst&userPage=" & prevPage & "'><img src='leftOne.gif' border='0'></a>")
            End If
            
            pr ("<input align='right' type='text' name='userPage' value='         " & userPage & "' size='4'>")
            
            If nextPage > pageCount Then
                pr ("<img src='rightOne_.gif' border='0'><img src='rightEnd_.gif' border='0'>")
            Else
                pr ("<a href='modUser.asp?mode=web_lst&userPage=" & nextPage & "'><img src='rightOne.gif' border='0'></a>")
                pr ("<a href='modUser.asp?mode=web_lst&userPage=" & pageCount & "'><img src='rightEnd.gif' border='0'></a>")
            End If
            
            pr ("&nbsp;of " & int(pageCount))
            pr ("&nbsp;&nbsp;&nbsp;&nbsp;<font size='1'>Total Records:&nbsp;" & int(RecCount) & "</font>")
            pr ("</td></tr>")
            
            pr ("</table>")
            pr ("</form>")
        End If
        
        pr ("</body>")
        pr ("</html>")
    End If
  
    
    If Request("mode") = "frm" Then
        Session("currUserID") = Request("UserID")
		if Request("UserID") = "0" then	'--- New
			strTitle = "Add New Logon"
			mode = "add"
			if Request("override") = "true" then
				FirstName = Session("_FirstName")
				LastName = Session("_LastName")
				Username = Session("_Username")
				Password = Session("_Password")
				emp_Access = Session("_emp_Access")
			end if
		else	'--- Edit
			strTitle = "Edit Logon"
			mode = "edit"
	        Set RS0 = Conn.Execute("SELECT * FROM  dbp_UserInfos  WHERE UserID= " & Request("UserID"))
	        If RS0.EOF Then
	            Response.Redirect "modUser.asp?mode=web_lst&msg=Status:+User+does+not+exist."
	        End If
			FirstName = RS0("FirstName")
			LastName = RS0("LastName")
			Username = RS0("Username")
			'Password = RS0("Password")
			emp_Access = GetUserDefault(Request("UserID"),"emp_Access","")
		end if
        pr ("<html>")
        pr ("<head>")
        pr ("<title>" & strTitle & "</title>")
        pr ("</head>")
        pr ("<script language='JavaScript'>")
        pr ("function javSubmit() {")
        pr ("    if (document.frm.FirstName.value.length == 0) {")
        pr ("        alert (""FirstName required."")")
        pr ("        document.frm.FirstName.focus()")
        pr ("        return")
        pr ("    }")
		
        if mode = "add" then
			pr ("    if (document.frm.Username.value.length == 0) {")
	        pr ("        alert (""Username required."")")
	        pr ("        document.frm.Username.focus()")
	        pr ("        return")
	        pr ("    }")
        end if
		
		pr ("    document.frm.submit()")
        pr ("}")
            
        pr ("function javCancel() {")
        pr ("    location.href = 'modUser.asp?mode=web_lst'")
        pr ("}")
        
		if Username = "Admin" then
	        pr ("function javDelete() {")
	        pr ("   alert('Cannot delete admin account.')")
	        pr ("}")
		else
	        pr ("function javDelete() {")
	        pr ("   if (confirm('Delete user?')) {")
	        pr ("       location.href = 'modUser.asp?mode=delete&UserID=" & Request("UserID") & "'")
	        pr ("   } else {")
	        pr ("       return")
	        pr ("   }")
	        pr ("}")
		end if
        
        pr ("</script>")
		
        pr ("<link rel='stylesheet' href='" & session("stylesheet") & "'>")
        pr ("<body background='Background.asp' onLoad='document.frm.FirstName.focus()'>")
            
		
        '***************  TitleBar
        pr ("<table border='0' width='100%' bgcolor='" & GetUserDefault(Session("web_UserID"),"usr_bgcolor","navy") & "' cellpadding='1' cellspacing='0'><tr><td>")
        pr ("<table border='0' width='100%' cellpadding='1' cellspacing='0'>")
        pr ("<tr valign='bottom'>")
        pr ("<td align='left' nowrap>")
        pr ("&nbsp;<font size='3' color='" & GetUserDefault(Session("web_UserID"),"usr_color","white") & "'><b>" & strTitle & "</b></font>")
        pr ("</td>")
        pr ("</tr>")
        pr ("</table>")
        pr ("</td></tr></table>")
    
        If Len(Request("msg")) > 0 Then
            pr ("<br><center><font color='red'><b>" & Request("msg") & "</b></font></center><br>")
        End If
        
        pr ("<form name='frm' action='modUser.asp'>")
        pr ("<input type='hidden' name='mode' value='" & mode & "'>")
        pr ("<input type='hidden' name='UserID' value='" & Request("UserID") & "'>")
            
        pr ("<table border='0'>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>FirstName</font></td>")
        pr ("<td><input type='text' name='FirstName' size='20' value='" & FirstName & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>LastName</font></td>")
        pr ("<td><input type='text' name='LastName' size='20' value='" & LastName & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>Username</font></td>")
		if mode = "add" then
        	pr ("<td><input type='text' name='Username' size='20' value='" & Username & "'></td>")
        else
			pr ("<td><font color='gray'><b>" & Username & "</b></td>")
		end if
		pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>Password</font></td>")
        pr ("<td><input type='password' name='Password' size='20' value='" & Password & "'></td>")
        pr ("</tr>")
		 
		'--- Set access types
		rayAccess = split("Read Only,Read/Write,Administrator",",")
        pr ("<tr>")
        pr ("<td align='right'><font size='2'>Access Type</font></td>")
        pr ("<td>")
        pr ("<select name='emp_Access'>")
		for i = 0 to ubound(rayAccess)
			if rayAccess(i) = emp_Access then
        		pr ("<option value='" & rayAccess(i) & "' selected>" & rayAccess(i))
			else
				pr ("<option value='" & rayAccess(i) & "'>" & rayAccess(i))
			end if
		next
        pr ("</select>")
        pr ("</td>")
        pr ("</tr>")

		
        pr ("<tr><td>&nbsp;</td><td>&nbsp;</td></tr>")

 
        pr ("<tr>")
        pr ("<td>&nbsp;</td>")
        pr ("<td>")
        pr ("<input type='button' value='Update' onClick='javSubmit()'>&nbsp;")
		if mode = "edit" then
        	pr ("<input type='button' value='Delete' onClick='javDelete()'>&nbsp;")
		end if
        pr ("<input type='button' value='Cancel' onClick='javCancel()'></td>")
        pr ("</tr>")
            
        pr ("</table>")
            
            
        pr ("</form>")
            
        pr ("</body>")
        pr ("</html>")
    End If
    
    If Request("mode") = "add" Then
        Set RS0 = Conn.Execute("SELECT * FROM dbp_Users WHERE UserName = '" & safeEntry(Request("Username")) & "'")
        If RS0.EOF Then
            strSQL = "INSERT INTO Users (FirstName,LastName,UserName,Password,LogonCount,DateCreate) VALUES ("
            strSQL = strSQL & "'" & safeEntry(Request("FirstName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("LastName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("UserName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("Password")) & "',"
			strSQL = strSQL & "" & "0" & ","
            strSQL = strSQL & "'" & Date & "')"
            Conn.Execute (strSQL)
			
            Set RSNew = Conn.Execute("SELECT Max(UserID) AS NewID FROM dbp_Users")
            Session("currUserID") = RSNew("NewID")
			
			Call SetUserDefault(RSNew("NewID"),"emp_Access",Request("emp_Access"))
			
            pg = "modUser.asp?mode=web_lst&msg=New+logon+account+added."
        Else
            Session("_FirstName") = Request("FirstName")
            Session("_LastName") = Request("LastName")
            Session("_Username") = ""
			Session("_Password") = Request("Password")
            Session("_emp_Access") = Request("emp_Access")
            pg = "modUser.asp?override=true&mode=frm&UserID=0&msg=Username+already+exists."
        End If
        Response.Redirect pg
    End If
    
    If Request("mode") = "edit" Then
		strSQL = "UPDATE dbp_Users SET "
        strSQL = strSQL & "UserFirst = '" & safeEntry(Request("FirstName")) & "',"
        strSQL = strSQL & "UserLast = '" & safeEntry(Request("LastName")) & "' "
        strSQL = strSQL & "WHERE UserID = " & Request("UserID")
 
        Conn.Execute (strSQL)
		Call SetUserDefault(Request("UserID"),"emp_Access",Request("emp_Access"))
        pg = "modUser.asp?mode=web_lst&msg=Logon+account+updated."
        Response.Redirect pg
    End If
    
    If Request("mode") = "delete" Then
        Conn.Execute ("DELETE FROM Users WHERE UserID = " & Request("UserID"))
		Conn.Execute ("DELETE FROM dbp_defaultUserSettings WHERE UserID = " & Request("UserID"))
        pg = "modUser.asp?mode=web_lst&msg=Logon+account+deleted."
        Response.Redirect pg
    End If
    
    Conn.Close
%>
