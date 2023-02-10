<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<!--#INCLUDE FILE="con_topmenu.asp"-->
<%  
set Rs=conn.execute("SELECT BaseDefault.Path FROM BaseDefault;")
Set DbMenu = Server.CreateObject("ADODB.Connection") 
DbMenu.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Rs("Path")


	session("candidat_CurrentPage") = "modUser.asp"
    If Request("mode") = "web_lst" Then
        pr ("<html>")
        pr ("<head>")
        pr ("<title>Manage Users</title>")
            
        pr ("<script language='JavaScript'>")
        pr ("function javNewUser() {")
        pr ("    location.href = 'modUser.asp?mode=frmNew&resetSession=true'")
        pr ("}")
        pr ("</script>")
        pr ("</head>")
		
	    pr ("<script language='VBscript'>")
	    pr ("function vbMouseOver(a)")
	    pr ("   a.style.backgroundcolor = ""#C1E8F7""")
	    pr ("end function")
		
	    pr ("function vbMouseOut(a)")
	    pr ("   a.style.backgroundcolor = ""#EEEEEE""")
	    pr ("end function")
		
	    pr ("function vbEditUsr(ID)")
	    pr ("   location.href = ""modUser.asp?mode=frmEdit&UserID="" & ID")
	    pr ("end function")
		
	    pr ("</script>")
		
		
        pr ("<link rel='stylesheet' href='PMainStyle1.asp'>")
        pr ("<body background='background.asp'>")

        Set RS1 = DbMenu.Execute("SELECT Count(*) as RecCount FROM Users")
        
        '***************  TitleBar
        pr ("<table border='2' width='100%' bgcolor='" & GetDefaultUser("usr_bgcolor","#678DB8") & "' cellpadding='1' cellspacing='0'><tr><td>")
        pr ("<table border='0' width='100%' cellpadding='1' cellspacing='0'>")
        pr ("<tr valign='bottom'>")
		pr ("<form name='frm'>")
        pr ("<td align='left' nowrap>")
        pr ("<img src='user.gif' border='0'>&nbsp;<font class='smallheader' color='" & GetDefaultUser("usr_color","navy")  & "'><b>User Management</b></font>")
        pr ("</td>")
        pr ("<td align='right' nowrap>")
        pr ("<input type='button'  class='cmdflat'   value='New Logon' onClick='javNewUser()'>")
        pr ("</td>")
        pr ("</tr>")
        pr ("</table>")
        pr ("</td></tr></table>")
        
        If Len(Request("msg")) > 0 Then
            pr ("<table border='0'><tr><td><font color='red'><b>" & Request("msg") & "</b></font></td></tr></table>")
        End If
        
        If Len(Request("userLetter")) > 0 Then
            session("candidat_userLetter") = Request("userLetter")
        End If
        If Request("userLetter") = "All" Then
            session("candidat_userLetter") = ""
        End If
    
        If Len(session("candidat_userLetter")) > 0 Then
            strSQL = "SELECT * FROM Users WHERE LastName LIKE '" & session("candidat_userLetter") & "%' "
            strCOUNT = "SELECT Count(*) AS RecCount FROM Users WHERE LastName LIKE '" & session("candidat_userLetter") & "%'"
        Else
            strSQL = "SELECT * FROM Users "
            strCOUNT = "SELECT Count(*) AS RecCount FROM Users "
        End If
        
        '********* Order by
        If Len(Request("UserSortBy")) > 0 Then
            session("candidat_UserSortBy") = Request("UserSortBy")
            strSQL = strSQL & " ORDER BY " & Request("UserSortBy")
        ElseIf Len(session("candidat_UserSortBy")) > 0 Then
            strSQL = strSQL & " ORDER BY " & session("candidat_UserSortBy")
        Else
            session("candidat_UserSortBy") = "UserID"
            strSQL = strSQL & " ORDER BY UserID"
        End If
    
        Set RS0 = DbMenu.Execute(strSQL)
        Set RSCount = DbMenu.Execute(strCOUNT)
        RecCount = RSCount("RecCount")
        
        pr ("<table bgcolor='#CFCEDB' cellspacing='1' width='100%'>")
        pr ("<tr>")
        pr ("<td colspan='100%' nowrap align='left'>")
        If Len(session("candidat_userLetter")) > 0 Then
            pr ("<a href='modUser.asp?mode=web_lst&userPage=1&userLetter=All'>[All]</a>&nbsp;&nbsp;")
        Else
            pr ("<a href='modUser.asp?mode=web_lst&userPage=1&userLetter=All'><font class='smallheader' color='red'><b>[All]</b></font></a>&nbsp;&nbsp;")
        End If
        
        For t = 1 To 26
            ltr = getLetter(t)
            ltr = UCase(ltr)
            If session("candidat_userLetter") = ltr Then
                pr ("<font class='smallheader' color='red'><b>" & ltr & "</b></font>&nbsp;")
            Else
                pr ("<a href='modUser.asp?mode=web_lst&userPage=1&userLetter=" & ltr & "'>" & ltr & "</a>&nbsp;")
            End If
        Next
        
        pr ("</td></tr>")
        pr ("</table>")
        
        pr ("<table bgcolor='#BBCCEC' cellspacing='1' width='100%'>")
        
        pr ("<tr bgcolor='#004080'>")
        
        If InStr(session("candidat_UserSortBy"), "UserID") Then
            If InStr(session("candidat_UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserID'><img src='SortDesc.gif' border='0'><font color='white'><b>ID</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserID+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>ID</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserID'><font color='white'><b>ID</b></font></a></td>")
        End If
        
        If InStr(session("candidat_UserSortBy"), "FirstName") Then
            If InStr(session("candidat_UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=FirstName'><img src='SortDesc.gif' border='0'><font color='white'><b>First Name</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=FirstName+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>First Name</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=FirstName'><font color='white'><b>First Name</b></font></a></td>")
        End If
        
        If InStr(session("candidat_UserSortBy"), "LastName") Then
            If InStr(session("candidat_UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LastName'><img src='SortDesc.gif' border='0'><font color='white'><b>Last Name</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LastName+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Last Name</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=LastName'><font color='white'><b>Last Name</b></font></a></td>")
        End If
        
        
        If InStr(session("candidat_UserSortBy"), "Username") Then
            If InStr(session("candidat_UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=Username'><img src='SortDesc.gif' border='0'><font color='white'><b>Username</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=Username+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Username</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=Username'><font color='white'><b>Username</b></font></a></td>")
        End If
		
        If InStr(session("candidat_UserSortBy"), "UserType") Then
            If InStr(session("candidat_UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserType'><img src='SortDesc.gif' border='0'><font color='white'><b>User Type</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserType+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>User Type</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=UserType'><font color='white'><b>User Type</b></font></a></td>")
        End If
		
        If InStr(session("candidat_UserSortBy"), "DateCreate") Then
            If InStr(session("candidat_UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateCreate'><img src='SortDesc.gif' border='0'><font color='white'><b>Create Date</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateCreate+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Create Date</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateCreate'><font color='white'><b>Create Date</b></font></a></td>")
        End If
        
        If InStr(session("candidat_UserSortBy"), "DateLastAccess") Then
            If InStr(session("candidat_UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateLastAccess'><img src='SortDesc.gif' border='0'><font color='white'><b>Last Logon</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateLastAccess+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Last Logon</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='modUser.asp?mode=web_lst&UserSortBy=DateLastAccess'><font color='white'><b>Last Logon</b></font></a></td>")
        End If
		
        If InStr(session("candidat_UserSortBy"), "LogonCount") Then
            If InStr(session("candidat_UserSortBy"), "Desc") Then
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
            session("candidat_userPage") = intPage
        End If
        
        If Len(session("candidat_userPage")) > 0 Then
            userPage = CInt(session("candidat_userPage"))
            endItem = userPage * 20
            beginItem = endItem - (20 - 1)
        End If
    
        Do While Not RS0.EOF
            i = i + 1
            If i >= beginItem And i <= endItem Then
                col = ""
	            If CStr(RS0("UserID")) = CStr(session("candidat_currUserID")) Then
	                pr ("<tr id='tr" & RS0("UserID") & "' style='background:#C1E8F7;' valign='top' OnMouseOver='vbMouseOver(tr" & RS0("UserID") & ")' OnMouseOut='vbMouseOut(tr" & RS0("UserID") & ")' OnClick='vbEditUsr(" & RS0("UserID") & ")'>")
	            Else
	                pr ("<tr id='tr" & RS0("UserID") & "' style='background:#F5F7FE;' valign='top' OnMouseOver='vbMouseOver(tr" & RS0("UserID") & ")' OnMouseOut='vbMouseOut(tr" & RS0("UserID") & ")' OnClick='vbEditUsr(" & RS0("UserID") & ")'>")
	            End If

                pr ("<td><a href='modUser.asp?mode=frmEdit&UserID=" & RS0("UserID") & "'><font class='smallertext'>" & RS0("UserID") & "</font></a></td>")
                pr ("<td><a href='modUser.asp?mode=frmEdit&UserID=" & RS0("UserID") & "'><font class='smallertext'>" & RS0("FirstName") & "</font></a></td>")
                pr ("<td><a href='modUser.asp?mode=frmEdit&UserID=" & RS0("UserID") & "'><font class='smallertext'>" & RS0("LastName") & "</font></a></td>")
                pr ("<td><a href='modUser.asp?mode=frmEdit&UserID=" & RS0("UserID") & "'><font class='smallertext'>" & RS0("UserName") & "</font></a></td>")
                pr ("<td><a href='modUser.asp?mode=frmEdit&UserID=" & RS0("UserID") & "'><font class='smallertext'>" & RS0("UserType") & "</font></a></td>")
				pr ("<td><a href='modUser.asp?mode=frmEdit&UserID=" & RS0("UserID") & "'><font class='smallertext'>" & funY2K(RS0("DateCreate")) & "</font></a></td>")
                pr ("<td><a href='modUser.asp?mode=frmEdit&UserID=" & RS0("UserID") & "'><font class='smallertext'>" & funY2K(RS0("DateLastAccess")) & "</font></a></td>")
                pr ("<td align='center'><a href='modUser.asp?mode=frmEdit&UserID=" & RS0("UserID") & "'><font class='smallertext'>" & RS0("LogonCount") & "</font></a></td>")          
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
            pr ("<form name='frmUser' action='" & objectPage & "'>")
            pr ("<input type='hidden' name='mode' value='modUser'>")
            pr ("<input type='hidden' name='sub' value='lst'>")
            brCount = 20 - iCounter
            For j = 0 To brCount
                pr ("<br>")
            Next
            prevPage = userPage - 1
            nextPage = userPage + 1
            
            pr ("<hr>")
            pr ("<table border='0' cellspacing='0'>")
            pr ("<tr><td>&nbsp;</td><td nowrap>Page:&nbsp;")
            
            '******************* Previous Page *************** -->
            
            
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
            
            pr ("&nbsp;of " & pageCount)
            pr ("&nbsp;&nbsp;&nbsp;&nbsp;<font class='td2'>Total Records:&nbsp;" & RecCount & "</font>")
            pr ("</td></tr>")
            
            pr ("</table>")
            pr ("</form>")
        End If
        
        pr ("</body>")
        pr ("</html>")
    End If
    
    If Request("mode") = "frmNew" Then
        If Request("resetSession") = "true" Then
            session("candidat_FirstName") = ""
            session("candidat_LastName") = ""
            session("candidat_Username") = ""
            session("candidat_Password") = ""
			session("candidat_UserType") = ""
        End If
        pr ("<html>")
        pr ("<head>")
        pr ("<title></title>")
        pr ("</head>")
        pr ("<script language='JavaScript'>")
        pr ("function javSubmit() {")
        pr ("    if (document.frm.FirstName.value.length == 0) {")
        pr ("        alert (""FirstName required."")")
        pr ("        document.frm.FirstName.focus()")
        pr ("        return")
        pr ("    }")
        pr ("    if (document.frm.UserName.value.length == 0) {")
        pr ("        alert (""Username required."")")
        pr ("        document.frm.UserName.focus()")
        pr ("        return")
        pr ("    }")
        pr ("    document.frm.submit()")
        pr ("}")
            
        pr ("function javCancel() {")
        pr ("    location.href = 'modUser.asp?mode=web_lst'")
        pr ("}")
        pr ("</script>")
        pr ("<link rel='stylesheet' href='PMainStyle1.asp'>")
        pr ("<body onLoad='document.frm.FirstName.focus()'>")
        
        '***************  TitleBar
        pr ("<table border='2' width='100%' bgcolor='" & GetDefaultUser("usr_bgcolor","#678DB8") & "' cellpadding='1' cellspacing='0'><tr><td>")
        pr ("<table border='0' width='100%' cellpadding='1' cellspacing='0'>")
        pr ("<tr valign='bottom'>")
		 pr ("<form name='frm' action='modUser.asp'>")
        pr ("<td align='left' nowrap>")
        pr ("&nbsp;<font class='smallheader' ><b>New Logon</b></font>")
        pr ("</td>")
        pr ("<td align='right' nowrap>")
        pr ("<input type='button'  class='cmdflat'   value='Submit' onClick='javSubmit()'>&nbsp;")
        pr ("<input type='button'  class='cmdflat'   value='Cancel' onClick='javCancel()'>&nbsp;")
        pr ("</td>")
        pr ("</tr>")
        pr ("</table>")
        pr ("</td></tr></table>")
    
        If Len(Request("msg")) > 0 Then
            pr ("<center><font color='red'><b>" & Request("msg") & "</b></font></center>")
        End If
        
       
        pr ("<input type='hidden' name='mode' value='subNew'>")
            
        pr ("<table border='0'>")
            
        pr ("<tr>")
        pr ("<td align='right'  class='smallerheader'>Prénom</td>")
        pr ("<td class='td2'><input type='text'  name='FirstName' size='20' value='" & session("candidat_FirstName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'  class='smallerheader'>Nom</td>")
        pr ("<td class='td2' ><input type='text' name='LastName' size='20' value='" & session("candidat_LastName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right' class='smallerheader'>Login</td>")
        pr ("<td class='td2'><input type='text'  name='UserName' size='20' value='" & session("candidat_UserName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'  class='smallerheader'>Mot de passe</td>")
        pr ("<td class='td2'><input type='Password' name='Password' size='20' value='" & session("candidat_Password") & "'></td>")
        pr ("</tr>")
        
        pr ("<tr>")
        pr ("<td align='right' class='smallerheader'>Profil : </td>")
        pr ("<td>")
        pr ("<select  class='tbflat'  name='UserType'>")
		pr ("<option value='Standard User' " & sel1 & ">Utilisateur")
		pr ("<option value='Administrator' " & sel2 & ">Administrateur")
        pr ("</select>")
        pr ("</td>")
        pr ("</tr>")
     
        pr ("</table>")
            
            
        pr ("</form>")
            
        pr ("</body>")
        pr ("</html>")
    End If
    
    If Request("mode") = "frmEdit" Then
        session("candidat_currUserID") = Request("UserID")
        Set RS0 = DbMenu.Execute("SELECT * FROM Users WHERE UserID= " & Request("UserID"))
        If RS0.EOF Then
            Response.Redirect "modUser.asp?mode=web_lst&msg=Status:+User+does+not+exist."
        End If
        pr ("<html>")
        pr ("<head>")
        pr ("<title></title>")
        pr ("</head>")
        pr ("<script language='JavaScript'>")
        pr ("function javSubmit() {")
        pr ("    if (document.frm.FirstName.value.length == 0) {")
        pr ("        alert (""FirstName required."")")
        pr ("        document.frm.FirstName.focus()")
        pr ("        return")
        pr ("    }")
        pr ("    document.frm.submit()")
        pr ("}")
            
        pr ("function javCancel() {")
        pr ("    location.href = 'modUser.asp?mode=web_lst'")
        pr ("}")
        
		if RS0("Username") <> "Admin" then
	        pr ("function javDelete() {")
	        pr ("   if (confirm('Delete user?')) {")
	        pr ("       location.href = 'modUser.asp?mode=subDelete&UserID=" & Request("UserID") & "'")
	        pr ("   } else {")
	        pr ("       return")
	        pr ("   }")
	        pr ("}")
		else
	        pr ("function javDelete() {")
	        pr ("   alert('Cannot delete admin account.')")
	        pr ("}")
		end if
        
        pr ("</script>")
		
        pr ("<link rel='stylesheet' href='PMainStyle1.asp'>")
        pr ("<body onLoad='document.frm.FirstName.focus()'>")
            
		
        '***************  TitleBar
        pr ("<table border='2' width='100%' bgcolor='" & GetDefaultUser("usr_bgcolor","#678DB8") & "' cellpadding='1' cellspacing='0'><tr><td>")
        pr ("<table border='0' width='100%' cellpadding='1' cellspacing='0'>")
        pr ("<tr valign='bottom'>")
        pr ("<td align='left' nowrap>")
        pr ("&nbsp;<font class='smallheader' color='" & GetDefaultUser("usr_color","navy") & "'><b>Edit Logon</b></font>")
        pr ("</td>")
        pr ("</tr>")
        pr ("</table>")
        pr ("</td></tr></table>")
    
        If Len(Request("msg")) > 0 Then
            pr ("<table border='0'><tr><td>" & Request("msg") & "</td></tr></table>")
        End If
        
        pr ("<form name='frm' action='modUser.asp' method='post'>")
        pr ("<input type='hidden' name='mode' value='subEdit'>")
        pr ("<input type='hidden' name='UserID' value='" & Request("UserID") & "'>")
            
        pr ("<table border='0'>")
            
        pr ("<tr>")
        pr ("<td align='right' class='smallerheader'>Prénom : </td>")
        pr ("<td class='td2'><input type='text' name='FirstName' size='20' value='" & RS0("FirstName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right' class='smallerheader'>Nom : </td>")
        pr ("<td class='td2'><input type='text' name='LastName' size='20' value='" & RS0("LastName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right' class='smallerheader'>Login : </td>")
        pr ("<td >" & RS0("UserName") & "</td>")
        pr ("</tr>")
            
        
		if RS0("UserType") = "Administrator" then
			sel2 = "selected"
		else
			sel1 = "selected"
		end if
        pr ("<tr>")
        pr ("<td align='right' class='smallerheader'>Profil : </td>")
        pr ("<td>")
        pr ("<select  class='tbflat'  name='UserType'>")
		pr ("<option value='Standard User' " & sel1 & ">Utilisateur")
		pr ("<option value='Administrator' " & sel2 & ">Administrateur")
        pr ("</select>")
        pr ("</td>")
        pr ("</tr>")

		
        pr ("<tr><td>&nbsp;</td><td>&nbsp;</td></tr>")

 
        pr ("<tr>")
        pr ("<td>&nbsp;</td>")
        pr ("<td>")
        pr ("<input type='button'  class='cmdflat'   value='Update' onClick='javSubmit()'>&nbsp;")
        pr ("<input type='button'  class='cmdflat'   value='Delete' onClick='javDelete()'>&nbsp;")
        pr ("<input type='button'  class='cmdflat'   value='Cancel' onClick='javCancel()'></td>")
        pr ("</tr>")
            
        pr ("</table>")
            
            
        pr ("</form>")
            
        pr ("</body>")
        pr ("</html>")
    End If
    
    If Request("mode") = "subNew" Then
        Set RS0 = DbMenu.Execute("SELECT * FROM Users WHERE UserName = '" & safeEntry(Request("Username")) & "'")
        If RS0.EOF Then
            strSQL = "INSERT INTO Users (FirstName,LastName,UserName,Password,UserType,LogonCount,DateCreate) VALUES ("
            strSQL = strSQL & "'" & safeEntry(Request("FirstName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("LastName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("UserName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("Password")) & "',"
			strSQL = strSQL & "'" & Request("UserType") & "',"
			strSQL = strSQL & "" & "0" & ","
            strSQL = strSQL & "'" & Date & "')"
            Conn.Execute (strSQL)
			
            Set RSNew = DbMenu.Execute("SELECT Max(UserID) AS NewID FROM Users")
            session("candidat_currUserID") = RSNew("NewID")
			
            pg = "modUser.asp?mode=web_lst&msg=New+record+successful."
        Else
            session("candidat_FirstName") = Request("FirstName")
            session("candidat_LastName") = Request("LastName")
            session("candidat_Username") = Request("Username")
            pg = "modUser.asp?mode=frmNew&msg=Username+already+exists."
        End If
        Response.Redirect pg
    End If
    
    If Request("mode") = "subEdit" Then
		strSQL = "UPDATE Users SET "
        strSQL = strSQL & "FirstName = '" & safeEntry(Request("FirstName")) & "',"
        strSQL = strSQL & "LastName = '" & safeEntry(Request("LastName")) & "', "
		strSQL = strSQL & "UserType = '" & Request("UserType") & "' "
        strSQL = strSQL & "WHERE UserID = " & Request("UserID")
        DbMenu.Execute (strSQL)
        pg = "modUser.asp?mode=web_lst&msg=Update+successful."
        Response.Redirect pg
    End If
    
    If Request("mode") = "subDelete" Then
		Set RS0 = DbMenu.Execute("SELECT * FROM Users WHERE UserID = " & Request("UserID"))
        if not RS0.EOF then
            DbMenu.Execute ("DELETE FROM Users WHERE UserID = " & Request("UserID"))
			Conn.Execute ("DELETE FROM DefaultUsers WHERE UserID = " & Request("UserID"))
			Conn.Execute ("DELETE FROM con_contacts WHERE UserID = " & Request("UserID"))
            pg = "modUser.asp?mode=web_lst&msg=Delete+successful."
		else
			pg = "modUser.asp?mode=web_lst&msg=User+deleted+already."
		end if
        Response.Redirect pg
    End If
    DbMenu.Close
    conn.close
    set DbMenu=nothing
    set conn=nothing
%>
