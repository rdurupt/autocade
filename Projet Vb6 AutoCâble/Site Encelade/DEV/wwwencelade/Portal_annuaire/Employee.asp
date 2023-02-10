<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<% 
Conn.close
Call Main()


Public Sub Main()
    'Set Application = ASPsc.Application
    'Set Request = ASPsc.Request
    'Set Session = ASPsc.Session
    'Set Response = ASPsc.Response
    'Set Server = ASPsc.Server
    
    '**************************  License ****
    Session("IsEvaluation") = "true"
    Session("LicensedTo") = "Sam Hurdowar"
    Session("ProductID") = "EMP-09292"
    Session("DateLicensed") = "5/6/1999"
    '******************************************
	
    If Request("mode") = "lst" Then
        Call lst
    ElseIf Request("mode") = "chtOrganization" Then
        Call chtOrganization
    ElseIf Request("mode") = "subQuery" Then
        Call subQuery
    ElseIf Request("mode") = "modLogon" Then
        Call modLogon
    ElseIf Request("mode") = "modUser" Then
        Call modUser
    ElseIf Request("mode") = "subSetting" Then
        Call subSetting
    ElseIf Request("mode") = "dlgDate" Then
        Call dlgDate
    ElseIf Request("mode") = "dlgAbout" Then
        Call dlgAbout
    ElseIf Request("mode") = "subEmp" Then
        Call subEmp
    End If
End Sub



Public Function subEmp()
    Response.Expires = 0
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Mode = 3
    Conn.Open Session("ADOEmployee")
	
	pg = "Employee.asp?mode=lst"
	if Request("sub") = "edit" then
		SearchBlob = ""
		strSQL = "UPDATE Employees SET "
		Set RS1 = Conn.Execute("SELECT * FROM Employees WHERE UserID = " & Request("UserID"))
		if not RS1.EOF then
			for i = 0 to RS1.fields.count-1
				if RS1(i).name <> "UserID" and RS1(i).name <> "SearchBlob" then
					if RS1(i).name = "BossID" then		'Numeric
						strSQL = strSQL & RS1(i).name & " = " & safeEntry(Request(RS1(i).name)) & ","
					else
						strSQL = strSQL & RS1(i).name & " = '" & safeEntry(Request(RS1(i).name)) & "',"
						SearchBlob = SearchBlob & " " & safeEntry(Request(RS1(i).name))
					end if
				end if
			next
			strSQL = strSQL & "SearchBlob = '" & SearchBlob & "' WHERE UserID = " & Request("UserID")
		    Conn.Execute(strSQL)
			pg = "Employee.asp?mode=lst&msg=Status:+Employee+updated."
	   	end if
	end if
    
	if Request("sub") = "new" then
		SearchBlob = ""
		strINSERT = "INSERT INTO Employees (BossID"
		strVALUES = " VALUES(" & Request("BossID")
		Set RS1 = Conn.Execute("SELECT * FROM defFields ORDER BY FieldID")
		do while not RS1.EOF
			if len(RS1("FieldAlias")) > 0 then
				strINSERT = strINSERT & "," & RS1("FieldName")
				strVALUES = strVALUES & ",'" & safeEntry(Request(RS1("FieldName"))) & "'"
				SearchBlob = SearchBlob & " " & safeEntry(Request(RS1("FieldName")))
			end if
			RS1.movenext
		loop
		strINSERT = strINSERT & ",SearchBlob) "
		strVALUES = strVALUES & ",'" & SearchBlob & "')"
		strSQL = strINSERT & " " & strVALUES
		Conn.Execute(strSQL)
		
		Set RS1 = Conn.Execute("SELECT Max(UserID) AS NewID FROM Employees")
		Session("currUserID") = RS1("NewID")
		pg = "Employee.asp?mode=lst&msg=Status:+Employee+added."
	end if
	
	if Request("sub") = "delete" then
		strSQL = strSQL & "DELETE FROM Employees WHERE UserID = " & Request("UserID")
	    Conn.Execute(strSQL)
		pg = "Employee.asp?mode=lst&msg=Status:+Employee+deleted."
	end if
	
    Conn.Close
    Response.Redirect pg
End Function

Public Function chtOrganization()
	Set Conn = Server.CreateObject("ADODB.Connection")  
	Conn.Mode = 3
	Conn.Open Session("ADOEmployee") 

	if len(Request("sub")) > 0 then
		Call SetDefaultUser("OrgListType",Request("sub"))
	end if
	pr ("<html>")
	pr ("<head>")
	pr ("<style>")
	pr ("A {text-decoration:none; font-family : Arial;font-size:10pt;}")
	pr ("A:Active {text-decoration:none; font-family : Arial;font-size:10pt;}")
	pr ("A:Visited {text-decoration:none; font-family : Arial;font-size:10pt;}")
	pr ("A:Hover {text-decoration:none; color : Red;font-family : Arial;font-size:10pt;}")
	pr ("</style>")
	pr ("<title>Organizational Chart</title>")
	pr ("</head>")
	
	
	pr ("<body leftmargin='0'>")
	
    '******** TitleBar
    pr ("<table width='100%' border='0' bgcolor='" & Session("TitleBarColor") & "' cellpadding='1' cellspacing='0'>")
    pr ("<tr valign='bottom'>")
    pr ("<td>")
    pr ("<table width='100%' border='0' cellpadding='1' cellspacing='0'>")
    pr ("<tr valign='bottom'>")
	
    pr ("<td align='left'>")
    pr ("<img src='chart.gif' border='0'>")
    pr ("<font size='4' color='" & Session("TitleTextColor") & "'>&nbsp;<b>Organizational Chart</b></font>")
    pr ("</td>")
	
    pr ("<td align='right'>")
	pr ("<a href='Employee.asp?mode=chtOrganization&sub=name'><b>[List by Name]</b></a>")
	pr ("<a href='Employee.asp?mode=chtOrganization&sub=title'><b>[List by Title]</b></a>")
    pr ("</td>")
	
    pr ("</tr>")
    pr ("</table>")
    pr ("</td>")
    pr ("</tr>")
    pr ("</table>")
	

	'**** Get the bosses of UserID
	Session("clickOnNode") = ""
	Session("IndexCount") = 0
	Session("lstBossID") = ""
	if len(Request("UserID")) > 0 then
		Set RS0 = Conn.Execute("SELECT * FROM Employees WHERE UserID = " & Request("UserID")) 
		if not RS0.EOF then
			BossID = RS0("BossID")
			Session("lstBossID") = ":" & BossID & ":"
			brk = 0
			do while BossID <> 0 and brk < 50
				Set RS1 = Conn.Execute("SELECT * FROM Employees WHERE UserID = " & BossID) 
				if not RS1.EOF then
					BossID = RS1("BossID")
					Session("lstBossID") = Session("lstBossID") & ":" & BossID & ":"
				else
					BossID = 0
				end if
				brk = brk + 1
			loop
		end if
	end if
	
	pr ("<br>")
	pr ("<script src='uis_tree.js'></script>")
	pr ("<script language='javascript'>")
	
	pr ("foldersTree = gFld(""<font color='navy'><b>" & GetDefault("CompanyName","DemoTech, Inc.") & "</b></font>"", """")")
	
	Set RS0 = Conn.Execute("SELECT * FROM Employees WHERE BossID = 0 ORDER BY LastName")  
	do while not RS0.EOF
		Session("IndexCount") = Session("IndexCount") + 1

		Set RS2 = Conn.Execute("SELECT Count(*) as RecCount FROM Employees WHERE BossID = " & RS0("UserID")) 
		if GetDefaultUser("OrgListType", "name") = "name" then
			strName = RS0("LastName") & ", " & RS0("FirstName") & " - "
			if len(strName) < 7 then
				strLINK = "<font color='red'>OPEN</font> - " & RS0("Title")
			else
				strLINK = strName & RS0("Title")
			end if
		else
			strName = " - " & RS0("LastName") & ", " & RS0("FirstName") 
			if len(strName) < 7 then
				strLINK = RS0("Title") & " - <font color='red'>OPEN</font>" 
			else
				strLINK = RS0("Title") & strName
			end if
		end if
		if RS2("RecCount") > 0 then
			if instr(Session("lstBossID"), ":" & RS0("UserID") & ":") > 0 then
				Session("clickOnNode") = Session("clickOnNode") & vbcrlf & "clickOnNode(" & Session("IndexCount") & ")"
			end if
			pr ("n" & RS0("UserID") & " = insFld(foldersTree, gFld(""<a href='javascript:javEmp(" & RS0("UserID") & ")'>" & strLINK & "</a>"", """"))")
			Call getChildren(RS0("UserID"))
		else
			imageFile = "userEnd.gif"	
			javLink = "javascript:javEmp(" & RS0("UserID") & ")"
			pr ("insDoc(foldersTree, gLnk(""" & imageFile & """,""" & javLink & """,""<a href='javascript:javEmp(" & RS0("UserID") & ")'>" & strLINK & "</a>""))")
		end if
		RS0.movenext 
	loop 
	
	pr ("</script>")
	pr ("<script language='javascript'>")
	pr ("initializeDocument()")
	'pr ("clickOnNode(2)")
	pr (Session("clickOnNode"))
	pr ("function javEmp(ID) {")
	pr ("	window.open('dlgEmp.asp?UserID=' + ID,'dlgemp','resizable=yes,status=no,top=75,left=400,width=500,height=400')")
	pr ("}")
	pr ("</script>")
	
	pr ("</body>")
	pr ("</html>")
	Conn.Close 

end function

Function getChildren(ParentID) 
	Set Conn1 = Server.CreateObject("ADODB.Connection")  
	Conn1.Mode = 3
	Conn1.Open Session("ADOEmployee") 
	strSQL = "SELECT * FROM Employees WHERE BossID = " & ParentID & " ORDER BY LastName"  
	Set RS00 = Conn1.Execute(strSQL) 
	Do While Not RS00.EOF 
		Session("IndexCount") = Session("IndexCount") + 1
		if GetDefaultUser("OrgListType", "name") = "name" then
			strName = RS00("LastName") & ", " & RS00("FirstName") & " - "
			if len(strName) < 7 then
				strLINK = "<font color='red'>OPEN</font> - " & RS00("Title")
			else
				strLINK = strName & RS00("Title")
			end if
		else
			strName = " - " & RS00("LastName") & ", " & RS00("FirstName") 
			if len(strName) < 7 then
				strLINK = RS00("Title") & " - <font color='red'>OPEN</font>" 
			else
				strLINK = RS00("Title") & strName
			end if
		end if
		
		Set RS22 = Conn1.Execute("SELECT Count(*) as RecCount FROM Employees WHERE BossID = " & RS00("UserID")) 
		if RS22("RecCount") > 0 then 
			if instr(Session("lstBossID"), ":" & RS00("UserID") & ":") > 0 then
				Session("clickOnNode") = Session("clickOnNode") & vbcrlf & "clickOnNode(" & Session("IndexCount") & ")"
			end if
			pr ("n" & RS00("UserID") & " = insFld(n" & ParentID & ", gFld(""<a href='javascript:javEmp(" & RS00("UserID") & ")'>" & strLINK & "</a>"", """"))")
			Call getChildren(RS00("UserID"))
		else
			imageFile = "userEnd.gif"	
			javLink = "javascript:javEmp(" & RS00("UserID") & ")"
			pr ("insDoc(n" & ParentID & ", gLnk(""" & imageFile & """,""" & javLink & """,""<a href='javascript:javEmp(" & RS00("UserID") & ")'>" & strLINK & "</a>""))")
		end if 
	    RS00.Movenext 
	Loop 
	Conn1.close
end function 

Public Function lst()
    Response.Expires = 0
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Mode = 3
    Conn.Open Session("ADOEmployee")

    If Len(Session("TitleBarColor")) < 2 Then
        Session("TitleBarColor") = GetDefaultUser("TitleBarColor","#6CBFD0")
        Session("TitleTextColor") = GetDefaultUser("TitleTextColor","navy")
    End If
    
    pr ("<html>")
    pr ("<head>")
    pr ("<title>" & GetDefault("AppTitle","Employee Manager") & "</title>")
    
    pr ("</head>")
    pr ("<script language='javascript'>")
    pr ("function javResetSearch() {")
    pr ("   location.href = 'Employee.asp?mode=lst&reset=true'")
    pr ("}")

    pr ("</script>")
	
    pr ("<script language='VBscript'>")
    pr ("function vbMouseOver(a)")
    pr ("   a.style.backgroundcolor = ""#C1E8F7""")
    pr ("end function")
	
    pr ("function vbMouseOut(a)")
    pr ("   a.style.backgroundcolor = ""#F5F7FE""")
    pr ("end function")
	
    pr ("function vbEditEmp(ID)")
    pr ("   location.href = ""frmEmp.asp?UserID="" & ID")
    pr ("end function")
	
    pr ("</script>")
	
    pr ("<link rel='stylesheet' href='StyleSheet.css'>")
    pr ("<body>")
    
	
	
    If Len(Request("employeeSearch")) > 0 Then
        Session("employeePage") = ""
		Session("employeeQuery") = ""
        Session("employeeSearch") = Request("employeeSearch")
    End If
    If Len(Request("employeeQuery")) > 0 Then
        Session("employeePage") = ""
		Session("employeeSearch") = ""
        Session("employeeQuery") = Request("employeeQuery")
		Session("qryFirstName") = Request("qryFirstName")
		Session("qryLastName") = Request("qryLastName")
		Session("qryTitle") = Request("qryTitle")
    End If
	
    If Len(Request("employeeLetter")) > 0 Then
        Session("employeePage") = ""
        Session("employeeLetter") = Request("employeeLetter")
        Session("employeeSearch") = ""
		Session("employeeQuery") = ""
    End If
    If Request("employeeLetter") = "All" Then
        Session("employeePage") = ""
		Session("employeeSearch") = ""
		Session("employeeQuery") = ""
        Session("employeeLetter") = ""
    End If
    If Request("reset") = "true" Then
        Session("employeePage") = ""
        Session("employeeSearch") = ""
		Session("employeeQuery") = ""
        Session("employeeLetter") = ""
    End If

	CaptionStr = "<font size='3' color='" & Session("TitleTextColor") & "'>Employee List</font>"
	'**** Get select string
	strSQL = "SELECT UserID"
	Set RS1 = Conn.Execute("SELECT defFields.FieldName FROM defFields,Fields WHERE defFields.FieldID = Fields.FieldID AND UserID = " & Session("webUserID") & " ORDER BY Fields.FieldOrder,Fields.FieldID") 
	if RS1.EOF then
		strSQL = "SELECT UserID,FirstName,LastName,Title,Phone,Extension,Email"
	end if
	do while not RS1.EOF
		strSQL = strSQL & "," & RS1("FieldName") 
		RS1.movenext
	loop
	strSQL = strSQL & " FROM Employees WHERE UserID <> null "
    sqlCount = "SELECT Count(*) as RecCount FROM Employees WHERE UserID <> null "
    
    '***** Search
    If Len(Session("employeeSearch")) > 0 Then
        strSQL = strSQL & "AND (SearchBlob Like '%" & Session("employeeSearch") & "%') "
        sqlCount = sqlCount & "AND (SearchBlob Like '%" & Session("employeeSearch") & "%') "
    
        CaptionStr = "<font size='3' color='" & Session("TitleTextColor") & "'>Search Results for: '" & Session("employeeSearch") & "'</font>"
    End If
    
	
	'****** Letter
    If Len(Session("employeeLetter")) > 0 Then
        strSQL = strSQL & " AND (LastName LIKE '" & Session("employeeLetter") & "%') "
        sqlCount = sqlCount & " AND (LastName LIKE '" & Session("employeeLetter") & "%') "
		CaptionStr = "<font size='3' color='" & Session("TitleTextColor") & "'>Last Name Beginning With: '" & Session("employeeLetter") & "'</font>"
    End If
    
	
	'****** Query
    If Len(Session("employeeQuery")) > 0 Then
		CaptionStr = "<font size='3' color='" & Session("TitleTextColor") & "'>QUERY:</font> "
		cutit = false
		if len(Session("qryFirstName")) > 0 then
	        strSQL = strSQL & " AND (FirstName LIKE '" & Session("qryFirstName") & "%') "
	        sqlCount = sqlCount & " AND (FirstName LIKE '" & Session("qryFirstName") & "%') "
			CaptionStr = CaptionStr & "<font size='2' color='" & Session("TitleTextColor") & "'>First Name Like '" & Session("qryFirstName") & "*'</font> and "
			cutit = true
		end if
		if len(Session("qryLastName")) > 0 then
	        strSQL = strSQL & " AND (LastName LIKE '" & Session("qryLastName") & "%') "
	        sqlCount = sqlCount & " AND (LastName LIKE '" & Session("qryLastName") & "%') "
			CaptionStr = CaptionStr & "<font size='2' color='" & Session("TitleTextColor") & "'>Last Name Like '" & Session("qryLastName") & "*'</font> and "
			cutit = true
		end if
		if len(Session("qryTitle")) > 0 then
	        strSQL = strSQL & " AND (Title LIKE '" & Session("qryTitle") & "%') "
	        sqlCount = sqlCount & " AND (Title LIKE '" & Session("qryTitle") & "%') "
			CaptionStr = CaptionStr & "<font size='2' color='" & Session("TitleTextColor") & "'>Title Like '" & Session("qryTitle") & "*'</font> and "
			cutit = true
		end if
		if cutit = true then 
			CaptionStr = left(CaptionStr, len(CaptionStr) - 4)
		end if
    End If
	
    '********* Order by
    If Len(Request("employeeSortBy")) > 0 Then
        Session("employeeSortBy") = Request("employeeSortBy")
        strSQL = strSQL & " ORDER BY " & Request("employeeSortBy")
    ElseIf Len(Session("employeeSortBy")) > 0 Then
        strSQL = strSQL & " ORDER BY " & Session("employeeSortBy")
    Else
        Session("employeeSortBy") = "FirstName"
        strSQL = strSQL & " ORDER BY FirstName"
    End If
    
    Set RS0 = Conn.Execute(strSQL)
    
    '*********** Get record count
    Set RS11 = Conn.Execute(sqlCount)
    RecCount = RS11("RecCount")
    
	'*********  Menu
    Call GetMenu()
	
	
    '******** TitleBar
    pr ("<table width='100%' border='0' bgcolor='" & Session("TitleBarColor") & "' cellpadding='1' cellspacing='0'>")
    pr ("<tr valign='bottom'>")
    pr ("<td>")
    pr ("<table width='100%' border='0' cellpadding='1' cellspacing='0'>")
    pr ("<tr valign='bottom'>")
    pr ("<form name='frmSearch' action='Employee.asp'>")
    pr ("<input type='hidden' name='mode' value='lst'>")
    pr ("<td align='left'>")
    pr ("&nbsp;<b>" & CaptionStr & "</b>")
    pr ("</td>")
    pr ("<td align='right'>")
    pr ("<input type='text' name='employeeSearch' size='20' value='" & Session("employeeSearch") & "'>")
    pr ("<input type='submit' value='Search'>&nbsp;")
    pr ("</td>")
    pr ("</tr>")
    pr ("</table>")
    pr ("</td>")
    pr ("</tr>")
    pr ("</table>")
    
    '***************** MESSAGE **********
    If Len(Request("msg")) > 0 Then
        pr ("<table bgcolor='#F8FCA7' width='100%' border='0'><tr><td><font size='-1' color='maroon'>" & Request("msg") & "</font></td></tr></table>")
    End If
    
	
	'**************************  EVALUATION TOGGLE ************
	if Session("IsEvaluation") = "true" then
		Set rs7 = Conn.Execute("SELECT Count(*) AS RecCount FROM Employees")
		if rs7("RecCount") > 35 then
        	pr ("<table bgcolor='#F8FCA7' width='100%' border='0'><tr><td>")
			pr ("<font size='-1' color='maroon'>")
			pr ("Evaluation version.&nbsp;&nbsp;")
			pr ("Please purchase at <a href='http://www.ASPintranet.com' target='_top'>")
			pr ("www.ASPintranet.com</a> or email <u>sam_hurdowar@yahoo.com</u></font>")
			pr ("</td></tr></table>")
		end if
	end if
	
    '***************** Command Bar **********
    pr ("<table bgcolor='#CFCEDB' width='100%'>")
    pr ("<tr>")
    pr ("<td nowrap>")
    
    If Len(Session("employeeLetter")) > 0 Then
        pr ("<a href='Employee.asp?mode=lst&employeeLetter=All'><font size='2'>[All]</font></a>&nbsp;")
    Else
        pr ("<a href='Employee.asp?mode=lst&employeeLetter=All'><font size='3' color='red'><b>[All]</b></font></a>&nbsp;")
    End If
    
    For t = 1 To 26
        ltr = getLetter(t)
        ltr = UCase(ltr)
        If Session("employeeLetter") = ltr Then
            pr ("<font size='3' color='red'><b>" & ltr & "</b></font>&nbsp;")
        Else
            pr ("<a href='Employee.asp?mode=lst&employeeLetter=" & ltr & "'><font size='2'>" & ltr & "</font></a>&nbsp;")
        End If
    Next
    
    pr ("</td></tr>")
    pr ("</table>")
    
	
	pr ("<table border='0' width='100%' cellspacing='1' bgcolor='silver'>")
    
	'************** Headers ******
    pr ("<tr bgcolor='#FC8D7A'>")
	for i = 1 to (RS0.fields.count-1)    
        If InStr(Session("employeeSortBy"), RS0(i).name) Then
            If InStr(Session("employeeSortBy"), "Desc") Then
                pr ("<td nowrap><a href='Employee.asp?mode=lst&employeeSortBy=" & RS0(i).name & "'><img src='SortDesc.gif' border='0'><b>" & getHeading(RS0(i).name) & "</b></a></td>")
            Else
                pr ("<td nowrap><a href='Employee.asp?mode=lst&employeeSortBy=" & RS0(i).name & "+Desc'><img src='SortAss.gif' border='0'><b>" & getHeading(RS0(i).name) & "</b></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='Employee.asp?mode=lst&employeeSortBy=" & RS0(i).name & "'><b>" & getHeading(RS0(i).name) & "</b></a></td>")
        End If
    next
    pr ("</tr>")
    
    '************  Initiate page counter
    currPageNum = GetDefaultUser("PageNum","15")
    currPageNum = CInt(currPageNum)
    employeePage = 1
    endItem = currPageNum
    beginItem = 1
    i = 0
    iCounter = 0
    
    intPage = Trim(Request("employeePage"))
    If Len(intPage) > 0 And IsNumeric(intPage) Then
        Session("employeePage") = intPage
    End If
    
    If Len(Session("employeePage")) > 0 Then
        employeePage = CInt(Session("employeePage"))
        endItem = employeePage * currPageNum
        beginItem = endItem - (currPageNum - 1)
    End If
    
    Do While Not RS0.EOF
        i = i + 1
        If i >= beginItem And i <= endItem Then
            If CStr(RS0("UserID")) = CStr(Session("currUserID")) Then
                pr ("<tr id='tr" & RS0("UserID") & "' style='background:#C1E8F7;' valign='top' OnMouseOver='vbMouseOver(tr" & RS0("UserID") & ")' OnMouseOut='vbMouseOut(tr" & RS0("UserID") & ")' OnClick='vbEditEmp(" & RS0("UserID") & ")'>")
            Else
                pr ("<tr id='tr" & RS0("UserID") & "' style='background:#F5F7FE;' valign='top' OnMouseOver='vbMouseOver(tr" & RS0("UserID") & ")' OnMouseOut='vbMouseOut(tr" & RS0("UserID") & ")' OnClick='vbEditEmp(" & RS0("UserID") & ")'>")
            End If
            
			for i = 1 to (RS0.fields.count-1)  
            	pr ("<td nowrap><a href='frmEmp.asp?UserID=" & RS0("UserID") & "'>" & RS0(i) & "</a></td>")	
			next

            pr ("</tr>")
            iCounter = iCounter + 1
        End If
        If i = endItem Then
            Exit Do
        End If
    
        RS0.MoveNext
    Loop
    
    pr ("</table>")
    pr ("</form>")
    '***********************************  Page Count  *****************
    pageCount = RecCount / currPageNum
    
    If RecCount = 0 Then
        pageCount = 0
    ElseIf pageCount < 1 Then
        pageCount = 1
    ElseIf InStr(pageCount, ".") Then
        intLeft = Left(pageCount, InStr(pageCount, "."))
        pageCount = intLeft + 1
    End If
    
    If iCounter > 0 Then
        pr ("<form name='frmUser' action='Employee.asp'>")
        pr ("<input type='hidden' name='mode' value='modUser'>")
        pr ("<input type='hidden' name='sub' value='lst'>")
        brCount = currPageNum - iCounter
        For j = 0 To brCount
            pr ("<br>")
        Next
        prevPage = employeePage - 1
        nextPage = employeePage + 1
        
        pr ("<hr>")
        pr ("<table border='0' cellspacing='0'>")
        pr ("<tr><td>&nbsp;</td><td nowrap>Page:&nbsp;")
        
        '******************* Previous Page *************** -->
        If prevPage < 1 Then
            pr ("<img src='leftEnd_.gif' border='0'><img src='leftOne_.gif' border='0'>")
        Else
            pr ("<a href='Employee.asp?mode=lst&employeePage=1'><img src='leftEnd.gif' border='0'></a>")
            pr ("<a href='Employee.asp?mode=lst&employeePage=" & prevPage & "'><img src='leftOne.gif' border='0'></a>")
        End If
        
        pr ("<input align='right' type='text' name='employeePage' value='         " & employeePage & "' size='4'>")
        
        If nextPage > pageCount Then
            pr ("<img src='rightOne_.gif' border='0'><img src='rightEnd_.gif' border='0'>")
        Else
            pr ("<a href='Employee.asp?mode=lst&employeePage=" & nextPage & "'><img src='rightOne.gif' border='0'></a>")
            pr ("<a href='Employee.asp?mode=lst&employeePage=" & pageCount & "'><img src='rightEnd.gif' border='0'></a>")
        End If
        
        pr ("&nbsp;of " & pageCount)
        pr ("<font size='1'>&nbsp;&nbsp;&nbsp;&nbsp;Total Records:&nbsp;" & RecCount & "</font>")
        pr ("</td></tr>")
        
        pr ("</table>")
        pr ("</form>")
    End If
    
    
    pr ("</body>")
    pr ("</html>")
    Conn.Close
End Function



Public Function modUser()
    Response.Expires = 0
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Mode = 3
    Conn.Open Session("ADOEmployee")
    objectPage = "Employee.asp"
    
    TitleBarColor = Session("TitleBarColor")
    TitleTextColor = Session("TitleTextColor")
    If Len(Session("TitleBarColor")) < 2 Then
        TitleBarColor = "#DFE9EC"
        TitleTextColor = "#004080"
    End If
    
    If Request("sub") = "lst" Then
		Session("thisGroupID") = ""
		Session("thisLogoName") = ""
        pr ("<html>")
        pr ("<head>")
        pr ("<title>Manage Users</title>")
            
        pr ("<script language='JavaScript'>")
        pr ("function javNewUser() {")
        pr ("    location.href = '" & objectPage & "?mode=modUser&sub=frmNew&resetSession=true'")
        pr ("}")
        pr ("function javCancel() {")
        pr ("    location.href = '" & objectPage & "?mode=lst'")
        pr ("}")
        pr ("</script>")
        pr ("</head>")
        pr ("<link rel='stylesheet' href='StyleSheet.css'>")
        pr ("<body>")
        
		Call GetMenu()
		
        Set RS1 = Conn.Execute("SELECT Count(*) as RecCount FROM Users")
        
        '***************  TitleBar
        pr ("<table border='0' width='100%' bgcolor='" & TitleBarColor & "' cellpadding='1' cellspacing='0'><tr><td>")
        pr ("<table border='0' width='100%' cellpadding='1' cellspacing='0'>")
        pr ("<tr valign='bottom'>")
        pr ("<td align='left' nowrap>")
        pr ("<img src='user.gif' border='0'>&nbsp;<font size='3' color='" & TitleTextColor & "'><b>User Management</b></font>")
        pr ("</td>")
        pr ("<form name='frm'>")
        pr ("<input type='hidden' name='sub' value='frmNew'>")
        pr ("<td align='right'>")
        pr ("<input type='button' value='New User' onClick='javNewUser()'>")
        pr ("</td>")
        pr ("</form>")
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
            strSQL = "SELECT * FROM Users WHERE LastName LIKE '" & Session("userLetter") & "%' "
            strCOUNT = "SELECT Count(*) AS RecCount FROM Users WHERE LastName LIKE '" & Session("userLetter") & "%'"
        Else
            strSQL = "SELECT * FROM Users "
            strCOUNT = "SELECT Count(*) AS RecCount FROM Users "
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
            pr ("<a href='" & objectPage & "?mode=modUser&sub=lst&userPage=1&userLetter=All'>[All]</a>&nbsp;&nbsp;")
        Else
            pr ("<a href='" & objectPage & "?mode=modUser&sub=lst&userPage=1&userLetter=All'><font size='3' color='red'><b>[All]</b></font></a>&nbsp;&nbsp;")
        End If
        
        For t = 1 To 26
            ltr = getLetter(t)
            ltr = UCase(ltr)
            If Session("userLetter") = ltr Then
                pr ("<font size='3' color='red'><b>" & ltr & "</b></font>&nbsp;")
            Else
                pr ("<a href='" & objectPage & "?mode=modUser&sub=lst&userPage=1&userLetter=" & ltr & "'>" & ltr & "</a>&nbsp;")
            End If
        Next
        
        pr ("</td></tr>")
        pr ("</table>")
        
        pr ("<table bgcolor='#BBCCEC' cellspacing='1' width='100%'>")
        
        pr ("<tr bgcolor='#004080'>")
        
        If InStr(Session("UserSortBy"), "UserID") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=UserID'><img src='SortDesc.gif' border='0'><font color='white'><b>ID</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=UserID+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>ID</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=UserID'><font color='white'><b>ID</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "FirstName") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=FirstName'><img src='SortDesc.gif' border='0'><font color='white'><b>First Name</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=FirstName+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>First Name</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=FirstName'><font color='white'><b>First Name</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "LastName") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=LastName'><img src='SortDesc.gif' border='0'><font color='white'><b>Last Name</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=LastName+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Last Name</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=LastName'><font color='white'><b>Last Name</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "AccessLevel") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=AccessLevel'><img src='SortDesc.gif' border='0'><font color='white'><b>Access Level</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=AccessLevel+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Access Level</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=AccessLevel'><font color='white'><b>Access Level</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "Username") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=Username'><img src='SortDesc.gif' border='0'><font color='white'><b>Username</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=Username+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Username</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=Username'><font color='white'><b>Username</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "DateCreate") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=DateCreate'><img src='SortDesc.gif' border='0'><font color='white'><b>Create Date</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=DateCreate+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Create Date</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=DateCreate'><font color='white'><b>Create Date</b></font></a></td>")
        End If
        
        If InStr(Session("UserSortBy"), "DateLastAccess") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=DateLastAccess'><img src='SortDesc.gif' border='0'><font color='white'><b>Last Logon</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=DateLastAccess+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Last Logon</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=DateLastAccess'><font color='white'><b>Last Logon</b></font></a></td>")
        End If
		
        If InStr(Session("UserSortBy"), "LogonCount") Then
            If InStr(Session("UserSortBy"), "Desc") Then
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=LogonCount'><img src='SortDesc.gif' border='0'><font color='white'><b>Log Count</b></font></a></td>")
            Else
                pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=LogonCount+Desc'><img src='SortAss.gif' border='0'><font color='white'><b>Log Count</b></font></a></td>")
            End If
        Else
            pr ("<td nowrap><a href='" & objectPage & "?mode=modUser&sub=lst&UserSortBy=LogonCount'><font color='white'><b>Log Count</b></font></a></td>")
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
                    col = "white"
                    pr ("<tr bgcolor='navy' valign='top'>")
                Else
                    pr ("<tr bgcolor='#FBF9FF' valign='top'>")
                End If
            
				employeeAccess = 1
				Set RSUser = Conn.Execute("SELECT defValue FROM dbp_defaultUserSettings WHERE defName = 'employeeAccess' AND UserID = " & RS0("UserID"))
                if not RSUser.EOF then
					employeeAccess = cInt(RSUser("defValue"))
				end if
				If employeeAccess = 1 Then
                    AccessLevel = "Read Only"
                ElseIf employeeAccess = 2 Then
                    AccessLevel = "Read/Write"
                ElseIf employeeAccess = 3 Then
                    AccessLevel = "Admin"
                Else
                    AccessLevel = "Undetermined"
                End If
                
                pr ("<td><a href='" & objectPage & "?mode=modUser&sub=frmEdit&UserID=" & RS0("UserID") & "'><font size='1' color='" & col & "'>" & RS0("UserID") & "</font></a></td>")
                pr ("<td><a href='" & objectPage & "?mode=modUser&sub=frmEdit&UserID=" & RS0("UserID") & "'><font size='1' color='" & col & "'>" & RS0("FirstName") & "</font></a></td>")
                pr ("<td><a href='" & objectPage & "?mode=modUser&sub=frmEdit&UserID=" & RS0("UserID") & "'><font size='1' color='" & col & "'>" & RS0("LastName") & "</font></a></td>")
                pr ("<td><a href='" & objectPage & "?mode=modUser&sub=frmEdit&UserID=" & RS0("UserID") & "'><font size='1' color='" & col & "'>" & AccessLevel & "</font></a></td>")
                pr ("<td><a href='" & objectPage & "?mode=modUser&sub=frmEdit&UserID=" & RS0("UserID") & "'><font size='1' color='" & col & "'>" & RS0("UserName") & "</font></a></td>")
                pr ("<td><a href='" & objectPage & "?mode=modUser&sub=frmEdit&UserID=" & RS0("UserID") & "'><font size='1' color='" & col & "'>" & funY2K(RS0("DateCreate")) & "</font></a></td>")
                pr ("<td><a href='" & objectPage & "?mode=modUser&sub=frmEdit&UserID=" & RS0("UserID") & "'><font size='1' color='" & col & "'>" & funY2K(RS0("DateLastAccess")) & "</font></a></td>")
                pr ("<td align='center'><a href='" & objectPage & "?mode=modUser&sub=frmEdit&UserID=" & RS0("UserID") & "'><font size='1' color='" & col & "'>" & RS0("LogonCount") & "</font></a></td>")          
                pr ("</tr>")
                iCounter = iCounter + 1
            End If
            If i = endItem Then
                Exit Do
            End If
            RS0.movenext
        Loop
            
        pr ("</table>")
        
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
                pr ("<a href='" & objectPage & "?mode=modUser&sub=lst&userPage=1'><img src='leftEnd.gif' border='0'></a>")
                pr ("<a href='" & objectPage & "?mode=modUser&sub=lst&userPage=" & prevPage & "'><img src='leftOne.gif' border='0'></a>")
            End If
            
            pr ("<input align='right' type='text' name='userPage' value='         " & userPage & "' size='4'>")
            
            If nextPage > pageCount Then
                pr ("<img src='rightOne_.gif' border='0'><img src='rightEnd_.gif' border='0'>")
            Else
                pr ("<a href='" & objectPage & "?mode=modUser&sub=lst&userPage=" & nextPage & "'><img src='rightOne.gif' border='0'></a>")
                pr ("<a href='" & objectPage & "?mode=modUser&sub=lst&userPage=" & pageCount & "'><img src='rightEnd.gif' border='0'></a>")
            End If
            
            pr ("&nbsp;of " & pageCount)
            pr ("&nbsp;&nbsp;&nbsp;&nbsp;<font size='1'>Total Records:&nbsp;" & RecCount & "</font>")
            pr ("</td></tr>")
            
            pr ("</table>")
            pr ("</form>")
        End If
        
        pr ("</body>")
        pr ("</html>")
    End If
    
    If Request("sub") = "frmNew" Then
        If Request("resetSession") = "true" Then
            Session("FirstName") = ""
            Session("LastName") = ""
            Session("Username") = ""
            Session("Password") = ""
            Session("AccessLevel") = ""
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
        pr ("    location.href = '" & objectPage & "?mode=modUser&sub=lst'")
        pr ("}")
        pr ("</script>")
        pr ("<link rel='stylesheet' href='StyleSheet.css'>")
        pr ("<body onLoad='document.frm.FirstName.focus()'>")
        

        call GetMenu()
        '***************  TitleBar
        pr ("<table border='0' width='100%' bgcolor='" & TitleBarColor & "' cellpadding='1' cellspacing='0'><tr><td>")
        pr ("<table border='0' width='100%' cellpadding='1' cellspacing='0'>")
        pr ("<tr valign='bottom'>")
        pr ("<td align='left' nowrap>")
        pr ("&nbsp;<font size='3' color='" & TitleTextColor & "'><b>New User</b></font>")
        pr ("</td>")
        pr ("</tr>")
        pr ("</table>")
        pr ("</td></tr></table>")
    
        If Len(Request("msg")) > 0 Then
            pr ("<table border='0'><tr><td>" & Request("msg") & "</td></tr></table>")
        End If
        
        pr ("<form name='frm' action='" & objectPage & "'>")
        pr ("<input type='hidden' name='sub' value='subNew'>")
        pr ("<input type='hidden' name='mode' value='modUser'>")
            
        pr ("<table border='0'>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>FirstName</font></td>")
        pr ("<td><input type='text' name='FirstName' size='20' value='" & Session("FirstName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>LastName</font></td>")
        pr ("<td><input type='text' name='LastName' size='20' value='" & Session("LastName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>Username</font></td>")
        pr ("<td><input type='text' name='UserName' size='20' value='" & Session("UserName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>Password</font></td>")
        pr ("<td><input type='Password' name='Password' size='20' value='" & Session("Password") & "'></td>")
        pr ("</tr>")
        
        sel1 = ""
        sel2 = ""
        sel3 = ""
        If CStr(Session("AccessLevel")) = "3" Then
            sel3 = "selected"
        ElseIf CStr(Session("AccessLevel")) = "2" Then
            sel2 = "selected"
        Else
            sel1 = "selected"
        End If
        pr ("<tr>")
        pr ("<td align='right'><font size=2>Access Level</font></td>")
        pr ("<td>")
        pr ("<select name='AccessLevel'>")
        pr ("<option value='1' " & sel1 & ">Read Only")
        pr ("<option value='2' " & sel2 & ">Read/Write")
        pr ("<option value='3' " & sel3 & ">Admin")
        pr ("</select>")
        pr ("</td>")
        pr ("</tr>")
		
        pr ("<tr><td>&nbsp;</td><td>&nbsp;</td></tr>")
		
	
        pr ("<tr>")
        pr ("<td>&nbsp;</td>")
        pr ("<td>")
        pr ("<input type='button' value='Submit' onClick='javSubmit()'>&nbsp;")
        pr ("<input type='button' value='Cancel' onClick='javCancel()'></td>")
        pr ("</tr>")
            
        pr ("</table>")
            
            
        pr ("</form>")
            
        pr ("</body>")
        pr ("</html>")
    End If
    
    If Request("sub") = "frmEdit" Then
        Session("currUserID") = Request("UserID")
        Set RS0 = Conn.Execute("SELECT * FROM Users WHERE UserID= " & Request("UserID"))
        If RS0.EOF Then
            Response.Redirect "" & objectPage & "?mode=modUser&sub=lst&msg=Status:+User+does+not+exist."
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
        pr ("    location.href = '" & objectPage & "?mode=modUser&sub=lst'")
        pr ("}")
        
        pr ("function javDelete() {")
        pr ("   if (confirm('Delete user?')) {")
        pr ("       location.href = '" & objectPage & "?mode=modUser&sub=subDelete&UserID=" & Request("UserID") & "'")
        pr ("   } else {")
        pr ("       return")
        pr ("   }")
        pr ("}")
        
        pr ("</script>")
        pr ("<link rel='stylesheet' href='StyleSheet.css'>")
        pr ("<body onLoad='document.frm.FirstName.focus()'>")
        

            
		call GetMenu()
        '***************  TitleBar
        pr ("<table border='3' width='100%' bgcolor='" & TitleBarColor & "' cellpadding='1' cellspacing='0'><tr><td>")
        pr ("<table border='0' width='100%' cellpadding='1' cellspacing='0'>")
        pr ("<tr valign='bottom'>")
        pr ("<td align='left' nowrap>")
        pr ("&nbsp;<font size='3' color='" & TitleTextColor & "'><b>Edit User</b></font>")
        pr ("</td>")
        pr ("</tr>")
        pr ("</table>")
        pr ("</td></tr></table>")
    
        If Len(Request("msg")) > 0 Then
            pr ("<table border='0'><tr><td>" & Request("msg") & "</td></tr></table>")
        End If
        
        pr ("<form name='frm' action='" & objectPage & "'>")
        pr ("<input type='hidden' name='sub' value='subEdit'>")
        pr ("<input type='hidden' name='UserID' value='" & Request("UserID") & "'>")
        pr ("<input type='hidden' name='mode' value='modUser'>")
            
        pr ("<table border='0'>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>FirstName</font></td>")
        pr ("<td><input type='text' name='FirstName' size='20' value='" & RS0("FirstName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>LastName</font></td>")
        pr ("<td><input type='text' name='LastName' size='20' value='" & RS0("LastName") & "'></td>")
        pr ("</tr>")
            
        pr ("<tr>")
        pr ("<td align='right'><font size=2>Username</font></td>")
        pr ("<td bgcolor='silver'>" & RS0("UserName") & "</td>")
        pr ("</tr>")
            
        
        sel1 = ""
        sel2 = ""
        sel3 = ""

		employeeAccess = 1
		Set RSUser = Conn.Execute("SELECT defValue FROM dbp_defaultUserSettings WHERE defName = 'employeeAccess' AND UserID = " & RS0("UserID"))
	    if not RSUser.EOF then
			employeeAccess = cInt(RSUser("defValue"))
		end if
		If employeeAccess = 3 Then
            sel3 = "selected"
        ElseIf employeeAccess = 2 Then
            sel2 = "selected"
        Else
            sel1 = "selected"
        End If
        pr ("<tr>")
        pr ("<td align='right'><font size=2>Access Level</font></td>")
        pr ("<td>")
        pr ("<select name='AccessLevel'>")
        pr ("<option value='1' " & sel1 & ">Read Only")
        pr ("<option value='2' " & sel2 & ">Read/Write")
        pr ("<option value='3' " & sel3 & ">Admin")
        pr ("</select>")
        pr ("</td>")
        pr ("</tr>")

		
        pr ("<tr><td>&nbsp;</td><td>&nbsp;</td></tr>")

 
        pr ("<tr>")
        pr ("<td>&nbsp;</td>")
        pr ("<td>")
        pr ("<input type='button' value='Update' onClick='javSubmit()'>&nbsp;")
        pr ("<input type='button' value='Delete' onClick='javDelete()'>&nbsp;")
        pr ("<input type='button' value='Cancel' onClick='javCancel()'></td>")
        pr ("</tr>")
            
        pr ("</table>")
            
            
        pr ("</form>")
            
        pr ("</body>")
        pr ("</html>")
    End If
    
    If Request("sub") = "subNew" Then
        Set RS0 = Conn.Execute("SELECT * FROM Users WHERE UserName = '" & safeEntry(Request("Username")) & "'")
        If RS0.EOF Then
            strSQL = "INSERT INTO Users (FirstName,LastName,UserName,Password,LogonCount,DateCreate) VALUES ("
            strSQL = strSQL & "'" & safeEntry(Request("FirstName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("LastName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("UserName")) & "',"
            strSQL = strSQL & "'" & safeEntry(Request("Password")) & "',"
			strSQL = strSQL & "" & "0" & ","
            strSQL = strSQL & "'" & Date & "')"
            Conn.Execute (strSQL)
			
            Set RSNew = Conn.Execute("SELECT Max(UserID) AS NewID FROM Users")
            Session("currUserID") = RSNew("NewID")
			
			'User defaults 
			Conn.Execute("INSERT INTO dbp_defaultUserSettings(UserID,defName,defValue) VALUES(" & RSNew("NewID") & ",'employeeAccess','" & Request("AccessLevel") & "')") 
			Conn.Execute("INSERT INTO dbp_defaultUserSettings(UserID,defName,defValue) VALUES(" & RSNew("NewID") & ",'MenuWidth','20')") 
			Conn.Execute("INSERT INTO dbp_defaultUserSettings(UserID,defName,defValue) VALUES(" & RSNew("NewID") & ",'PageNum','20')") 
            pg = "" & objectPage & "?mode=modUser&sub=lst&msg=Status:+User+created."
        Else
            Session("FirstName") = Request("FirstName")
            Session("LastName") = Request("LastName")
            Session("Username") = Request("Username")
            Session("Password") = Request("Password")
            Session("AccessLevel") = Request("AccessLevel")
            pg = "" & objectPage & "?mode=modUser&sub=frmNew&msg=Status:+Username+already+exists."
        End If
        Response.Redirect pg
    End If
    
    If Request("sub") = "subEdit" Then
		pg = "" & objectPage & "?mode=modUser&sub=lst&msg=Status:+User+updated."
        Set RS1 = Conn.Execute("SELECT * FROM dbp_defaultUserSettings WHERE defName = 'employeeAccess' and UserID = " & Request("UserID")) 
		if RS1.EOF then
			Conn.Execute("INSERT INTO dbp_defaultUserSettings(UserID,defName,defValue) VALUES(" & Request("UserID") & ",'employeeAccess','" & Request("AccessLevel") & "')")
		else
			strSQL = "UPDATE Users SET "
	        strSQL = strSQL & "FirstName = '" & safeEntry(Request("FirstName")) & "',"
	        strSQL = strSQL & "LastName = '" & safeEntry(Request("LastName")) & "' "
	        strSQL = strSQL & "WHERE UserID = " & Request("UserID")
	        Conn.Execute (strSQL)

			'***** Default
			'Conn.Execute("UPDATE dbp_defaultUserSettings SET defValue = '" & Request("GroupID") & "' WHERE defName = 'GroupID' AND UserID = " & Request("UserID"))
			'Conn.Execute("UPDATE dbp_defaultUserSettings SET defValue = '" & Request("LogoName") & "' WHERE defName = 'LogoName' AND UserID = " & Request("UserID"))
			Set RS0 = Conn.Execute("SELECT * FROM Users WHERE UserID = " & Request("UserID"))
			if RS0("UserName") <> "Admin" then 
				Conn.Execute("UPDATE dbp_defaultUserSettings SET defValue = '" & Request("AccessLevel") & "' WHERE defName = 'employeeAccess' AND UserID = " & Request("UserID"))
			end if
		end if        
        
        Response.Redirect pg
    End If
    
    If Request("sub") = "subDelete" Then
		Set RS0 = Conn.Execute("SELECT * FROM Users WHERE UserID = " & Request("UserID"))
        if not RS0.EOF then
			If RS0("Username") = "Admin" Then
	            pg = "" & objectPage & "?mode=modUser&sub=lst&msg=Cannot+delete+admin+account."
	        Else
	            strSQL = "DELETE FROM Users WHERE UserID = " & Request("UserID")
	            Conn.Execute (strSQL)
				Conn.Execute ("DELETE FROM dbp_defaultUserSettings WHERE UserID = " & Request("UserID"))
	            pg = "" & objectPage & "?mode=modUser&sub=lst&msg=User+deleted."
	        End If
		else
			pg = "" & objectPage & "?mode=modUser&sub=lst&msg=Unable+to+delete+user."
		end if
        Response.Redirect pg
    End If
    
    Conn.Close
End Function

public function modLogon()
	Response.Expires = 0
	Set Conn = Server.CreateObject("ADODB.Connection")  
	Conn.Mode = 3
	Conn.Open Session("ADOEmployee") 
	objectPage = "Employee.asp"
	if Request("sub") = "logon" then
		Set RS1 = Conn.Execute("SELECT * FROM Users WHERE Password = '" & safeEntry(Request("p")) & "' AND Username = '" & safeEntry(Request("u")) & "'")
		If Not RS1.EOF Then
		    Session("webUserID") = RS1("UserID")
		    Session("employeeAccess") = cInt(GetDefaultUser("employeeAccess","0"))
			
			if isnumeric(RS1("LogonCount")) then
				LogonCount = RS1("LogonCount") + 1
			else
				LogonCount = 1
			end if
			Conn.Execute("UPDATE Users SET DateLastAccess = '" & now() & "',LogonCount = " & LogonCount & " WHERE UserID = " & RS1("UserID"))
		    pg = "frameset.asp"
		Else
		    pg = objectPage & "?mode=modLogon&msg=Status:+Not+authorized."
		End If
		Conn.Close 
		Response.redirect pg
	elseif Request("sub") = "chgpass" then
		Set RS1 = Conn.Execute("SELECT * FROM Users WHERE Password = '" & safeEntry(Request("p1")) & "' AND UserID = " & Session("webUserID"))
		If Not RS1.EOF Then
			if Session("IsEvaluation") = "true" then
				pg = "frmPassword.asp?msg=Evaluation+Version:+Unable+to+change+password."
			else
				Conn.Execute("UPDATE Users SET Password = '" & safeEntry(Request("p2")) & "' WHERE UserID = " & Session("webUserID"))
			    pg = "Employee.asp?mode=lst&menuEmployee=EmployeeList&msg=Password+changed."
			end if
		Else
		    pg = "frmPassword.asp?mode=modLogon&sub=frmpass&msg=Invalid+password."
		End If
		Conn.Close
		Response.redirect pg
	else	
		pr ("<html>")
		pr ("<head>")
		if Request("sub") = "frmpass" then
			pr ("<title>Change Password</title>")
		else
			pr ("<title>" & GetDefault("AppTitle","Employee Manager") & "</title>")
		end if
		pr ("</head>")
		pr ("<script language='javascript'>")
		pr ("function javCancel() {")
		pr ("	location.href = '" & objectPage & "?mode=lst'")
		pr ("}")
		pr ("</script>")
		pr ("<link rel='stylesheet' href='StyleSheet.css'>")
		if Request("sub") = "frmpass" then
			pr ("<body onLoad='document.frm.p1.focus()'>")
		else
			pr ("<body onLoad='document.frm.u.focus()'>")
		end if
		
		pr ("<center>")
		pr ("<br><br><br><br><br><br>")

		pr ("<form name='frm' action='" & objectPage & "'>")
		pr ("<input type='hidden' name='mode' value='modLogon'>")
		
		if Request("sub") = "frmpass" then
			pr ("<input type='hidden' name='sub' value='chgpass'>")
			pr ("<input type='hidden' name='sid' value='" & Request("sid") & "'>")
		else
			pr ("<input type='hidden' name='sub' value='logon'>")
		end if
		pr ("<table bgcolor='silver' border='0' cellspacing='1' bordercolor='gray'>")
		
		if Request("sub") = "frmpass" then
			pr ("<tr bgcolor='navy'><td><font color='white'><b>Change Password</b></font></td></tr>")	
		else
			pr ("<tr bgcolor='navy'><td><font color='white'><b>" & GetDefault("AppTitle","Employee Manager") & "&nbsp;Login</b></font></td></tr>")	
		end if
		pr ("<tr><td>")
		pr ("<table border='0'>")
		if len(Request("msg")) > 0 then 
			pr ("<tr><td colspan='2' align='center'><font color='maroon'>" & Request("msg") & "</font></td></tr>")
		end if 
		pr ("<tr><td>&nbsp;</td><td>&nbsp;</td></tr>")
		
		if Request("sub") = "frmpass" then
			pr ("<tr>")
			pr ("<td align='right'><font size='2'><b>Old Password:</b>&nbsp;</font></td>")
			pr ("<td align='left'><input type='password' name='p1' size='20' value=''></td>")
			pr ("</tr>")
			
			pr ("<tr>")
			pr ("<td align='right'><font size='2'><b>New Password:</b>&nbsp;</font></td>")
			pr ("<td align='left'><input type='password' name='p2' size='20' value=''></td>")
			pr ("</tr>")
		else
			strAdmin = ""
			strPass = ""
			if Session("IsEvaluation") = "true" then
				strAdmin = "Admin"
				strPass = "new" 
			end if
			pr ("<tr>")
			pr ("<td align='right'><font size='2'><b>User name:</b>&nbsp;</font></td>")
			pr ("<td align='left'><input type='text' name='u' size='20' value='" & strAdmin & "'></td>")
			pr ("</tr>")
			
			pr ("<tr>")
			pr ("<td align='right'><font size='2'><b>Password:</b>&nbsp;</font></td>")
			pr ("<td align='left'><input type='password' name='p' size='20' value='" & strPass & "'></td>")
			pr ("</tr>")
		end if
		pr ("<tr>")
		pr ("<td align='right'>&nbsp;</td>")
		pr ("<td align='left'>")
		pr ("<input type='submit' value='OK'>&nbsp;")
		pr ("</tr>")
		
		pr ("<tr><td>&nbsp;</td><td>&nbsp;</td></tr>")
		
		pr ("</table>")
		pr ("</td></tr>")
		pr ("</table>")
		
		pr ("</form>")
		
		pr ("</center>")
		
		pr ("</body>")
		pr ("</html>")
		Conn.Close
	end if

end function



Public Function dlgAbout()
    pr ("<html>")
    pr ("<head><title>About " & GetDefault("AppTitle", "Employee Manager") & "</title></head>")
    pr ("<link rel='stylesheet' href='StyleSheet.css'>")
    pr ("<body>")
    pr ("<center>")
    pr ("<br><br><br><br>")
    If Session("IsEvaluation") = "true" Then
        pr ("<font size='2'>This software is an evaluation version.<br>")
        pr ("Please purchase at <a href='http://www.aspintranet.com' target='_top'><u>www.ASPIntranet.com</u></a>.<br>")
    Else
        pr ("<font size='2'>This software is licensed to " & Session("LicensedTo") & ".&nbsp;&nbsp;")
        pr (Session("DateLicensed") & "<br>Product ID:" & Session("ProductID") & "<br>")
    End If
    pr ("Copyright &copy; 1999 <a href='mailto:sam_hurdowar@yahoo.com'>Sam Hurdowar</a>.</font>")
    pr ("</center>")
    pr ("</body>")
    pr ("</html>")
End Function
	
Public Function dlgDate()
    pr ("<html>")
    pr ("<head>")
    pr ("<title>Calendar</title>")
    pr ("</head>")
    pr ("<script Language='javascript'>")
    pr ("function javIt(d) {")
    pr ("   window.opener.document." & Request("frm") & "." & Request("fld") & ".value = d")
    pr ("}")
    pr ("function javCal(m,y) {")
    pr ("   location.href = 'Employee.asp?mode=dlgDate&calMonth=' + m + '&calYear=' + y + '&frm=" & Request("frm") & "&fld=" & Request("fld") & "'")
    pr ("}")
    pr ("</script>")
    pr ("<body bgcolor='white'>")
    pr ("<center>")
    
    Const cSUN = 1, cMON = 2, cTUE = 3, cWED = 4, cTHU = 5, cFRI = 6, cSAT = 7
    
    intThisDay = Day(Date)
    datToday = Date
    
    If Request("calMonth") = "" Then
      intThisMonth = Month(datToday)
    Else
      intThisMonth = CInt(Request("calMonth"))
    End If
    
    If IsEmpty(Request("calYear")) Or Not IsNumeric(Request("calYear")) Then
      datToday = Date
      intThisYear = Year(datToday)
    Else
      intThisYear = CInt(Request("calYear"))
    End If
    
    strMonthName = MonthName(intThisMonth)
    datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
    intFirstWeekDay = WeekDay(datFirstDay, vbSunday)
    intLastDay = GetLastDay(intThisMonth)
    
    IntPrevMonth = intThisMonth - 1
    If IntPrevMonth = 0 Then
        IntPrevMonth = 12
        intPrevYear = intThisYear - 1
    Else
        intPrevYear = intThisYear
    End If
    
    IntNextMonth = intThisMonth + 1
    If IntNextMonth > 12 Then
        IntNextMonth = 1
        intNextYear = intThisYear + 1
    Else
        intNextYear = intThisYear
    End If
    
    LastMonthDate = GetLastDay(intLastMonth) - intFirstWeekDay + 2
    NextMonthDate = 1
    intPrintDay = 1
    
    dFirstDay = intThisMonth & "/1/" & intThisYear
    dLastDay = intThisMonth & "/" & intLastDay & "/" & intThisYear
    
    pr ("<TABLE border='1' bordercolor='#004080' cellspacing='0'>")
    pr ("<TR><TD>")
    pr ("<table border='0' bgcolor='gray' cellspacing='1' cellpadding='0'>")
    pr ("<tr bgcolor='gray'>")
    pr ("<td width='40' align='left'><a href='javascript:javCal(" & IntPrevMonth & "," & intPrevYear & ")'><img src='leftOne.gif' border='0'></a></td>")
    pr ("<form name='frmCal' action='Employee.asp'>")
    pr ("<input type='hidden' name='mode' value='dlgDate'>")
    pr ("<input type='hidden' name='frm' value='" & Request("frm") & "'>")
    pr ("<input type='hidden' name='fld' value='" & Request("fld") & "'>")
    pr ("<td colspan='5' align='center'>")
    pr ("<select name='calMonth'>")
    For i = 1 To 12
        Mon = MonthName(i)
        If i = intThisMonth Then
            pr ("<option value='" & i & "' selected>" & Mon)
        Else
            pr ("<option value='" & i & "'>" & Mon)
        End If
    Next
    pr ("</select>")
    
    a = Year(Date) - 1
    b = Year(Date) + 10
    pr ("&nbsp;<select name='calYear'>")
    For i = a To b
        If i = intThisYear Then
            pr ("<option value='" & i & "' selected>" & i)
        Else
            pr ("<option value='" & i & "'>" & i)
        End If
    Next
    pr ("</select>")
    pr ("<input type='submit' value='Go!'>")
    pr ("</td>")
    
    pr ("<td width='40' align='right'><a href='javascript:javCal(" & IntNextMonth & "," & intNextYear & ")'><img src='rightOne.gif' border='0'></a></td>")
    pr ("</tr>")
    
    '*****************************  Day Label  *********************
    pr ("<tr bgcolor='#E1E1E1'>")
    pr ("<td width='40' align='left' valign='top'>Sun</td>")
    pr ("<td width='40' align='left' valign='top'>Mon</td>")
    pr ("<td width='40' align='left' valign='top'>Tue</td>")
    pr ("<td width='40' align='left' valign='top'>Wed</td>")
    pr ("<td width='40' align='left' valign='top'>Thu</td>")
    pr ("<td width='40' align='left' valign='top'>Fri</td>")
    pr ("<td width='40' align='left' valign='top'>Sat</td>")
    pr ("</tr>")
    
    '*****************************  Days  *********************
    EndRows = False
    Do While EndRows = False
       pr ("<tr>")
    
       For intLoopDay = cSUN To cSAT
            If intFirstWeekDay > cSUN Then
                sColor = "silver"
                sValue = LastMonthDate
                pr ("<td width='40' bgcolor='" & sColor & "'>" & sValue & "</td>")
                LastMonthDate = LastMonthDate + 1
                intFirstWeekDay = intFirstWeekDay - 1
            Else
    
                If intPrintDay > intLastDay Then
                    sColor = "silver"
                    sValue = NextMonthDate
                    pr ("<td width='40' bgcolor='" & sColor & "'>" & sValue & "</td>")
                    NextMonthDate = NextMonthDate + 1
                    EndRows = True
                Else
                    If intPrintDay = intLastDay Then
                        EndRows = True
                    End If
                    strDate = intThisMonth & "/" & intPrintDay & "/" & intThisYear
                    strLink = "<a href=""javascript:javIt('" & strDate & "')"">" & intPrintDay & "</a>"
                    sColor = "white"
                    sValue = strLink
                    pr ("<td width='40' bgcolor='" & sColor & "'>" & sValue & "</td>")
                End If
    
                intPrintDay = intPrintDay + 1
            End If
        
        Next
        
        pr ("</tr>")
    Loop
    pr ("</table>")
    pr ("</TD></TR>")
    pr ("<TR><TD align='center' bgcolor='#F2F7A4'>Today is <a href='javascript:javCal(" & Month(Date) & "," & Year(Date) & ")'>" & Date & "</a></TD></TR>")
    pr ("</TABLE>")
    pr ("<br>")
    pr ("<input type='button' value='Close' onClick='window.close();'>")
    pr ("</form>")
    pr ("</center>")
    pr ("</body>")
    pr ("</html>")
    
End Function

Public Function subSetting()
	if Request("userAction") = "Set to Default" then
		if Session("employeeAccess") > 2 then
			Call SetDefault("AppTitle","Employee Manager")
			Call SetDefault("CompanyName","DemoTech, Inc.")
		end if
		Call SetDefaultUser("MenuWidth","150")
		Call SetDefaultUser("PageNum","15")
		Call SetDefaultUser("TitleBarColor","#6CBFD0")
		Call SetDefaultUser("TitleTextColor","navy")
		pg = "frmSetting.asp?msg=Settings+set+to+default."
	else
		if Session("employeeAccess") > 2 then
			Call SetDefault("AppTitle",safeEntry(Request("AppTitle")))
			Call SetDefault("CompanyName",safeEntry(Request("CompanyName")))
		end if
		Call SetDefaultUser("MenuWidth",Request("MenuWidth"))
		Call SetDefaultUser("PageNum",Request("PageNum"))
		Call SetDefaultUser("TitleBarColor",safeEntry(Request("TitleBarColor")))
		Call SetDefaultUser("TitleTextColor",safeEntry(Request("TitleTextColor")))
		pg = "frmSetting.asp?msg=Settings+updated."
	end if
	
	Session("TitleBarColor") = GetDefaultUser("TitleBarColor","#6CBFD0")
	Session("TitleTextColor") = GetDefaultUser("TitleTextColor","navy")
	response.redirect pg
end function

Function SetDefault(fld, val)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOEmployee")
    ConnSP.Execute ("UPDATE dbp_defaultSettings SET defValue = '" & val & "' WHERE defName = '" & fld & "'")
    ConnSP.Close
End Function

Function GetDefault(fld,def)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOEmployee")
    Set RS100 = ConnSP.Execute("SELECT * FROM dbp_defaultSettings WHERE defName = '" & fld & "'")
    If Not RS100.EOF Then
        GetDefault = trim(RS100("defValue"))
	else
		ConnSP.Execute("INSERT INTO dbp_defaultSettings(defName,defValue) VALUES('" & fld & "','" & def & "')")
		GetDefault = def
    End If
    ConnSP.Close
End Function

Function SetDefaultUser(fld, val)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOEmployee")
    ConnSP.Execute ("UPDATE dbp_defaultUserSettings SET defValue = '" & safeEntry(val) & "' WHERE defName = '" & fld & "' AND UserID = " & Session("webUserID"))
    ConnSP.Close
End Function

Function GetDefaultUser(fld,def)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open Session("ADOEmployee")
    Set RSSpec = ConnSP.Execute("SELECT defValue FROM dbp_defaultUserSettings WHERE defName = '" & fld & "' AND UserID = " & Session("webUserID"))
    If Not RSSpec.EOF Then
        GetDefaultUser = trim(RSSpec("defValue"))
    Else
		ConnSP.Execute("INSERT INTO dbp_defaultUserSettings(UserID,defName,defValue) VALUES(" & Session("webUserID") & ",'" & fld & "','" & def & "')")
        GetDefaultUser = def
    End If
    ConnSP.Close
End Function



Function pr(strPrint)
    Response.Write strPrint & vbCrLf
End Function

Function funY2K(d)
    strDate = Trim(d)
    If InStr(strDate, " ") Then
        strDate = Left(strDate, InStr(strDate, " "))
        trailer = Right(d, Len(d) - InStr(d, " "))
    End If
    If IsDate(strDate) Then
        dateY2K = strDate
        If InStr(strDate, "/") = 2 Then
            strMonth = Left(strDate, 1)
            If InStr(3, strDate, "/") = 4 Then
                strDay = Mid(strDate, 3, 1)
            Else
                strDay = Mid(strDate, 3, 2)
            End If
            strYear = Right(strDate, Len(strDate) - InStr(3, strDate, "/"))
        ElseIf InStr(strDate, "/") = 3 Then
            strMonth = Left(strDate, 2)
            If InStr(4, strDate, "/") = 5 Then
                strDay = Mid(strDate, 4, 1)
            Else
                strDay = Mid(strDate, 4, 2)
            End If
            strYear = Right(strDate, Len(strDate) - InStr(4, strDate, "/"))
        End If
        intYear = CInt(strYear)
        If intYear >= 0 And intYear < 51 Then
            strYear = "20" & strYear
        ElseIf intYear > 50 And intYear < 100 Then
            strYear = "19" & strYear
        End If
        
        funY2K = strMonth & "/" & strDay & "/" & strYear & " " & trailer
    Else
        funY2K = ""
    End If
End Function

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

function getHeading(strName)
    Set Conn1 = Server.CreateObject("ADODB.Connection")
    Conn1.Mode = 3
    Conn1.Open Session("ADOEmployee")
	Set RS111 = Conn1.Execute("SELECT FieldAlias FROM defFields WHERE FieldName = '" & strName & "'")
	if not RS111.EOF then
		getHeading = RS111("FieldAlias")
	else
		getHeading = strName
	end if	
	if strName = "ContactID" then
		getHeading = "ID"
	end if
	Conn1.Close
end function

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

Public Function safeEntry(strField)
    strSafe = Trim(strField)
    strSafe = funReplace(strSafe, "'", "`")
    strSafe = funReplace(strSafe, "<", "&lt;")
    strSafe = funReplace(strSafe, ">", "&gt;")
    safeEntry = strSafe
End Function

public Function funReplace(a,b,c)
	funReplace = replace(a,b,c)
end function

%>
