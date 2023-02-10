<% 

Function ChkString(str)
     If str = "" Then
        str = " "
     End If
     
     ChkString = server.htmlencode(str)

End Function

Sub HashIt(DataToHash)
            Randomize
            Salt = ""
            For i = 1 To 10
                '65 is ASCII for "A"
                Salt = Salt & Chr(Int(Rnd * 26) + 65)
            Next
            ' Calculate Hash of (Password & Salt)
            Set CM = Server.CreateObject("AspCrypt.Crypt")
            Session("HashData") = CM.Crypt(Salt, DataToHash)
            Session("HashSecure") = Salt
End Sub

Sub SendEmail(EmailFrom, MailTo, EmailSubject, EmailText)
		Set Mail = Server.CreateObject("CDONTS.NewMail")
		Mail.BodyFormat = 1	' Text Only, 0 for HTML
		Mail.Subject = EmailSubject
		Mail.To = MailTo
		Mail.Body = EmailText
		Mail.Send(emailfrom)
		Set Mail = Nothing
End Sub



    Response.Expires = 0
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Mode = 3
    Conn.Open Session("DSN")
    
    SQL = "select * from dbp_Users order by UserId"
    Set RS1 = Server.CreateObject("ADODB.Recordset")
    
    RS1.Open SQL, Conn
    
    Do Until RS1.EOF
        NewUserId = RS1("UserId")
        RS1.MoveNext
    Loop
    NewUserId = NewUserId + 1
    RS1.Close
    Set RS1 = Nothing

    ' Insertion dans dbp_Users
    ' génération mot de passe aléatoire
    HashIt (Request.Form("Email"))
    NewPass = LCase(Left(Session("HashData"), 4)) & (Int((Rnd(Timer) * 10000)))
    NewPass = ucase(Left(NewPass, 8))
    HashIt (NewPass)

    pg = "ASPIntranet.asp?mode=emp_lst"
	
	' insertion base user portail
	strSQL = "INSERT INTO dbp_Users ("
        strVAL = " VALUES("
	strSQL= strSQL & "UserID,"
	strVAL= strVAL & NewUserId & ","
	strSQL= strSQL & "Cid,"
	strVAL= strVAL & session("Application") & ","
	strSQL= strSQL & "UserLast,"
	strVAL= strVAL & "'" & chkstring(request("nachname")) & "',"
	strSQL= strSQL & "UserFirst,"
	strVAL= strVAL & "'" & chkstring(request("vorname")) & "',"
	strSQL= strSQL & "UserLogin,"
	strVAL= strVAL & "'" & chkstring(request("email")) & "',"
	strSQL= strSQL & "UserEmail,"
	strVAL= strVAL & "'" & chkstring(request("email")) & "',"
	strSQL= strSQL & "UserPass,"
	strVAL= strVAL & "'" & Session("HashData") & "',"
	strSQL= strSQL & "UserSecure,"
	strVAL= strVAL & "'" & Session("HashSecure") & "',"
	strSQL= strSQL & "UserDisplayName) "
	strVAL= strVAL & "'" & chkstring(request("vorname")) &  " " & chkstring(request("nachname")) & "')"    
	
	conn.execute(strsql & strval)
	response.write newpass & "<br>"
	response.write  strSQL & strVAL & "<br><br><br>"

	' insertion dans userinfos
	SearchBlob = ""
        strSQL = "INSERT INTO dbp_UserInfos ("
        strVAL = " VALUES("
	strSQL= strSQL & "UserID,"
	strVAL= strVAL & NewUserId & ","
	strSQL= strSQL & "LastName,"
	strVAL= strVAL & "'" & chkstring(request("nachname")) & "',"
	strSQL= strSQL & "FirstName,"
	strVAL= strVAL & "'" & chkstring(request("vorname")) & "',"
	strSQL= strSQL & "Title,"
	strVAL= strVAL & "'" & chkstring(request("Title")) & "',"
	strSQL= strSQL & "fld6,"
	strVAL= strVAL & "'" & chkstring(request("company")) & "',"
	strSQL= strSQL & "email,"
	strVAL= strVAL & "'" & chkstring(request("email")) & "',"
	strSQL= strSQL & "fld1,"
	strVAL= strVAL & "'" & chkstring(request("telefon")) & "',"
	strSQL= strSQL & "fld3,"
	strVAL= strVAL & "'" & chkstring(request("cellphone")) & "',"
	strSQL= strSQL & "fld2,"
	strVAL= strVAL & "'" & chkstring(request("telefax")) & "',"
	strSQL= strSQL & "fld5,"
	strVAL= strVAL & "'" & chkstring(request("strasse")) & "',"
	strSQL= strSQL & "fld8,"
	strVAL= strVAL & "'" & chkstring(request("ort")) & "',"
	strSQL= strSQL & "fld13,"
	strVAL= strVAL & "'" & chkstring(request("plz")) & "',"
	strSQL= strSQL & "country,"
	strVAL= strVAL & "'" & chkstring(request("land")) & "',"
	strSQL= strSQL & "fld7,"
	strVAL= strVAL & "'" & chkstring(request("staat")) & "',"
	strSQL= strSQL & "fld9,"
	strVAL= strVAL & "'" & chkstring(request("interest1")) & "',"
	strSQL= strSQL & "fld10,"
	strVAL= strVAL & "'" & chkstring(request("interest2")) & "',"
	strSQL= strSQL & "fld11,"
	strVAL= strVAL & "'" & chkstring(request("newsletter")) & "',"
	strSQL= strSQL & "fld14,"
	strVAL= strVAL & "'" & chkstring(request("memberSIM")) & "', "
	strSQL= strSQL & "fld12) "
	strVAL= strVAL & "'" & chkstring(request("accept")) & "')"
	response.write  strSQL & strVAL & "<br>"
	conn.execute(strsql & strval)

        ' Inscription dans le groupe Registerd visitor = 19
	SQL = "INSERT INTO dbp_GroupsPermission (Cid, GroupId, ObjTypeId, ObjPermId, ObjId, ObjPermType)"
        SQL = SQL & " Values (" & Session("APPLICATION") & ",19,50,1," & NewUserId & ",10)"
        
	Conn.Execute (SQL)
        
	response.write  SQL 

        
	MyEmail = Request("Email")
        EmailFrom = "Webmaster Simalliance.org<webmaster@simalliance.org>"
        EmailSubject = "Your simalliance.org registration account"
        EmailText = "Votre compte est activé" & vbCrLf & "Login = " & request("email") & vbcrlf & "Votre mot de passe est : " & NewPass & vbCrLf
        SendEmail EmailFrom, MyEmail, EmailSubject, EmailText

	if request("memberSIM")="member" then 
	        Myemail="lwagneur@euxia.net"
		EmailFrom = "Webmaster Simalliance.org<webmaster@simalliance.org>"
        	EmailSubject = "coucou y'a un membre"
        	EmailText = "Votre compte est activé" & vbCrLf & "Login = " & request("email") & vbcrlf & "Votre mot de passe est : " & NewPass & vbCrLf
        	SendEmail EmailFrom, MyEmail, EmailSubject, EmailText
	end if
	
%>
	


