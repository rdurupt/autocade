
<%
	' This Include check for existing user session
	' If no current session then login
If Session("DSN")="" then
		Response.write ("<THead>Session expirée ...<Br>Vous devez vous ré-identifier<Br><H2><a href='../default.asp' target=_top>Cliquez ici</a>")
		Response.end
end if
'============ On définit le DSN du fichier pour tout le Portail =========
set my_conn= Server.CreateObject("ADODB.Connection")
my_Conn.Open Session("DSN")
'============ Cryptage de données en MD4-5 ==================
sub HashIt(DataToHash)			
			Randomize
			Salt = ""
			For i = 1 to 10
				'65 is ASCII for "A"
				Salt = Salt & chr(int(Rnd * 26) + 65)
			Next
			' Calculate Hash of (Password & Salt)
			Set CM = Server.CreateObject("AspCrypt.Crypt")
			Session("HashData")=CM.Crypt (Salt, DataToHash)
			Session("HashSecure")=Salt
end sub

Public Function CheckHash(DataToCheck,CryptedData,Salt)
			Set CM = Server.CreateObject("AspCrypt.Crypt")
			If CM.Crypt (Salt, DataToCheck)=CryptedData then
				CheckHash=1
			else
				CheckHash=0
			end If
end function

sub CheckAdminRights
	If session("ADMIN")<>1 then
				Session("USERId")=0
				Response.Write "<p id=alarm>Vous n'avez pas la permission d'accéder à cette zone" & vbcrlf
				Response.Write "Vous avez été déconnecté </p>" & vbcrlf
				Response.write ("<THead>Session expirée ...<Br>Vous devez vous ré-identifier<Br><H2>")
				Response.end
	end if
end sub 

Sub SendEmail(EmailFrom,EmailTo,EmailSubject,EmailText)
	set myemail =  server.createobject("Dynu.Email")
	myemail.Smtp = session("MAILSERVER")
	Response.Write session("MAILSERVER") & "<br>"
	Response.Write session("MAILMASTER") & "<br>"
    Response.Write EmailFrom & "<br>"
	Response.Write EmailTo & "<br>"
	Response.Write EmailSubject & "<br>"
	Response.Write EmailText & "<br>"
	result = myemail.Send(EmailFrom,EmailTo,EmailSubject, EmailText)
	Response.write result
	set myemail = nothing

End sub

	 if Session("USERId")=0 then
		Response.write ("<THead>Session expirée ...<Br>Vous devez vous ré-identifier<Br><H2>")
		Response.end
		
    end if
	
	' Open Current DSN 
	function doCode(str, oTag, cTag, roTag, rcTag)
	tx = split(str, cTag)
	t = ""
	for i = 0 to ubound(tx)

	  if lcase(oTag) =  "[a]" then
		p = instr(1, tx(i), "[a]", 1) 
		if p <> 0 then
			tmp = mid(tx(i), p)
			url = mid(tmp, 4)
			if lcase(left(url, 5)) = "http:" then
				tmp1 = Replace(tmp, "[a]"&url, "<A href='" & url & "' Target=_Blank>Link</a>", 1, -1, 1)			
			else
				tmp1 = Replace(tmp, "[a]"&url, "<A href='http://" & url & "' Target=_Blank>Link</a>"	, 1, -1, 1)	
			end if
			t =t & Replace(tx(i), tmp, tmp1)
		else
			t = t & tx(i)
		end if
	  else
		cnt = instr(1,tx(i), oTag,1)
		select case cnt 
			case 0
				t=t&tx(i) & " " 
			case else 
				t = t & Replace(tx(i), oTag, roTag,1,1,1)
				t = t & " " & rcTag & " "
		end select
	  end if
	next
	doCode = t
end function
	Function ChkString(str)
	 if str = "" then 
		str = " "
	 Else
		if BadWordFiler = "true" then
		  bwords = split(BadWords, "|")
		  for i = 0 to ubound(bwords)
			str= replace(str, bwords(i), string(len(bwords(i)),"*"), 1,-1,1) 
		  next
        End if
	 End If
	 
	 '  Do ASP Forum Code
	 str = doCode(str, "[b]", "[/b]", "<b>", "</b>")
	 str = doCode(str, "[i]", "[/i]", "<i>", "</i>") 
	 str = doCode(str, "[quote]", "[/quote]", "<BLOCKQUOTE><font size=1 face=arial>quote:<hr height=1 noshade>", "<hr height=1 noshade></BLOCKQUOTE></font><font face='" & DefaultFontFace & "' size=2>") 
	 str = doCode(str, "[a]", "[/a]", "<a>", "</a>")
	 str = doCode(str, "[code]", "[/code]", "<pre>", "</pre>")
	 
	 if smiles = "true" then str= smile(str)
	 
	 str = Replace(str, "'", "''")
	 str = Replace(str, "|", "/")
	 
	 ChkString = str
End Function



%>

