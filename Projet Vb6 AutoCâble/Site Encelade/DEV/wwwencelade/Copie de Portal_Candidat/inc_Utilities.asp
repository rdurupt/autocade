<%

Function SetDefault(fld, val)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open session("candidat_ADOContact")
    ConnSP.Execute ("UPDATE Defaults SET defValue = '" & val & "' WHERE defName = '" & fld & "'")
    ConnSP.Close
	Set ConnSP = Nothing
End Function

Function GetDefault(fld,def)
    Set ConnSP2 = Server.CreateObject("ADODB.Connection")
    ConnSP2.Open Session("candidat_ADOContact")
    Set RS100 = Server.CreateObject("ADODB.RecordSet")
    SQL = "SELECT * FROM Defaults WHERE defName = '" & fld & "'"
    RS100.Open SQL, ConnSP2, 3, 1, 1
    'Set RS100 = ConnSP2.Execute("SELECT * FROM Defaults WHERE defName = '" & fld & "'")
    If Not RS100.EOF Then
        GetDefault = trim(RS100("defValue"))
        RS100.Close
	else
		ConnSP2.Execute("INSERT INTO Defaults(defName,defValue) VALUES('" & fld & "','" & def & "')")
		GetDefault = def
    End If
    ConnSP2.Close
    Set RS100 = Nothing
	Set ConnSP2 = Nothing
End Function

Function SetDefaultUser(fld, val)
    Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open session("candidat_ADOContact")
    ConnSP.Execute ("UPDATE DefaultUsers SET defValue = '" & safeEntry(val) & "' WHERE defName = '" & fld & "' AND UserID = " & session("candidat_web_UserID"))
    ConnSP.Close
	Set ConnSP = Nothing
End Function

Function GetDefaultUser(fld,def)
  Set ConnSP = Server.CreateObject("ADODB.Connection")
    ConnSP.Open session("candidat_ADOContact")
    Set RSSpec = ConnSP.Execute("SELECT defValue FROM DefaultUsers WHERE defName = '" & fld & "' AND UserID = " & session("candidat_web_UserID"))
    If Not RSSpec.EOF Then
        GetDefaultUser = trim(RSSpec("defValue"))
    Else
		ConnSP.Execute("INSERT INTO DefaultUsers(UserID,defName,defValue) VALUES(" & session("candidat_web_UserID") & ",'" & fld & "','" & def & "')")
        GetDefaultUser = def
    End If
    ConnSP.Close
	Set ConnSP = Nothing
	 Set RSSpec =Nothing
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
    Conn1.Open session("candidat_ADOContact")
	Set RS111 = Conn1.Execute("SELECT FieldAlias FROM con_FieldDefs WHERE FieldName = '" & strName & "'")
	if not RS111.EOF then
		getHeading = RS111("FieldAlias")
	else
		getHeading = strName
	end if	
	if strName = "ContactID" then
		getHeading = "ID"
	end if
	Conn1.Close
	 Set RS111=Nothing
	 set Conn1=Nothing
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
