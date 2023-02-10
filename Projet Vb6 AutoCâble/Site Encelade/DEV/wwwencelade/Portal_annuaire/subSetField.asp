<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<%  
Set RS1 = Conn.Execute("SELECT FieldID FROM defFields")
do while not RS1.EOF 
	key = "key" & RS1("FieldID")

	strSQL = "UPDATE defFields SET "
	strSQL = strSQL & "FieldAlias = '" & safeEntry(Request(key)) & "' "
	strSQL = strSQL & "WHERE FieldID = " & RS1("FieldID")
	Conn.Execute(strSQL)
	RS1.movenext
loop
Conn.Close

pg = "frmSetField.asp?msg=Fields+updated."
response.redirect pg
%>