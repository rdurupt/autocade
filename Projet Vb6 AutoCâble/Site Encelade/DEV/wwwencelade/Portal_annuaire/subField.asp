<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<%  
Conn.Execute("DELETE FROM dbp_dirFields WHERE UserID = " & Session("UserID"))
strArray = split(Request("lstFieldID"), ",")

for i = 0 to ubound(strArray)
	order = "order" & trim(strArray(i))
	FieldOrder = Request(order)
	if not isnumeric(FieldOrder) then
		FieldOrder = 0
	end if
	Conn.Execute("INSERT INTO dbp_dirFields(UserID,FieldID,FieldOrder) VALUES(" & Session("UserID") & "," & trim(strArray(i)) & "," & FieldOrder & ")")
next
Conn.Close

pg = "frmField.asp?msg=Fields+updated."
response.redirect pg
%>