<% 
Response.Expires = 0 
Session("web_UserID") = Request("sid")
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<% 

'**********  Get Field Labels
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields")
do while not RS0.EOF
	Session("emp_" & RS0("FieldName")) = RS0("FieldAlias")
	RS0.movenext
loop 

if Request("UserID") <> "0" then
	Set RS0 = Conn.Execute("SELECT * FROM dbp_UserInfos WHERE UserID = " & Request("UserID") ) 
	if not RS0.EOF then
		strName = RS0("FirstName") & " " & RS0("LastName")
		BossID = RS0("BossID")
		FirstName = RS0("FirstName")
		LastName = RS0("LastName")
		Title = RS0("Title")
		WorkPhone = RS0("WorkPhone")
		WorkExt = RS0("WorkExt")
		Email = RS0("Email")
		Session("Subordinates") = ":" & Request("UserID") & ":"
		Call GetSubordinates(Request("UserID"))
		exclusive = replace(Session("Subordinates"), ":" & Request("UserID") & ":", "")
	
	End if
	editMode = true
else
	strName = "New Member"
	editMode = false
end if




if Request("mode") = "edit" then
	strSQL = "UPDATE dbp_UserInfos SET "
	strSQL = strSQL & "BossID = " & Request("BossID") 
	strSQL = strSQL & " WHERE UserID = " & Request("UserID")
	Conn.Execute(strSQL)
	response.redirect "dlgStatus.asp?toPage=ASPIntranet.asp&key=mode&keyValue=emp_chtOrganization&key2=UserID&keyValue2=" & Request("UserID") & "&key3=sid&keyValue3=" & Request("sid") & "&msg=Mise+à+jour+en+cours..."
end if




%>
<html>
<head>
<title>Fiche de <%= strName %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "ASPIntranet.asp?mode=emp_chtOrganization"
}
</script>
<link rel="stylesheet" href="<%=session("stylesheet")%>">
<body background="Background.asp" >
<% '************* Titlebar %>
<table width="100%" border="1" bgcolor="<%= GetDefault("emp_bgcolor","#678DB8") %>" cellpadding="1" cellspacing="0">
<tr>
<td>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
<tr>

<form name="frm" action="emp_dlgEmp.asp">
<%
if editmode then
%><input type="hidden" name="mode" value="edit">
<%
end if
%>
<input type="hidden" name="UserID" value="<%= Request("UserID") %>">
<input type="hidden" name="sid" value="<%= Request("sid") %>">

<td align="left">
<font size="3" color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>"><b><%= strName %></b>
</td>
<td align="right">
<% if Session("emp_Access") = "Administrator" OR Session("emp_Access") = "Read/Write"  then %>
	<input class="cmdflat" type="Submit" value="Valider">&nbsp;

<% end if %>
<input class="cmdflat" type="button" onClick="window.close()" value="Fermer">&nbsp;
</td>

</tr>
</table>
</td>
</tr>
</table>

<br>
<table border="0" width="100%">

<tr class="smallerheader">
<td align="right"><%= Session("emp_Title") %> : </td>
<td align="left"><%= Title %></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("emp_FirstName") %> : </td>
<td align="left"><%= FirstName %></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("emp_LastName") %> : </td>
<td align="left"><%= LastName %></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("emp_WorkPhone") %> : </td>
<td align="left"><%= WorkPhone %></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("emp_WorkExt") %> : </td>
<td align="left"><%= WorkExt %></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("emp_Email") %> : </td>
<td align="left"><a href="mailto:<%= Email %>"><%= Email %></a></td>
</tr>

<tr class="smallerheader">
<td align="right">Responsable : </td>
<td align="left">
<select name="BossID">
<option value='0'>Non affecté
<%
Set RS1 = Conn.Execute("SELECT * FROM dbp_UserInfos  where userid<>2 ORDER BY LastName")
do while not RS1.EOF
	strName = RS1("LastName") & ", " & RS1("FirstName") & " - "
	if len(strName) < 7 then
		strName = "OPEN - " & RS1("Title")
	else
		strName = strName & RS1("Title")
	end if

	if instr(Session("Subordinates"), ":" & RS1("UserID") & ":") = 0 then 
		if BossID = RS1("UserID") then
			pr ("<option value='" & RS1("UserID") & "' selected>" & strName)	
		else
			pr ("<option value='" & RS1("UserID") & "'>" & strName)	
		end if
	end if

	RS1.movenext
loop
%>
</select>
</td>
</tr>

</table>
</form>

</body>
</html>
<% Conn.Close %>

