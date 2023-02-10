<% 
Response.Expires = 0 
Session("CurrentPage") = "emp_frmEmp.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<!--#INCLUDE FILE="inc_Menus.asp"-->
<% 

'**********  Get Field Labels
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields")
do while not RS0.EOF
	Session(RS0("FieldName")) = RS0("FieldAlias")
	RS0.movenext
loop 

if Request("UserID") <> "0" then
	Set RS0 = Conn.Execute("SELECT * FROM dbp_UserInfos WHERE UserID = " & Request("UserID") ) 
	Session("currUserID") = Request("UserID")
	if not RS0.EOF then
		strName = "Membre : " & RS0("FirstName") & " " & RS0("LastName")
		BossID = RS0("BossID")
		FirstName = RS0("FirstName")
		MiddleName = RS0("MiddleName")
		LastName = RS0("LastName")
		Address = RS0("Address")
		City = RS0("City")
		State = RS0("State")
		Zip = RS0("Zip")
		Country = RS0("Country")
		Title = RS0("Title")
		WorkPhone = RS0("WorkPhone")
		WorkExt = RS0("WorkExt")
		Fax = RS0("Fax")
		Email = RS0("Email")
		HomePhone = RS0("HomePhone")
		MobilePhone = RS0("MobilePhone")
		Notes = RS0("Notes")
	end if
	editMode = true
else
	strName = "New Member"
	editMode = false
end if

%>
<html>
<head>
<title><%= GetDefault("emp_AppTitle", "Member Manager") %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "ASPIntranet.asp?mode=emp_lst"
}	
function javDelete() {
	if (confirm("Delete Member ?")) {
		location.href = "ASPIntranet.asp?mode=emp_subEmp&sub=delete&UserID=<%= Request("UserID") %>"
	} else {
		return
	}
}
function javLogon() {

	location.href = "../Portal_Asp/Portal.asp?mode=UserUpdate&user=<%= Request("UserID") %>"
}		
</script>
<link rel='stylesheet' href='../Portal_styles/PMainStyle1.asp'>
<body background="../Portal_Html/Images/Background.asp" bgcolor="#F5F7FE" onLoad="document.frm.FirstName.focus()">
<% Call GetMenu() %>
<% '************* Titlebar %>
<table width="100%" border="1" bgcolor="<%= GetUserDefault(Session("web_UserID"),"emp_bgcolor","#678DB8") %>" cellpadding="1" cellspacing="0">
<tr>
<td>
<table width="100%" border="0" cellpadding="1" cellspacing="0">
<tr>

<form name="frm" action="ASPIntranet.asp">
<input type="hidden" name="mode" value="emp_subEmp">
<input type="hidden" name="UserID" value="<%= Request("UserID") %>">

<% if editMode = true then %>	
	<input type="hidden" name="sub" value="edit">
<% else %>
	<input type="hidden" name="sub" value="new">
<% end if %>

<td align="left">
<font size="3" color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>"><b><%= strName %></b>
</td>
<td align="right">
<% if Session("emp_Access") = "Administrator" OR Session("emp_Access") = "Read/Write" then %>
	<input class="cmdflat" type="submit" value="Submit">&nbsp;
	<% if editMode = true then %>	
		<input class="cmdflat" type="button" value="Delete" onClick="javDelete()" >&nbsp;
		<input class="cmdflat" type="button" value="logon Info" onClick="javLogon()" >&nbsp;
	<% end if %>
<% end if %>
<input class="cmdflat" type="button" onClick="javCancel()" value="Cancel">&nbsp;
</td>
</tr>
</table>
</td>
</tr>
</table>

<table border="0" width="100%">
<tr>
<td align="right"><font size="2"><%= Session("BossID") %></td>
<td align="left" colspan="3">
<select class="tbflat" name="BossID">
<option value='0'>Head of Company
<%
if editMode = true then
	Session("Subordinates") = ":" & Request("UserID") & ":"
	Call GetSubordinates(Request("UserID"))
end if
Set RS1 = Conn.Execute("SELECT * FROM dbp_UserInfos ORDER BY LastName")
do while not RS1.EOF
	strName = RS1("LastName") & ", " & RS1("FirstName") & " - "
	if len(strName) < 7 then
		strName = "OPEN - " & RS1("Title")
	else
		strName = strName & RS1("Title")
	end if
	if editMode = true then
		if instr(Session("Subordinates"), ":" & RS1("UserID") & ":") = 0 then 
			if BossID = RS1("UserID") then
				pr ("<option value='" & RS1("UserID") & "' selected>" & strName)	
			else
				pr ("<option value='" & RS1("UserID") & "'>" & strName)	
			end if
		end if
	else
		pr ("<option value='" & RS1("UserID") & "'>" & strName)	
	end if
	RS1.movenext
loop
%>
</select>
</td>
</tr>
<tr class="smallerheader">
<td  align="right"><%= Session("FirstName") %></td>
<td align="left"><input class="smallerheaderW type="text" name="FirstName" size="30" tabindex="1" value="<%= FirstName %>"></td>
<td align="right"><%= Session("Title") %></td>
<td align="left"><input class="smallerheaderW type="text" name="Title" size="30" tabindex="9" value="<%= Title %>"></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("MiddleName") %></td>
<td align="left"><input class="smallerheaderW  type="text" name="MiddleName" size="30" tabindex="2" value="<%= MiddleName %>"></td>
<td align="right"><%= Session("WorkPhone") %></td>
<td align="left"><input class="smallerheaderW type="text" name="WorkPhone" size="30" tabindex="10" value="<%= WorkPhone %>"></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("LastName") %></td>
<td align="left"><input class="smallerheaderW type="text" name="LastName" size="30" tabindex="3" value="<%= LastName %>"></td>
<td align="right"><%= Session("WorkExt") %></td>
<td align="left"><input class="smallerheaderW type="text" name="WorkExt" size="30" tabindex="11" value="<%= WorkExt %>"></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("Address") %></td>
<td align="left"><textarea class="smallerheaderW name="Address" cols="25" rows="2" tabindex="4"><%= Address %></textarea></td>
<td align="right"><%= Session("Fax") %></td>
<td align="left"><input class="smallerheaderW type="text" name="Fax" size="30" tabindex="14" value="<%= Fax %>"></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("City") %></td>
<td align="left"><input class="smallerheaderW type="text" name="City" size="30" tabindex="5" value="<%= City %>"></td>
<td align="right"><%= Session("Email") %></td>
<td align="left"><input class="smallerheaderW type="text" name="Email" size="30" tabindex="15" value="<%= Email %>"></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("State") %></td>
<td align="left"><input class="smallerheaderW type="text" name="State" size="30" tabindex="6" value="<%= State %>"></td>
<td align="right"><%= Session("HomePhone") %></td>
<td align="left"><input class="smallerheaderW type="text" name="HomePhone" size="30" tabindex="12" value="<%= HomePhone %>"></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("Zip") %></td>
<td align="left"><input class="smallerheaderW type="text" name="Zip" size="30" tabindex="7" value="<%= Zip %>"></td>
<td align="right"><%= Session("MobilePhone") %></td>
<td align="left"><input class="smallerheaderW type="text" name="MobilePhone" size="30" tabindex="13" value="<%= MobilePhone %>"></td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("Country") %></td>
<td align="left"><input class="smallerheaderW type="text" name="Country" size="30" tabindex="8" value="<%= Country %>"></td>
<td align="right">&nbsp;</td>
<td align="left">&nbsp;</td>
</tr>

<tr class="smallerheader">
<td align="right"><%= Session("Notes") %></td>
<td colspan="3" align="left"><textarea class="smallerheaderW name="Notes" cols="50" rows="4" tabindex="17"><%= Notes %></textarea></td>
</tr>

<!-------------------------  Custom Fields ------------------- -->
<% 
Set RS1 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldAlias <> '' ORDER BY FieldID") 
pr ("<tr><td>&nbsp;</td><td colspan='3'>&nbsp;</td></tr>")
do while not RS1.EOF
	if left(RS1("FieldName"),3) = "fld" and len(RS1("FieldAlias")) > 0 then
		strValue = ""
		if editMode = true then
			key = RS1("FieldName")
			strValue = RS0(key)
		end if
		pr ("<tr>")
		pr ("<td align='right'>" & RS1("FieldAlias") & "</td>")
		pr ("<td colspan='3'><input type='text' name='" & RS1("FieldName") & "' value='" & strValue & "' size='60'></td>")
		pr ("</tr>")
	end if
	RS1.movenext
loop
%>

</table>
</form>

</body>
</html>
<% Conn.Close %>

