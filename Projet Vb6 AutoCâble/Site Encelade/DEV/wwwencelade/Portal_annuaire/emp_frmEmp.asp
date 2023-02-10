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
	Session("emp_" & RS0("FieldName")) = RS0("FieldAlias")
	RS0.movenext
loop 

if Request("UserID") <> "0" then
	Set RS0 = Conn.Execute("SELECT * FROM dbp_UserInfos WHERE  UserID = " & Request("UserID") ) 
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
		
</script>
<link rel='stylesheet' href='<%=session("stylesheet")%>'>
<body background="Background.asp" bgcolor="#F5F7FE" onLoad="document.frm.FirstName.focus()">
<% Call GetMenu() %>
<% '************* Titlebar %><form name="frm" action="ASPIntranet.asp">
<input type="hidden" name="mode" value="emp_subEmp">
<input type="hidden" name="UserID" value="<%= Request("UserID") %>">

<% if editMode = true then %>	
	<input type="hidden" name="sub" value="edit">
<% end if %>
  <table width="617" border="0" align="center" cellpadding="0" cellspacing="0" background="bg.jpg">
    <tr> 
      <td><img src="haut.gif" width="617" height="24"></td>
    </tr>
    <tr> 
      <td><div align="right"></div>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="50%"><div align="right"><img src="haut2.gif" width="25" height="18"></div></td>
            <td width="50%" background="haut_bg.gif"><div align="center"><font color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>" size="2" face="Arial, Helvetica, sans-serif"><b><%= strName %></b> 
                </font></div></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td>
<table width="90%" border="0" align="center">
          <tr class="smallerheader"> 
            <td align="right" >&nbsp;</td>
            <td align="left" colspan="3">&nbsp; </td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right"><%= Session("emp_BossID") %></td>
            <td align="left"> 
              <% if Session("emp_Access") = "Administrator" OR Session("emp_Access") = "Read/Write" then 

pr("<select name=""BossID"">")
pr("<option value='0'>Non affecté")

if editMode = true then
	Session("Subordinates") = ":" & Request("UserID") & ":"
	Call GetSubordinates(Request("UserID"))
end if
Set RS1 = Conn.Execute("SELECT * FROM dbp_UserInfos where userid<>2 ORDER BY LastName") 
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

pr("</select>")
else
Set RS1 = Conn.Execute("SELECT * FROM dbp_UserInfos where userid<>2 ORDER BY LastName")
my_ok=0
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
				my_ok=1
				pr("<input type=""hidden""  name=""BossID"" value=" & RS1("UserID") & ">")
				pr ( strName)	
			end if
		end if
	else
		pr ( strName)	
	end if
	RS1.movenext
loop
if my_ok=0 then
	pr("<input type=""hidden""  name=""BossID"" value=""0"">")
	'pr ( strName)
end if
end if
%>
            </td>
            <td  align="right"><%= Session("emp_FirstName") %></td>
            <td align="left"><input class="champ" type="text" name="FirstName" size="30" tabindex="1" value="<%= FirstName %>"></td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right"><%= Session("emp_Title") %></td>
            <td align="left"><input class="champ" type="text" name="Title" size="30" tabindex="9" value="<%= Title %>"></td>
            <td  align="right"><%= Session("emp_Fax") %></td>
            <td align="left"><input class="champ" type="text" name="Fax" size="30" tabindex="14" value="<%= Fax %>"></td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right"><%= Session("emp_LastName") %></td>
            <td align="left"><input class="champ" type="text" name="LastName" size="30" tabindex="3" value="<%= LastName %>"></td>
            <td align="right"><%= Session("emp_MobilePhone") %></td>
            <td align="left"><input class="champ" type="text" name="MobilePhone" size="30" tabindex="13" value="<%= MobilePhone %>"></td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right"><%= Session("emp_WorkPhone") %></td>
            <td align="left"><input class="champ" type="text" name="WorkPhone" size="30" tabindex="10" value="<%= WorkPhone %>"></td>
            <td align="right">&nbsp;</td>
            <td align="left">&nbsp;</td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right"><%= Session("emp_WorkExt") %></td>
            <td align="left"><input class="champ" type="text" name="WorkExt" size="30" tabindex="11" value="<%= WorkExt %>"></td>
            <td align="right">&nbsp;</td>
            <td align="left">&nbsp;</td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right" valign="top"><%= Session("emp_Email") %></td>
            <td align="left"><input class="champ" type="text" name="Email" size="30" tabindex="15" value="<%= Email %>"></td>
            <td align="right"><%= Session("emp_Zip") %></td>
            <td align="left"><input class="champ" type="text" name="Zip" size="30" tabindex="7" value="<%= Zip %>"></td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right"><%= Session("emp_City") %></td>
            <td align="left"><input class="champ" type="text" name="City" size="30" tabindex="5" value="<%= City %>"></td>
            <td align="right"><%= Session("emp_Country") %></td>
            <td align="left"><input class="champ" type="text" name="Country" size="30" tabindex="8" value="<%= Country %>"></td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right"><%= Session("emp_Notes") %></td>
            <td align="left"><textarea class="champ" name="Notes" cols="25" rows="3" tabindex="17"><%= Notes %></textarea></td>
            <td align="right"><%= Session("emp_Address") %></td>
            <td align="left"><textarea class="champ" name="Address" cols="25" rows="3" tabindex="4"><%= Address %></textarea></td>
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
        <div align="center"><br>
          <% if Session("emp_Access") = "Administrator" OR Session("emp_Access") = "Read/Write" or cint(Request("UserID"))=cint(session("web_userid")) then %>
          <input name="Submit" type="Submit" class="cmdflat" value="Valider">
          &nbsp; 
          <% end if %>
          <input name="button" type="button" class="cmdflat" onClick="javCancel()" value="Retour">
          &nbsp; </div></td>
    </tr>
    <tr> 
      <td><img src="bas.gif"></td>
    </tr>
  </table>
  </form>

</body>
</html>
<% Conn.Close %>

