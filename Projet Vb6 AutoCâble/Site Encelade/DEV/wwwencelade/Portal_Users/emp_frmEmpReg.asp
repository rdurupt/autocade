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

strName = "Créer votre compte I-Graal en ligne"
editMode = false


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

	location.href = "../Portal_Asp/Portal_User_Edit.asp?user=<%= Request("UserID") %>"
}		
</script>
<link rel="stylesheet" href="../Portal_styles/PMainStyle1.asp">
<body background="../Portal_Html/Images/Background.asp" onLoad="document.frm.FirstName.focus()">
<form name="frm" action="ASPIntranet.asp">
	<input type="hidden" name="mode" value="emp_subEmpReg">
	<input type="hidden" name="sub" value="new">
<table border="0" NOREPEAT align="left" width="600px">
<tr>
<td align="right" class="smallerheader1">&nbsp;<BR><BR><BR></font></td>
<td align="left" class="smallerheader1">&nbsp;<BR><BR><BR></td>
</tr>
<tr>
<td align="right"><b>Créer votre compte </b></td>
<td align="left"><b>I-Graal en ligne :</b>&nbsp;</td>
</tr>
<tr>
<td align="right"><font size="2"><%= Session("BossID") %></font></td>
<td align="left" colspan="3">
</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("FirstName") %></font></td>
<td align="left" class="smallerheader1"><input type="text" class="cmd1flat" name="FirstName" size="30" tabindex="1" value="<%= FirstName %>"><font color="red"> * obligatoire</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("LastName") %></font></td>
<td align="left" class="smallerheader1"><input type="text"  class="cmd1flat" name="LastName" size="30" tabindex="2" value="<%= LastName %>"><font color="red"> * obligatoire</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("MiddleName") %></font></td>
<td align="left" class="smallerheader1"><input type="text" class="cmd1flat" name="MiddleName" size="30" tabindex="3" value="<%= MiddleName %>"><font color="red"> * obligatoire</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("Email") %></font></td>
<td align="left" class="smallerheader1"><input type="text"  class="cmd1flat" name="Email" size="30" tabindex="4" value="<%= Email %>"><font color="red"> * obligatoire</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("Address") %></font></td>
<td align="left" class="smallerheader1"><textarea name="Address"  class="cmd1flat" cols="25" rows="2" tabindex="5"><%= Address %></textarea><font color="red"> * obligatoire</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("Zip") %></font></td>
<td align="left" class="smallerheader1"><input type="text"  class="cmd1flat" name="Zip" size="7" tabindex="6" value="<%= Zip %>"><font color="red"> * obligatoire</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("City") %></font></td>
<td align="left" class="smallerheader1"><input type="text"  class="cmd1flat" name="City" size="30" tabindex="7" value="<%= City %>"><font color="red"> * obligatoire</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("Country") %></font></td>
<td align="left" class="smallerheader1"><input type="text"  class="cmd1flat" name="Country" size="30" tabindex="8" value="<%= Country %>"><font color="red"> * obligatoire</td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("WorkPhone") %></font></td>
<td align="left" class="smallerheader1"><input type="text" class="cmd1flat" name="WorkPhone" size="30" tabindex="9" value="<%= WorkPhone %>"></td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("MobilePhone") %></font></td>
<td align="left" class="smallerheader1"><input type="text"  class="cmd1flat" name="MobilePhone" size="30" tabindex="10" value="<%= MobilePhone %>"></td>
</tr><tr>
<td align="right" class="smallerheader1"><%= Session("Title") %></font></td>
<td align="left" class="smallerheader1"><input type="text" class="cmd1flat" name="Title" size="30" tabindex="11" value="<%= Title %>"></td>
</tr>
<tr>
<td align="right" class="smallerheader1">Informations sur votre e-broker</font></td>
<td colspan="3" align="left"><textarea name="Conditions" class="cmd1flat" cols="50" rows="5" tabindex="12"><%= Notes %></textarea></td>
</tr>
<tr>
<td align="right"><input type="submit"  class="cmd1flat" value="S'inscrire">&nbsp;</td>
<td align="left"><input type="button"  class="cmd1flat" onClick="javCancel()" value="Abandonner">&nbsp;</td>
</tr>
</table>
</form>

</body>
</html>
<% Conn.Close %>

