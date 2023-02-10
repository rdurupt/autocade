<% 
Response.Expires = 0 
Session("CurrentPage") = "emp_frmQuery.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->

<!--#INCLUDE FILE="emp_TopMenu.asp"-->
<html>
<head>
<title><%= GetDefault("emp_AppTitle","Employee Manager") %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "ASPIntranet.asp?mode=emp_lst"
}

</script>
<link rel="stylesheet" href="<%=session("stylesheet")%>">
<body background="Background.asp" bgcolor="#F5F7FE" onLoad="document.frm.qryFirstName.focus()">
<% 
Call GetMenu() 
'**********  Get Field Labels
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields")
do while not RS0.EOF
	Session("emp_" & RS0("FieldName")) = RS0("FieldAlias")
	RS0.movenext
loop 
%>
<% '****** Header ***** %>
<form name="frm" action="ASPIntranet.asp">
<input type="hidden" name="mode" value="emp_lst">
<input type="hidden" name="employeeQuery" value="true">
  <br><table width="617" border="0" align="center" cellpadding="0" cellspacing="0" background="bg.jpg">
    <tr>
      <td><img src="haut.gif" width="617" height="24"></td>
    </tr>
    <tr>
      <td><div align="right"></div>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><div align="right"><img src="haut2.gif" width="25" height="18"></div></td>
            <td width="50%" background="haut_bg.gif"><div align="center"><font color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>" size="2" face="Arial, Helvetica, sans-serif"><b>Recherche 
                dans l'<%=GetDefault("emp_AppTitle","Employee Manager")%></b></font> 
              </div></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td><br><table align="center" cellspacing="0" cellpadding="0" border="0">
          <tr> 
            <td nowrap class="smallerheader">&nbsp;&nbsp;<%= Session("emp_FirstName") %>&nbsp;</td>
            <td><input type="text" class="champ" size="30" name="qryFirstName" value="<%= Session("qryFirstName") %>"></td>
          </tr>
          <tr> 
            <td nowrap class="smallerheader">&nbsp;&nbsp;<%= Session("emp_LastName") %>&nbsp;</td>
            <td><input type="text" class="champ" size="30" name="qryLastName" value="<%= Session("qryLastName") %>"></td>
          </tr>
          <tr> 
            <td nowrap class="smallerheader">&nbsp;&nbsp;<%= Session("emp_Title") %>&nbsp;</td>
            <td><input type="text" class="champ" size="30" name="qryTitle" value="<%= Session("qryTitle") %>"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td> <input name="submit" type="submit" class="cmdflat" value=" Chercher "> 
              <input name="button" type="button" class="cmdflat" onClick="javCancel()" value="Annuler"> 
            </td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td><img src="bas.gif"></td>
    </tr>
  </table>
</form>

</body>
</html>

<% Conn.Close %>