<% 
Response.Expires = 0 
Session("CurrentPage") = "emp_frmSetField.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->

<!--#INCLUDE FILE="emp_TopMenu.asp"-->
<%
if Request("mode") = "execute" then
	Set RS1 = Conn.Execute("SELECT FieldID FROM dbp_DirDefFields")
	do while not RS1.EOF 
		key = "key" & RS1("FieldID")
		strSQL = "UPDATE dbp_DirDefFields SET "
		strSQL = strSQL & "FieldAlias = '" & safeEntry(Request(key)) & "' "
		strSQL = strSQL & "WHERE FieldID = " & RS1("FieldID")
		Conn.Execute(strSQL)
		if len(Request(key)) = 0 then
			Conn.Execute("DELETE FROM dbp_DirFields WHERE FieldID = " & RS1("FieldID"))
		end if
		
		RS1.movenext
	loop
	msg = "Champs mis à jour."
end if
%>
<html>
<head>
<title><%= GetDefault("emp_AppTitle","Employee Manager") %></title>
</head>
<script language="javascript">
function javCancel() {
	location.href = "ASPIntranet.asp?mode=emp_lst&menuEmployee=EmployeeList"
}
</script>
<link rel="stylesheet" href="<%=session("stylesheet")%>">
<body background="Background.asp">



<% '****** Header ***** %><br>
<form name="frm" action="emp_frmSetField.asp">
<table width="617" border="0" align="center" cellpadding="0" cellspacing="0" background="bg.jpg">
          <tr> 
            <td><img src="haut.gif" width="617" height="24"></td>
          </tr>
          <tr> 
            <td><div align="right"></div>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="50%"><div align="right"><img src="haut2.gif" width="25" height="18"></div></td>
                  <td width="50%" background="haut_bg.gif"><div align="center"><font color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>" size="2" face="Arial, Helvetica, sans-serif"><b>Champs</b></font> 
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td><center>
                <center>
                  <p><font color="red"><b><%= msg %></b></font></p>
                  </center><br>
                <center>
                  <TABLE width="95%" border="0" align="center" cellpadding="3" cellspacing="0">
                    <TR valign="top"> 
                      
                <TD align="center"> <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr class="smallheader"> 
                      <td><div align="right">CHAMP</div></td><td width="5%">&nbsp;</td>
                      <td><div align="center">ALIAS</div></td>
                    </tr>
                    <% 
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields ORDER BY FieldAlias") 
do while not RS0.EOF
	if left(RS0("FieldName"),3) <> "fld" then 
		pr ("<tr>")
		pr ("<td align='right'><font class='smallerheader'>" & RS0("FieldName") & "</font></td>")
		pr ("<td>&nbsp;</td>")
		pr ("<td align='center'><input type='text' class='champ' name='key" & RS0("FieldID") & "' value='" & RS0("FieldAlias") &"' size='30'></td>")
		pr ("</tr>")
	end if
	RS0.movenext
loop
%>
                  </table></TD>
                    </TR>
                  </table>
            <br>
            <input name="submit" type="submit" class="cmdflat" value="Valider">
          </center>
        </center></td>
          </tr>
          <tr> 
            <td><img src="bas.gif"></td>
          </tr>
        </table>
        

</form>
</body>
</html>
<% Conn.Close %>
