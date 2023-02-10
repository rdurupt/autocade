<% 
Response.Expires = 0 
Session("CurrentPage") = "emp_frmField.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->

<!--#INCLUDE FILE="emp_TopMenu.asp"-->
<%
if Request("mode") = "execute" then
	Conn.Execute("DELETE FROM dbp_DirFields")
	strArray = split(Request("lstFieldID"), ",")
	
	for i = 0 to ubound(strArray)
		order = "order" & trim(strArray(i))
		FieldOrder = Request(order)
		if not isnumeric(FieldOrder) then
			FieldOrder = 0
		end if
		Conn.Execute("INSERT INTO dbp_DirFields(UserID,FieldID,FieldOrder) VALUES(0," & trim(strArray(i)) & "," & FieldOrder & ")")
	next
	
	'**** Reindex
	'i = 0
	'Set RS0 = Conn.Execute("SELECT * FROM dbp_DirFields ORDER BY FieldOrder") 
	'do while not RS0.EOF
	'	Set RS1 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldID = " & RS0("FieldID")) 
	'	if not RS1.EOF then
	'		i = i + 1
	'		Conn.Execute("UPDATE dbp_DirFields SET FieldOrder = " & i & " WHERE UserID = " & Session("web_UserID") & " AND FieldID = " & RS0("FieldID")) 
	'	end if
	'	RS0.movenext
	'loop
	msg = "Affichage mis à jour."
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
<% '****** Header ***** %>
<br>
<form name="frm" action="emp_frmField.asp">

      <td> <table width="617" border="0" align="center" cellpadding="0" cellspacing="0" background="bg.jpg">
          <tr> 
            <td><img src="haut.gif" width="617" height="24"></td>
          </tr>
          <tr> 
            <td><div align="right"></div>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="50%"><div align="right"><img src="haut2.gif" width="25" height="18"></div></td>
                  <td width="50%" background="haut_bg.gif"><div align="center"><font color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy")  %>" size="2" face="Arial, Helvetica, sans-serif"><b>Affichage</b></font> 
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td><center>
                <font color="red"><b><%= msg %></b></font>
              </center>
              
          <div align="center" class="smallheader"><br>
            &nbsp;&nbsp;S&eacute;lectionnez les champs que vous voulez voir apparaître 
            dans la liste. <font color="#FF6600"><br>
            </font><br>
          </div>
          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr class="smallheader" > 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td align="left">CHAMP</td>
                  <td>ORDRE</td>
                </tr>
                <% 
Set RS0 = Conn.Execute("SELECT * FROM dbp_DirFields ORDER BY FieldOrder") 
do while not RS0.EOF
	Set RS1 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldID = " & RS0("FieldID")) 
	if not RS1.EOF then
		pr ("<tr>")
		pr ("<td>&nbsp;</td>")
		pr ("<td align='right'><input type='checkbox' name='lstFieldID' value='" & RS1("FieldID") &"' checked></td>")
		pr ("<td><font size='1' class='smallerheader'>" & RS1("FieldAlias") & "</font></td>")
		pr ("<td><input type='text' class='champ' name='order" & RS1("FieldID") & "' value='" & RS0("FieldOrder") &"' size='4'></td>")
		pr ("</tr>")
	end if
	RS0.movenext
loop


Set RS0 = Conn.Execute("SELECT * FROM dbp_DirDefFields WHERE FieldAlias <> '' ORDER BY FieldID") 
do while not RS0.EOF
	Set RS1 = Conn.Execute("SELECT * FROM dbp_DirFields WHERE FieldID = " & RS0("FieldID")) 
	if RS1.EOF then
		pr ("<tr>")
		pr ("<td>&nbsp;</td>")
		pr ("<td align='right'><input type='checkbox' name='lstFieldID' value='" & RS0("FieldID") &"'></td>")
		pr ("<td><font class='smallerheader'>" & RS0("FieldAlias") & "</font></td>")
		pr ("<td><input type='text' class='champ' name='order" & RS0("FieldID") & "' value='' size='4'></td>")
		pr ("</tr>")
	end if
	RS0.movenext
loop
%>
              </table>
              <div align="center"><br>
                <input name="submit" type="submit" class="cmdflat" value="Valider">
              </div></td>
          </tr>
          <tr> 
            <td><img src="bas.gif"></td>
          </tr>
        </table>
       
        



</form>
</body>
</html>
<% Conn.Close %>
