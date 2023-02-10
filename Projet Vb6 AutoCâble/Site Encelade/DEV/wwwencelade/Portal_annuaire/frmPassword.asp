<% 
Response.Expires = 0 
Session("CurrentPage") = "frmPassword.asp"
%>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->


<html>
<head>
<!--#INCLUDE FILE="emp_TopMenu.asp"-->
<title><%= GetDefault("web_AppTitle","ASP Intranet Suite") %></title>
</head>
<script language="javascript">
function javSubmit() {
	if (document.frm.PASSW.value !== document.frm.PASSW1.value) {
		alert("Les mots de passe ne correspondent pas.")
		return
	} 
	document.frm.submit()
}

</script>
<link rel="stylesheet" href="<%=Session("StyleSheet")%>">
<body onLoad="document.frm.Password1.focus()">
<% '****** Header ***** %>

<br><form name="frm" action="aspintranet.asp">
<input type="hidden" name="mode" value="PasswordUpdate">
<input type="hidden" name="userid" value="<%=Session("web_UserID")%>">

  <table width="617" border="0" align="center" cellpadding="0" cellspacing="0" background="bg.jpg">
    <tr> 
      <td><img src="haut.gif" width="617" height="24"></td>
    </tr>
    <tr> 
      <td><div align="right"></div>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="50%"><div align="right"><img src="haut2.gif" width="25" height="18"></div></td>
            <td width="50%" background="haut_bg.gif"><div align="center"><font  color="<%= GetUserDefault(Session("web_UserID"),"emp_color","navy") %>" size="2" face="Arial, Helvetica, sans-serif"><b>Changement 
                de votre mot de passe</b></font> </div></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td>
        <table border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td>&nbsp;</td>
            <td><font color="red"><b><%= msg %></b></font></td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right">Ancien Mot de Passe&nbsp;</td>
            <td align="left"><input type="password" class="champ" name="Password1" value=""></td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right">Nouveau Mot de Passe&nbsp;</td>
            <td align="left"><input type="password" class="champ" name="PASSW" value=""></td>
          </tr>
          <tr class="smallerheader"> 
            <td align="right">Confirmer ce Mot de Passe&nbsp;</td>
            <td align="left"><input type="password" class="champ" name="PASSW1" value=""></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td align="left"> <input name="button" type="button" class="cmdflat" onClick="javSubmit()" value="Valider"> 
            </td>
          </tr>
        </table> </td>
    </tr>
    <tr> 
      <td><img src="bas.gif"></td>
    </tr>
  </table>
  </form>
</center>
</body>
</html>
<% Conn.Close %>