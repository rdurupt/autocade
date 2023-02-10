<% Response.Expires = 0 

%>
<!--#INCLUDE FILE="Portal_Common_db.asp"-->
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="inc_Utilities.asp"-->
<%
   
Set RS1 = Conn.Execute("SELECT * FROM dbp_UserInfos WHERE UserID =" & session("USERID"))
If Not RS1.EOF Then
    session("strTmplMain")=""
    Session("web_UserID") = RS1("UserID")
    Session("emp_Access") = GetUserDefault(Session("web_UserID"),"emp_Access","0")
    Session("FullName") = RS1("FirstName") & " " & RS1("LastName")
	session("Category")=Request.QueryString ("Category")
        pg = "emp_frameset.asp?nomenu=1&msg=Status:+Login+successful."
Else
    pg = "default.asp?msg=Status:+Invalid+password."
End If

Conn.Close
%>

<html>
<script language="JavaScript">
function GoThere() {
	location.href = "<%= pg %>"
}
</script>
<body onLoad="GoThere()">
</body>
</html>