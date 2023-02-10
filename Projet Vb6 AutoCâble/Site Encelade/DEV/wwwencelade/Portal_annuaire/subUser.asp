<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<%
function funSafeEntry(strField)
	strSafe = replace(strField,"'","`")
	strSafe = replace(strSafe,"<","&lt;")
	strSafe = replace(strSafe,">","&gt;")
	funSafeEntry = strSafe
end function

'Edit
If Request("editMode") = "Edit" Then
    FirstName = funSafeEntry(Request("FirstName"))
    LastName = funSafeEntry(Request("LastName"))
    UserName = funSafeEntry(Request("UserName"))

    sql = "UPDATE Users SET "
    sql = sql & "FirstName = '" & FirstName & "',"
    sql = sql & "LastName = '" & LastName & "',"
    sql = sql & "UserName = '" & UserName & "',"
    sql = sql & "AccessLevel = " & Request("AccessLevel") & ","
    sql = sql & "WHERE UserID = " & Request("UserID")
    Conn.Execute (sql)
    
    pg = "lstUser.asp?msg=Status:+Record+updated."
End If

'Add
If Request("editMode") = "New" Then
    FirstName = funSafeEntry(Request("FirstName"))
    LastName = funSafeEntry(Request("LastName"))
    UserName = funSafeEntry(Request("UserName"))
    Password = funSafeEntry(Request("Password"))

    sql = "INSERT INTO Users (FirstName,LastName,UserName,Password,AccessLevel,CreateDate) VALUES ("
    sql = sql & "'" & FirstName & "',"
    sql = sql & "'" & LastName & "',"
    sql = sql & "'" & UserName & "',"
    sql = sql & "'" & Password & "',"
    sql = sql & "" & Request("AccessLevel") & ","
    sql = sql & "'" & Date & "')"
    Conn.Execute (sql)
    
    Set RS12 = Conn.Execute("SELECT Max(UserID) as [NewID] FROM Users")
    Session("CurrUserID") = RS12("NewID")
        
    pg = "lstUser.asp?msg=Status:+Record+added."
End If

'Delete
If Request("editMode") = "Delete" Then
    If CInt(Request("UserID")) <> 1 Then
        Conn.Execute("DELETE FROM Users WHERE UserID = " & Request("UserID"))
    End If
    Session("CurrUserID") = 0
    pg = "lstUser.asp?msg=Status:+Record+deleted"
    If CInt(Request("UserID")) = 1 Then
        pg = "lstUser.asp?msg=Status:+Cannot+delete+admin+account"
    End If
End If
%>

<html>
<script language="javascript">
function goThere() {
location.href = "<%= pg %>"
}
</script>
<body onload="goThere()">
</html>

<% Conn.Close %>
