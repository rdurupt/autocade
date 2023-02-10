<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Dim MyShoix As New PageMain

    'Protected Sub Menu1_MenuItemClick(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MenuEventArgs)
       
    'End Sub
</script>

<html >
<head runat="server">
    <title>Page sans titre</title>
</head>

<body background="ImageSysteme/Num.gif">
    &nbsp;<form id="form1" runat="server" method="post">
        
    <div>
        &nbsp;
        



        <%
            
            Select Case Request("MyModule")
                Case "Module"
                    Response.Write(MyShoix.MyModule())
                Case "Nomenclature"
                    Response.Write("Nomenclature")
                Case "Autre"
                    Response.Write("Autre")
    
                Case "2"
                    Response.Write("2")
                             
                Case Else
                   
                    Response.Write(MyShoix.MyModule())
            End Select%>
        <!--<input name="Hidden1" type="hidden" value=<%'= Request("Hidden1")%> />-->
    </div>
    </form>
</body>
</html>
