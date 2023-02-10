<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server"  >
    'Public Sub SampleTreeView_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs)
        
    'Dim a = Me.Eval("form1.text")
    'a.text = "toto"
    'Dim myXml As New XmlDocument
    'myXml = CType(XmlSource.GetXmlDocument(), XmlDataDocument)

    'Dim iterator As String = SampleTreeView.SelectedNode.Expanded = "true"
    'Dim myNode As XmlNode = myXml.SelectSingleNode(iterator)

    'myNode.InnerText = "ThisIsATest"
    'XmlSource.Save()
    'TreeView1.DataBind()
    'TreeView1.ExpandAll()
    'End Sub

   
    
</script>

<html>
<head runat="server">
    <title>Page sans titre</title>
</head>
<script language='javascript'>
function v(){
alert('a');
}


</script>
<body background="ImageSysteme/Num.gif">

  
   

   <form id="form1" runat="server" action="//eboutique/"  Target="FrmCentre"  method="post">
       &nbsp;
       <asp:Menu ID="Menu1" runat="server" Target="gauche2">
           <Items>
              
                   <asp:MenuItem NavigateUrl="~/G2.aspx?MyModule=Nomenclature" Text="Nomenclature" Value="Nomenclature">
                   </asp:MenuItem>
                   <asp:MenuItem NavigateUrl="~/G2.aspx?MyModule=Autre" Text="Autre" Value="Autre"></asp:MenuItem>
              
           </Items>
       </asp:Menu>

    
    </form>
  
</body>
</html>
