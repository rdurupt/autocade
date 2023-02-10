<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server" >
    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
Select Case Request("MyModule")
Case "Nomenclature"

    ' Create a new TreeView control.
        Dim NewTree As TreeView
        Dim con As New Adodb
        Dim RsDepartement As Object
        Dim RsCliant As Object
        Dim a As New PageMain
        Dim Node_Departement As Object
        Dim Node_Root As Object
        Dim Node_Cliant As Object
        Dim Node_Activite As Object
        Dim Node_CleAc As Object
        Dim Node_Type As Object
        Dim RS_Type As Object
        Dim RS_CleAc As Object
        Dim SaveClient As String
        Dim T_Type As Collection
        Dim IndexColec As Long
        con.OpenConnection("sql_numero")
        con.Sql = "SELECT     dbo.T_Departement.*"
        con.Sql = con.Sql & "FROM T_Departement order by Departement"
        RsDepartement = con.OpenRecordset()
        NewTree = TreeView1
        
    ' Set the properties of the TreeView control.
    

    ' Create the tree node binding relationship.

        ' Create the root node binding.
        Node_Root = a.SetTree( "Departement", "PagePrincipale.aspx?MyModule=Nomenclature&Page=Departement")
        While RsDepartement.eof = False
            
            Node_Departement = a.SetTree( RsDepartement("Departement").value, "PagePrincipale.aspx?MyModule=Nomenclature&Page=Departement&Id=" & RsDepartement("Id").value)
            con.Sql = "SELECT     TOP 100 PERCENT Client, Id_Departement "
            con.Sql = con.Sql & "FROM          dbo.Activité "
            con.Sql = con.Sql & "WHERE     Client <> '' AND Id_Departement = " & RsDepartement("id").value & " "
            con.Sql = con.Sql & "GROUP BY Client, Id_Departement  "
            con.Sql = con.Sql & "ORDER BY dbo.Activité.[Client] "
            
           
            
            
            'con.Sql = "SELECT     TOP 100 PERCENT Activité.* "
            'con.Sql = con.Sql & "dbo.Activité LEFT OUTER JOIN "
            'con.Sql = con.Sql & "dbo.Chrono ON dbo.Activité.[Clé ac] = dbo.Chrono.[Clé ac] "
            'con.Sql = con.Sql & "WHERE dbo.Activité.Id_Departement = " & RsDepartement("id").value & " "
            'con.Sql = con.Sql & "dbo.Chrono.[Date] DESC, dbo.Activité.Client "
     
            RsCliant = con.OpenRecordset()
            SaveClient = ""
            While RsCliant.eof = False
               
                        Node_Cliant = a.SetTree(RsCliant("Client").value, "PagePrincipale.aspx?MyModule=Nomenclature&Page=Client")
                Node_Activite = a.SetTree( "Avtivitée", "PagePrincipale.aspx?MyModule=Nomenclature&Page=Avtivitée")
                con.Sql = "SELECT     TOP 100 PERCENT [Clé ac],Id_Departement "
                con.Sql = con.Sql & "FROM  dbo.Activité "
                con.Sql = con.Sql & "WHERE     Id_Departement = " & RsCliant("Id_Departement").value & " AND Client = '" & RsCliant("Client").value & "' "
                con.Sql = con.Sql & "GROUP BY [Clé ac],Id_Departement "
                con.Sql = con.Sql & "ORDER BY [Clé ac] DESC"
                
                RS_CleAc = con.OpenRecordset()
                While RS_CleAc.eof = False
                    Node_CleAc = a.SetTree( "" & RS_CleAc("Clé ac").value, "")
                   
                    
                    con.Sql = "SELECT     TOP 100 PERCENT dbo.Chrono.* "
                    con.Sql = con.Sql & "FROM dbo.Chrono "
                    con.Sql = con.Sql & "WHERE     [Clé ac] = " & RS_CleAc("Clé ac").value & " "
                    con.Sql = con.Sql & "ORDER BY [Date] DESC, [Clé ty]"
                    RS_Type = con.OpenRecordset()
                    T_Type = New Collection
                    While RS_Type.eof = False
                        Node_Type = a.SetTree( "" & RS_Type("Clé ty").value, "")
                        On Error Resume Next
                        T_Type.Add(Node_Type, "Ck_" & RS_Type("Clé ty").value)
                        T_Type("Ck_" & RS_Type("Clé ty").value).ChildNodes.Add(a.SetTree(T_Type("Ck_" & RS_Type("Clé ty").value), "" & RS_Type("Clé ty").value & "_" & RS_Type("Clé ac").value & "_" & RS_Type("Année").value & "_" & RS_Type("Clé Ch").value & "_" & RS_Type("Rév").value, ""))
                        On Error GoTo 0
                        RS_Type.movenext()
                    End While
                    For IndexColec = 1 To T_Type.Count
                        Node_CleAc.ChildNodes.Add(T_Type(IndexColec))
                    Next
                    Node_Activite.ChildNodes.Add(Node_CleAc)
                    RS_CleAc.movenext()
                End While
                
                
                Node_Cliant.ChildNodes.Add(Node_Activite)
                
               
                Node_Departement.ChildNodes.Add(Node_Cliant)
                RsCliant.movenext()
            End While
            Node_Root.ChildNodes.Add(Node_Departement)
            RsDepartement.movenext()
        End While
        NewTree.Nodes.Add(Node_Root)
       
       
        con.CloseConnection()
        end select
    End Sub


</script>

<html>
<head runat="server">
    <title>Page sans titre</title>
</head>
<body background="ImageSysteme/Num.gif">
<form id="form1" runat="server" action="//eboutique/"  method="post">         


  
   <%  Select Case Request("MyModule")
           Case "Nomenclature"
               %>
             
 <asp:TreeView ID="TreeView1" runat="server" Target="FrmCentre" ShowLines="True">
     
        </asp:TreeView>

   

      
       <script language="javascript">
            parent.frames['FrmCentre'].location='PagePrincipale.aspx?MyModule=Nomenclature'
        </script>
       
 <%  
      Case "Autre"
        %>page à dévelop
    per
        <script language="javascript">
            parent.frames['FrmCentre'].location='PagePrincipale.aspx?MyModule=Autre'
        </script>
        
        <%
                               
 Case Else
         %>
    Séle ct ionez le module<%
                            End Select
                                
                           
                                %>&nbsp;
                        
            
    </form>
 
</body>
</html>
