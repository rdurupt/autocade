<%@ Page Language="VB" %>

<script runat="server">

  Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

    ' Create a new TreeView control.
        Dim NewTree As TreeView
        Dim con As New Adodb
        Dim RsDepartement As Object
        Dim ValK As String
        Dim ChampValK As String
        Dim ChampM1ValK As String
        
        
        Dim a As New PageMain
        Dim Node_Departement As Object
        Dim Node_Root As Object
       
        Dim KPrim As Integer
        Dim KPrim2 As Integer
        
        Dim C_CleAc As Collection
        Dim C_CleAc_Name As New Collection
        Dim C_Client As New Collection
        Dim C_Client_Name As New Collection
        
        Dim C_Type As New Collection
        Dim C_Type_Name As New Collection
        Dim C_Departement As New Collection
        Dim C_Departement_Name As New Collection
        Dim C_Activite As New Collection
        Dim C_C_Activite_Name As New Collection
        Dim C_Arber As New Collection
        Dim C_Arber_Name As New Collection
        Dim IndexColec As Integer
        Dim Test As Object
        
        con.OpenConnection("sql_numero")
        con.Sql = "SELECT     TOP 100 PERCENT dbo.T_Departement.Departement , dbo.Activité.*, dbo.Chrono.* "
        con.Sql = con.Sql & "FROM         dbo.T_Departement LEFT OUTER JOIN "
        con.Sql = con.Sql & "dbo.Activité ON dbo.T_Departement.Id = dbo.Activité.Id_Departement LEFT OUTER JOIN "
        con.Sql = con.Sql & "dbo.Chrono ON dbo.Activité.[Clé ac] = dbo.Chrono.[Clé ac] "
        con.Sql = con.Sql & "ORDER BY dbo.T_Departement.Departement, dbo.Activité.[Clé ac] DESC, dbo.Chrono.[Clé Ch] DESC"
        
        con.Sql = "SELECT     TOP 100 PERCENT dbo.T_Departement.Id AS [§1],  "
        con.Sql = con.Sql & "'Département' AS [£NoN£1],dbo.T_Departement.Departement AS [£NoN£2], dbo.Activité.Client AS [£NoN£3], 'Activité' as Activites,dbo.Activité.[Clé ac], dbo.Chrono.[Clé Ch] "
        con.Sql = con.Sql & "FROM         dbo.T_Departement  LEFT OUTER JOIN "
        con.Sql = con.Sql & "dbo.Activité ON dbo.T_Departement.Id = dbo.Activité.Id_Departement LEFT OUTER JOIN "
        con.Sql = con.Sql & "dbo.Chrono ON dbo.Activité.[Clé ac] = dbo.Chrono.[Clé ac] "
        con.Sql = con.Sql & "ORDER BY dbo.T_Departement.Departement, dbo.Activité.[Clé ac] DESC, dbo.Chrono.[Clé Ch] DESC"
        
        
        
        RsDepartement = con.OpenRecordset()
        NewTree = TreeView1
        
        ' Set the properties of the TreeView control.
    

        ' Create the tree node binding relationship.
        On Error Resume Next
        ' Create the root node binding.
        'Node_Root = a.SetTree("Departement", "PagePrincipale.aspx?MyModule=Nomenclature&Page=Departement", NewTree, True)
        KPrim = -1
        For IndexColec = 0 To RsDepartement.Fields.count - 1
            If Left(RsDepartement(IndexColec).name, 1) <> "§" Then Exit For
            KPrim = KPrim + 1
        Next
        While RsDepartement.eof = False
            For IndexColec = KPrim + 1 To RsDepartement.Fields.count - 1
                ValK = ""
                For KPrim2 = 0 To KPrim
                    ValK = ValK & RsDepartement(KPrim2).value & "§"
                Next
                If Trim("" & RsDepartement(IndexColec).value) <> "" Then
                  
                    If Left(RsDepartement(IndexColec).name, Len("£NoN£")) = "£NoN£" Then
                        ChampValK = Trim("" & RsDepartement(IndexColec).value)
                    Else
                        ChampValK = Trim("" & ValK & RsDepartement(IndexColec).value)
                    End If
                    If Left(RsDepartement(IndexColec - 1).name, Len("£NoN£")) = "£NoN£" Then
                        
                        ChampM1ValK = Trim("" & RsDepartement(IndexColec - 1).value)
                    Else
                        
                        ChampM1ValK = Trim("" & ValK & RsDepartement(IndexColec - 1).value)
                    End If
                    Test = Nothing
                    Test = C_Arber_Name(ChampValK)
                    Err.Clear()
                    Test = C_Arber_Name(ChampValK)
                    If Err.Number <> 0 Then
                        Err.Clear()
                        If IndexColec = KPrim + 1 Then
                            C_Arber.Add(a.SetTree(RsDepartement(IndexColec).value, "PagePrincipale.aspx?MyModule=Nomenclature&Page=Departement&Id=" & RsDepartement(IndexColec).value, NewTree, True, False), ChampValK)
                        Else
                            C_Arber.Add(a.SetTree(RsDepartement(IndexColec).value, "PagePrincipale.aspx?MyModule=Nomenclature&Page=Departement&Id=" & RsDepartement(IndexColec).value, C_Arber(ChampM1ValK), False, True), ChampValK)
                        End If
                    End If
                    C_Arber_Name.Add(Trim("" & ChampValK), "" & ChampValK)
                End If
               
            Next
            'C_Arber
            'Test = Nothing
            'Err.Clear()
            'Test = C_Departement_Name(RsDepartement("Departement").value)
            'If Err.Number <> 0 Then
            '    Err.Clear()
            '    Node_Departement = a.SetTree(RsDepartement("Departement").value, "PagePrincipale.aspx?MyModule=Nomenclature&Page=Departement&Id=" & RsDepartement("Id_Departement").value, Node_Root, False, True)
            'End If
            'C_Departement.Add(Node_Departement, RsDepartement("Departement").value)
            'C_Departement_Name.Add(RsDepartement("Departement").value, RsDepartement("Departement").value)
            ''Node_Departement = a.SetTree("Activiée", "")
            
            'If Trim("" & RsDepartement("Client").value) <> "" Then
            '    Test = C_Client_Name(RsDepartement("Client").value)
            '    If Err.Number <> 0 Then
            '        Err.Clear()
            '        Node_Departement = a.SetTree(RsDepartement("Client").value, "", C_Departement(RsDepartement("Departement").value), False, True)
            '    End If
            '    'C_Client.Add(a.SetTree(RsDepartement("Client").value, ""), RsDepartement("Client").value)
            '    C_Client_Name.Add(RsDepartement("Departement").value, RsDepartement("Client"))
            'End If
            ''Client 
           
           
       
        
            RsDepartement.movenext()
        End While
        ''For IndexColec = 1 To C_Client_Name.Count
        ''    C_Departement(C_Client_Name(IndexColec)).ChildNodes.Add(C_Client(IndexColec))
        ''Next
        ''For IndexColec = 1 To C_Departement.Count
        ''    Node_Root.ChildNodes.Add(C_Departement(IndexColec))
        ''Next
        ''Node_Root.ChildNodes.Add(Node_Departement)
        'NewTree.Nodes.Add(Node_Root)
       
       
        con.CloseConnection()
    End Sub



    Protected Sub TreeView1_SelectedNodeChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub
</script>

<html>
  <body>
    <form id="Form1" runat="server">
    
      <h3>TreeView Constructor Example</h3>
      
      
   
      
      
      
      <br>
        <asp:TreeView ID="TreeView1" runat="server" Target="FrmCentre" ShowLines="True" OnSelectedNodeChanged="TreeView1_SelectedNodeChanged">
     
        </asp:TreeView>
        <br>
      
      <asp:Label id="Message" runat="server"/>
    
    </form>
  </body>
</html>