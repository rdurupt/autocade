Imports Microsoft.VisualBasic

Public Class PageMain
    Public Function MyModule() As String


        MyModule = "<table width=""70%"" border=""0"" align=""center"">"
        MyModule = MyModule & "<tr>"
        MyModule = MyModule & "<td><a href=""PagePrincipale.aspx?MyModule=Numerotation"">Numerotation</td>"
        MyModule = MyModule & "</tr>"
        MyModule = MyModule & "<tr>"
        MyModule = MyModule & "<td><a href=""PagePrincipale.aspx?MyModule=Stat"">Stat</td>"
        MyModule = MyModule & "</tr>"
        MyModule = MyModule & "</table>"
    End Function
    Public Function SetTree(ByVal MyValue As String, ByVal MyPage As String, Optional ByVal NewTree As Object = Nothing, Optional ByVal Paire As Boolean = False, Optional ByVal Fils As Boolean = False) As Object

        Dim RootBinding As TreeNode

        RootBinding = New TreeNode
        RootBinding.Text = Trim(MyValue)
        RootBinding.Value = RootBinding.Text
        RootBinding.NavigateUrl = MyPage
        RootBinding.Expanded = "false"
        If Fils = True Then
            NewTree.ChildNodes.Add(RootBinding)
        End If
        If Paire = True Then
            NewTree.Nodes.Add(RootBinding)
        End If
        SetTree = RootBinding

        RootBinding = Nothing
    End Function
    Public Sub RetourMenu(ByVal Rs)
        Dim Txt As String

        Txt = ""
        Txt = "<asp:TextBox ID=""TextBox1""   runat=""server"">toto</asp:TextBox>" & vbCrLf
        Txt = Txt & "<asp:TreeView id=""SampleTreeView""" & vbCrLf
        Txt = Txt & "runat=""server"" CollapseImageToolTip="""" " & vbCrLf
        Txt = Txt & "CollapseImageUrl=""~/ImageSysteme/icon_folder_open.gif"" " & vbCrLf
        Txt = Txt & "ExpandImageUrl=""~/ImageSysteme/icon_folder.gif"" " & vbCrLf
        Txt = Txt & "Font-Size=""15pt""" & vbCrLf
        Txt = Txt & "ForeColor=""Blue"" >" & vbCrLf
        Txt = Txt & "<Nodes>"


        'While Rs.eof = False


        Txt = Txt & "<asp:TreeNode Value=""Section 1"" " & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx"" " & vbCrLf
        Txt = Txt & "Text=""D&#233;partement""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre"" " & vbCrLf
        Txt = Txt & "Expanded=""False"">  " & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""" & Rs("id_departement").value & ">" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""DPSC""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""Renault""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & " NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""Activit&#233;es""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1"" " & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""868""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""Gestion""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False""> " & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""toto""" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "ext=""Production""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""PI""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""3150""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""Rev : 1""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre"" />" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""PL""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""3150""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1" & vbCrLf
        Txt = Txt & " NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""Rev : 1""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre"" />" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "</asp:TreeNode> " & vbCrLf
        Txt = Txt & " <asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""DAP""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre""" & vbCrLf
        Txt = Txt & "Expanded=""False"">" & vbCrLf
        Txt = Txt & "<asp:TreeNode Value=""Section 1""" & vbCrLf
        Txt = Txt & "NavigateUrl=""PagePrincipale.aspx""" & vbCrLf
        Txt = Txt & "Text=""toto""" & vbCrLf
        Txt = Txt & "Target=""FrmCentre"" /> " & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf
        Txt = Txt & "</asp:TreeNode>" & vbCrLf



        '    Rs.movenext()
        'End While





        Txt = Txt & "</Nodes>" & vbCrLf
        Txt = Txt & "</asp:TreeView>"

    End Sub

End Class
