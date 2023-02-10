<%
Dim foo
Dim leGroupId
Dim leCatId
Dim SQL
Dim oTools
Dim aGroups
Dim sGroups
Dim aCats
Dim sCats

sGroups = ""
sCats = ""

Set oTools = Server.CreateObject(Session("PortalObject") & ".PortalTools")

foo = Request("foo")
leGroupId = Request("h_groupid")
leCatId = Request("h_catid")

SQL = "SELECT DISTINCT GroupId"
SQL = SQL & " FROM dbp_GroupsPermission"
SQL = SQL & " WHERE CId = " & Session("PortalId")
SQL = SQL & " AND ObjTypeId = 100"
SQL = SQL & " AND ObjId = " & leCatId
SQL = SQL & " AND GroupId <> " & leGroupId & ";"

aGroups = oTools.dbSelect(SQL)

If Not IsArray(aGroups) Then
 sGroups = leGroupId
End If

SQL = "SELECT tbl1.CatId"
SQL = SQL & " FROM ((dbp_Categories AS tbl1"
SQL = SQL & " INNER JOIN dbp_GroupsPermission AS tbl2"
SQL = SQL & " ON (tbl2.CId = tbl1.CId"
SQL = SQL & " AND tbl2.ObjId= tbl1.CatId))"
SQL = SQL & " INNER JOIN dbp_CatSort AS tbl3"
SQL = SQL & " ON (tbl2.CId = tbl3.CId"
SQL = SQL & " AND tbl2.ObjId = tbl3.CatId"
SQL = SQL & " AND tbl2.GroupId = tbl3.GroupId))"
SQL = SQL & " WHERE tbl1.CId = " & Session("PortalId")
SQL = SQL & " AND tbl1.CatType = 0"
SQL = SQL & " AND tbl2.GroupId = " & leGroupId
SQL = SQL & " AND tbl2.ObjTypeId = 100"
SQL = SQL & " ORDER BY tbl1.CatParent, tbl3.CatSort;"

aCats = oTools.dbSelect(SQL)

If IsArray(aCats) Then
 For i = 0 To UBound(aCats, 2)
  sCats = sCats & aCats(0, i) & ","
 Next
 sCats = Left(sCats, Len(sCats) - 1)
Else
 sCats = leCatId
End If

If IsArray(aGroups) Then
 For i = 0 To Ubound(aGroups, 2)
  SQL = "SELECT tbl1.CatId"
  SQL = SQL & " FROM ((dbp_Categories AS tbl1"
  SQL = SQL & " INNER JOIN dbp_GroupsPermission AS tbl2"
  SQL = SQL & " ON (tbl2.CId = tbl1.CId"
  SQL = SQL & " AND tbl2.ObjId= tbl1.CatId))"
  SQL = SQL & " INNER JOIN dbp_CatSort AS tbl3"
  SQL = SQL & " ON (tbl2.CId = tbl3.CId"
  SQL = SQL & " AND tbl2.ObjId = tbl3.CatId"
  SQL = SQL & " AND tbl2.GroupId = tbl3.GroupId))"
  SQL = SQL & " WHERE tbl1.CId = " & Session("PortalId")
  SQL = SQL & " AND tbl1.CatType = 0"
  SQL = SQL & " AND tbl2.GroupId = " & aGroups(0, i)
  SQL = SQL & " AND tbl2.ObjTypeId = 100"
  SQL = SQL & " AND tbl2.ObjId IN (" & sCats & ")"
  SQL = SQL & " ORDER BY tbl1.CatParent, tbl3.CatSort;"
  aCats = oTools.dbSelect(SQL)
  If IsArray(aCats) Then
   sCats = ""
   For j = 0 To UBound(aCats, 2)
    sCats = sCats & aCats(0, j) & ","
   Next
   sCats = Left(sCats, Len(sCats) - 1)
  Else
   sCats = leCatId
   Exit For
  End If
 Next
Else
 SQL = "SELECT tbl1.CatId"
 SQL = SQL & " FROM ((dbp_Categories AS tbl1"
 SQL = SQL & " INNER JOIN dbp_GroupsPermission AS tbl2"
 SQL = SQL & " ON (tbl2.CId = tbl1.CId"
 SQL = SQL & " AND tbl2.ObjId= tbl1.CatId))"
 SQL = SQL & " INNER JOIN dbp_CatSort AS tbl3"
 SQL = SQL & " ON (tbl2.CId = tbl3.CId"
 SQL = SQL & " AND tbl2.ObjId = tbl3.CatId"
 SQL = SQL & " AND tbl2.GroupId = tbl3.GroupId))"
 SQL = SQL & " WHERE tbl1.CId = " & Session("PortalId")
 SQL = SQL & " AND tbl1.CatType = 0"
 SQL = SQL & " AND tbl2.GroupId = " & leGroupId
 SQL = SQL & " AND tbl2.ObjTypeId = 100"
 SQL = SQL & " AND tbl2.ObjId IN (" & sCats & ")"
 SQL = SQL & " ORDER BY tbl1.CatParent, tbl3.CatSort;"
 aCats = oTools.dbSelect(SQL)
 If IsArray(aCats) Then
  sCats = ""
  For i = 0 To UBound(aCats, 2)
   sCats = sCats & aCats(0, i) & ","
  Next
  sCats = Left(sCats, Len(sCats) - 1)
 Else
  sCats = leCatId
 End If
End If
SQL = "SELECT tbl1.CatId,"
SQL = SQL & " tbl1.CatParent,"
SQL = SQL & " tbl1.CatCaption,"
SQL = SQL & " tbl1.CatLayoutId,"
SQL = SQL & " tbl1.CatExtUrl,"
SQL = SQL & " tbl1.CatIsObject,"
SQL = SQL & " tbl1.CatIdObject,"
SQL = SQL & " tbl1.CatIdTemplate,"
SQL = SQL & " tbl1.CatMode"
SQL = SQL & " FROM ((dbp_Categories AS tbl1"
SQL = SQL & " INNER JOIN dbp_GroupsPermission AS tbl2"
SQL = SQL & " ON (tbl2.CId = tbl1.CId"
SQL = SQL & " AND tbl2.ObjId= tbl1.CatId))"
SQL = SQL & " INNER JOIN dbp_CatSort AS tbl3"
SQL = SQL & " ON (tbl2.CId = tbl3.CId"
SQL = SQL & " AND tbl2.ObjId = tbl3.CatId"
SQL = SQL & " AND tbl2.GroupId = tbl3.GroupId))"
SQL = SQL & " WHERE tbl1.CId = " & Session("PortalId")
SQL = SQL & " AND tbl1.CatType = 0"
SQL = SQL & " AND tbl2.GroupId = " & leGroupId
SQL = SQL & " AND tbl2.ObjTypeId = 100"
SQL = SQL & " AND tbl2.ObjId IN (" & sCats & ")"
SQL = SQL & " ORDER BY tbl1.CatParent, tbl3.CatSort;"
aCats = oTools.dbSelect(SQL)
aCats = oTools.SortIdIdPar(aCats, 0, 1, 0)
leString = ""
leStrInternal = ""
For i = 0 To UBound(aCats, 2)
 If CInt(aCats(1, i)) = 0 Then
  leLevel = 1
 Else
  leLevel = oTools.GetLevel(aCats, i, 0, 1, 0)
 End If
 laNode = UCase(oTools.GetNode(aCats, i, 0, 1, 0))
 Select Case laNode
  Case "LAST"
   img = "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_nl.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
  Case "NOTLAST"
   img = "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_nt.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
  Case "LASTNODE"
   img = "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_nl.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
  Case "NODE"
   img = "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_nt.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
 End Select

            '   - 4 : CatExtUrl
            '   - 5 : CatIsObject
            '   - 6 : CatIdObject
            '   - 7 : CatIdTemplate
            '   - 8 : CatMode
 If CInt(aCats(5, i)) = 0 Then
  leLink = "goToTemplate(" & aCats(7, i) & ", '');goToCat(" & aCats(0, i) & ", '');goToPage('" & aCats(8, i) & "','" & aCats(4, i) & "§§', '');"
 Else
  SQL = "SELECT ObjUrl,"
  SQL = SQL & " ObjMode,"
  SQL = SQL & " ObjTemplateId"
  SQL = SQL & " FROM dbp_ObjectTypes"
  SQL = SQL & " WHERE CId = " & Session("PortalId")
  SQL = SQL & " AND ObjTypeId = " & aCats(6, i) & ";"
  tempURL = oTools.dbSelect(SQL)
  leLink = "goToCat(" & aCats(0, i) & ", '');goToTemplate(" & tempURL(2, 0) & ", '');goToPage('" & tempURL(1, 0) & "','" & Session("PortalPath") & tempURL(0, 0) & "§§', '');"
 End If
 leStrInternal = leStrInternal & leString & img & "&nbsp;"
 leStrInternal = leStrInternal & "<a href=""javascript:SetLink('" & Replace(Replace(leLink, "§§", ""), "'", "\'") & "');"" border=""0"" class=""linkCat"">"
 leStrInternal = leStrInternal & aCats(2, i) & "</a><br>" & vbCrLf
 SQL = "SELECT tbl1.TopicId,"
 SQL = SQL & " tbl1.TopicTitle,"
 SQL = SQL & " tbl1.TopicIsObject,"
 SQL = SQL & " tbl1.TopicIdObject"
 SQL = SQL & " FROM dbp_Topics AS tbl1"
 SQL = SQL & " WHERE tbl1.CId = " & Session("PortalId")
 SQL = SQL & " AND tbl1.TopicCatId = " & aCats(0, i)
 SQL = SQL & " AND EXISTS ("
 SQL = SQL & "  SELECT tbl2.LayoutObjId"
 SQL = SQL & "  FROM dbp_LayoutItems AS tbl2"
 SQL = SQL & "  WHERE tbl2.CId = tbl1.CId"
 SQL = SQL & "  AND tbl2.LayoutId = " & aCats(3, i)
 SQL = SQL & "  AND tbl2.LayoutObjId = tbl1.TopicId"
 SQL = SQL & "  AND tbl2.LayoutObjType = '101-1'"
 SQL = SQL & " )"
 SQL = SQL & " ORDER BY tbl1.TopicTitle;"
 aTopics = oTools.dbSelect(SQL)
 If IsArray(aTopics) Then
  For j = 0 To Ubound(aTopics, 2)
   leStrInternal = leStrInternal & leString
   Select Case laNode
    Case "LAST"
     leStrInternal = leStrInternal & "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_blank.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
    Case "LASTNODE"
     leStrInternal = leStrInternal & "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_blank.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
    Case Else
     leStrInternal = leStrInternal & "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_nv.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
   End Select
   If aCats(0, i) = aCats(1, i + 1) And i <> Ubound(aCats, 2) Then
    leStrInternal = leStrInternal & "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_nv.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">" & "&nbsp;"
   Else
    leStrInternal = leStrInternal & "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_blank.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">" & "&nbsp;"
   End If
   If Cint(aTopics(2, j)) = 0 Then
    leStrInternal = leStrInternal & "<a href=""javascript:SetLink('" & Replace(Replace(leLink, "§§", "#" & aTopics(0, j)), "'", "\'") & "');"" border=""0"" class=""linkTopic"">"
   Else
    SQL = "SELECT ObjUrl,"
    SQL = SQL & " ObjMode,"
    SQL = SQL & " ObjTemplateId"
    SQL = SQL & " FROM dbp_ObjectTypes"
    SQL = SQL & " WHERE CId = " & Session("PortalId")
    SQL = SQL & " AND ObjTypeId = " & aTopics(3, j) & ";"
    tempOBJ = oTools.dbSelect(SQL)
    leStrInternal = leStrInternal & "<a href=""javascript:SetLink('" & Replace("goToTopic(" & aTopics(0, j) & ", '');goToCat(" & aCats(0, i) & ", '');goToTemplate(" & tempOBJ(2, 0) & ", '');goToPage('" & tempOBJ(1, 0) & "','" & Session("PortalPath") & tempOBJ(0, 0) & "', '');", "'", "\'") & "');"" border=""0"" class=""linkTopic"">"
   End If
   leStrInternal = leStrInternal & aTopics(1, j) & "</a><br>" & vbCrLf
  Next
 End If
 Select Case laNode
  Case "LAST"
   leString = leString & "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_blank.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
  Case "NOTLAST"
   leString = leString & "<img src=""" & Session("PortalPath") & "/Portal_Image/Admin/tree_nv.gif"" width=""16"" height=""22"" border=""0"" align=""absmiddle"">"
  Case "LASTNODE"
   If i + 1 > UBound(aCats, 2) Then
    leString = ""
   Else
    If CInt(aCats(1, i + 1)) = 0 Then
     leNextLevel = 1
    Else
     leNextLevel = oTools.GetLevel(aCats, (i + 1), 0, 1, 0)
    End If
    For j = leLevel To leNextLevel + 1 Step -1
     If leString <> "" Then
      leString = Left(leString, InStrRev(leString, "<img") - 1)
     End If
    Next
   End If
 End Select
Next

Set oTools = Nothing

%>
<html>
<head>
<title>Insérer / Modifier un lien</title>
<link rel="stylesheet" type="text/css" href="../Css/HE_Style.css">
<script language="JavaScript" src="../Js/HE_Link.js"></script>
</head>

<body>
<FORM METHOD="POST" name="linkForm">
<table width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout:fixed;">
 <tr>
  <td width="15"></td>
  <td class="TitrePopup">Gestionnaire de Liens</td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">Veuillez saisir les informations requises pour l'insertion ou la modification d'un lien.<br>Cliquer le boutton "Annuler" pour fermer cette fenetre.</td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">&nbsp;</td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">
   <table width="98%" border="0" cellspacing="0" cellpadding="0" class="TblPopupTitre">
    <tr>
     <td>&nbsp;&nbsp;Informations sur le lien</td>
    </tr>
   </table>
  </td>
 </tr>
 <tr>
  <td colspan="2" height="10"></td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">
   <table border="0" cellspacing="0" cellpadding="5" width="98%" class="TblPopupInt">
    <tr>
     <td class="Text" width="100">URL :</td>
     <td class="Text">
      <input type="text" name="link" value="" class="Text" style="width:220;">
     </td>
    </tr>
    <tr>
     <td class="Text">Target Window :</td>
     <td class="Text">
      <input type="text" name="targetWindow" value="" class="Text" style="width:90;">
      <select name="targetText" class="Text" onChange="targetWindow.value = targetText[targetText.selectedIndex].value; targetText.value = ''; targetWindow.focus();" style="width:90;">
       <option value=""></option>
       <option value="">None</option>
       <option value=_blank>_blank</option>
       <option value=_parent>_parent</option>
       <option value=_self>_self</option>
       <option value=_top>_top</option>
      </select>
     </td>
    </tr>
    <tr>
     <td class="Text">Ancre :</td>
     <td class="Text">
      <select name="targetAnchor" class="Text" onChange="link.value = targetAnchor[targetAnchor.selectedIndex].value; targetAnchor.value = ''; link.focus();" style="width:90;">
       <option value=""></option>
       <script>getAnchors()</script>
      </select>
     </td>
    </tr>
   </table>
  </td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">&nbsp;</td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">
   <table width="98%" border="0" cellspacing="0" cellpadding="0" class="TblPopupTitre">
    <tr>
     <td>&nbsp;&nbsp;Liens internes</td>
    </tr>
   </table>
  </td>
 </tr>
 <tr>
  <td colspan="2" height="10"></td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">
   <table border="0" cellspacing="0" cellpadding="5" width="98%" class="TblPopupInt">
    <tr>
     <td class="Text">
      <%Response.Write leStrInternal%>
     </td>
    </tr>
   </table>
  </td>
 </tr>
 <tr>
  <td colspan="2" height="10"></td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td>	
   <input type="button" name="insertLink" value="Insérer le Lien" class="Text" onClick="javascript:InsertLink();">
   <input type="button" name="removeLink" value="Supprimer le Lien" class="Text" onClick="javascript:RemoveLink();">
   <input type=button name="Cancel" value="Annuler" class="Text" onClick="javascript:window.close();">
  </td>
 </tr>
</table>
</form>
<script language="JavaScript">
getLink();
</script>
</body>
</html>