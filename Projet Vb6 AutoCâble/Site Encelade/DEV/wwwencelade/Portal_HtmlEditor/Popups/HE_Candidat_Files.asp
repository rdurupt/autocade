<%


lePathRoot = Session("PortalPath") & Session("lePathRootDoc")
leFormBack = Request("h_formback")
leImgChange = Request("h_imgchange")
Public Function FormatDateForOrder(inDATE)
 Dim sOut
 sOut = Year(inDATE) & _ 
        FormatTwoDigit(Month(inDATE)) & _ 
        FormatTwoDigit(Day(inDATE)) & _ 
        FormatTwoDigit(Hour(inDATE)) & _ 
        FormatTwoDigit(Minute(inDATE)) & _ 
        FormatTwoDigit(Second(inDATE))
 FormatDateForOrder = sOut
End Function

Public Function FormatTwoDigit(inDIGIT)
 Dim sOut
 If Len(inDIGIT) < 2 Then
  sOut = "0" & inDIGIT
 Else
  sOut = inDIGIT
 End If
 FormatTwoDigit = sOut
End Function

Public Function MultiSort(ByVal inARRAY, inCOL, inORDER)
 nbC = UBound(inARRAY, 1)
 nbR = UBound(inARRAY, 2)
 ReDim aOut(nbC, nbR)
 ' On passe le premier
 For i = 0 To nbC
  aOut(i, 0) = inARRAY(i, 0)
 Next
 lastIns = 1
 ' On passe tout le reste
 For i = 1 To nbR
  ' Ou on insere ?
  swapPos = -1
  ' Dans quel sens on tri ?
  If UCase(inORDER) = "ASC" Then
   For j = 0 To nbR
    If StrComp(inARRAY(inCOL, i), aOut(inCOL, j), vbTextCompare) < 0 Then
     swapPos = j
     Exit For
    End If
   Next
  Else
   For j = 0 To nbR
    If StrComp(inARRAY(inCOL, i), aOut(inCOL, j), vbTextCompare) > 0 Then
     swapPos = j
     Exit For
    End If
   Next
  End If
  If swapPos = -1 Then
   For j = 0 To nbC
    aOut(j, lastIns) = inARRAY(j, i)
   Next
  Else
   For j = nbR To swapPos Step -1
    For k = 0 To nbC
     If j <> 0 Then
      aOut(k, j) = aOut(k, j - 1)
     End If
    Next
   Next
   For j = 0 To nbC
    aOut(j, swapPos) = inARRAY(j, i)
   Next
  End If
  lastIns = lastIns + 1
 Next
 MultiSort = aOut
End Function

Dim FSO
Dim lePathEnCours
Dim leRep

If Request("f_PathTo").Count > 0 Then
 lePathEnCours = Request("f_PathTo") & "/"
Else
 lePathEnCours = lePathRoot
End If

If Request("f_OrderBy").Count > 0 Then
 leOrderBy = CInt(Request("f_OrderBy"))
Else
 leOrderBy = 0
End If

If Request("f_OrderStr").Count > 0 Then
 leOrderStr = Request("f_OrderStr")
Else
 leOrderStr = "ASC"
End If

Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set leRep = FSO.GetFolder(Server.MapPath(lePathEnCours))

If leRep.SubFolders.Count > 0 Then
 Redim aReps(2, leRep.SubFolders.Count - 1)
 indx = 0
 For Each sFold In leRep.SubFolders
  aReps(0, indx) = sFold.Name
  aReps(1, indx) = sFold.DateLastModified
  aReps(2, indx) = FormatDateForOrder(sFold.DateLastModified)
  indx = indx + 1
 Next
Else
 aReps = -1
End If

If leRep.Files.Count > 0 Then
 Redim aFiles(2, leRep.Files.Count - 1)
 indx = 0
 For Each sFile In leRep.Files
  aFiles(0, indx) = sFile.Name
  aFiles(1, indx) = sFile.DateLastModified
  aFiles(2, indx) = FormatDateForOrder(sFile.DateLastModified)
  indx = indx + 1
 Next
Else
 aFiles = -1
End If

Set leRep = Nothing
Set FSO = Nothing

If IsArray(aReps) Then
 aRep = MultiSort(aReps, leOrderBy, leOrderStr)
Else
 aRep = -1
End If

If IsArray(aFiles) Then
 aFile = MultiSort(aFiles, leOrderBy, leOrderStr)
Else
 aFile = -1
End If
%>
<html>
<head>
<title>Insérer unlien vers un fichier</title>
<link rel="stylesheet" type="text/css" href="../Css/HE_Style.css">
<script language="JavaScript">
function jsTrim(monItem) {
 var monTexte = new String("");
 monTexte = monItem.value;
 while (monTexte.charAt(0) == ' ') {
  monTexte = monTexte.substring(1,monTexte.length);
 }
 while (monTexte.charAt(monTexte.length - 1) == ' ') {
  monTexte = monTexte.substring(0, (monTexte.length - 1));
 }
 monItem.value = monTexte;
};
function GoToFolder(inFolder) {
 document.ficForm.f_PathTo.value = document.ficForm.f_PathTo.value + inFolder;
 document.ficForm.action = "?";
 document.ficForm.submit();
};
function GoToParent() {
 var sFold = document.ficForm.f_PathTo.value;
 sFold = sFold.substring(0, sFold.length - 1);
 sFold = sFold.substring(0, sFold.lastIndexOf("/"));
 document.ficForm.f_PathTo.value = sFold;
 document.ficForm.action = "?";
 document.ficForm.submit();
};
function GoToRoot() {
 document.ficForm.f_PathTo.value = "<%Response.Write Left(lePathRoot, Len(lePathRoot) - 1)%>";
 document.ficForm.action = "?";
 document.ficForm.submit();
};
function GoToNav(inFolder) {
 document.ficForm.f_PathTo.value = inFolder;
 document.ficForm.action = "?";
 document.ficForm.submit();
};
function GoToOrder( inORDERBY, inORDERSTR) {
 var str = document.ficForm.f_PathTo.value;
 if (str.substring(str.length - 1, str.length) == "/") {
  str = str.substring(0, str.length - 1);
 }
 document.ficForm.f_PathTo.value = str;
 document.ficForm.f_OrderBy.value = inORDERBY;
 document.ficForm.f_OrderStr.value = inORDERSTR;
 document.ficForm.action = "?";
 document.ficForm.submit();
};
function CheckFolderName() {
 var myError = 0;
 var str = document.ficForm.f_NewFolder.value;
 if (str.indexOf(String.fromCharCode(92)) >= 0) {
  myError = 1;
 } else if (str.indexOf(String.fromCharCode(47)) >= 0) {
  myError = 1;
 } else if (str.indexOf(":") >= 0) {
  myError = 1;
 } else if (str.indexOf("*") >= 0) {
  myError = 1;
 } else if (str.indexOf("?") >= 0) {
  myError = 1;
 } else if (str.indexOf(String.fromCharCode(34)) >= 0) {
  myError = 1;
 } else if (str.indexOf("<") >= 0) {
  myError = 1;
 } else if (str.indexOf(">") >= 0) {
  myError = 1;
 } else if (str.indexOf("|") >= 0) {
  myError = 1;
 } else if (str.indexOf(".") >= 0) {
  myError = 1;
 }
 if (myError == 1) {
  alert('Un nom de dossier ne peut contenir l\'un des caractères suivants :' + String.fromCharCode(13) + String.fromCharCode(92) + ' / : * ? " < > | .');
  document.ficForm.f_NewFolder.value = "";
  return false;
 } else {
  return true;
 }
};
function AddFolder() {
 if ("" + document.ficForm.f_NewFolder.value != "") {
  jsTrim(document.ficForm.f_NewFolder);
  if (CheckFolderName()) {
   document.ficForm.h_mode.value = "AddFolder";
   document.ficForm.action = "HE_Class.asp";
   document.ficForm.submit();
  }
 } else {
  alert('Veuillez preciser un nom de dossier');
 }
};
function AddFile() {
 if (document.uplForm.f_NewFile.value + "" != "") {
  var myError = 0;
  var str = document.uplForm.f_NewFile.value;
  str = str.substring(str.lastIndexOf(String.fromCharCode(92)) + 1, str.length);
  str = str.substring(0, str.indexOf("."));
  if (str.indexOf(String.fromCharCode(92)) >= 0) {
   myError = 1;
  } else if (str.indexOf(String.fromCharCode(47)) >= 0) {
   myError = 1;
  } else if (str.indexOf(":") >= 0) {
   myError = 1;
  } else if (str.indexOf("*") >= 0) {
   myError = 1;
  } else if (str.indexOf("?") >= 0) {
   myError = 1;
  } else if (str.indexOf(String.fromCharCode(34)) >= 0) {
   myError = 1;
  } else if (str.indexOf("<") >= 0) {
   myError = 1;
  } else if (str.indexOf(">") >= 0) {
   myError = 1;
  } else if (str.indexOf("|") >= 0) {
   myError = 1;
  } else if (str.indexOf(".") >= 0) {
   myError = 1;
  } else if (str.indexOf(" ") >= 0) {
   myError = 1;
  }
  if (myError == 1) {
   alert('Un nom de fichier ne peut contenir l\'un des caractères suivants :' + String.fromCharCode(13) + String.fromCharCode(92) + ' / : * ? " < > | . et espace');
  } else {
   var myHref = ""
   myHref += "HE_Class.asp?h_mode=AddFile&saveto=disk";
   myHref += "&f_PathTo=" + document.ficForm.f_PathTo.value;
   myHref += "&f_OrderBy=" + document.ficForm.f_OrderBy.value;
   myHref += "&f_OrderStr=" + document.ficForm.f_OrderStr.value;
   myHref += "&foo=" + document.ficForm.foo.value;
   myHref += "&h_page=" + document.ficForm.h_page.value;
   myHref += "&h_imgchange=" + document.ficForm.h_imgchange.value;
   myHref += "&h_formback=" + document.ficForm.h_formback.value;
   document.uplForm.action = myHref;
   document.uplForm.submit();
  }
 } else {
  alert('Veuillez préciser un fichier');
 }
};
function DelFolder(inFOLDER) {
 if (confirm("Attention : Vous allez effacer le dossier " + inFOLDER + " ainsi que tout son contenu (fichiers et sous-repertoires)." + String.fromCharCode(13) + String.fromCharCode(13) + "Voulez vous continuer ?")) {
  document.ficForm.h_mode.value = "DelFolder";
  document.ficForm.action = "HE_Class.asp?f_folder=" + inFOLDER;
  document.ficForm.submit();
 }
};
function DelFile(inFILE) {
 if (confirm("Attention : Vous allez effacer le fichier " + inFILE + "." + String.fromCharCode(13) + String.fromCharCode(13) + "Voulez vous continuer ?")) {
  document.ficForm.h_mode.value = "DelFile";
  document.ficForm.action = "HE_Class.asp?f_file=" + inFILE;
  document.ficForm.submit();
 }
};
function SetFile(inFILE) {
 var lHost = "<%Response.Write lePathRoot%>";
 window.opener.<%Response.Write leFormBack%>.value = inFILE.replace(lHost, "");
 self.close();
};
</script>
</head>

<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" style="table-layout:fixed;">
 <tr>
  <td width="15"></td>
  <td class="TitrePopup">Bibliothèque de fichiers</td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">
   <table border="0" width="100%" style="table-layout:fixed;">
    <tr>
     <td valign="top" class="Text">
      Veuillez choisir un fichier.
      <br>
      Cliquer le boutton "Annuler" pour fermer cette fenetre.
     </td>
     <td align="right">
      <input type="button" name="close" value="Annuler" class="Text" onClick="self.close();">
     </td>
    </tr>
   </table>
  </td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">
   <table border="0" width="98%" cellpadding="0" cellspacing="5" style="table-layout:fixed;">
    <tr>
     <td class="Text">
      <form method="post" name="ficForm" action="?">
      Nouveau Dossier :<br>
      <input type="text" name="f_NewFolder" class="Text" style="width:220;" onBlur="jsTrim(this);CheckFolderName();">&nbsp;<input type="button" name="Add" value="Ajouter" class="Text" onClick="AddFolder();">
      <input type="hidden" name="h_page" value="cand_fic">
      <input type="hidden" name="h_mode" value="">
      <input type="hidden" name="foo" value="<%Response.Write Request("foo")%>">
      <input type="hidden" name="f_PathTo" value="<%Response.Write lePathEnCours%>">
      <input type="hidden" name="f_OrderBy" value="<%Response.Write leOrderBy%>">
      <input type="hidden" name="f_OrderStr" value="<%Response.Write leOrderStr%>">
      
      <input type="hidden" name="h_formback" value="<%Response.Write leFormBack%>">
      <input type="hidden" name="h_imgchange" value="<%Response.Write leImgChange%>">
      
            <input type="hidden" name="f_OrderStr" value="<%Response.Write leOrderStr%>">
      </form>
     </td>
     <td class="Text">
      <form name="uplForm" method="post" enctype="multipart/form-data" action="?">
      Ajouter un fichier :<br>
      <input type="file" name="f_NewFile" class="Text" style="width:220;">&nbsp;<input type="button" name="f_upl" value="Ajouter" class="Text" onClick="AddFile();">
      </form>
     </td>
    </tr>
   </table>
  </td>
 </tr>
<%
If Request("f_msg").Count > 0 Then
%>
 <tr>
  <td>&nbsp;</td>
  <td class="Text" style="color:#FF0000;font-weight:bold;">
<%
Response.Write Request("f_msg")
%>
  </td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">&nbsp;</td>
 </tr>
<%
End If
%>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">
   <table border="0" width="98%" cellspacing="0" cellpadding="0" class="TblPopupTitre">
    <tr>
     <td>
      &nbsp;&nbsp;
      <a href="javascript:GoToRoot();" border="0" onMouseOver="window.status=' ';return true;" onMouseOut="window.status=' ';return true;" style="font-family:Arial,Verdana;color:#FFFFFF;text-decoration:none;font-size:12px;">
      Fichiers
      </a>&nbsp;/&nbsp;
      <%
      If lePathEnCours <> lePathRoot Then
       leStrNav = ""
       leHref = lePathRoot
       str = Replace(lePathEnCours, lePathRoot, "")
       Do While Instr(1, str, "/", vbTextCompare) > 0
        temp = Left(str, Instr(1, str, "/", vbTextCompare))
        str = Right(str, Len(str) - Len(temp))
        leHref = leHref & temp
        temp = Replace(temp, "/", "")
        leStrNav = leStrNav & "<a href=""javascript:GoToNav('" & Left(leHref, Len(leHref) - 1) & "');"" border=""0"" onMouseOver=""window.status=' ';return true;"" onMouseOut=""window.status=' ';return true;"" style=""font-family:Arial,Verdana;color:#FFFFFF;text-decoration:none;font-size:12px;"">" & vbCrLf
        leStrNav = leStrNav & temp & vbCrLf
        leStrNav = leStrNav & "</a>&nbsp;/&nbsp;" & vbCrLf
       Loop
       Response.Write leStrNav
      End If
      %>
     </td>
     <td align="right">
     Trier par&nbsp;&nbsp;
     <a href="javascript:GoToOrder(0, 'ASC');" border="0" onMouseOver="window.status=' ';return true;" onMouseOut="window.status=' ';return true;">
     <img src="../Img/Uploader/tri_<%If leOrderBy = 0 AND leOrderStr = "ASC" Then Response.Write "o" End If%>up.gif" border="0" width="10" height="9">
     </a>
     Nom
     <a href="javascript:GoToOrder(0, 'DESC');" border="0" onMouseOver="window.status=' ';return true;" onMouseOut="window.status=' ';return true;">
     <img src="../Img/Uploader/tri_<%If leOrderBy = 0 AND leOrderStr = "DESC" Then Response.Write "o" End If%>down.gif" border="0" width="10" height="9">
     </a>
     &nbsp;&nbsp;ou&nbsp;&nbsp;
     <a href="javascript:GoToOrder(2, 'ASC');" border="0" onMouseOver="window.status=' ';return true;" onMouseOut="window.status=' ';return true;">
     <img src="../Img/Uploader/tri_<%If leOrderBy = 2 AND leOrderStr = "ASC" Then Response.Write "o" End If%>up.gif" border="0" width="10" height="9">
     </a>
     Date
     <a href="javascript:GoToOrder(2, 'DESC');" border="0" onMouseOver="window.status=' ';return true;" onMouseOut="window.status=' ';return true;">
     <img src="../Img/Uploader/tri_<%If leOrderBy = 2 AND leOrderStr = "DESC" Then Response.Write "o" End If%>down.gif" border="0" width="10" height="9">
     </a>
     &nbsp;&nbsp;
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
  <td class="Text">
<table border="0" width="98%" cellspacing="0" cellpadding="5" class="TblPopupInt" style="table-layout:fixed;">
<%
Dim leStr

leStr = ""
If lePathEnCours <> lePathRoot Then
 leStr = leStr & "<tr>" & vbCrLf
 leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">"
 leStr = leStr & "&nbsp;"
 leStr = leStr & "</td>"
 leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">"
 leStr = leStr & "<a href=""javascript:GoToParent();"" border=""0"" onMouseOver=""window.status='Repertoire parent';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
 leStr = leStr & "<img src=""../Img/Uploader/up.gif"" width=""16"" height=""16"" border=""0"" align=""absmiddle"">"
 leStr = leStr & "</a>"
 leStr = leStr & "</td>"
 leStr = leStr & "<td class=""Text"">"
 leStr = leStr & "<a href=""javascript:GoToParent();"" border=""0"" onMouseOver=""window.status='Repertoire parent';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
 leStr = leStr & "Repertoire parent"
 leStr = leStr & "</a>"
 leStr = leStr & "</td>" & vbCrLf
 leStr = leStr & "</tr>" & vbCrLf
End If

If IsArray(aRep) Then
 For i = 0 To Ubound(aRep, 2)
  leStr = leStr & "<tr>" & vbCrLf
  leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">"
  leStr = leStr & "<a href=""javascript:DelFolder('" & aRep(0, i) & "');"" border=""0"" onMouseOver=""window.status='Supprimer';return true;"" onMouseOut=""window.status=' ';return true;"">"
  leStr = leStr & "<img src=""../Img/Uploader/trash.gif"" width=""14"" height=""15"" border=""0"" align=""absmiddle"">"
  leStr = leStr & "</a>"
  leStr = leStr & "</td>"
  leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">"
  leStr = leStr & "<a href=""javascript:GoToFolder('" & aRep(0, i) & "');"" border=""0"" onMouseOver=""window.status='Parcourir';return true;"" onMouseOut=""window.status=' ';return true;"">"
  leStr = leStr & "<img src=""../Img/Uploader/fc.gif"" width=""16"" height=""16"" border=""0"" align=""absmiddle"">"
  leStr = leStr & "</a>"
  leStr = leStr & "</td>"
  leStr = leStr & "<td class=""Text"">"
  leStr = leStr & "<a href=""javascript:GoToFolder('" & aRep(0, i) & "');"" border=""0"" onMouseOver=""window.status='Parcourir';return true;"" onMouseOut=""window.status=' ';return true;"">"
  leStr = leStr & aRep(0, i)
  leStr = leStr & "</a>"
  leStr = leStr & "</td>" & vbCrLf
  leStr = leStr & "</tr>" & vbCrLf
 Next
 If IsArray(aFile) Then
  For i = 0 To Ubound(aFile, 2)
   leStr = leStr & "<tr>" & vbCrLf
   leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">"
   leStr = leStr & "<a href=""javascript:DelFile('" & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Supprimer';return true;"" onMouseOut=""window.status=' ';return true;"">"
   leStr = leStr & "<img src=""../Img/Uploader/trash.gif"" width=""14"" height=""15"" border=""0"" title=""Supprimer le fichier"" align=""absmiddle"">"
   leStr = leStr & "</a>"
   leStr = leStr & "</td>"
   leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">"
   leStr = leStr & "<a href=""javascript:SetFile('" & lePathEnCours & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Insérer le lien';return true;"" onMouseOut=""window.status=' ';return true;"">"
   leStr = leStr & "<img src=""../Img/Uploader/attach.gif"" width=""12"" height=""20"" border=""0"" title=""Insérer le lien"" align=""absmiddle"">"
   leStr = leStr & "</a>"
   leStr = leStr & "</td>"
   leStr = leStr & "<td class=""Text"">"
   leStr = leStr & "<a href=""javascript:SetFile('" & lePathEnCours & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Insérer le lien';return true;"" onMouseOut=""window.status=' ';return true;"">"
   leStr = leStr & aFile(0, i)
   leStr = leStr & "</a>"
   leStr = leStr & "</td>" & vbCrLf
   leStr = leStr & "</tr>" & vbCrLf
  Next
 End If
Else
 If IsArray(aFile) Then
  For i = 0 To Ubound(aFile, 2)
   leStr = leStr & "<tr>" & vbCrLf
   leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">"
   leStr = leStr & "<a href=""javascript:DelFile('" & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Supprimer';return true;"" onMouseOut=""window.status=' ';return true;"">"
   leStr = leStr & "<img src=""../Img/Uploader/trash.gif"" width=""14"" height=""15"" border=""0"" title=""Supprimer le fichier"" align=""absmiddle"">"
   leStr = leStr & "</a>"
   leStr = leStr & "</td>"
   leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">"
   leStr = leStr & "<a href=""javascript:SetFile('" & lePathEnCours & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Insérer le lien';return true;"" onMouseOut=""window.status=' ';return true;"">"
   leStr = leStr & "<img src=""../Img/Uploader/attach.gif"" width=""12"" height=""20"" border=""0"" title=""Insérer le lien"" align=""absmiddle"">"
   leStr = leStr & "</a>"
   leStr = leStr & "</td>"
   leStr = leStr & "<td class=""Text"">"
   leStr = leStr & "<a href=""javascript:SetFile('" & lePathEnCours & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Insérer le lien';return true;"" onMouseOut=""window.status=' ';return true;"">"
   leStr = leStr & aFile(0, i)
   leStr = leStr & "</a>"
   leStr = leStr & "</td>" & vbCrLf
   leStr = leStr & "</tr>" & vbCrLf
  Next
 Else
  leStr = leStr & "<tr>" & vbCrLf
  leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">&nbsp;</td>"
  leStr = leStr & "<td class=""Text"" align=""center"" width=""25"" valign=""absmiddle"">&nbsp;</td>"
  leStr = leStr & "<td class=""Text"">"
  leStr = leStr & "Il n'y a pas de fichiers pour le moment" & vbCrLf
  leStr = leStr & "</td>" & vbCrLf
  leStr = leStr & "</tr>" & vbCrLf
 End If
End If

Response.Write leStr
%>
</table>
  </td>
 </tr>
 <tr>
  <td colspan="2" height="10"></td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td>	
  </td>
 </tr>
</table>
</body>
</html>