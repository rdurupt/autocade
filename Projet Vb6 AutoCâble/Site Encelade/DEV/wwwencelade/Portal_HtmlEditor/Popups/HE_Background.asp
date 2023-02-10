<%
lePathRoot = Session("PortalPath") & "Portal_Upload/Backgrounds/"

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
<title>Insérer / Modifier une image</title>
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
 document.imgForm.f_PathTo.value = document.imgForm.f_PathTo.value + inFolder;
 document.imgForm.action = "?";
 document.imgForm.submit();
};
function GoToParent() {
 var sFold = document.imgForm.f_PathTo.value;
 sFold = sFold.substring(0, sFold.length - 1);
 sFold = sFold.substring(0, sFold.lastIndexOf("/"));
 document.imgForm.f_PathTo.value = sFold;
 document.imgForm.action = "?";
 document.imgForm.submit();
};
function GoToRoot() {
 document.imgForm.f_PathTo.value = "<%Response.Write Left(lePathRoot, Len(lePathRoot) - 1)%>";
 document.imgForm.action = "?";
 document.imgForm.submit();
};
function GoToNav(inFolder) {
 document.imgForm.f_PathTo.value = inFolder;
 document.imgForm.action = "?";
 document.imgForm.submit();
};
function GoToOrder( inORDERBY, inORDERSTR) {
 var str = document.imgForm.f_PathTo.value;
 if (str.substring(str.length - 1, str.length) == "/") {
  str = str.substring(0, str.length - 1);
 }
 document.imgForm.f_PathTo.value = str;
 document.imgForm.f_OrderBy.value = inORDERBY;
 document.imgForm.f_OrderStr.value = inORDERSTR;
 document.imgForm.action = "?";
 document.imgForm.submit();
};
function CheckFolderName() {
 var myError = 0;
 var str = document.imgForm.f_NewFolder.value;
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
  alert('Un nom de dossier ne peut contenir l\'un des caractères suivants :' + String.fromCharCode(13) + String.fromCharCode(92) + ' / : * ? " < > | . et espace');
  document.imgForm.f_NewFolder.value = "";
  return false;
 } else {
  return true;
 }
};
function AddFolder() {
 if ("" + document.imgForm.f_NewFolder.value != "") {
  jsTrim(document.imgForm.f_NewFolder);
  if (CheckFolderName()) {
   document.imgForm.h_mode.value = "AddFolder";
   document.imgForm.action = "HE_Class.asp";
   document.imgForm.submit();
  }
 } else {
  alert('Veuillez preciser un nom de dossier');
 }
};
function AddImage() {
 if (document.uplForm.f_NewFile.value + "" != "") {
  var myError = 0;
  var imgInput = document.uplForm.f_NewFile.value.substring(document.uplForm.f_NewFile.value.lastIndexOf(String.fromCharCode(92)) + 1, document.uplForm.f_NewFile.value.length);
  if (document.images) {
   for (var i = 0; i < document.images.length; i++) {
    imgName = document.images[i].src.substring(document.images[i].src.lastIndexOf("/") + 1, document.images[i].src.length);
    if (imgInput.toUpperCase() == imgName.toUpperCase()) {
     myError = 1;
     break;
    }
   };
  }
  if (myError == 1) {
   alert('Ce nom d\'image est déjà utilisé');
  } else {
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
    myHref += "HE_Class.asp?h_mode=AddImage&saveto=disk";
    myHref += "&f_PathTo=" + document.imgForm.f_PathTo.value;
    myHref += "&f_OrderBy=" + document.imgForm.f_OrderBy.value;
    myHref += "&f_OrderStr=" + document.imgForm.f_OrderStr.value;
    myHref += "&h_page=" + document.imgForm.h_page.value;
    document.uplForm.action = myHref;
    document.uplForm.submit();
   }
  }
 } else {
  alert('Veuillez préciser une image');
 }
};
function AfficheMaxi(inPATH) {
 var i1 = new Image;
 i1.src = inPATH;
 html = '<HTML><HEAD><TITLE>Image</TITLE></HEAD><BODY LEFTMARGIN=0 MARGINWIDTH=0 TOPMARGIN=0 MARGINHEIGHT=0><CENTER><IMG SRC="'+inPATH+'" BORDER=0 NAME=imageTest onLoad="window.resizeTo(document.imageTest.width+14,document.imageTest.height+32)"></CENTER></BODY></HTML>';
 var popupImage = window.open('','_blank','toolbar=0,location=0,directories=0,menuBar=0,scrollbars=0,resizable=1,top=100,left=100');
 popupImage.document.open();
 popupImage.document.write(html);
 popupImage.document.close()
};
function DelFolder(inFOLDER) {
 if (confirm("Attention : Vous allez effacer le dossier " + inFOLDER + " ainsi que tout son contenu (fichiers et sous-repertoires)." + String.fromCharCode(13) + String.fromCharCode(13) + "Voulez vous continuer ?")) {
  document.imgForm.h_mode.value = "DelFolder";
  document.imgForm.action = "HE_Class.asp?f_folder=" + inFOLDER;
  document.imgForm.submit();
 }
};
function DelImage(inIMG) {
 if (confirm("Attention : Vous allez effacer l\'image " + inIMG + "." + String.fromCharCode(13) + String.fromCharCode(13) + "Voulez vous continuer ?")) {
  document.imgForm.h_mode.value = "DelImage";
  document.imgForm.action = "HE_Class.asp?f_image=" + inIMG;
  document.imgForm.submit();
 }
};
function SetBgImage(inIMG) {
 window.opener.document.formContent.f_CatBackGround.value = inIMG;
 self.close();
};
</script>
</head>

<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" style="table-layout:fixed;">
 <tr>
  <td width="15"></td>
  <td class="TitrePopup">Bibliothèque d'images</td>
 </tr>
 <tr>
  <td>&nbsp;</td>
  <td class="Text">
   <table border="0" width="100%" style="table-layout:fixed;">
    <tr>
     <td valign="top" class="Text">
      Veuillez choisir une image.
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
      <form method="post" name="imgForm" action="?">
      Nouveau Dossier :<br>
      <input type="text" name="f_NewFolder" class="Text" style="width:220;" onBlur="jsTrim(this);CheckFolderName();">&nbsp;<input type="button" name="Add" value="Ajouter" class="Text" onClick="AddFolder();">
      <input type="hidden" name="h_page" value="bgd">
      <input type="hidden" name="h_mode" value="">
      <input type="hidden" name="f_PathTo" value="<%Response.Write lePathEnCours%>">
      <input type="hidden" name="f_OrderBy" value="<%Response.Write leOrderBy%>">
      <input type="hidden" name="f_OrderStr" value="<%Response.Write leOrderStr%>">
      </form>
     </td>
     <td class="Text">
      <form name="uplForm" method="post" enctype="multipart/form-data" action="?">
      Ajouter une image :<br>
      <input type="file" name="f_NewFile" class="Text" style="width:220;">&nbsp;<input type="button" name="f_upl" value="Ajouter" class="Text" onClick="AddImage();">
      </form>
     </td>
    </tr>
   </table>
  </td>
 </tr>
<!--
 <tr>
  <td>&nbsp;</td>
  <td class="Text">&nbsp;</td>
 </tr>
-->
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
      Backgrounds
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
Dim nbCol
Dim tdDisp
Dim leStr

nbCol = 4
tdDisp = 0
leStr = ""
If lePathEnCours <> lePathRoot Then
 leStr = leStr & " <tr>" & vbCrLf
 leStr = leStr & "  <td class=""Text"" align=""center"" width=""25%"">" & vbCrLf
 leStr = leStr & "   <a href=""javascript:GoToParent();"" border=""0"" onMouseOver=""window.status='Repertoire parent';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
 leStr = leStr & "   <img src=""../Img/Uploader/folder_up.gif"" width=""70"" height=""70"" border=""0"">" & vbCrLf
 leStr = leStr & "   </a>" & vbCrLf
 leStr = leStr & "   <br>" & vbCrLf
 leStr = leStr & "   Repertoire parent" & vbCrLf
 leStr = leStr & "  </td>" & vbCrLf
 tdDisp = 1
Else
 tdDisp = 0
End If

If IsArray(aRep) Then
 For i = 0 To Ubound(aRep, 2)
  If tdDisp Mod nbCol = 0 Then
   If tdDisp <> 0 Then
    leStr = leStr & " </tr>" & vbCrLf
    tdDisp = 0
   End If
   leStr = leStr & " <tr>" & vbCrLf
  End If
  leStr = leStr & "  <td class=""Text"" align=""center"" width=""25%"">" & vbCrLf
  leStr = leStr & "   <table border=""0"" cellpadding=""0"" cellspacing=""2"">" & vbCrLf
  leStr = leStr & "    <tr>" & vbCrLf
  leStr = leStr & "     <td class=""Text"" align=""center"">" & vbCrLf
  leStr = leStr & "      <a href=""javascript:GoToFolder('" & aRep(0, i) & "');"" border=""0"" onMouseOver=""window.status='Parcourir';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
  leStr = leStr & "      <img src=""../Img/Uploader/folder.gif"" width=""70"" height=""70"" border=""0"">" & vbCrLf
  leStr = leStr & "      </a>" & vbCrLf
  leStr = leStr & "     </td>" & vbCrLf
  leStr = leStr & "     <td class=""Text"" align=""center"">" & vbCrLf
  leStr = leStr & "      <a href=""javascript:DelFolder('" & aRep(0, i) & "');"" border=""0"" onMouseOver=""window.status='Supprimer';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
  leStr = leStr & "      <img src=""../Img/Uploader/trash.gif"" width=""14"" height=""15"" border=""0"">"
  leStr = leStr & "      </a>" & vbCrLf
  leStr = leStr & "     </td>" & vbCrLf
  leStr = leStr & "    </tr>" & vbCrLf
  leStr = leStr & "    <tr>" & vbCrLf
  leStr = leStr & "     <td class=""Text"" align=""center"" colspan=""2"">" & vbCrLf
  leStr = leStr & "      " & aRep(0, i) & vbCrLf
  leStr = leStr & "     </td>" & vbCrLf
  leStr = leStr & "    </tr>" & vbCrLf
  leStr = leStr & "   </table>" & vbCrLf
  leStr = leStr & "  </td>" & vbCrLf
  tdDisp = tdDisp + 1
 Next
 If IsArray(aFile) Then
  indx = 0
  For i = tdDisp To nbCol - 1
   If indx =< UBound(aFile, 2) Then
    leStr = leStr & "  <td class=""Text"" align=""center"" width=""25%"">" & vbCrLf
    leStr = leStr & "   <table border=""0"" cellpadding=""0"" cellspacing=""2"">" & vbCrLf
    leStr = leStr & "    <tr>" & vbCrLf
    leStr = leStr & "     <td class=""Text"" align=""center"">" & vbCrLf
    leStr = leStr & "      <img src=""" & lePathEnCours & aFile(0, indx) & """ width=""70"" height=""70"" border=""0"" style=""cursor:hand;"" onClick=""AfficheMaxi('" & lePathEnCours & aFile(0, indx) & "');"">" & vbCrLf
    leStr = leStr & "     </td>" & vbCrLf
    leStr = leStr & "     <td width=""20"" align=""center"" valign=""top"">" & vbCrLf
    leStr = leStr & "      <a href=""javascript:SetBgImage('" & lePathEnCours & aFile(0, indx) & "');"" border=""0"" onMouseOver=""window.status='Insérer l\'image comme background';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
    leStr = leStr & "      <img src=""../Img/Uploader/img_back.gif"" width=""20"" height=""21"" border=""0"" title=""Insérer l'image comme background"">" & vbCrLf
    leStr = leStr & "      </a>" & vbCrLf
    leStr = leStr & "      <a href=""javascript:DelImage('" & aFile(0, indx) & "');"" border=""0"" onMouseOver=""window.status='Supprimer';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
    leStr = leStr & "      <img src=""../Img/Uploader/trash.gif"" width=""14"" height=""15"" border=""0"" title=""Supprimer l'image"">" & vbCrLf
    leStr = leStr & "      </a>" & vbCrLf
    leStr = leStr & "     </td>" & vbCrLf
    leStr = leStr & "    </tr>" & vbCrLf
    leStr = leStr & "    <tr>" & vbCrLf
    leStr = leStr & "     <td class=""Text"" align=""center"" colspan=""2"">" & vbCrLf
    leStr = leStr & "      " & aFile(0, indx) & vbCrLf
    leStr = leStr & "     </td>" & vbCrLf
    leStr = leStr & "    </tr>" & vbCrLf
    leStr = leStr & "   </table>" & vbCrLf
    leStr = leStr & "  </td>" & vbCrLf
   Else
    leStr = leStr & "  <td class=""Text"" width=""25%"">" & vbCrLf
    leStr = leStr & "   &nbsp;" & vbCrLf
    leStr = leStr & "  </td>" & vbCrLf
   End If
   indx = indx + 1
  Next
  leStr = leStr & " </tr>" & vbCrLf
  tdDisp = 0
  For i = indx To Ubound(aFile, 2)
   If tdDisp Mod nbCol = 0 Then
    If tdDisp <> 0 Then
     leStr = leStr & " </tr>" & vbCrLf
     tdDisp = 0
    End If
    leStr = leStr & " <tr>" & vbCrLf
   End If
   leStr = leStr & "  <td class=""Text"" align=""center"" width=""25%"">" & vbCrLf
   leStr = leStr & "   <table border=""0"" cellpadding=""0"" cellspacing=""2"">" & vbCrLf
   leStr = leStr & "    <tr>" & vbCrLf
   leStr = leStr & "     <td class=""Text"" align=""center"">" & vbCrLf
   leStr = leStr & "      <img src=""" & lePathEnCours & aFile(0, i) & """ width=""70"" height=""70"" border=""0"" style=""cursor:hand;"" onClick=""AfficheMaxi('" & lePathEnCours & aFile(0, i) & "');"">" & vbCrLf
   leStr = leStr & "     </td>" & vbCrLf
   leStr = leStr & "     <td width=""20"" align=""center"" valign=""top"">" & vbCrLf
   leStr = leStr & "      <a href=""javascript:SetBgImage('" & lePathEnCours & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Insérer l\'image comme background';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
   leStr = leStr & "      <img src=""../Img/Uploader/img_back.gif"" width=""20"" height=""21"" border=""0"" title=""Insérer l'image comme background"">" & vbCrLf
   leStr = leStr & "      </a>" & vbCrLf
   leStr = leStr & "      <a href=""javascript:DelImage('" & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Supprimer';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
   leStr = leStr & "      <img src=""../Img/Uploader/trash.gif"" width=""14"" height=""15"" border=""0"" title=""Supprimer l'image"">" & vbCrLf
   leStr = leStr & "      </a>" & vbCrLf
   leStr = leStr & "     </td>" & vbCrLf
   leStr = leStr & "    </tr>" & vbCrLf
   leStr = leStr & "    <tr>" & vbCrLf
   leStr = leStr & "     <td class=""Text"" align=""center"" colspan=""2"">" & vbCrLf
   leStr = leStr & "      " & aFile(0, i) & vbCrLf
   leStr = leStr & "     </td>" & vbCrLf
   leStr = leStr & "    </tr>" & vbCrLf
   leStr = leStr & "   </table>" & vbCrLf
   leStr = leStr & "  </td>" & vbCrLf
   tdDisp = tdDisp + 1
  Next
  For i = tdDisp To nbCol - 1
   leStr = leStr & "  <td class=""Text"" width=""25%"">" & vbCrLf
   leStr = leStr & "   &nbsp;" & vbCrLf
   leStr = leStr & "  </td>" & vbCrLf
   indx = indx + 1
  Next
  leStr = leStr & " </tr>" & vbCrLf
 Else
  For i = tdDisp To nbCol - 1
   leStr = leStr & "  <td class=""Text"" align=""center"" width=""25%"">" & vbCrLf
   leStr = leStr & "   &nbsp;" & vbCrLf
   leStr = leStr & "  </td>" & vbCrLf
   indx = indx + 1
  Next
  leStr = leStr & " </tr>" & vbCrLf
 End If
Else
 If IsArray(aFile) Then
  For i = 0 To Ubound(aFile, 2)
   If tdDisp Mod nbCol = 0 Then
    If tdDisp <> 0 Then
     leStr = leStr & " </tr>" & vbCrLf
     tdDisp = 0
    End If
    leStr = leStr & " <tr>" & vbCrLf
   End If
   leStr = leStr & "  <td class=""Text"" align=""center"" width=""25%"">" & vbCrLf
   leStr = leStr & "   <table border=""0"" cellpadding=""0"" cellspacing=""2"">" & vbCrLf
   leStr = leStr & "    <tr>" & vbCrLf
   leStr = leStr & "     <td class=""Text"" align=""center"">" & vbCrLf
   leStr = leStr & "      <img src=""" & lePathEnCours & aFile(0, i) & """ width=""70"" height=""70"" border=""0"" style=""cursor:hand;"" onClick=""AfficheMaxi('" & lePathEnCours & aFile(0, i) & "');"">" & vbCrLf
   leStr = leStr & "     </td>" & vbCrLf
   leStr = leStr & "     <td width=""20"" align=""center"" valign=""top"">" & vbCrLf
   leStr = leStr & "      <a href=""javascript:SetBgImage('" & lePathEnCours & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Insérer l\'image comme background';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
   leStr = leStr & "      <img src=""../Img/Uploader/img_back.gif"" width=""20"" height=""21"" border=""0"" title=""Insérer l'image comme background"">" & vbCrLf
   leStr = leStr & "      </a>" & vbCrLf
   leStr = leStr & "      <a href=""javascript:DelImage('" & aFile(0, i) & "');"" border=""0"" onMouseOver=""window.status='Supprimer';return true;"" onMouseOut=""window.status=' ';return true;"">" & vbCrLf
   leStr = leStr & "      <img src=""../Img/Uploader/trash.gif"" width=""14"" height=""15"" border=""0"" title=""Supprimer l'image"">" & vbCrLf
   leStr = leStr & "      </a>" & vbCrLf
   leStr = leStr & "     </td>" & vbCrLf
   leStr = leStr & "    </tr>" & vbCrLf
   leStr = leStr & "    <tr>" & vbCrLf
   leStr = leStr & "     <td class=""Text"" align=""center"" colspan=""2"">" & vbCrLf
   leStr = leStr & "      " & aFile(0, i) & vbCrLf
   leStr = leStr & "     </td>" & vbCrLf
   leStr = leStr & "    </tr>" & vbCrLf
   leStr = leStr & "   </table>" & vbCrLf
   leStr = leStr & "  </td>" & vbCrLf
   tdDisp = tdDisp + 1
  Next
  For i = tdDisp To nbCol - 1
   leStr = leStr & "  <td class=""Text"" align=""center"" width=""25%"">" & vbCrLf
   leStr = leStr & "   &nbsp;" & vbCrLf
   leStr = leStr & "  </td>" & vbCrLf
   indx = indx + 1
  Next
  leStr = leStr & " </tr>" & vbCrLf
 Else
  leStr = leStr & " </tr>" & vbCrLf
  leStr = leStr & " <tr>" & vbCrLf
  leStr = leStr & "  <td class=""Text"" align=""center"">" & vbCrLf
  leStr = leStr & "   Il n'y a pas d'images pour le moment" & vbCrLf
  leStr = leStr & "  </td>" & vbCrLf
  leStr = leStr & " </tr>" & vbCrLf
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