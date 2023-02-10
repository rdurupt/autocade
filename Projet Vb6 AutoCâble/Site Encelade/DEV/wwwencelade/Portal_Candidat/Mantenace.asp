<html>
<head>
<style type="text/css">
<!--
.Style1 {
	color: #A4C1AC;
	font-weight: bold;
	font-size: 36px;
}
.Style2 {color: #00FF66}
.Style3 {color: #99FFCC; }
.Style4 {color: #66FF99; }
.Style6 {color: #99CCAB; }
.Style7 {color: #A4C1AC}
-->
</style>
<link rel="stylesheet" href="PMainStyle1.asp">
</head>

<link rel="stylesheet" href="PMainStyle1.asp">
</head>

<body background="background.gif"><br><br><br><br>
<script language="JavaScript">
var NoOffFirstLineMenus = 3;
var BaseHref = "";
Menu1 = new Array("Type de pièce", "", "", 8, 18, 124, "", "", "", "", "", "", -1, -1, -1, "", "Type de pièce");
Menu1_1 = new Array("A-CONNECTEURS", "switchbase.asp?CatId=6", "", 0, 18, 196, "", "", "", "", "", "", -1, -1, -1, "", "A-CONNECTEURS");
Menu1_2 = new Array("B-CONNECTIQUE", "switchbase.asp?CatId=13", "", 0, 18, 124, "", "", "", "", "", "", -1, -1, -1, "", "B-CONNECTIQUE");
Menu1_3 = new Array("C-CAPOTS_&_VERROUX", "switchbase.asp?CatId=15", "", 0, 18, 164, "", "", "", "", "", "", -1, -1, -1, "", "C-CAPOTS_&_VERROUX");
Menu1_4 = new Array("D-JOINTS_&_BOUCHONS", "switchbase.asp?CatId=55", "", 0, 18, 172, "", "", "", "", "", "", -1, -1, -1, "", "D-JOINTS_&_BOUCHONS");
Menu1_5 = new Array("E-SUPPORTS_&_FIXATIONS", "switchbase.asp?CatId=58", "", 0, 18, 196, "", "", "", "", "", "", -1, -1, -1, "", "E-SUPPORTS_&_FIXATIONS");
Menu1_6 = new Array("F-FILS_ET_COMPOSANTS", "switchbase.asp?CatId=65", "", 0, 18, 196, "", "", "", "", "", "", -1, -1, -1, "", "F-FILS_ET_COMPOSANTS");
Menu1_7 = new Array("G-HABILLAGES", "switchbase.asp?CatId=66", "", 0, 18, 196, "", "", "", "", "", "", -1, -1, -1, "", "G-HABILLAGES");
Menu1_8 = new Array("H-BAGUES", "switchbase.asp?CatId=69", "", 0, 18, 196, "", "", "", "", "", "", -1, -1, -1, "", "H-BAGUES");
Menu2 = new Array("Outils", "", "", 2, 18, 68, "", "", "", "", "", "", -1, -1, -1, "", "Outils");
Menu2_1 = new Array("Recherche", "Contact.asp?mode=search", "", 0, 18, 156, "", "", "", "", "", "", -1, -1, -1, "", "Recherche");
Menu2_2 = new Array("Historique Client", "javascript:parent.frames['main'].location='Contact.asp?mode=EspaceClient&NumFrm=1'", "", 0, 18, 156, "", "", "", "", "", "", -1, -1, -1, "", "Historique Client");
Menu3 = new Array("Administration", "", "", 5, 18, 132, "", "", "", "", "", "", -1, -1, -1, "", "Administration");
Menu3_1 = new Array("Listes", "javascript:parent.frames['main'].location='con_lstCategory.asp'", "", 0, 18, 132, "", "", "", "", "", "", -1, -1, -1, "", "Listes");
Menu3_2 = new Array("Paramétrages", "javascript:parent.frames['main'].location='Contact.asp?mode=con_frmSetting'", "", 0, 18, 116, "", "", "", "", "", "", -1, -1, -1, "", "Paramétrages");
Menu3_3 = new Array("Configuration", "javascript:parent.frames['main'].location='Contact.asp?mode=Config'", "", 0, 18, 124, "", "", "", "", "", "", -1, -1, -1, "", "Configuration");
Menu3_4 = new Array("Gestion Client", "javascript:parent.frames['main'].location='Contact.asp?mode=ConfigCLI'", "", 0, 18, 132, "", "", "", "", "", "", -1, -1, -1, "", "Gestion Client");
Menu3_5 = new Array("Historiques", "javascript:parent.frames['main'].location='Contact.asp?mode=Historiques'", "", 0, 18, 132, "", "", "", "", "", "", -1, -1, -1, "", "Historiques");
</script>
<script language="JavaScript">function Go(){return}</script>
<script language="JavaScript" src="Portal_Menu_Format.js"></script>
<script language="JavaScript" src="Portal_Menu.js"></script>
<script language='javascript'>
function javCadyRempli() {
myForm = this.document.forms[0];
myForm.mode.value = 'Candidat_caddy';
myForm.action='Contact.asp?mode=Candidat_caddy';
myForm.submit();
}
function javSendEmail() {
   document.frm.submit()
}
function javSelectAll(chk) {
   len = document.frm.elements.length;
   var i=0;
   for( i=0; i<len; i++) {
       if (document.frm.elements[i].name=='lstTO') {
           document.frm.elements[i].checked=chk;
       }
   }
}
function javNewContact() {
   location.href = 'contact.asp?mode=con_frmCandidat&ContactID=0'
}
function javResetSearch() {
   location.href = 'Contact.asp?mode=con_lst&reset=true'
}
function javCat() {
   var ID = document.frmSearch.CatID.options[document.frmSearch.CatID.selectedIndex].value
   location.href = 'Contact.asp?mode=con_lst&CatID=' + ID
}
function javAddToCat() {
   window.open('con_dlgAddToCat.asp?CatID=&sid=10&ListType=','dlgcat','resizable=yes,status=no,top=200,left=200,width=300,height=400')
}
function javRemoveFromCat() {
   window.open('con_dlgRemoveFromCat.asp?CatID=&sid=10&ListType=','dlgcat','resizable=yes,status=no,top=200,left=200,width=300,height=400')
}
function javDelete(ContactID) {
    if (confirm("Supprimer cette fiche?")) {
        location.href = "Contact.asp?mode=con_subContact&sub=delete&ContactID=+ContactID"
    } else {
        return
    }
}
</script>
<script language="VBscript">
function vbMouseOver(a)
   a.style.backgroundcolor = "#AFB4DC"
end function
function vbMouseOut(a)
   a.style.backgroundcolor = "#DBDEF1"
end function
function vbEditContact(ID)
   location.href = "contact.asp?mode=con_frmCandidat&ContactID=" & ID
end function
</script>
<table width="100%" border="2" bgcolor="#8486B6" cellpadding="1" cellspacing="0">
<tr valign="bottom">
<td>
<table width="100%" border="0">
<tr valign="bottom">
<form name="frmSearch" action="Contact.asp" method="get">
<input type="hidden" name="mode" value="con_lst">
<td align="left" valign="middle">
 Type de pièce: A-CONNECTEURS
<font class="header"></font>
</td>
<td align="right" class="smallerheader" valign="middle">
<input type="text"   name="contactSearch" size="17" value="">
<input type="submit"  class="cmdflat"   value="Chercher">
<input type="button"  class="cmdflat"   value="Nouveau" onClick="javNewContact()">
</td>
</form>
</tr>
</table>
</td>
</tr>
</table>
<table bgcolor="#989AC9" width="100%">
<tr>
<td nowrap>
</td></tr>
</table>
<table width="52%" border="0" align="center" cellpadding="0">
  <tr>
    <td><p align="center" class="Style1">Site Arr&ecirc;t&eacute; pour maintenance </p></td>
  </tr>
  <tr>
    <td><p class="Style6">&nbsp;  </p>
      <p class="Style6">&nbsp; &nbsp; &nbsp; &nbsp; <span class="Style7">&nbsp;ENCELADE&nbsp;vous    informe que son site eboutique a &eacute;t&eacute; interrompu pour des raisons de maintenance.</span></p>
      <p class="Style7"><br>
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Nous    faisons de notre mieux rendre le indisponible de nouveau dans les plus bref    d&eacute;lais. </p>
      <p class="Style4"><span class="Style7"><br>
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;Nous vous prions de bien vouloir nous    excuser pour le g&egrave;ne occasionn&eacute;.<br>
  &nbsp; </span></p></td>
  </tr>
  <tr>
    <td><p><span class="Style2"><span class="Style3"><span class="Style4"></span></span></span></p></td>
  </tr>
</table>
<p class="Style4">
<script language="JavaScript">





</body>
</html>
</p>
