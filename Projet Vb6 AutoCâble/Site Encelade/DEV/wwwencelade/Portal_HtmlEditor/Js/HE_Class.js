var HEArray = new Array();

function Html_Editor( inWIN, inDOC, inMODE ) {
 this.uID = document.uniqueID;
 this.heMode = inMODE;
 HEArray[this.uID] = this;
 this.bb = new Html_BarreBouttons(this.uID, this.heMode);
 this.cm = new Html_ContextMenu(this.uID);

 this.doc = inDOC;
 this.win = inWIN;
 this.gHTML = "";
 this.back = "";
 this.groupid = 0;
 this.catid = 0;
 this.html = "";
 this.opener = null;
 this.oPopup = this.win.createPopup();
 this.oPopWin = null;
 this.tableDefault = 1;
 this.borderShown = "no";
 this.mode = 1;
 this.editModeOn = true;
 this.editModeView = true;
 this.sourceModeOn = false;
 this.sourceModeView = false;
 this.previewModeView = false;
 this.isEditingHTMLPage = 1;
 this.colorType = 0;
 this.toggleWasOn = null;
 this.fontFamily = null;
 this.fontSize = null;
 this.text = null;
 this.bgColor = null;
 this.background = null;
 this.fontChange = null;
 this.linkWin = null;
 this.anchorWin = null;
 this.emailWin = null;
 this.imageWin = null;
 this.insertTableWin = null;
 this.editTableWin = null;
 this.editCellWin = null;
 this.editPageWin = null;
 this.uplWin = null;
 this.selectedTD = null;
 this.selectedTR = null;
 this.selectedTBODY = null;
 this.selectedTable = null;
 this.selectedImage = null;
 this.sDeb = "";
 this.sFin = "";

 this.GetHtmlEditor = function()
 {
  var str = "";
  str += "<table border=\"1\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" height=\"90%\">";
  str += "<tr><td height=\"1\">";
  str += this.bb.Init();
  str += "</td></tr><tr height=\"100%\"><td>";
  str += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" height=\"100%\" class=\"TblIFrame\">";
  str += "<tr height=\"100%\"><td>";
  str += "<iframe id=\"des_" + this.uID + "\" style=\"width:100%;height:100%;\" contenteditable=\"true\" security=\"restricted\" onBlur=\"UpdateValue();\"></iframe>";
  str += "<iframe id=\"pre_" + this.uID + "\" style=\"width:100%;height:100%;display:none;\" onBlur=\"UpdateValue();\"></iframe>";
  str += "<input type=\"hidden\" name=\"he_html_" + this.uID + "\" value=\"\">";
  str += "</td></tr></table>";
  str += "</td></tr><tr><td>";
  str += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" class=\"TblStatus\">";
  str += "<tr>";
  str += "<td background=\"" + HE_IMGPATH + "Editor/he_border.gif\" height=\"22\">";
  str += "<img src=\"" + HE_IMGPATH + "Editor/he_des_up.gif\" id=\"ImgEdit\" border=\"0\" width=\"98\" height=\"22\" style=\"cursor:hand;\" onClick=\"ModeDesign();\" onselectstart=\"return false;\" onselect=\"return false;\" ondragstart=\"return false;\"><img src=\"" + HE_IMGPATH + "Editor/he_src.gif\" id=\"ImgSource\" border=\"0\" width=\"98\" height=\"22\" style=\"cursor:hand;\" onClick=\"ModeSource();\" onselectstart=\"return false;\" onselect=\"return false;\" ondragstart=\"return false;\"><img src=\"" + HE_IMGPATH + "Editor/he_pre.gif\" id=\"ImgPreview\" border=\"0\" width=\"98\" height=\"22\" style=\"cursor:hand;\" onClick=\"ModePreview();\" onselectstart=\"return false;\" onselect=\"return false;\" ondragstart=\"return false;\">";
  str += "</td><td background=\"" + HE_IMGPATH + "Editor/he_border.gif\" height=\"22\" align=\"right\" id=\"sbt_" + this.uID + "\">&nbsp;</td></tr></table>";
  str += "</td></tr></table>";
  return str;
 };
 this.SetBack = function( inBACK )
 {
  this.back = inBACK;
 };
 this.GetBack = function()
 {
  return this.back;
 };
 this.SetGroupId = function( inGROUP )
 {
  this.groupid = inGROUP;
 };
 this.SetCatId = function( inCAT )
 {
  this.catid = inCAT;
 };
 this.SetOpener = function( inOPENER )
 {
  this.opener = inOPENER;
 };
 this.GetOpener = function()
 {
  return this.opener;
 };
 this.CloseAllWin = function()
 {
  if (this.linkWin) {
   this.linkWin.close();
   this.linkWin = null;
  }
  if (this.anchorWin) {
   this.anchorWin.close();
   this.anchorWin = null;
  }
  if (this.emailWin) {
   this.emailWin.close();
   this.emailWin = null;
  }
  if (this.imageWin) {
   this.imageWin.close();
   this.imageWin = null;
  }
  if (this.insertTableWin) {
   this.insertTableWin.close();
   this.insertTableWin = null;
  }
  if (this.editTableWin) {
   this.editTableWin.close();
   this.editTableWin = null;
  }
  if (this.editCellWin) {
   this.editCellWin.close();
   this.editCellWin = null;
  }
  if (this.editPageWin) {
   this.editPageWin.close();
   this.editPageWin = null;
  }
  if (this.uplWin) {
   this.uplWin.close();
   this.uplWin = null;
  }
 };
 this.DrawHE = function( inDIV )
 {
  var oDiv = this.doc.getElementById(inDIV);
  oDiv.innerHTML = this.gHTML;
 };
 this.ReInit = function( inWIN, inDOC )
 {
  this.doc = inDOC;
  this.win = inWIN;
  this.CloseAllWin();
  this.back = "";
  this.groupid = 0;
  this.catid = 0;
  this.html = "";
  this.opener = null;
  this.oPopup = null;
  this.oPopup = this.win.createPopup();
  this.oPopWin = null;
  this.tableDefault = 1;
  this.borderShown = "no";
  this.mode = 1;
  this.editModeOn = true;
  this.editModeView = true;
  this.sourceModeOn = false;
  this.sourceModeView = false;
  this.previewModeView = false;
  this.isEditingHTMLPage = 1;
  this.colorType = 0;
  this.toggleWasOn = null;
  this.fontFamily = null;
  this.fontSize = null;
  this.text = null;
  this.bgColor = null;
  this.background = null;
  this.fontChange = null;
  this.linkWin = null;
  this.anchorWin = null;
  this.emailWin = null;
  this.imageWin = null;
  this.insertTableWin = null;
  this.editTableWin = null;
  this.editCellWin = null;
  this.editPageWin = null;
  this.uplWin = null;
  this.selectedTD = null;
  this.selectedTR = null;
  this.selectedTBODY = null;
  this.selectedTable = null;
  this.selectedImage = null;
  this.sDeb = "";
  this.sFin = "";
  this.doc.recalc(true);
 };
 this.Init = function()
 {
  var str = "";
  str += this.GetHtmlEditor();
  str += this.cm.gStr;
  this.gHTML = str;
//  oDiv.innerHTML = str;
//  this.DoLoad();
 };
 this.Init();
};