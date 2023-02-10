// Instancier PuID en page
var PuID = null;

var HE_CHARS = 
[
 "&copy;",
 "&reg;",
 "&#153;",
 "&pound;",
 "&#151;",
 "&divide;",
 "&yen;",
 "&euro;",
 "&#147;",
 "&#148;",
 "&#149;",
 "&para;"
];

function GUI( inDELAY ) {
 if (JS_HE.editModeView) {
  UpdateGUI(inDELAY);
 }
};

GetModeles = function()
{
 var oSlt = document.all["Modeles_" + PuID], i = 0, oLength = oSlt.length, uneOption;
 for (i = 0; i < oLength; i++) {
  oSlt.options[0] = null;
 };
 for (i = 0; i < HE_MODELES.length; i++) {
  uneOption = new Option(HE_MODELES[i][0], HE_MODELES[i][1], false, false);
  oSlt.options[oSlt.length] = uneOption;
 };
 oSlt.options[0].selected = true;
};

IsAllowed = function()
{
 var foo = eval("des_" + PuID);
 foo.focus();
 return true;
};

IsSelection = function()
{
 var foo = eval("des_" + PuID);
 if ((foo.document.selection.type == "Text") || (foo.document.selection.type == "Control")) {
  return true;
 } else {
  return false;
 }
};

IsTextSelected = function()
{
 var foo = eval("des_" + PuID);
 if (foo.document.selection.type == "Text") {
  return true;
 } else {
  return false;
 }
};

IsImageSelected = function()
{
 var foo = eval("des_" + PuID);
 if (foo.document.selection.type == "Control") {
  var oControlRange = foo.document.selection.createRange();
  if (oControlRange(0).tagName.toUpperCase() == "IMG") {
   JS_HE.selectedImage = foo.document.selection.createRange()(0);
   return true;
  }
 }
};

IsCursorInTableCell = function()
{
 if (IsAllowed()) {
  var foo = eval("des_" + PuID);
  if (foo.document.selection.type != "Control") {
   var elem = foo.document.selection.createRange().parentElement();
   while (elem && elem.tagName.toUpperCase() != "TD") {
    elem = elem.parentElement;
    if (elem == null) {break;}
   };
   if (elem) {
    JS_HE.selectedTD = elem;
    JS_HE.selectedTR = JS_HE.selectedTD.parentElement;
    JS_HE.selectedTBODY = JS_HE.selectedTR.parentElement;
    JS_HE.selectedTable = JS_HE.selectedTBODY.parentElement;
    return true;
   }
  }
 }
};

IsTableSelected = function()
{
 if (IsAllowed()) {
  var foo = eval("des_" + PuID);
  if (foo.document.selection.type == "Control") {
   var oControlRange = foo.document.selection.createRange();
   if (oControlRange(0).tagName.toUpperCase() == "TABLE") {
    JS_HE.selectedTable = foo.document.selection.createRange()(0);
    return true;
   }
  }
 }
};

ModeDesign = function()
{
 if (JS_HE.editModeView == false) {
  JS_HE.isEditingHTMLPage = 1;
  var foo = document.getElementById("des_" + PuID);
  var pre = document.getElementById("pre_" + PuID);
  document.getElementById("toolbar-edit_" + PuID).className = "OnBouttons";
  document.getElementById("toolbar-source_" + PuID).className = "Off";
  document.getElementById("toolbar-preview_" + PuID).className = "Off";
  foo.style.display = "";
  pre.style.display = "none";
  if (JS_HE.sourceModeOn) {
   SwitchMode();
  }
  JS_HE.editModeView = true;
  JS_HE.sourceModeView = false;
  JS_HE.previewModeView = false;
  JS_HE.sourceModeOn = false;
  JS_HE.editModeOn = true;
  document.getElementById("ImgEdit").src = HE_IMGPATH + "Editor/he_des_up.gif";
  document.getElementById("ImgSource").src = HE_IMGPATH + "Editor/he_src.gif";
  document.getElementById("ImgPreview").src = HE_IMGPATH + "Editor/he_pre.gif";
  ShowStatus();
  InitFoo();
  foo.focus();
  UpdateGUI(null);
 }
};

SwitchMode = function()
{
 var str = "", maRegExp;
 var foo = eval("des_" + PuID);
 if (JS_HE.mode == 1) {
  if (JS_HE.borderShown == "yes") {
   ToggleBorders();
   JS_HE.toggleWasOn = "yes";
  } else {
   JS_HE.toggleWasOn = "no";
  }
  document.getElementById("toolbar-edit_" + PuID).className = "Off";
  document.getElementById("toolbar-source_" + PuID).className = "OnBouttons";
  JS_HE.fontFamily = foo.document.body.style.fontFamily;
  JS_HE.fontSize = foo.document.body.style.fontSize;
  JS_HE.text = foo.document.body.text;
  JS_HE.bgColor = foo.document.body.bgColor;
  JS_HE.background = foo.document.body.background;
  if (JS_HE.isEditingHTMLPage == 0) {
   str = foo.document.body.innerHTML;
  } else {
   str = foo.document.documentElement.outerHTML;
  }
  maRegExp = /&amp;/g;
  str = str.replace(maRegExp,'&');
  if (JS_HE.heMode == "body") {
   JS_HE.sDeb = str.substring(0, str.toUpperCase().indexOf(">", str.toUpperCase().indexOf("<BODY")) + 1);
   JS_HE.sFin = str.substring(str.indexOf("</BODY>"), str.length);
   str = str.replace(JS_HE.sDeb, "");
   str = str.replace(JS_HE.sFin, "");
  }
  foo.document.body.innerText = str;
  foo.document.body.style.fontFamily = "Courier";
  foo.document.body.style.fontSize = "10pt";
  foo.document.body.bgColor = '#FFFFFF';
  foo.document.body.text = '#000000';
  foo.document.body.background = '';
  JS_HE.fontChange = true;
  JS_HE.mode = 2;
 } else {
  str = JS_HE.sDeb + foo.document.body.innerText + JS_HE.sFin;
  JS_HE.sDeb = "";
  JS_HE.sFin = "";
  foo.document.write(str);
  foo.document.close();
  if (JS_HE.fontChange) {
   if (JS_HE.fontFamily != "" && JS_HE.fontFamily != null) {
    foo.document.body.style.fontFamily = JS_HE.fontFamily;
   } else {
    foo.document.body.style.removeAttribute("fontFamily");
   }
   if (JS_HE.fontSize != "" && JS_HE.fontSize != null) {
    foo.document.body.style.fontSize = JS_HE.fontSize;
   } else {
    foo.document.body.style.removeAttribute("fontSize");
   }
   if (JS_HE.bgColor != "" && JS_HE.bgColor != null) {
    foo.document.body.bgColor = JS_HE.bgColor;
   } else {
    foo.document.body.removeAttribute("bgColor");
   }
   if (JS_HE.text != "" && JS_HE.text != null) {
    foo.document.body.text = JS_HE.text;
   } else {
    foo.document.body.removeAttribute("text");
   }
   if (JS_HE.background != "" && JS_HE.background != null) {
    foo.document.body.background = JS_HE.background;
   } else {
    foo.document.body.removeAttribute("background");
   }
  }
  document.getElementById("toolbar-edit_" + PuID).className = "OnBouttons";
  document.getElementById("toolbar-source_" + PuID).className = "Off";
  JS_HE.mode = 1;
  if (JS_HE.toggleWasOn == "yes") {
   ToggleBorders();
   JS_HE.toggleWasOn = "no";
  }
 }
 ShowStatus();
};

ShowStatus = function()
{
 var str = "";
 if (JS_HE.borderShown == "yes") {
  str = "Bordures : ON  ";
 } else {
  str = "Bordures : OFF";
 }
 var status = eval("sbt_" + PuID);
 status.innerHTML = "<span class=\"Text\">" + str + "</span>&nbsp;";
};

InitFoo = function()
{
 var iFrames = document.all.tags("IFRAME");
 var el = iFrames[0];
 var uID = PuID;
 el.frameWindow = document.frames[el.id];
 el.frameWindow.document.oncontextmenu = function ()
 {
  if (!el.frameWindow.event.ctrlKey){
   ShowContextMenu(el.frameWindow.event);
   return false;
  }
 };
 el.frameWindow.document.onerror = function ()
 {
  return true;
 };
 el.frameWindow.document.onkeypress  = function() { GUI(null); }
 el.frameWindow.document.onkeyup     = function() { GUI(null); }
 el.frameWindow.document.onmouseup   = function() { GUI(null); }
 el.frameWindow.document.ondrop      = function() { GUI(100); }
 el.frameWindow.document.oncut       = function() { GUI(100); }
 el.frameWindow.document.onpaste     = function() { GUI(100); }
 el.frameWindow.document.onblur      = function() { GUI(-1); }
};

UpdateGUI = function( inDELAY )
{
 var tooSoon = null, queue = null, aBut, i = 0, oBut, oCmd, oSlt, fName, fSize, fOK, str = "", pre;
 RemoveGarbage();
 var foo = eval("des_" + PuID);
 if (inDELAY == null) {
  var runDelay = 0;
 } else {
  var runDelay = inDELAY;
 }
 if (runDelay > 0) {
  return setTimeout(function(){ UpdateGUI(); }, runDelay);
 }
 if (tooSoon == 1 && runDelay > 0) {
   queue = 1;
   return;
 }
 tooSoon = 1;
 setTimeout(function(){tooSoon = 0; if(queue==1){UpdateGUI(-1);} queue = 0;}, 333);
 aBut = new Array("Bold","Italic","Underline","JustifyLeft","JustifyCenter","JustifyRight","InsertOrderedList","InsertUnorderedList");
 for (i = 0; i < aBut.length; i++) {
  oBut = document.all["b_" + aBut[i] + "_" + PuID];
  if (oBut == null) {
   continue;
  }
  oCmd = foo.document.queryCommandState(aBut[i]);
  if (!oCmd) {
   if (oBut.className != "Boutton") {
    oBut.className = "Boutton";
   }
   if (oBut.disabled != false) {
    oBut.disabled = false;
   }
  } else if (oCmd) {
   if (oBut.className != "BouttonActif") {
    oBut.className = "BouttonActif";
   }
   if (oBut.disabled != false) {
    oBut.disabled = false;
   }
  }
 };
 fName = foo.document.queryCommandValue('FontName');
 if (fName != null && fName != "") { fName = fName.toLowerCase();}
 fSize = foo.document.queryCommandValue('FontSize');
 oSlt = document.all["FontFamily_" + PuID];
 fOK = 0;
 if (fName != null && fName != "") {
  for (i = 0; i < oSlt.length; i++) {
   if (oSlt[i].text.toLowerCase() == fName) {
    oSlt.selectedIndex = i;
    fOK = 1;
    break;
   }
  };
 }
 if (fOK == 0) {
  oSlt.selectedIndex = 0;
 }
 oSlt = document.all["FontSize_" + PuID];
 fOK = 0;
 if (fSize != null && fSize != "") {
  for (i = 0; i < oSlt.length; i++) {
   if (parseInt(oSlt[i].text.toLowerCase(), 10) == parseInt(fSize, 10)) {
    oSlt.selectedIndex = i;
    fOK = 1;
    break;
   }
  };
 }
 if (fOK == 0) {
  oSlt.selectedIndex = 0;
 }
};

ToggleBorders = function()
{
 var foo = eval("des_" + PuID);
 if (foo.document.body) {
  var i = 0, j = 0, k = 0, allRows, allCellsInRow;
  var allTables = foo.document.body.getElementsByTagName("TABLE");
  var allLinks = foo.document.body.getElementsByTagName("A");
  for (i = 0; i < allTables.length; i++) {
   if (JS_HE.borderShown == "no") {
    allTables[i].style.border = "1px dotted #BFBFBF";
   } else {
    allTables[i].removeAttribute("style");
   }
   allRows = allTables[i].rows;
   for (j = 0; j < allRows.length; j++) {
    allCellsInRow = allRows[j].cells;
    for (k = 0; k < allCellsInRow.length; k++) {
     if (JS_HE.borderShown == "no") {
      allCellsInRow[k].style.border = "1px dotted #BFBFBF";
     } else {
      allCellsInRow[k].removeAttribute("style");
     }
    };
   };
  };
  for (i = 0; i < allLinks.length; i++) {
   if (JS_HE.borderShown == "no") {
    if (allLinks[i].href.toUpperCase() == "") {
     allLinks[i].style.border = "1px dashed #000000";;
     allLinks[i].style.width = "20px";
     allLinks[i].style.height = "16px";
     allLinks[i].style.backgroundColor = "#FFFFCC";
     allLinks[i].style.color = "#FFFFCC";
    }
   } else {
    allLinks[i].removeAttribute("style");
   }
  };
  foo.document.body.innerHTML = foo.document.body.innerHTML;
 }
 if (JS_HE.borderShown == "no") {
  JS_HE.borderShown = "yes";
 } else {
  JS_HE.borderShown = "no";
 }
 ShowStatus();
};

RemoveGarbage = function()
{
 var inTOK = 0;
 var foo = eval("des_" + PuID);
 if (foo.document.body) {
  str = foo.document.body.innerHTML;
  if (str.search(/<(FONT face="")>([^>]*)<\/FONT>/gi) > 0) {
   str = str.replace(/<(FONT face="")>([^>]*)<\/FONT>/gi, "$2");
   inTOK = 1;
  }
  if (str.search(/<(FONT)>([^>]*)<\/FONT>/gi) > 0) {
   str = str.replace(/<(FONT)>([^>]*)<\/FONT>/gi, "$2");
   inTOK = 1;
  }
  if (inTOK == 1) {
   if (str.substring(str.length - 4, str.length).toUpperCase() == "</P>") {
    str = str.substring(0, str.lastIndexOf("<P>")) + str.substring(str.lastIndexOf("<P>") + 3, str.length);
    str = str.substring(0, str.length - 4);
   }
   foo.document.body.innerHTML = str;
  }
 }
};

DoLoad = function()
{
 if (document.readyState == "complete") {
  var foo = eval("des_" + PuID);
  foo.document.designMode = 'On';
  SetValue();
  if (JS_HE.tableDefault != 0) {
   ToggleBorders();
  }
  ShowStatus();
  foo.focus();
  InitFoo();
  UpdateGUI(null);
  foo.focus();
 } else {
  setTimeout("DoLoad();", 100);
 }
};

DoFiles = function()
{
 if (IsAllowed()) {
  if (IsSelection()) {
   if (!IsTableSelected()) {
    var leftPos = (screen.availWidth-700) / 2;
    var topPos = (screen.availHeight-500) / 2;
    JS_HE.CloseAllWin();
    JS_HE.uplWin = window.open(HE_PATH + 'Popups/HE_Files.asp?foo=' + PuID,'','width=700,height=500,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
   } else {
    alert('Il est impossible de créer un lien sur cet élément');
   }
  } else {
   alert('Il n\'y a pas de selection en cours');
  }
 }
};

DoImage = function()
{
 if (IsAllowed()) {
  if (IsImageSelected()) {
   var leftPos = (screen.availWidth-500) / 2;
   var topPos = (screen.availHeight-320) / 2;
   JS_HE.CloseAllWin();
   JS_HE.imageWin = window.open(HE_PATH + 'Popups/HE_EditImage.htm?foo=' + PuID,'','width=500,height=320,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
  } else {
   var leftPos = (screen.availWidth-700) / 2;
   var topPos = (screen.availHeight-500) / 2;
   JS_HE.CloseAllWin();
   JS_HE.imageWin = window.open(HE_PATH + 'Popups/HE_Images.asp?mode=editor&foo=' + PuID + '&_display=' + JS_HE.heMode,'','width=700,height=500,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
  }
 }
};

DoImageCell = function( inFORM )
{
 var leftPos = (screen.availWidth-700) / 2;
 var topPos = (screen.availHeight-500) / 2;
 JS_HE.imageWin = window.open(HE_PATH + 'Popups/HE_Images.asp?mode=opener&f_form=' + inFORM + '&f_win=editCellWin&foo=' + PuID,'','width=700,height=500,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
};

DoImageTable = function( inFORM )
{
 var leftPos = (screen.availWidth-700) / 2;
 var topPos = (screen.availHeight-500) / 2;
 JS_HE.imageWin = window.open(HE_PATH + 'Popups/HE_Images.asp?mode=opener&f_form=' + inFORM + '&f_win=editTableWin&foo=' + PuID,'','width=700,height=500,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
};

DoImageNewTable = function( inFORM )
{
 var leftPos = (screen.availWidth-700) / 2;
 var topPos = (screen.availHeight-500) / 2;
 JS_HE.imageWin = window.open(HE_PATH + 'Popups/HE_Images.asp?mode=opener&f_form=' + inFORM + '&f_win=insertTableWin&foo=' + PuID,'','width=700,height=500,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
};

DoImagePage = function( inFORM )
{
 var leftPos = (screen.availWidth-700) / 2;
 var topPos = (screen.availHeight-500) / 2;
 JS_HE.imageWin = window.open(HE_PATH + 'Popups/HE_Images.asp?mode=opener&f_form=' + inFORM + '&f_win=editPageWin&foo=' + PuID,'','width=700,height=500,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
};

DoCommand = function( inCMD )
{
 if (IsAllowed()) {
  document.execCommand(inCMD)
  var foo = eval("des_" + PuID);
  foo.focus();
 }
};

DoFontFamily = function( inSLT )
{
 var foo = eval("des_" + PuID);
 if (IsAllowed()) {
  foo.document.execCommand('FontName', false, inSLT.options[inSLT.selectedIndex].value);
 } else {
  inSLT.selectedIndex = 0;
 }
 foo.focus();
};

DoFontSize = function( inSLT )
{
 var foo = eval("des_" + PuID);
 if (IsAllowed()) {
  foo.document.execCommand('FontSize', true, inSLT.options[inSLT.selectedIndex].value);
 } else {
  inSLT.selectedIndex = 0;
 }
 foo.focus();
};

DoModele = function( inSLT )
{
 var foo = eval("des_" + PuID);
 if (IsAllowed()) {
  if (inSLT.options[inSLT.selectedIndex].value != -1) {
   if (confirm("Attention : L\'insertion d\'un modèle effacera tout le contenu actuel." + String.fromCharCode(13) + "Etes-vous sur de vouloir continuer ?")) {
    foo.document.body.innerHTML = inSLT.options[inSLT.selectedIndex].value;
    InitFoo();
    UpdateGUI(null);
   }
  }
 }
 inSLT.selectedIndex = 0;
 foo.focus();
};

ShowMenu = function( inMENU, inW, inH)
{
 var lefter = event.clientX;
 var leftoff = event.offsetX;
 var topper = event.clientY;
 var topoff = event.offsetY;
 var oPopBody = JS_HE.oPopup.document.body;
 var HTMLContent = eval(inMENU).innerHTML
 oPopBody.innerHTML = HTMLContent
 oPopBody.onselectstart = function() {return false;};
 oPopBody.ondragstart = function() {return false;};
 oPopBody.oncontextmenu = function() {return false;};
 JS_HE.oPopup.show(lefter - leftoff - 2, topper - topoff + 22, inW, inH, document.body);
 return false;
};

ShowContextMenu = function( inEVENT )
{
 var foo = eval("des_" + PuID);
 var menu = "ContextMenu_" + PuID;
 var myW = 182;
 var myH = 67;
 var lefter = inEVENT.clientX;
 var topper = inEVENT.clientY;
 var oPopBody = JS_HE.oPopup.document.body;
 var HTMLContent = "<table style='BORDER-LEFT: buttonface 1px solid; BORDER-TOP: buttonface 1px solid; BORDER-RIGHT: buttonshadow 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;' cellpadding=0 cellspacing=0><tr><td>";
 HTMLContent += eval(menu).innerHTML;
 if (IsImageSelected()) {
  HTMLContent += eval("CmImg_" + PuID).innerHTML;
  HTMLContent += eval("CmLink_" + PuID).innerHTML;
  myH = myH + 46;
 }
 if (IsTextSelected() && JS_HE.sourceModeView != true) {
  HTMLContent += eval("CmLink_" + PuID).innerHTML;
  myH = myH + 23;
 }
 if (IsTableSelected() || IsCursorInTableCell()) {
  HTMLContent += eval("CmTableMenu_" + PuID).innerHTML;
  myH = myH + 23;
 }
 if (IsCursorInTableCell()) {
  HTMLContent += eval("CmTableCell_" + PuID).innerHTML;
  HTMLContent += eval("CmTableCols_" + PuID).innerHTML;
  HTMLContent += eval("CmTableRows_" + PuID).innerHTML;
  HTMLContent += eval("CmTableDel_" + PuID).innerHTML;
  myH = myH + 145;
 }
 HTMLContent += "</td></tr></table>";
 oPopBody.innerHTML = HTMLContent;
 oPopBody.onselectstart = function() {return false;};
 oPopBody.ondragstart = function() {return false;};
 oPopBody.oncontextmenu = function() {return false;};
 JS_HE.oPopup.show(lefter + 2, topper + 2, myW, myH, foo.document.body);
};

InsertTable = function()
{
 if (IsAllowed()) {
  var leftPos = (screen.availWidth-500) / 2;
  var topPos = (screen.availHeight-350) / 2;
  JS_HE.CloseAllWin();
  JS_HE.insertTableWin = window.open(HE_PATH + 'Popups/HE_InsertTable.htm?foo=' + PuID + '&borderShown=' + JS_HE.borderShown,'','width=500,height=350,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
 }
};

ModifyTable = function()
{
 if (IsAllowed()) {
  var leftPos = (screen.availWidth-500) / 2;
  var topPos = (screen.availHeight-300) / 2;
  JS_HE.CloseAllWin();
  JS_HE.editTableWin = window.open(HE_PATH + 'Popups/HE_EditTable.htm?foo=' + PuID + '&borderShown=' + JS_HE.borderShown,'','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
 }
};

ModifyCell = function()
{
 if (IsAllowed()) {
  var leftPos = (screen.availWidth-500) / 2;
  var topPos = (screen.availHeight-300) / 2;
  JS_HE.CloseAllWin();
  JS_HE.editCellWin = window.open(HE_PATH + 'Popups/HE_EditCell.htm?foo=' + PuID + '&borderShown=' + JS_HE.borderShown,'','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
 }
};

InsertColRight = function()
{
 if (IsCursorInTableCell()) {
  var rowcount, position, newCell;
  var moveFromEnd = (JS_HE.selectedTR.cells.length - 1) - (JS_HE.selectedTD.cellIndex);
  var aRows = JS_HE.selectedTable.rows;
  for (var i = 0; i < aRows.length; i++) {
   rowCount = aRows[i].cells.length - 1;
   position = rowCount - moveFromEnd;
   if (position < 0) {position = 0;}
   newCell = aRows[i].insertCell(position + 1);
   newCell.innerHTML = "&nbsp;";
   if (JS_HE.borderShown == "yes") {newCell.style.border = "1px dotted #BFBFBF";}
  };
  JS_HE.oPopup.hide();
 }
};

InsertColLeft = function()
{
 if (IsCursorInTableCell()) {
  var rowcount, position, newCell;
  var moveFromEnd = (JS_HE.selectedTR.cells.length - 1) - (JS_HE.selectedTD.cellIndex);
  var aRows = JS_HE.selectedTable.rows;
  for (var i = 0; i < aRows.length; i++) {
   rowCount = aRows[i].cells.length - 1;
   position = rowCount - moveFromEnd;
   if (position < 0) {position = 0;}
   newCell = aRows[i].insertCell(position);
   newCell.innerHTML = "&nbsp;";
   if (JS_HE.borderShown == "yes") {newCell.style.border = "1px dotted #BFBFBF";}
  };
  JS_HE.oPopup.hide();
 }
};

InsertRowBefore = function()
{
 if (IsCursorInTableCell()) {
  var newTD;
  var numCols = 0;
  var aCells = JS_HE.selectedTR.cells;
  for (var i = 0; i < aCells.length; i++) {
   numCols = numCols + aCells[i].getAttribute("colSpan");
  };
  var newTR = JS_HE.selectedTable.insertRow(JS_HE.selectedTR.rowIndex);
  for (i = 0; i < numCols; i++) {
   newTD = newTR.insertCell();
   newTD.innerHTML = "&nbsp;";
   if (JS_HE.borderShown == "yes") {newTD.style.border = "1px dotted #BFBFBF";}
  };
  JS_HE.oPopup.hide();
 }
};

InsertRowAfter = function()
{
 if (IsCursorInTableCell()) {
  var newTD;
  var numCols = 0;
  var aCells = JS_HE.selectedTR.cells;
  for (var i = 0; i < aCells.length; i++) {
   numCols = numCols + aCells[i].getAttribute("colSpan");
  };
  var newTR = JS_HE.selectedTable.insertRow(JS_HE.selectedTR.rowIndex + 1);
  for (i = 0; i < numCols; i++) {
   newTD = newTR.insertCell();
   newTD.innerHTML = "&nbsp;";
   if (JS_HE.borderShown == "yes") {newTD.style.border = "1px dotted #BFBFBF";}
  };
  JS_HE.oPopup.hide();
 }
};

DeleteRow = function()
{
 if (IsCursorInTableCell()) {
  if (JS_HE.selectedTable.rows.length > 1) {
   JS_HE.selectedTable.deleteRow(JS_HE.selectedTR.rowIndex);
  } else {
   JS_HE.selectedTable.removeNode(true);
  }
  JS_HE.oPopup.hide();
 }
};

DeleteCol = function()
{
 if (IsCursorInTableCell()) {
  var endOfRow, position, aCellRows;
  var moveFromEnd = (JS_HE.selectedTR.cells.length - 1) - (JS_HE.selectedTD.cellIndex);
  var aRows = JS_HE.selectedTable.rows;
  if (aRows.length > 1 || (aRows.length == 1 && aRows[0].cells.length > 1)) {
   for (var i = 0; i < aRows.length; i++) {
    endOfRow = aRows[i].cells.length - 1;
    position = endOfRow - moveFromEnd;
    if (position < 0) {position = 0;}
    aCellRows = aRows[i].cells;
    if (aCellRows[position].colSpan > 1) {
     aCellRows[position].colSpan = aCellRows[position].colSpan - 1;
    } else {
     aRows[i].deleteCell(position);
    }
   };
  } else {
   JS_HE.selectedTable.removeNode(true);
  }
  JS_HE.oPopup.hide();
 }
};

DoLink = function()
{
 if (IsAllowed()) {
  if (IsSelection()) {
   if (!IsTableSelected()) {
    var leftPos = (screen.availWidth-500) / 2;
    var topPos = (screen.availHeight-500) / 2;
    JS_HE.CloseAllWin();
    JS_HE.linkWin = window.open(HE_PATH + 'Popups/HE_Link.asp?foo=' + PuID + '&h_groupid=' + JS_HE.groupid + '&h_catid=' + JS_HE.catid, '', 'width=500,height=500,scrollbars=yes,resizable=yes,titlebar=0,top=' + topPos + ',left=' + leftPos);
   } else {
    alert('Il est impossible de créer un lien sur cet élément');
   }
  } else {
   alert('Il n\'y a pas de selection en cours');
  }
 }
};

DoAnchor = function()
{
 if (IsAllowed()) {
  var foo = eval("des_" + PuID);
  var leftPos = (screen.availWidth-500) / 2;
  var topPos = (screen.availHeight-300) / 2;
  if ((foo.document.selection.type == "Control") && (foo.document.selection.createRange()(0).tagName == "A") && (foo.document.selection.createRange()(0).href == "")) {
   JS_HE.CloseAllWin();
   JS_HE.anchorWin = window.open(HE_PATH + 'Popups/HE_Anchor.htm?foo_mode=Edit&foo=' + PuID,'','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
  } else {
   JS_HE.CloseAllWin();
   JS_HE.anchorWin = window.open(HE_PATH + 'Popups/HE_Anchor.htm?foo_mode=Add&foo=' + PuID,'','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
  }
 }
};

DoEmail = function()
{
 if (IsAllowed()) {
  if (IsSelection()) {
   var leftPos = (screen.availWidth-500) / 2;
   var topPos = (screen.availHeight-300) / 2;
   JS_HE.CloseAllWin();
   JS_HE.emailWin = window.open(HE_PATH + 'Popups/HE_Email.htm?foo=' + PuID,'','width=500,height=300,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
  } else {
   alert('Il n\'y a pas de selection en cours');
  }
 }
};

DoChars = function()
{
 if (IsAllowed()) {
  var nRows = parseInt((HE_CHARS.length / 4), 10);
  if (HE_CHARS.length % 4 != 0) {
   nRows++;
  }
  ShowMenu('CharMenu_' + PuID, 104, (28 * nRows) + 2);
 } else {
  var foo = eval("des_" + PuID);
  foo.focus();
 }
};

DoProperties = function()
{
 if (IsAllowed()) {
  var leftPos = (screen.availWidth-500) / 2;
  var topPos = (screen.availHeight-450) / 2;
  JS_HE.CloseAllWin();
  JS_HE.editPageWin = window.open(HE_PATH + 'Popups/HE_Properties.htm?foo=' + PuID + '&borderShown=' + JS_HE.borderShown,'','width=500,height=450,scrollbars=no,resizable=no,titlebar=0,top=' + topPos + ',left=' + leftPos);
 }
};

InsertChar = function( inCHAR )
{
 var foo = eval("des_" + PuID);
 var sel = foo.document.selection;
 var rng = sel.createRange();
 rng.pasteHTML(inCHAR.innerHTML)
 JS_HE.oPopup.hide()
};

DoTable = function()
{
 if (IsAllowed()) {
  if (!IsImageSelected()) {
   if (IsCursorInTableCell() || IsTableSelected()) {
    document.getElementById("TR_ModifyTable_" + PuID).disabled = false;
    document.getElementById("TD_ModifyTable_" + PuID).disabled = false;
    document.getElementById("TD_ModifyTable_" + PuID).style.filter = "";
    document.getElementById("IMG_ModifyTable_" + PuID).style.filter = "";
   } else {
    document.getElementById("TR_ModifyTable_" + PuID).disabled = true;
    document.getElementById("TD_ModifyTable_" + PuID).disabled = true;
    document.getElementById("TD_ModifyTable_" + PuID).style.filter = "Alpha(Opacity=100)";
    document.getElementById("IMG_ModifyTable_" + PuID).style.filter = "Alpha(Opacity=20)";
   }
   if (IsCursorInTableCell()) {
    var IsValid = false;
    var FilterTD = "";
    var FilterIMG = "";
   } else {
    var IsValid = true;
    var FilterTD = "Alpha(Opacity=100)";
    var FilterIMG = "Alpha(Opacity=20)";
   }
   var aTable = new Array("ModifyCell", "InsertColRight", "InsertColLeft", "InsertRowBefore", "InsertRowAfter", "DeleteRow", "DeleteCol");
   for (var i = 0; i < aTable.length; i++) {
    document.getElementById("TR_" + aTable[i] + "_" + PuID).disabled = IsValid;
    document.getElementById("TD_" + aTable[i] + "_" + PuID).disabled = IsValid;
    document.getElementById("TD_" + aTable[i] + "_" + PuID).style.filter = FilterTD;
    document.getElementById("IMG_" + aTable[i] + "_" + PuID).style.filter = FilterIMG;
   };
   ShowMenu('TableMenu_' + PuID, 200, 230);
  } else {
   alert("Il est impossible de créer une table sur cet élément");
  }
 } else {
  var foo = eval("des_" + PuID);
  foo.focus();
 }
};

DoTableColor = function()
{
 JS_HE.colorType = 3;
 var lefter = JS_HE.insertTableWin.event.clientX;
 var topper = JS_HE.insertTableWin.event.clientY;
 JS_HE.oPopWin = JS_HE.insertTableWin.createPopup();
 var oPopBody = JS_HE.oPopWin.document.body;
 var HTMLContent = eval('ColorMenuOpener_' + PuID).innerHTML
 oPopBody.innerHTML = HTMLContent;
 oPopBody.onselectstart = function() {return false;};
 oPopBody.ondragstart = function() {return false;};
 oPopBody.oncontextmenu = function() {return false;};
 JS_HE.oPopWin.show(lefter + 25, topper - 160, 240, 320, JS_HE.insertTableWin.document.body);
 return false;
};

DoTableColorEdit = function()
{
 JS_HE.colorType = 4;
 var lefter = JS_HE.editTableWin.event.clientX;
 var topper = JS_HE.editTableWin.event.clientY;
 JS_HE.oPopWin = JS_HE.editTableWin.createPopup();
 var oPopBody = JS_HE.oPopWin.document.body;
 var HTMLContent = eval('ColorMenuOpener_' + PuID).innerHTML
 oPopBody.innerHTML = HTMLContent
 oPopBody.onselectstart = function() {return false;};
 oPopBody.ondragstart = function() {return false;};
 oPopBody.oncontextmenu = function() {return false;};
 JS_HE.oPopWin.show(lefter + 25, topper - 160, 240, 320, JS_HE.editTableWin.document.body);
 return false;
};

DoCellColorEdit = function()
{
 JS_HE.colorType = 5;
 var lefter = JS_HE.editCellWin.event.clientX;
 var topper = JS_HE.editCellWin.event.clientY;
 JS_HE.oPopWin = JS_HE.editCellWin.createPopup();
 var oPopBody = JS_HE.oPopWin.document.body;
 var HTMLContent = eval('ColorMenuOpener_' + PuID).innerHTML
 oPopBody.innerHTML = HTMLContent
 oPopBody.onselectstart = function() {return false;};
 oPopBody.ondragstart = function() {return false;};
 oPopBody.oncontextmenu = function() {return false;};
 JS_HE.oPopWin.show(lefter + 25, topper - 160, 240, 320, JS_HE.editCellWin.document.body);
 return false;
};

DoPageColorBg = function()
{
 JS_HE.colorType = 6;
 var lefter = JS_HE.editPageWin.event.clientX;
 var topper = JS_HE.editPageWin.event.clientY;
 JS_HE.oPopWin = JS_HE.editPageWin.createPopup();
 var oPopBody = JS_HE.oPopWin.document.body;
 var HTMLContent = eval('ColorMenuOpener_' + PuID).innerHTML
 oPopBody.innerHTML = HTMLContent
 oPopBody.onselectstart = function() {return false;};
 oPopBody.ondragstart = function() {return false;};
 oPopBody.oncontextmenu = function() {return false;};
 JS_HE.oPopWin.show(lefter + 25, topper - 160, 240, 320, JS_HE.editPageWin.document.body);
 return false;
};

DoPageColorText = function()
{
 JS_HE.colorType = 7;
 var lefter = JS_HE.editPageWin.event.clientX;
 var topper = JS_HE.editPageWin.event.clientY;
 JS_HE.oPopWin = JS_HE.editPageWin.createPopup();
 var oPopBody = JS_HE.oPopWin.document.body;
 var HTMLContent = eval('ColorMenuOpener_' + PuID).innerHTML
 oPopBody.innerHTML = HTMLContent
 oPopBody.onselectstart = function() {return false;};
 oPopBody.ondragstart = function() {return false;};
 oPopBody.oncontextmenu = function() {return false;};
 JS_HE.oPopWin.show(lefter + 25, topper - 160, 240, 320, JS_HE.editPageWin.document.body);
 return false;
};

DoPageColorLink = function()
{
 JS_HE.colorType = 8;
 var lefter = JS_HE.editPageWin.event.clientX;
 var topper = JS_HE.editPageWin.event.clientY;
 JS_HE.oPopWin = JS_HE.editPageWin.createPopup();
 var oPopBody = JS_HE.oPopWin.document.body;
 var HTMLContent = eval('ColorMenuOpener_' + PuID).innerHTML
 oPopBody.innerHTML = HTMLContent
 oPopBody.onselectstart = function() {return false;};
 oPopBody.ondragstart = function() {return false;};
 oPopBody.oncontextmenu = function() {return false;};
 JS_HE.oPopWin.show(lefter + 25, topper - 160, 240, 320, JS_HE.editPageWin.document.body);
 return false;
};

DoFontColor = function()
{
 if (IsAllowed()) {
  JS_HE.colorType = 1;
  ShowMenu('ColorMenu_' + PuID, 240, 320);
 } else {
  var foo = eval("des_" + PuID);
  foo.focus();
 }
};

DoBackColor = function()
{
 if (IsAllowed()) {
  JS_HE.colorType = 2;
  ShowMenu('ColorMenu_' + PuID, 240, 320);
 } else {
  var foo = eval("des_" + PuID);
  foo.focus();
 }
};

ShowColor = function( inTDC, inTDT, inTD, inTXT)
{
 var str = "";
 if (inTXT != "") {
  inTDT.innerHTML = inTD.style.backgroundColor.toUpperCase() + " " + inTXT;
 } else {
  inTDT.innerHTML = inTD.style.backgroundColor.toUpperCase().replace("BUTTONFACE", "");
 }
 inTDC.style.backgroundColor = inTD.style.backgroundColor;
};

DoCleanCode = function()
{
 if (confirm("Etes-vous sur de vouloir nettoyer le code HTML ?")){
  if (JS_HE.borderShown == "yes") {
   ToggleBorders();
   JS_HE.toggleWasOn = "yes";
  } else {
   JS_HE.toggleWasOn = "no";
  }
  var foo = eval("des_" + PuID);
  foo.document.body.innerHTML = CleanHTMLCode(foo.document.body.innerHTML);
  if (JS_HE.toggleWasOn == "yes") {
   ToggleBorders();
   JS_HE.toggleWasOn = "no";
  }
 }
};

ToggleColor = function( inCOLORON, inCOLOROFF, inCOLOROFF_2 )
{
 inCOLORON.style.display = "";
 inCOLOROFF.style.display = "none";
 inCOLOROFF_2.style.display = "none";
};

DoColor = function( inCOLOR )
{
 var foo = eval("des_" + PuID);
 if (JS_HE.colorType == 1) {
  foo.document.execCommand('ForeColor', false, inCOLOR);
  JS_HE.oPopup.hide();
 } else if (JS_HE.colorType == 2) {
  foo.document.execCommand('BackColor', false, inCOLOR);
  JS_HE.oPopup.hide();
 } else if (JS_HE.colorType == 3) {
  JS_HE.insertTableWin.focus();
  JS_HE.insertTableWin.document.tableForm.bgcolor.value = inCOLOR;
  JS_HE.oPopWin.hide();
 } else if (JS_HE.colorType == 4) {
  JS_HE.editTableWin.focus();
  JS_HE.editTableWin.document.tableForm.bgcolor.value = inCOLOR;
  JS_HE.oPopWin.hide();
 } else if (JS_HE.colorType == 5) {
  JS_HE.editCellWin.focus();
  JS_HE.editCellWin.document.tableForm.bgcolor.value = inCOLOR;
  JS_HE.oPopWin.hide();
 } else if (JS_HE.colorType == 6) {
  JS_HE.editPageWin.focus();
  JS_HE.editPageWin.document.pageForm.bgcolor.value = inCOLOR;
  JS_HE.oPopWin.hide();
 } else if (JS_HE.colorType == 7) {
  JS_HE.editPageWin.focus();
  JS_HE.editPageWin.document.pageForm.textcolor.value = inCOLOR;
  JS_HE.oPopWin.hide();
 } else if (JS_HE.colorType == 8) {
  JS_HE.editPageWin.focus();
  JS_HE.editPageWin.document.pageForm.linkcolor.value = inCOLOR;
  JS_HE.oPopWin.hide();
 }
 JS_HE.colorType = 0;
};

CleanHTMLCode = function( inCODE )
{
 var code = "";
 code = inCODE;
 // removes all Class attributes on a tag eg. '<p class=asdasd>xxx</p>' returns '<p>xxx</p>'
 code = code.replace(/<([\w]+) class=([^ |>]*)([^>]*)/gi, "<$1$3");
 // removes all style attributes eg. '<tag style="asd asdfa aasdfasdf" something else>' returns '<tag something else>'
 code = code.replace(/<([\w]+) style="([^"]*)"([^>]*)/gi, "<$1$3");
 // gets rid of all xml stuff... <xml>,<\xml>,<?xml> or <\?xml>
 code = code.replace(/<\\?\??xml[^>]>/gi, "");
 // get rid of ugly colon tags <a:b> or </a:b>
 code = code.replace(/<\/?\w+:[^>]*>/gi, "");
 // removes all empty <p> tags
 code = code.replace(/<p([^>])*>(&nbsp;)*\s*<\/p>/gi,"");
 // removes all empty span tags
 code = code.replace(/<span([^>])*>(&nbsp;)*\s*<\/span>/gi,"");
 return code;
};

RemoveGuidelines = function( inFRA )
{
 var i = 0, j = 0, k = 0, allRows, allCellsInRow;
 var allTables = inFRA.document.body.getElementsByTagName("TABLE");
 var allLinks = inFRA.document.body.getElementsByTagName("A");
 for (i = 0; i < allTables.length; i++) {
  allTables[i].removeAttribute("style");
  allRows = allTables[i].rows;
  for (j = 0; j < allRows.length; j++) {
   allCellsInRow = allRows[j].cells;
   for (k = 0; k < allCellsInRow.length; k++) {
    allCellsInRow[k].removeAttribute("style");
   };
  };
 };
 for (i = 0; i < allLinks.length; i++) {
  allLinks[i].removeAttribute("style");
 };
 inFRA.document.body.innerHTML = inFRA.document.body.innerHTML;
};

UpdateValue = function()
{
 if (document.activeElement) {
  if (document.activeElement.parentElement.id == "he") {
   return false;
  } else {
   var myForm = eval("document.all.he_html_" + PuID);
   myForm.value = SaveHTMLPage();
  }
 }
};

GetHTMLValue = function()
{
 UpdateValue();
 var myForm = eval("document.all.he_html_" + PuID);
 return myForm.value;
};

SetHTMLValue = function( inVALUE )
{
 ModeDesign();
 JS_HE.html = inVALUE;
 DoLoad();
};

SetValue = function()
{
 var foo;
 foo = eval("des_" + PuID);
 foo.document.write(JS_HE.html);
 foo.document.close();
};

SaveHTMLPage = function()
{
 var str = "", maRegExp, pre, foo;
 if (JS_HE.previewModeView) {
  pre = eval("pre_" + PuID);
  if (JS_HE.isEditingHTMLPage == 1) {
   str = pre.document.documentElement.outerHTML;
  } else {
   str = pre.document.body.innerHTML;
  }
  if (JS_HE.heMode == "body") {
   var sDeb = str.substring(0, str.toUpperCase().indexOf(">", str.toUpperCase().indexOf("<BODY")) + 1);
   var sFin = str.substring(str.indexOf("</BODY>"), str.length);
   str = str.replace(sDeb, "");
   str = str.replace(sFin, "");
  }
 }
 if (JS_HE.sourceModeView) {
  foo = eval("des_" + PuID);
  str = foo.document.body.innerText;
 }
 if (JS_HE.editModeView) {
  foo = eval("des_" + PuID);
  if (JS_HE.isEditingHTMLPage == 1) {
   str = foo.document.documentElement.outerHTML;
  } else {
   str = foo.document.body.innerHTML;
  }
  if (JS_HE.heMode == "body") {
   var sDeb = str.substring(0, str.toUpperCase().indexOf(">", str.toUpperCase().indexOf("<BODY")) + 1);
   var sFin = str.substring(str.indexOf("</BODY>"), str.length);
   str = str.replace(sDeb, "");
   str = str.replace(sFin, "");
  }
 }
 maRegExp = /&amp;/g;
 str = str.replace(maRegExp,'&');
 str = str.replace(/<(table|td|a|form.*?) style="([^"]*)"([^>]*)/gi, "<$1$3");
 return str;
};

ModeSource = function()
{
 if (JS_HE.sourceModeView == false) {
  JS_HE.isEditingHTMLPage = 1;
  var foo = document.getElementById("des_" + PuID);
  var pre = document.getElementById("pre_" + PuID);
  document.getElementById("toolbar-edit_" + PuID).className = "Off";
  document.getElementById("toolbar-source_" + PuID).className = "OnBouttons";
  document.getElementById("toolbar-preview_" + PuID).className = "Off";
  foo.style.display = "";
  pre.style.display = "none";
  if (JS_HE.editModeOn) {
   SwitchMode();
  }
  JS_HE.sourceModeView = true;
  JS_HE.editModeView = false;
  JS_HE.previewModeView = false;
  JS_HE.editModeOn = false;
  JS_HE.sourceModeOn = true;
  document.getElementById("ImgEdit").src = HE_IMGPATH + "Editor/he_des.gif";
  document.getElementById("ImgSource").src = HE_IMGPATH + "Editor/he_src_up.gif";
  document.getElementById("ImgPreview").src = HE_IMGPATH + "Editor/he_pre.gif";
  foo = eval("des_" + PuID);
  foo.focus();
  foo.document.body.innerText = foo.document.body.innerText;
 }
};

ModePreview = function()
{
 if (JS_HE.previewModeView == false) {
  JS_HE.isEditingHTMLPage = 0;
  var foo = document.getElementById("des_" + PuID);
  var pre = document.getElementById("pre_" + PuID);
  document.getElementById("toolbar-edit_" + PuID).className = "Off";
  document.getElementById("toolbar-source_" + PuID).className = "Off";
  document.getElementById("toolbar-preview_" + PuID).className = "OnBouttons";
  foo.style.display = "none";
  pre.style.display = "";
  if (JS_HE.sourceModeOn) {
   ShowPreview(1);
  } else {
   ShowPreview(0);
  }
  JS_HE.sourceModeView = false;
  JS_HE.editModeView = false;
  JS_HE.previewModeView = true;
  document.getElementById("ImgEdit").src = HE_IMGPATH + "Editor/he_des.gif";
  document.getElementById("ImgSource").src = HE_IMGPATH + "Editor/he_src.gif";
  document.getElementById("ImgPreview").src = HE_IMGPATH + "Editor/he_pre_up.gif";
  pre = eval("pre_" + PuID);
  pre.document.write(pre.document.documentElement.innerHTML);
  pre.document.close();
  pre.focus();
 }
};

ShowPreview = function( inSRC )
{
 var previewHTML;
 var foo = eval("des_" + PuID);
 var pre = eval("pre_" + PuID);
 if (inSRC == 1) {
  if (JS_HE.heMode == "body") {
   previewHTML = JS_HE.sDeb + foo.document.body.innerText + JS_HE.sFin;
  } else {
   previewHTML = foo.document.body.innerText;
  }
 } else {
  previewHTML = foo.document.documentElement.outerHTML;
 }
 pre.document.write(previewHTML);
 pre.document.close();
 RemoveGuidelines(pre);
 var status = eval("sbt_" + PuID);
 status.innerHTML = "<span class=\"Text\">Bordures : OFF</span>&nbsp;";
};

function ContextOver( inCONTEXT ) {
 inCONTEXT.runtimeStyle.backgroundColor = "Highlight";
 if (inCONTEXT.state) {
  inCONTEXT.runtimeStyle.color = "GrayText";
 } else {
  inCONTEXT.runtimeStyle.color = "HighlightText";
 }
};

function ContextOut( inCONTEXT ) {
 inCONTEXT.runtimeStyle.backgroundColor = "";
 inCONTEXT.runtimeStyle.color = "";
};

function BouttonOver( inBoutton ) {
 inBoutton.style.borderBottom = "1px solid buttonshadow";
 inBoutton.style.borderRight = "1px solid buttonshadow";
 inBoutton.style.borderTop = "1px solid buttonhighlight";
 inBoutton.style.borderLeft = "1px solid buttonhighlight";
};

function BouttonOut( inBoutton ) {
 inBoutton.style.border = "1px solid buttonface";
};

function BouttonDown( inBoutton ) {
 inBoutton.style.borderBottom = "1px solid buttonhighlight";
 inBoutton.style.borderRight = "1px solid buttonhighlight";
 inBoutton.style.borderTop = "1px solid buttonshadow";
 inBoutton.style.borderLeft = "1px solid buttonshadow";
};

function CharsOut( inBoutton ) {
 inBoutton.style.border = "1px solid #666666";
};