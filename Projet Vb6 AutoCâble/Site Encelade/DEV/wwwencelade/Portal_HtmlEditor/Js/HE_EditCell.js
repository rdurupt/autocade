function getQuery(inINDX){
 var queryString = location.search;
 var data = queryString.slice(1,queryString.length);
 var aData = data.split("&");
 var aOut = aData[inINDX].split("=");
 return aOut[1];
};

var uID = getQuery(0);

var foo = eval("window.opener.des_" + uID);

var myTable = eval("window.opener.JS_HE");

var cellBgColor = myTable.selectedTD.bgColor;
var cellWidth = myTable.selectedTD.width;
var cellAlign = myTable.selectedTD.align;
var cellvAlign = myTable.selectedTD.vAlign;
var cellBackground = myTable.selectedTD.background;

function setValues() {
 document.tableForm.bgcolor.value = cellBgColor;
 document.tableForm.width.value = cellWidth;
 document.tableForm.background.value = cellBackground;
 for (var i = 0; i < document.tableForm.align.length; i++) {
  if (document.tableForm.align.options[i].value.toUpperCase() == cellAlign.toUpperCase()) {
   document.tableForm.align.options[i].selected = true;
   break;
  }
 };
 for (i = 0; i < document.tableForm.valign.length; i++) {
  if (document.tableForm.valign.options[i].value.toUpperCase() == cellvAlign.toUpperCase()) {
   document.tableForm.valign.options[i].selected = true;
   break;
  }
 };
 this.focus();
};

document.onkeydown = function () { 
 if (event.keyCode == 13) {
  EditCell();
 }
};

document.onkeypress = onkeyup = function () {
 if (event.keyCode == 13) {
  event.cancelBubble = true;
  event.returnValue = false;
  return false;			
 }
};

function DisplayColor() {
 window.opener.DoCellColorEdit();
};

function DisplayImage( inFORM ) {
 window.opener.DoImageCell( inFORM );
};

function EditCell() {
 var error = 0;
 if (document.tableForm.width.value < 0) {
  alert("Veuillez vérifier la taille");
  document.tableForm.width.select();
  document.tableForm.width.focus();
  error = 1;
 }
 if (error != 1) {
  myTable.selectedTD.width = document.tableForm.width.value;
  if (document.tableForm.bgcolor.value != "") {
   myTable.selectedTD.bgColor = document.tableForm.bgcolor.value;
  } else {
   myTable.selectedTD.removeAttribute('bgColor', 0);
  }
  if (document.tableForm.align.options[document.tableForm.align.selectedIndex].value != "") {
   myTable.selectedTD.align = document.tableForm.align.options[document.tableForm.align.selectedIndex].value;
  } else {
   myTable.selectedTD.removeAttribute('align', 0);
  }
  if (document.tableForm.valign.options[document.tableForm.valign.selectedIndex].value != "") {
   myTable.selectedTD.vAlign = document.tableForm.valign.options[document.tableForm.valign.selectedIndex].value;
  } else {
   myTable.selectedTD.removeAttribute('vAlign', 0);
  }
  if (document.tableForm.background.value != "") {
   myTable.selectedTD.background = document.tableForm.background.value;
  } else {
   myTable.selectedTD.removeAttribute('background', 0)
  }
  if (window.opener.JS_HE.imageWin) {
   window.opener.JS_HE.imageWin.close();
   window.opener.JS_HE.imageWin = null;
  }
  self.close();        
 }
};