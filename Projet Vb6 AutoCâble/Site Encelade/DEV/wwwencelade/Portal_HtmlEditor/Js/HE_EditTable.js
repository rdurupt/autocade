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

var tableBgColor = myTable.selectedTable.bgColor;
var tableSpacing = myTable.selectedTable.cellSpacing;
var tablePadding = myTable.selectedTable.cellPadding;
var tableBorder = myTable.selectedTable.border;
var tableWidth = myTable.selectedTable.width;
var tableBackground = myTable.selectedTable.background;

function setValues() {
 if (tableSpacing == "") {tableSpacing = 2;}
 if (tablePadding == "") {tablePadding = 1;}
 if (tableBorder == "") {tableBorder = 0;}
 document.tableForm.bgcolor.value = tableBgColor;
 document.tableForm.padding.value = tablePadding;
 document.tableForm.spacing.value = tableSpacing;
 document.tableForm.border.value = tableBorder;
 document.tableForm.width.value = tableWidth;
 document.tableForm.background.value = tableBackground;
 this.focus();
};

document.onkeydown = function () { 
 if (event.keyCode == 13) {
  EditTable();
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
 window.opener.DoTableColorEdit();
};

function DisplayImage( inFORM ) {
 window.opener.DoImageTable( inFORM );
};

function EditTable() {
 var error = 0;
 var border = document.tableForm.border.value;
 var padding = document.tableForm.padding.value;
 var spacing = document.tableForm.spacing.value;
 var width = document.tableForm.width.value;
 var bgcolor = document.tableForm.bgcolor.value;
 var background = document.tableForm.background.value;
 if (width < 0 || width == "") {
  alert("Veuillez vérifier la taille");
  document.tableForm.width.select();
  document.tableForm.width.focus();
  error = 1;
 } else if (isNaN(padding) || padding < 0 || padding == "") {
  alert("Veuillez vérifier l\'espacement du texte");
  document.tableForm.padding.select();
  document.tableForm.padding.focus();
  error = 1;
 } else if (isNaN(spacing) || spacing < 0 || spacing == "") {
  alert("Veuillez vérifier l\'espacement des cellules");
  document.tableForm.spacing.select();
  document.tableForm.spacing.focus();
  error = 1;
 } else if (isNaN(border) || border < 0 || border == "") {
  alert("Veuillez vérifier les bordures");
  document.tableForm.border.select();
  document.tableForm.border.focus();
  error = 1;
 }
 if (error != 1) {
  myTable.selectedTable.cellPadding = padding;
  myTable.selectedTable.cellSpacing = spacing;
  myTable.selectedTable.border = border;
  myTable.selectedTable.width = width;
  if (bgcolor != "") {
   myTable.selectedTable.bgColor = bgcolor;
  } else {
   myTable.selectedTable.removeAttribute("bgColor", 0);
  }
  if (background != "") {
   myTable.selectedTable.background = background;
  } else {
   myTable.selectedTable.removeAttribute('background', 0)
  }
  if (window.opener.JS_HE.imageWin) {
   window.opener.JS_HE.imageWin.close();
   window.opener.JS_HE.imageWin = null;
  }
  self.close();
 }
};