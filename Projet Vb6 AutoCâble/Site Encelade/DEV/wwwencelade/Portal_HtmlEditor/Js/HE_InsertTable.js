function getQuery(inINDX){
 var queryString = location.search;
 var data = queryString.slice(1,queryString.length);
 var aData = data.split("&");
 var aOut = aData[inINDX].split("=");
 return aOut[1];
};

var uID = getQuery(0);

var borderShown = getQuery(1);

var foo = eval("window.opener.des_" + uID);

document.onkeydown = function () { 
 if (event.keyCode == 13) {
  InsertTable();
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
 window.opener.DoTableColor();
};

function DisplayImage( inFORM ) {
 window.opener.DoImageNewTable( inFORM );
};

function InsertTable() {
 error = 0;
 var sel = foo.document.selection;
 if (sel!=null) {
  var rng = sel.createRange();
  if (rng!=null) {
   var border = document.tableForm.border.value;
   var columns = document.tableForm.columns.value;
   var padding = document.tableForm.padding.value;
   var rows = document.tableForm.rows.value;
   var spacing = document.tableForm.spacing.value;
   var width = document.tableForm.width.value;
   var bgcolor = document.tableForm.bgcolor.value;
   var background = document.tableForm.background.value;
   if (isNaN(rows) || rows < 0 || rows == "") {
    alert("Veuillez vérifier le nombre de lignes");
    document.tableForm.rows.select();
    document.tableForm.rows.focus();
    error = 1;
   } else if (isNaN(columns) || columns < 0 || columns == "") {
    alert("Veuillez vérifier le nombre de colonnes");
    document.tableForm.columns.select();
    document.tableForm.columns.focus();
    error = 1;
   } else if (width < 0 || width == "") {
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
    if (bgcolor != "") {
     bgcolor = " bgcolor = " + bgcolor;
    } else {
     bgcolor = "";
    }
    if (borderShown == "yes") {
     var style = ' style="BORDER-RIGHT:1px dotted #BFBFBF; BORDER-TOP:1px dotted #BFBFBF; BORDER-LEFT:1px dotted #BFBFBF; BORDER-BOTTOM:1px dotted #BFBFBF;"';
    } else {
     var style = "";
    }
    if (background != "") {
     background = " background= " + background;
    } else {
     background = "";
    }
    var HTMLTable = "<Table width=" + width + " border=" + border + " cellpadding=" + padding + " cellspacing=" + spacing + bgcolor + style + background + ">";
    for (var i = 0; i < rows; i++) {
     HTMLTable += "<tr>";
     for (var j = 0; j < columns; j++) {
      HTMLTable += "<td " + style + ">&nbsp</td>";
     };
     HTMLTable += "</tr>";
    };
    HTMLTable += "</table>";
    foo.focus();
    rng.pasteHTML(HTMLTable);
   }
  }
 }
 if (error != 1) {
  if (window.opener.JS_HE.imageWin) {
   window.opener.JS_HE.imageWin.close();
   window.opener.JS_HE.imageWin = null;
  }
  self.close();
 }
};