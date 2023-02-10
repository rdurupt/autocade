// PREVOIR COLSPAN / ROWSPAN - INCREASE / DECREASE SUR MENU TABLE

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

var HE_NAMED =
[
 ["#FF0000", "red"],
 ["#FFFF00", "yellow"],
 ["#00FF00", "lime"],
 ["#00FFFF", "cyan"],
 ["#0000FF", "blue"],
 ["#FF00FF", "magenta"],
 ["#FFFFFF", "white"],
 ["#F5F5F5", "whitesmoke"],
 ["#DCDCDC", "gainsboro"],
 ["#D3D3D3", "lightgrey"],
 ["#C0C0C0", "silver"],
 ["#A9A9A9", "darkgray"],
 ["#808080", "gray"],
 ["#696969", "dimgray"],
 ["#000000", "black"],
 ["#2F4F4F", "darkslategray"],
 ["#708090", "slategray"],
 ["#778899", "lightslategray"],
 ["#4682B4", "steelblue"],
 ["#4169E1", "royalblue"],
 ["#6495ED", "cornflowerblue"],
 ["#B0C4DE", "lightsteelblue"],
 ["#7B68EE", "mediumslateblue"],
 ["#6A5ACD", "slateblue"],
 ["#483D8B", "darkslateblue"],
 ["#191970", "midnightblue"],
 ["#000080", "navy"],
 ["#00008B", "darkblue"],
 ["#0000CD", "mediumblue"],
 ["#1E90FF", "dodgerblue"],
 ["#00BFFF", "deepskyblue"],
 ["#87CEFA", "lightskyblue"],
 ["#87CEEB", "skyblue"],
 ["#ADD8E6", "lightblue"],
 ["#B0E0E6", "powderblue"],
 ["#F0FFFF", "azure"],
 ["#E0FFFF", "lightcyan"],
 ["#AFEEEE", "paleturquoise"],
 ["#48D1CC", "mediumturquoise"],
 ["#20B2AA", "lightseagreen"],
 ["#008B8B", "darkcyan"],
 ["#008080", "teal"],
 ["#5F9EA0", "cadetblue"],
 ["#00CED1", "darkturquoise"],
 ["#00FFFF", "aqua"],
 ["#40E0D0", "turquoise"],
 ["#7FFFD4", "aquamarine"],
 ["#66CDAA", "mediumaquamarine"],
 ["#8FBC8F", "darkseagreen"],
 ["#3CB371", "mediumseagreen"],
 ["#2E8B57", "seagreen"],
 ["#006400", "darkgreen"],
 ["#008000", "green"],
 ["#228B22", "forestgreen"],
 ["#32CD32", "limegreen"],
 ["#00FF00", "lime"],
 ["#7FFF00", "chartreuse"],
 ["#7CFC00", "lawngreen"],
 ["#ADFF2F", "greenyellow"],
 ["#98FB98", "palegreen"],
 ["#90EE90", "lightgreen"],
 ["#00FF7F", "springgreen"],
 ["#00FA9A", "mediumspringgreen"],
 ["#556B2F", "darkolivegreen"],
 ["#6B8E23", "olivedrab"],
 ["#808000", "olive"],
 ["#BDB76B", "darkkhaki"],
 ["#B8860B", "darkgoldenrod"],
 ["#DAA520", "goldenrod"],
 ["#FFD700", "gold"],
 ["#F0E68C", "khaki"],
 ["#EEE8AA", "palegoldenrod"],
 ["#FFEBCD", "blanchedalmond"],
 ["#FFE4B5", "moccasin"],
 ["#F5DEB3", "wheat"],
 ["#FFDEAD", "navajowhite"],
 ["#DEB887", "burlywood"],
 ["#D2B48C", "tan"],
 ["#BC8F8F", "rosybrown"],
 ["#A0522D", "sienna"],
 ["#8B4513", "saddlebrown"],
 ["#D2691E", "chocolate"],
 ["#CD853F", "peru"],
 ["#F4A460", "sandybrown"],
 ["#8B0000", "darkred"],
 ["#800000", "maroon"],
 ["#A52A2A", "brown"],
 ["#B22222", "firebrick"],
 ["#CD5C5C", "indianred"],
 ["#F08080", "lightcoral"],
 ["#FA8072", "salmon"],
 ["#E9967A", "darksalmon"],
 ["#FFA07A", "lightsalmon"],
 ["#FF7F50", "coral"],
 ["#FF6347", "tomato"],
 ["#FF8C00", "darkorange"],
 ["#FFA500", "orange"],
 ["#FF4500", "orangered"],
 ["#DC143C", "crimson"],
 ["#FF0000", "red"],
 ["#FF1493", "deeppink"],
 ["#FF00FF", "fuchsia"],
 ["#FF69B4", "hotpink"],
 ["#FFB6C1", "lightpink"],
 ["#FFC0CB", "pink"],
 ["#DB7093", "palevioletred"],
 ["#C71585", "mediumvioletred"],
 ["#800080", "purple"],
 ["#8B008B", "darkmagenta"],
 ["#9370DB", "mediumpurple"],
 ["#8A2BE2", "blueviolet"],
 ["#4B0082", "indigo"],
 ["#9400D3", "darkviolet"],
 ["#9932CC", "darkorchid"],
 ["#BA55D3", "mediumorchid"],
 ["#DA70D6", "orchid"],
 ["#EE82EE", "violet"],
 ["#DDA0DD", "plum"],
 ["#D8BFD8", "thistle"],
 ["#E6E6FA", "lavender"],
 ["#F8F8FF", "ghostwhite"],
 ["#F0F8FF", "aliceblue"],
 ["#F5FFFA", "mintcream"],
 ["#F0FFF0", "honeydew"],
 ["#FAFAD2", "lightgoldenrodyellow"],
 ["#FFFACD", "lemonchiffon"],
 ["#FFF8DC", "cornsilk"],
 ["#FFFFE0", "lightyellow"],
 ["#FFFFF0", "ivory"],
 ["#FFFAF0", "floralwhite"],
 ["#FAF0E6", "linen"],
 ["#FDF5E6", "oldlace"],
 ["#FAEBD7", "antiquewhite"],
 ["#FFE4C4", "bisque"],
 ["#FFDAB9", "peachpuff"],
 ["#FFEFD5", "papayawhip"],
 ["#FFF5EE", "seashell"],
 ["#FFF0F5", "lavenderblush"],
 ["#FFE4E1", "mistyrose"],
 ["#FFFAFA", "snow"]
];

var HE_TABLE =
[
 ["IMG", "Bouttons/he_table.gif", "Insérer une table", "InsertTable();", "InsertTable"],
 ["IMG", "Bouttons/he_etable.gif", "Modifier une table", "ModifyTable();", "ModifyTable"],
 ["IMG", "Bouttons/he_ecell.gif", "Modifier une cellule", "ModifyCell();", "ModifyCell"],
 ["SEP", "Bouttons/he_hvert.gif", "", ""],
 ["IMG", "Bouttons/he_icdtable.gif", "Insérer une colonne à droite", "InsertColRight();", "InsertColRight"],
 ["IMG", "Bouttons/he_icgtable.gif", "Insérer une colonne à gauche", "InsertColLeft();", "InsertColLeft"],
 ["SEP", "Bouttons/he_hvert.gif", "", ""],
 ["IMG", "Bouttons/he_ilttable.gif", "Insérer une ligne au dessus", "InsertRowBefore();", "InsertRowBefore"],
 ["IMG", "Bouttons/he_ilbtable.gif", "Insérer une ligne au dessous", "InsertRowAfter();", "InsertRowAfter"],
 ["SEP", "Bouttons/he_hvert.gif", "", ""],
 ["IMG", "Bouttons/he_dltable.gif", "Supprimer la ligne", "DeleteRow();", "DeleteRow"],
 ["IMG", "Bouttons/he_dctable.gif", "Supprimer la colonne", "DeleteCol();", "DeleteCol"]
// ["SEP", "Bouttons/he_hvert.gif", "", ""],
// ["IMG", "Bouttons/he_icoltable.gif", "Increase Colspan", "IncreaseColspan();", "IncreaseColspan"],
// ["IMG", "Bouttons/he_dcoltable.gif", "Decrease Colspan", "DecreaseColspan();", "DecreaseColspan"]
];

var HE_CONTEXT =
[
 ["ContextMenu", 
  [
   ["Couper", "parent.document.execCommand('Cut');"],
   ["Copier", "parent.document.execCommand('Copy');"],
   ["Coller", "parent.document.execCommand('Paste');"]
  ]
 ],
 ["CmTableMenu", 
  [
   ["Modifier la table", "parent.ModifyTable();"]
  ]
 ],
 ["CmTableCell", 
  [
   ["Modifier la cellule", "parent.ModifyCell();"]
  ]
 ],
 ["CmTableCols", 
  [
   ["Insérer une colonne à droite", "parent.InsertColRight();"],
   ["Insérer une colonne à gauche", "parent.InsertColLeft();"]
  ]
 ],
 ["CmTableRows", 
  [
   ["Insérer une ligne au dessus", "parent.InsertRowBefore();"],
   ["Insérer une ligne au dessous", "parent.InsertRowAfter();"]
  ]
 ],
 ["CmTableDel", 
  [
   ["Supprimer la ligne", "parent.DeleteRow();"],
   ["Supprimer la colonne", "parent.DeleteCol();"]
  ]
 ],
 ["CmImg", 
  [
   ["Modifier l'image", "parent.DoImage();"]
  ]
 ],
 ["CmLink", 
  [
   ["Insérer / Modifier un lien", "parent.DoLink();"]
  ]
 ]
];

function Html_ContextMenu( inID ) {
 this.GetMenuChars = function()
 {
  var str = "", i = 0, tdDisp = 0;
  str += "<DIV ID=\"CharMenu_" + inID + "\" STYLE=\"display:none;\">";
  str += "<table cellpadding=\"1\" cellspacing=\"5\" border=\"1\" bordercolor=\"#666666\" style=\"cursor: hand;font-family: Verdana; font-size: 14px; font-weight: bold; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;\" bgcolor=\"buttonface\">";
  str += "<tr>";
  for (i = 0; i < HE_CHARS.length; i++) {
   if (i % 4 == 0 && i != 0) {
    str += "</tr>";
    str += "<tr>";
    tdDisp = 0;
   }
   str += "<td style=\"width=15px; cursor: hand;\" onClick=\"parent.InsertChar(this);\" onMouseOver=\"parent.BouttonOver(this);\" onMouseOut=\"parent.CharsOut(this);\" onMouseDown=\"parent.BouttonDown(this);\">" + HE_CHARS[i] + "</td>";
   tdDisp++;
  };
  for (i = tdDisp; i < 4; i++) {
   str += "<td style=\"width=15px; border: 1px solid buttonface;\">&nbsp;</td>";
  };
  str += "</tr>";
  str += "</table>";
  str += "</DIV>";
  return str;
 };
 this.GetMenuColors = function()
 {
  var str = "", i = 0, j = 0, k = 0, tdDisp = 0, totDisp = 0, fDisp = 0, aHex;
  str += "<DIV ID=\"ColorMenu_" + inID + "\" STYLE=\"display:none;\">";
  str += "<table id=\"Named_" + inID + "\" cellpadding=\"1\" cellspacing=\"2\" border=\"1\" bordercolor=\"#666666\" style=\"font-family: Verdana; font-size: 7px; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid; table-layout; fixed;\" bgcolor=\"buttonface\" width=\"100%\" height=\"100%\">";
  str += "<tr><td align=\"center\" colspan=\"16\" style=\"height: 20px;border:1px solid buttonface;font-family: verdana; font-size:12px;\">";
  str += "<font color=\"#FF0000\">Named</font> - ";
  str += "<a style=\"cursor:hand;\" onClick=\"javascript:parent.ToggleColor(Gray_" + inID + ", Named_" + inID + ", Safe_" + inID + ");\">GrayScale</a> - ";
  str += "<a style=\"cursor:hand;\" onClick=\"javascript:parent.ToggleColor(Safe_" + inID + ", Named_" + inID + ", Gray_" + inID + ");\">Safety</a>";
  str += "</td></tr>";
  str += "<tr><td colspan=\"16\" id=\"NamedColor_" + inID + "\" style=\"height: 20px;font-family: verdana; font-size:12px;\">&nbsp;</td></tr>";
  str += "<tr><td width=\"100%\" colspan=\"16\" id=\"NamedName_" + inID + "\" style=\"height: 20px;border: 1px solid buttonface; font-family: verdana; font-size:9px;\" align=\"left\">&nbsp;</td>";
  str += "</tr>";
  str += "<tr><td colspan=\"16\" onMouseOver=\"parent.ShowColor(NamedColor_" + inID + ", NamedName_" + inID + ", this, '');\" style=\"height: 20px;font-family: verdana; font-size:10px;cursor: hand;\" align=\"center\" onClick=\"parent.DoColor('');\">Aucune</td></tr>";
  str += "<tr>";
  for (i = 0; i < HE_NAMED.length; i++) {
   if (i % 16 == 0 && i != 0) {
    str += "</tr>";
    str += "<tr>";
    tdDisp = 0;
   }
   str += "<td style=\"height:12px;width:12px;cursor: hand;background-color:'" + HE_NAMED[i][0] + "'; font-size: 0px;\"";
   str += " onMouseOver=\"parent.ShowColor(NamedColor_" + inID + ", NamedName_" + inID + ", this, '" + HE_NAMED[i][1] + "');\"";
   str += " onClick=\"parent.DoColor('" + HE_NAMED[i][0] + "');\">" + String.fromCharCode(160) + "</td>";
   tdDisp++;
   totDisp++;
  };
  for (i = tdDisp; i < 16; i++) {
   str += "<td style=\"height:12px;width:12px;border: 1px solid buttonface; background-color:buttonface; font-size: 0px;\"";
   str += " onMouseOver=\"parent.ShowColor(NamedColor_" + inID + ", NamedName_" + inID + ", this, '');\">" + String.fromCharCode(160) + "</td>";
   totDisp++;
  };
  str += "</tr>";
  str += "<tr>";
  fDisp = totDisp;
  for (i = totDisp; i < (16 * 16); i++) {
   if (i % 16 == 0 && i != fDisp) {
    str += "</tr>";
    str += "<tr>";
   }
   str += "<td style=\"height:12px;width:12px;border: 1px solid buttonface; background-color:buttonface; font-size: 0px;\"";
   str += " onMouseOver=\"parent.ShowColor(NamedColor_" + inID + ", NamedName_" + inID + ", this, '');\">" + String.fromCharCode(160) + "</td>";
  };
  str += "</tr>";
  str += "</table>";
  aHex = new Array("0","1","2","3","4","5","6","7","8","9","A","B","C","D","E","F");
  str += "<table id=\"Gray_" + inID + "\" cellpadding=\"1\" cellspacing=\"2\" border=\"1\" bordercolor=\"#666666\" style=\"font-family: Verdana; font-size: 7px; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid; table-layout; fixed;display: none;\" bgcolor=\"buttonface\" width=\"100%\" height=\"100%\">";
  str += "<tr><td align=\"center\" colspan=\"16\" style=\"height:20px;border:1px solid buttonface;font-family: verdana; font-size:12px;\">";
  str += "<a style=\"cursor:hand;\" onClick=\"javascript:parent.ToggleColor(Named_" + inID + ", Gray_" + inID + ", Safe_" + inID + ");\">Named</a> - ";
  str += "<font color=\"#FF0000\">GrayScale</font> - ";
  str += "<a style=\"cursor:hand;\" onClick=\"javascript:parent.ToggleColor(Safe_" + inID + ", Named_" + inID + ", Gray_" + inID + ");\">Safety</a>";
  str += "</td></tr>";
  str += "<tr><td colspan=\"16\" id=\"GrayColor_" + inID + "\" style=\"height:20px;font-family: verdana; font-size:12px;\">&nbsp;</td></tr>";
  str += "<tr><td colspan=\"16\" id=\"GrayName_" + inID + "\" style=\"height:20px;border: 1px solid buttonface; font-family: verdana; font-size:10px;\" align=\"left\">&nbsp;</td></tr>";
  str += "<tr><td onMouseOver=\"parent.ShowColor(GrayColor_" + inID + ", GrayName_" + inID + ", this, '');\" colspan=\"16\" style=\"height:20px;font-family: verdana; font-size:10px;cursor: hand;\" align=\"center\" onClick=\"parent.DoColor('');\">Aucune</td></tr>";
  str += "<tr>";
  for (i = 0; i < aHex.length; i++) {
   for (j = 0; j < aHex.length; j++) {
    str += "<td style=\"height:12px;width:12px;cursor: hand;background-color:'#" + aHex[i] + aHex[j] + aHex[i] + aHex[j] + aHex[i] + aHex[j] + "'; font-size: 0px;\"";
    str += " onMouseOver=\"parent.ShowColor(GrayColor_" + inID + ", GrayName_" + inID + ", this, '');\"";
    str += " onClick=\"parent.DoColor('#" + aHex[i] + aHex[j] + aHex[i] + aHex[j] + aHex[i] + aHex[j] + "');\">" + String.fromCharCode(160) + "</td>";
   };
   if (i < aHex.length - 1) {
     str += "</tr>";
     str += "<tr>";
   }
  };
  str += "</tr>";
  str += "</table>";
  aHex = new Array("00","33","66","99","CC","FF");
  tdDisp = 0;
  totDisp = 0;
  str += "<table id=\"Safe_" + inID + "\" cellpadding=\"1\" cellspacing=\"2\" border=\"1\" bordercolor=\"#666666\" style=\"font-family: Verdana; font-size: 7px; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid; table-layout; fixed;display: none;\" bgcolor=\"buttonface\" width=\"100%\" height=\"100%\">";
  str += "<tr><td align=\"center\" colspan=\"16\" style=\"height:20px;border:1px solid buttonface;font-family: verdana; font-size:12px;\">";
  str += "<a style=\"cursor:hand;\" onClick=\"javascript:parent.ToggleColor(Named_" + inID + ", Gray_" + inID + ", Safe_" + inID + ");\">Named</a> - ";
  str += "<a style=\"cursor:hand;\" onClick=\"javascript:parent.ToggleColor(Gray_" + inID + ", Named_" + inID + ", Safe_" + inID + ");\">GrayScale</a> - ";
  str += "<font color=\"#FF0000\">Safety</font>";
  str += "</td></tr>";
  str += "<tr><td colspan=\"16\" id=\"SafeColor_" + inID + "\" style=\"height:20px;font-family: verdana; font-size:12px;\">&nbsp;</td></tr>";
  str += "<tr><td colspan=\"16\" id=\"SafeName_" + inID + "\" style=\"height:20px;border: 1px solid buttonface; font-family: verdana; font-size:10px;\" align=\"left\">&nbsp;</td></tr>";
  str += "<tr><td colspan=\"16\" onMouseOver=\"parent.ShowColor(SafeColor_" + inID + ", SafeName_" + inID + ", this, '');\" style=\"height:20px;font-family: verdana; font-size:10px;cursor: hand;\" align=\"center\" onClick=\"parent.DoColor('');\">Aucune</td></tr>";
  str += "<tr>";
  for (i = 0; i < aHex.length; i++) {
   for (j = 0; j < aHex.length; j++) {
    for (k = 0; k < aHex.length; k++) {
     if (totDisp % 16 == 0 && totDisp != 0) {
      str += "</tr>";
      str += "<tr>";
      tdDisp = 0;
     }
     str += "<td style=\"height:12px;width:12px;cursor: hand;background-color:'#" + aHex[i] + aHex[j] + aHex[k] + "'; font-size: 0px;\"";
     str += " onMouseOver=\"parent.ShowColor(SafeColor_" + inID + ", SafeName_" + inID + ", this, '');\"";
     str += " onClick=\"parent.DoColor('#" + aHex[i] + aHex[j] + aHex[k] + "');\">" + String.fromCharCode(160) + "</td>";
     tdDisp++;
     totDisp++;
    };
   };
  };
  for (i = tdDisp; i < 16; i++) {
   str += "<td style=\"height:12px;width:12px;border: 1px solid buttonface; background-color:buttonface; font-size: 0px;\"";
   str += " onMouseOver=\"parent.ShowColor(SafeColor_" + inID + ", SafeName_" + inID + ", this, '');\">" + String.fromCharCode(160) + "</td>";
   totDisp++;
  };
  str += "</tr>";
  str += "<tr>";
  fDisp = totDisp;
  for (i = totDisp; i < (16 * 16); i++) {
   if (i % 16 == 0 && i != fDisp) {
    str += "</tr>";
    str += "<tr>";
   }
   str += "<td style=\"height:12px;width:12px;border: 1px solid buttonface; background-color:buttonface; font-size: 0px;\"";
   str += " onMouseOver=\"parent.ShowColor(SafeColor_" + inID + ", SafeName_" + inID + ", this, '');\">" + String.fromCharCode(160) + "</td>";
  };
  str += "</tr>";
  str += "</table>";
  str += "</DIV>";
  return str;
 };
 this.GetMenuTable = function()
 {
  var str = "", i = 0;
  str += "<DIV ID=\"TableMenu_" + inID + "\" STYLE=\"display:none;\">";
  str += "<table border=\"0\" cellspacing=\"0\" cellpadding=\"0\" width=\"200\" style=\"BORDER-LEFT: buttonhighlight 1px solid;BORDER-TOP: buttonhighlight 1px solid;BORDER-RIGHT: buttonshadow 1px solid;BORDER-BOTTOM: buttonshadow 1px solid;\" bgcolor=\"buttonface\">";
  for (i = 0; i < HE_TABLE.length; i++) {
   if (HE_TABLE[i][0] == "SEP") {
    str += "<tr height=\"10\">";
    str += "<td align=\"center\"><img src=\"" + HE_IMGPATH + HE_TABLE[i][1] + "\" width=\"180\" height=\"2\"></td>";
    str += "</tr>";
   } else {
    str += "<tr title=\"" + HE_TABLE[i][2] + "\" onClick=\"parent." + HE_TABLE[i][3] + "\" id=\"TR_" + HE_TABLE[i][4] + "_" + inID + "\">";
    str += "<td nowrap id=\"TD_" + HE_TABLE[i][4] + "_" + inID + "\" onMouseOver=\"parent.BouttonOver(this);\" onMouseOut=\"parent.BouttonOut(this);\" onMouseDown=\"parent.BouttonDown(this);\"";
    str += " style=\"BORDER-LEFT: buttonface 1px solid;BORDER-RIGHT: buttonface 1px solid;BORDER-TOP: buttonface 1px solid;BORDER-BOTTOM: buttonface 1px solid;font-family: Arial, Verdana;font-size:12px;cursor:hand;\">";
    str += "<img src=\"" + HE_IMGPATH + HE_TABLE[i][1] + "\" id=\"IMG_" + HE_TABLE[i][4] + "_" + inID + "\" width=\"21\" height=\"20\" align=\"absmiddle\" border=\"0\">&nbsp;" + HE_TABLE[i][2] + "&nbsp;</td></tr>";
   }
  };
  str += "</table>";
  str += "</DIV>";
  return str;
 };
 this.GetMenuContext = function()
 {
  var str = "", i = 0, j = 0, oClick = "";
  for (i = 0; i < HE_CONTEXT.length; i++) {
   str += "<DIV ID=\"" + HE_CONTEXT[i][0] + "_" + inID + "\" STYLE=\"display:none;\">";
   str += "<table border=\"0\" cellspacing=\"0\" cellpadding=\"2\" width=\"180\" style=\"BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 1px solid; BORDER-TOP: buttonhighlight 1px solid; BORDER-BOTTOM: buttonshadow 1px solid;\" bgcolor=\"buttonface\">";
   for (j = 0; j < HE_CONTEXT[i][1].length; j++) {
    if (HE_CONTEXT[i][1][j][1].indexOf("§inID§") >= 0) {
     oClick = HE_CONTEXT[i][1][j][1].replace("§inID§", inID);
    } else {
     oClick = HE_CONTEXT[i][1][j][1];
    }
    str += "<tr onClick=\"" + oClick + ";parent.JS_HE.oPopup.hide();\">";
    str += "<td style=\"cursor: default;BORDER-LEFT: buttonface 1px solid; BORDER-RIGHT: buttonface 1px solid; BORDER-TOP: buttonface 1px solid; BORDER-BOTTOM: buttonface 1px solid;font-family: Arial, Verdana;font-size:11px;\""
    str += " onMouseOver=\"parent.ContextOver(this);\" onMouseOut=\"parent.ContextOut(this);\">";
    str += "&nbsp&nbsp;&nbsp;&nbsp&nbsp;" + HE_CONTEXT[i][1][j][0] + "&nbsp;</td>"
    str += "</tr>";
   };
   str += "</table>";
   str += "</DIV>";
  };
  return str;
 };
 this.Init = function()
 {
  var str = "", sMenuColor = "";
  String.prototype.RemoveAll = function(inTOK, inREP) {
   var sOut = this;
   while (sOut.indexOf(inTOK) >= 0) {
    sOut = sOut.replace(inTOK, inREP);
   };
   return sOut;
  };
  sMenuColor = str = this.GetMenuColors();
  str = str.RemoveAll("parent.ToggleColor", "parent.window.opener.ToggleColor");
  str = str.RemoveAll("parent.ShowColor", "parent.window.opener.ShowColor");
  str = str.RemoveAll("parent.DoColor", "parent.window.opener.DoColor");
  str = str.RemoveAll("ColorMenu_" + inID, "ColorMenuOpener_" + inID);
  this.gStr = "";
  this.gStr += this.GetMenuChars();
  this.gStr += sMenuColor;
  this.gStr += str;
  this.gStr += this.GetMenuTable();
  this.gStr += this.GetMenuContext();
 };
 this.Init();
};