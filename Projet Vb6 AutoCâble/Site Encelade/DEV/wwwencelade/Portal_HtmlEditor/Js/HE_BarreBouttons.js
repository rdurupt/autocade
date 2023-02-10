var Barre_Source =
[
 ["IMG", "Bouttons/he_cut.gif", "Couper (Ctrl+X)", "DoCommand('Cut');"],
 ["IMG", "Bouttons/he_copy.gif", "Copier (Ctrl+C)", "DoCommand('Copy');"],
 ["IMG", "Bouttons/he_paste.gif", "Coller (Ctrl+V)", "DoCommand('Paste');"],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_undo.gif", "Annuler la frappe (Ctrl+Z)", "DoCommand('Undo');"],
 ["IMG", "Bouttons/he_redo.gif", "Rétablir la frappe (Ctrl+Y)", "DoCommand('Redo');"]
];

var Barre_EditTop =
[
 ["IMG", "Bouttons/he_cut.gif", "Couper (Ctrl+X)", "DoCommand('Cut');", ""],
 ["IMG", "Bouttons/he_copy.gif", "Copier (Ctrl+C)", "DoCommand('Copy');", ""],
 ["IMG", "Bouttons/he_paste.gif", "Coller (Ctrl+V)", "DoCommand('Paste');", ""],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_undo.gif", "Annuler la frappe (Ctrl+Z)", "DoCommand('Undo');", ""],
 ["IMG", "Bouttons/he_redo.gif", "Rétablir la frappe (Ctrl+Y)", "DoCommand('Redo');", ""],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_bold.gif", "Gras (Ctrl+B)", "DoCommand('Bold');", "Bold"],
 ["IMG", "Bouttons/he_underline.gif", "Souligné (Ctrl+U)", "DoCommand('Underline');", "Underline"],
 ["IMG", "Bouttons/he_italic.gif", "Italique (Ctrl+I)", "DoCommand('Italic');", "Italic"],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_ul_nombres.gif", "Insérer une liste de nombres", "DoCommand('InsertOrderedList');", "InsertOrderedList"],
 ["IMG", "Bouttons/he_ul.gif", "Insérer une liste", "DoCommand('InsertUnorderedList');", "InsertUnorderedList"],
 ["IMG", "Bouttons/he_outdent.gif", "Diminuer l'indentation", "DoCommand('Outdent');", ""],
 ["IMG", "Bouttons/he_indent.gif", "Augmenter l'indentation", "DoCommand('Indent');", ""],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_left.gif", "Aligner à gauche", "DoCommand('JustifyLeft');", "JustifyLeft"],
 ["IMG", "Bouttons/he_center.gif", "Aligner au centre", "DoCommand('JustifyCenter');", "JustifyCenter"],
 ["IMG", "Bouttons/he_right.gif", "Aligner à droite", "DoCommand('JustifyRight');", "JustifyRight"],
 ["IMG", "Bouttons/he_justify.gif", "Justifier", "DoCommand('JustifyFull');", "JustifyFull"],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_help.gif", "Aide", ""]
];

var Barre_EditBot =
[
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_cfonts.gif", "Couleur de la police", "DoFontColor();"],
 ["IMG", "Bouttons/he_cback.gif", "Couleur de fond", "DoBackColor();"],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_hr.gif", "Insérer une ligne horizontale", "DoCommand('InsertHorizontalRule');", ""],
 ["IMG", "Bouttons/he_link.gif", "Insérer / Modifier un lien", "DoLink();", ""],
 ["IMG", "Bouttons/he_anchor.gif", "Insérer / Modifier une ancre", "DoAnchor();", ""],
 ["IMG", "Bouttons/he_email.gif", "Insérer / Modifier un email", "DoEmail();", ""],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_itable.gif", "Tableau", "DoTable();"],
 ["SEP", "Bouttons/he_vert.gif", "", ""],
 ["IMG", "Bouttons/he_image.gif", "Insérer / Modifier une image", "DoImage();"],
 ["IMG", "Bouttons/he_attach.gif", "Insérer un lien vers un document", "DoFiles();"],
 ["IMG", "Bouttons/he_chars.gif", "Insérer un caractère", "DoChars();"],
 ["IMG", "Bouttons/he_code.gif", "Nettoyer le code", "DoCleanCode();"],
 ["IMG", "Bouttons/he_borders.gif", "Afficher / Masquer les bordures", "ToggleBorders();"]
];

function Html_BarreBouttons( inID, inMODE ) {
 if (inMODE != "body") {
  Barre_EditBot[Barre_EditBot.length] = ["IMG", "Bouttons/he_properties.gif", "Modifier les propriétés", "DoProperties();"];
  Barre_EditBot[Barre_EditBot.length] = ["IMG", "Bouttons/he_code.gif", "Nettoyer le code", "DoCleanCode();"];
 }
 this.Init = function()
 {
  var str = "", i = 0;
  str += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" class=\"BarreBouttons\" id=\"toolbar_" + inID + "\">";
  str += "<tr><td class=\"Text\" height=\"52\">";
  str += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" class=\"Off\" id=\"toolbar-preview_" + inID + "\">";
  str += "<tr id=\"he\"><td class=\"Text\" height=\"55\">";
  str += "&nbsp;&nbsp;&nbsp;<b>Prévisualisation</b>";
  str += "</td></tr></table>";
  str += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" class=\"Off\" id=\"toolbar-source_" + inID + "\">";
  str += "<tr><td class=\"Text\" height=\"22\">";
  str += "<table border=\"0\" cellspacing=\"0\" cellpadding=\"1\"><tr id=\"he\">";
  for (i = 0; i < Barre_Source.length; i++) {
   str += "<td>";
   if (Barre_Source[i][0].toUpperCase() == "IMG") {
    str += "<button unselectable=\"on\"";
    str += " class=\"Boutton\"";
    str += " onMouseOver=\"BouttonOver(this);\"";
    str += " onMouseOut=\"BouttonOut(this);\"";
    str += " onMouseDown=\"BouttonDown(this);\"";
    if (Barre_Source[i][3] != "") {
     str += " onClick=\"" + Barre_Source[i][3] + "UpdateGUI();\"";
    }
    str += "><img src=\"" + HE_IMGPATH + Barre_Source[i][1] + "\" border=\"0\"";
    str += " width=\"21\"";
    str += " height=\"20\"";
    str += " title=\"" + Barre_Source[i][2] + "\">";
    str += "</button>";
   } else if (Barre_Source[i][0].toUpperCase() == "SEP") {
    str += "<img src=\"" + HE_IMGPATH + Barre_Source[i][1] + "\" border=\"0\"";
    str += " width=\"2\"";
    str += " height=\"20\"";
    str += " class=\"Boutton\">";
   }    
   str += "</td>";
  };
  str += "</tr></table></td></tr><tr>";
  str += "<td width=\"100%\" height=\"1\" bgcolor=\"#000000\"></td>";
  str += "</tr><tr>";
  str += "<td width=\"100%\" height=\"30\" class=\"Text\">&nbsp;</td>";
  str += "</tr></table>";
  str += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" class=\"OnBouttons\" id=\"toolbar-edit_" + inID + "\">";
  str += "<tr><td class=\"Text\" height=\"22\">";
  str += "<table border=\"0\" cellspacing=\"0\" cellpadding=\"1\"><tr id=\"he\">";
  for (i = 0; i < Barre_EditTop.length; i++) {
   str += "<td>";
   if (Barre_EditTop[i][0].toUpperCase() == "IMG") {
    str += "<button unselectable=\"on\"";
    str += " class=\"Boutton\"";
    str += " onMouseOver=\"BouttonOver(this);\"";
    str += " onMouseOut=\"BouttonOut(this);\"";
    str += " onMouseDown=\"BouttonDown(this);\"";
    if (Barre_EditTop[i][3] != "") {
     str += " onClick=\"" + Barre_EditTop[i][3] + "UpdateGUI();\"";
    }
    if (Barre_EditTop[i][4] != "") {
     str += " id=\"b_" + Barre_EditTop[i][4] + "_" + inID + "\">";
    } else {
     str += ">";
    }
    str += "<img src=\"" + HE_IMGPATH + Barre_EditTop[i][1] + "\" border=\"0\"";
    str += " width=\"21\"";
    str += " height=\"20\"";
    str += " title=\"" + Barre_EditTop[i][2] + "\">";
    str += "</button>";
   } else if (Barre_EditTop[i][0].toUpperCase() == "SEP") {
    str += "<img src=\"" + HE_IMGPATH + Barre_EditTop[i][1] + "\" border=\"0\"";
    str += " width=\"2\"";
    str += " height=\"20\"";
    str += " class=\"Boutton\">";
   }    
   str += "</td>";
  };
  str += "</tr></table></td></tr><tr>";
  str += "<td width=\"100%\" height=\"1\" bgcolor=\"#000000\"></td>";
  str += "</tr><tr><td class=\"Text\" height=\"30\">";
  str += "<table border=\"0\" cellspacing=\"0\" cellpadding=\"1\"><tr id=\"he\">";
  str += "<td><select id=\"FontFamily_" + inID + "\" size=\"1\" unselectable=\"On\" class=\"Slt70px\" onChange=\"DoFontFamily(this);\">";
  str += "<option value=\"\" selected>Police";
  str += "<option value=\"Times New Roman\">Défaut";
  str += "<option value=\"Arial\">Arial";
  str += "<option value=\"Verdana\">Verdana";
  str += "<option value=\"Tahoma\">Tahoma";
  str += "<option value=\"Courier New\">Courier New";
  str += "<option value=\"Georgia\">Georgia";
  str += "</select></td>";
  str += "<td><select id=\"FontSize_" + inID + "\" size=\"1\" unselectable=\"On\" class=\"Slt60px\" onChange=\"DoFontSize(this);\">";
  str += "<option value=\"+0\" selected>Taille";
  str += "<option value=\"1\">1";
  str += "<option value=\"2\">2";
  str += "<option value=\"3\">3";
  str += "<option value=\"4\">4";
  str += "<option value=\"5\">5";
  str += "<option value=\"6\">6";
  str += "<option value=\"7\">7";
  str += "</select></td>";
  str += "<td><select id=\"Modeles_" + inID + "\" size=\"1\" unselectable=\"On\" class=\"Slt90px\" onChange=\"DoModele(this);\">";
  str += "</select></td>";
  for (i = 0; i < Barre_EditBot.length; i++) {
   str += "<td>";
   if (Barre_EditBot[i][0].toUpperCase() == "IMG") {
    str += "<button unselectable=\"on\"";
    str += " class=\"Boutton\"";
    str += " onMouseOver=\"BouttonOver(this);\"";
    str += " onMouseOut=\"BouttonOut(this);\"";
    str += " onMouseDown=\"BouttonDown(this);\"";
    if (Barre_EditBot[i][3] != "") {
     str += " onClick=\"" + Barre_EditBot[i][3] + "UpdateGUI();\"";
    }
    if (Barre_EditBot[i][4] != "") {
     str += " id=\"b_" + Barre_EditBot[i][4] + "_" + inID + "\">";
    } else {
     str += ">";
    }
    str += "<img src=\"" + HE_IMGPATH + Barre_EditBot[i][1] + "\" border=\"0\"";
    str += " width=\"21\"";
    str += " height=\"20\"";
    str += " title=\"" + Barre_EditBot[i][2] + "\">";
    str += "</button>";
   } else if (Barre_EditBot[i][0].toUpperCase() == "SEP") {
    str += "<img src=\"" + HE_IMGPATH + Barre_EditBot[i][1] + "\" border=\"0\"";
    str += " width=\"2\"";
    str += " height=\"20\"";
    str += " class=\"Boutton\">";
   }    
   str += "</td>";
  };
  str += "</tr></table>";
  str += "</td></tr></table>";
  str += "</td></tr></table>";
  return str;
 };
};