var CM =
[
  [
    "ALIGN", "Alignement&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC=\"\" ID=\"CM_ALIGN_IMG\" onselectstart=\"return false;\" ondragstart=\"return false;\" ALIGN=\"absmiddle\">",
      [
        ["<IMG SRC=\"" + LM_IMGPATH + "lm_left.gif\" onselectstart=\"return false;\" ondragstart=\"return false;\" ALIGN=\"absmiddle\">&nbsp;&nbsp;Aligner à gauche", "parent.DivAlign('left'"],
        ["<IMG SRC=\"" + LM_IMGPATH + "lm_center.gif\" onselectstart=\"return false;\" ondragstart=\"return false;\" ALIGN=\"absmiddle\">&nbsp;&nbsp;Aligner au centre", "parent.DivAlign('center'"],
        ["<IMG SRC=\"" + LM_IMGPATH + "lm_right.gif\" onselectstart=\"return false;\" ondragstart=\"return false;\" ALIGN=\"absmiddle\">&nbsp;&nbsp;Aligner à droite", "parent.DivAlign('right'"]
      ]
  ],
  [
    "HRSTYLE", "Style&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC=\"\" ID=\"CM_HRSTYLE\" onselectstart=\"return false;\" ondragstart=\"return false;\" ALIGN=\"absmiddle\">",
      [
        ["<CENTER>Aucun</CENTER>", "parent.HrStyle('none'"],
        ["<HR STYLE=\"WIDTH:100%;COLOR:BUTTONSHADOW;BORDER:SOLID;\">", "parent.HrStyle('solid'"],
        ["<HR STYLE=\"WIDTH:100%;COLOR:BUTTONSHADOW;BORDER:DOTTED;\">", "parent.HrStyle('dotted'"],
        ["<HR STYLE=\"WIDTH:100%;COLOR:BUTTONSHADOW;BORDER:DASHED;\">", "parent.HrStyle('dashed'"]
      ]
  ],
  [
    "HRCOLOR", "Couleur&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SPAN ID=\"CM_HRCOLOR\"></SPAN>",
      [
        ["<CENTER>Aucune</CENTER>", "parent.HrColor('none'"],
        ["<DIV STYLE=\"BACKGROUND-COLOR:#000000;COLOR:#FFFFFF;TEXT-ALIGN:CENTER;\">Noir</DIV>", "parent.HrColor('black'"],
        ["<IMG SRC=\"" + LM_IMGPATH + "lm_color.gif\" onselectstart=\"return false;\" ondragstart=\"return false;\" ALIGN=\"absmiddle\">&nbsp;&nbsp;Choisir ...", "parent.HrColor('choose'"]
      ]
  ],
  [
    "HRSIZE", "Taille&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SPAN ID=\"CM_HRSIZE\"></SPAN>",
      [
        ["<CENTER>10%</CENTER>", "parent.HrSize('10%'"],
        ["<CENTER>20%</CENTER>", "parent.HrSize('20%'"],
        ["<CENTER>30%</CENTER>", "parent.HrSize('30%'"],
        ["<CENTER>40%</CENTER>", "parent.HrSize('40%'"],
        ["<CENTER>50%</CENTER>", "parent.HrSize('50%'"],
        ["<CENTER>60%</CENTER>", "parent.HrSize('60%'"],
        ["<CENTER>70%</CENTER>", "parent.HrSize('70%'"],
        ["<CENTER>80%</CENTER>", "parent.HrSize('80%'"],
        ["<CENTER>90%</CENTER>", "parent.HrSize('90%'"],
        ["<CENTER>100%</CENTER>", "parent.HrSize('100%'"]
      ]
  ]
];

function DivAlign( inAlign, inID ) {
  LMArray[inID].cm.DivAlign(inAlign);
};

function HrStyle( inStyle, inID ) {
  LMArray[inID].cm.HrStyle(inStyle);
};

function HrColor( inColor, inID ) {
  LMArray[inID].cm.HrColor(inColor);
};

function HrSize( inSize, inID ) {
  LMArray[inID].cm.HrSize(inSize);
};

function ShowPopupAttached(inPopup, inID) {
  LMArray[inID].cm.ShowPopupAttached(inPopup);
};

function Layout_ContextMenu( inLM ) {
  this.lm = inLM;
  this.oPopup = window.createPopup();
  this.oPopupAttached = window.createPopup();
  this.sDiv = "";
  this.aDiv = new Array();
  this.Init = function()
  {
    var i, j, oDiv, iDiv;
    this.gDiv = document.createElement("<DIV ID=\"CM_" + this.lm.uID + "\">");
    this.aDiv[this.aDiv.length] = "CM_" + this.lm.uID;
    this.gDiv.style.display = "none";
    this.gDiv.oncontextmenu = "return false";
    this.gDiv.ondragstart = "return false";
    this.gDiv.onselectstart = "return false";
    for (i = 0; i < CM.length; i++) {
      oDiv = document.createElement("<DIV ID=\"CM_" + CM[i][0] + "_ELE_" + this.lm.uID + "\">");
      this.aDiv[this.aDiv.length] = "CM_" + CM[i][0] + "_ELE_" + this.lm.uID;
      oDiv.style.visibility = "hidden";
      oDiv.style.position = "relative";
      oDiv.style.top = 0;
      oDiv.style.left = 0;
      oDiv.style.backgroundColor = "buttonface";
      oDiv.style.border = "1 solid buttonshadow";
      oDiv.style.borderTop = "1 solid buttonshadow";
      oDiv.style.borderLeft = "1 solid buttonshadow";
      oDiv.oncontextmenu = "return false";
      oDiv.onselectstart = "return false";
      oDiv.ondragstart = "return false";
      this.gDiv.appendChild(oDiv);
      iDiv = document.createElement("DIV");
      iDiv.style.position = "relative";
      iDiv.style.top = 0;
      iDiv.style.left = 0;
      iDiv.style.backgroundColor = "buttonface";
      iDiv.style.border = "1 solid buttonshadow";
      iDiv.style.borderTop = "1 solid buttonhighlight";
      iDiv.style.borderLeft = "1 solid buttonhighlight";
      iDiv.style.height = 18;
      iDiv.style.color = "#000000";
      iDiv.style.fontFamily = "Arial";
      iDiv.style.fontSize = "x-small";
      iDiv.style.cursor = "hand";
      iDiv.style.padding = 2;
      iDiv.style.paddingLeft = 5;
      iDiv.style.paddingRight = 5;
      iDiv.onmouseover = "this.style.background='buttonhighlight'";
      iDiv.onmouseout = "this.style.background='buttonface'";
      iDiv.oncontextmenu = "return false";
      iDiv.ondragstart = "return false";
      iDiv.onselectstart = "return false";
      iDiv.innerHTML = CM[i][1];
      iDiv.onclick = "parent.ShowPopupAttached('" + CM[i][0] + "', '" + this.lm.uID + "');";
      oDiv.appendChild(iDiv);
      oDiv = document.createElement("<DIV ID=\"CM_" + CM[i][0] + "_ATT_" + this.lm.uID + "\">");
      this.aDiv[this.aDiv.length] = "CM_" + CM[i][0] + "_ATT_" + this.lm.uID;
      oDiv.style.visibility = "hidden";
      oDiv.style.position = "relative";
      oDiv.style.top = 0;
      oDiv.style.left = 0;
      oDiv.style.backgroundColor = "buttonface";
      oDiv.style.border = "1 solid buttonshadow";
      oDiv.style.borderTop = "1 solid buttonshadow";
      oDiv.style.borderLeft = "1 solid buttonshadow";
      oDiv.oncontextmenu = "return false";
      oDiv.onselectstart = "return false";
      oDiv.ondragstart = "return false";
      this.gDiv.appendChild(oDiv);
      for (j = 0; j < CM[i][2].length; j++) {
        iDiv = document.createElement("DIV");
        iDiv.style.position = "relative";
        iDiv.style.top = 0;
        iDiv.style.left = 0;
        iDiv.style.backgroundColor = "buttonface";
        iDiv.style.border = "1 solid buttonshadow";
        iDiv.style.borderTop = "1 solid buttonhighlight";
        iDiv.style.borderLeft = "1 solid buttonhighlight";
        iDiv.style.height = 18;
        iDiv.style.color = "#000000";
        iDiv.style.fontFamily = "Arial";
        iDiv.style.fontSize = "x-small";
        iDiv.style.cursor = "hand";
        iDiv.style.padding = 2;
        iDiv.style.paddingLeft = 5;
        iDiv.style.paddingRight = 5;
        iDiv.onmouseover = "this.style.background='buttonhighlight'";
        iDiv.onmouseout = "this.style.background='buttonface'";
        iDiv.oncontextmenu = "return false";
        iDiv.ondragstart = "return false";
        iDiv.onselectstart = "return false";
        iDiv.innerHTML = CM[i][2][j][0];
        iDiv.onclick = CM[i][2][j][1] + ", '" + this.lm.uID + "');";
        oDiv.appendChild(iDiv);
      };
    };
  };
  this.DivAlign = function( inAlign )
  {
    this.sDiv.align = inAlign;
    this.sDiv = null;
    this.oPopupAttached.hide();
  };
  this.HrStyle = function( inStyle )
  {
    var str = "", hrW = "", hrC = "", hrB = "";
    hrW = "WIDTH:" + this.sDiv.childNodes[0].style.width + ";";
    if (this.sDiv.childNodes[0].style.color + "" != "") {
      hrC = "COLOR:" + this.sDiv.childNodes[0].style.color + ";";
    } else {
      hrC = "COLOR:buttonshadow;";
    }
    if (inStyle != "none") {
      hrB = "BORDER:" + inStyle + ";";
    } else {
      if (hrC == "COLOR:buttonshadow;") {
        hrC = "";
      }
    }
    str = "<HR STYLE=\"" + hrW + hrC + hrB + "\">";
    this.sDiv.innerHTML = str;
    this.sDiv = null;
    this.oPopupAttached.hide();
  };
  this.HrColor = function ( inColor )
  {
    var str = "", hrW = "", hrC = "", hrB = "", sColor;
    hrW = "WIDTH:" + this.sDiv.childNodes[0].style.width + ";";
    hrC = this.sDiv.childNodes[0].style.color;
    hrB = "BORDER:" + this.sDiv.childNodes[0].style.border + ";";
    if (inColor == "choose") {
      if ((hrC + "" == "") || (hrC == "buttonshadow")) {
        sColor = document.getElementById("LM_COLOR").ChooseColorDlg();
      } else {
        sColor = document.getElementById("LM_COLOR").ChooseColorDlg(hrC);
      }
      if (sColor + "" != "0" || hrC == "#000000") {
        sColor = sColor.toString(16);
        if (sColor.length < 6) {
          hrC = "000000".substring(0, 6 - sColor.length);
          sColor = hrC.concat(sColor);
        }
        hrC = "COLOR:#" + sColor + ";";
      } else {
        if (hrB + "" != "BORDER:;") {
          hrC = "COLOR:buttonshadow;";
        } else {
          hrB = "";
          hrC = "";
        }
      }
    } else if (inColor == "none") {
      if (hrB + "" != "BORDER:;") {
        hrC = "COLOR:buttonshadow;";
      } else {
        hrB = "";
        hrC = "";
      }
    } else if (inColor == "black") {
      hrC = "COLOR:#000000;";
    }
    str = "<HR STYLE=\"" + hrW + hrC + hrB + "\">";
    this.sDiv.innerHTML = str;
    this.sDiv = null;
    this.oPopupAttached.hide();
  };
  this.HrSize = function( inSize )
  {
    var str = "", hrW = "", hrC = "", hrB = "";
    hrW = "WIDTH:" + inSize + ";";
    hrC = "COLOR:" + this.sDiv.childNodes[0].style.color + ";";
    hrB = "BORDER:" + this.sDiv.childNodes[0].style.border + ";";
    str = "<HR STYLE=\"" + hrW + hrC + hrB + "\">";
    this.sDiv.innerHTML = str;
    this.sDiv = null;
    this.oPopupAttached.hide();
  };
  this.ShowPopupAttached = function( inPopup )
  {
    var i, oMenu, oOld, height = 0, oClone;
    for (i = 0; i < this.aDiv.length; i++) {
      oMenu = document.getElementById(this.aDiv[i]);
      oMenu.style.display = "none";
      oMenu.style.visibility = "hidden";
    };
    oClone = document.getElementById("CM_" + this.lm.uID);
    oOld = document.getElementById("CM_" + inPopup + "_ELE_" + this.lm.uID);
    oMenu = document.getElementById("CM_" + inPopup + "_ATT_" + this.lm.uID);
    height += (oMenu.childNodes.length * 24) + 2;
    oOld.style.visibility = "hidden";
    oOld.style.display = "none";
    oMenu.style.visibility = "visible";
    oMenu.style.display = "";
    this.oPopupAttached.document.body.innerHTML = oClone.innerHTML;
    this.oPopupAttached.show(this.lm.oLeft, this.lm.oTop, 140, height);
  };
  this.Init();
};