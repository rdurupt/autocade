// Portal_Calendar v1.0 - A.VERLA
//
// v1.0 :           18/01/2003
//   - Portal_Calendar : v1.2
//   - DC_Calendar     : v1.0
//   - DC_Language     : v1.0
//   - DC_Format       : v1.0

var DCArray = new Array();

Portal_Calendar = function()
{

 this.uID = document.uniqueID;
 DCArray[this.uID] = this;
 this.gStr = "";
 this.dc = new Array();
 this.d = new Date();

 this.GetObject = function( inNAME )
 {
  var i = 0, oObj = null;
  for (i = 0; i < this.dc.length; i++) {
   if (this.dc[i].name == inNAME) {
    oObj = this.dc[i];
    break;
   }
  };
  return oObj;
 };
 this.DeleteAll = function()
 {
  document.getElementById(this.uID + "_a").innerHTML = "&nbsp;";
  document.getElementById(this.uID + "_m").innerHTML = "&nbsp;";
  for (i = 0; i < 42; i++) {
   document.getElementById(this.uID + "_d" + i).innerHTML = "&nbsp;";
  };
 };
 this.IsLeap = function( inYEAR )
 {
  if (inYEAR % 400 == 0) {
   return true;
  } else if ((inYEAR % 4 == 0) && (inYEAR % 100 != 0)){
   return true;
  } else {
   return false;
  }
 };
 this.DaysIn = function( inMONTH, inYEAR )
 {
  var m = 0;
  if (("§0§§2§§4§§6§§7§§9§§11§").indexOf("§" + inMONTH + "§") >= 0) {
   m = 31;
  } else if (("§3§§5§§8§§10§").indexOf("§" + inMONTH + "§") >= 0) {
   m = 30;
  } else {
   if (this.IsLeap(inYEAR)) {
    m = 29;
   } else {
    m = 28;
   }
  }
  return m;
 };

 this.SetColor = function( inNAME, inTYPE, inCOLOR )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   if (inTYPE.toUpperCase() == "BACKGROUND") {
    oObj.cBg = inCOLOR;
   } else if (inTYPE.toUpperCase() == "BORDER") {
    oObj.cBo = inCOLOR;
   }
  }
 };
 this.SetDateFormat = function( inNAME, inFMT )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   oObj.f = inFMT;
  }
 };
 this.SetImage = function( inNAME, inTYPE, inIMG )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   if (inTYPE.toUpperCase() == "NEXT") {
    oObj.next = inIMG;
   } else if (inTYPE.toUpperCase() == "PREV") {
    oObj.prev = inIMG;
   } else if (inTYPE.toUpperCase() == "CLOSE") {
    oObj.close = inIMG;
   }
  }
 };
 this.SetFirstDayOfWeek = function( inNAME, inDAY )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   oObj.fdw = inDAY;
  }
 };
 this.SetLanguage = function( inNAME, inLNG )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   oObj.lng = inLNG;
  }
 };
 this.SetPosition = function( inNAME, inTOP, inLEFT )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   oObj.top = inTOP;
   oObj.left = inLEFT;
  }
 };
 this.Show = function( inNAME )
 {
  var oObj = this.GetObject(inNAME), i = 0, oTd = null, indx = 0;
  if (oObj) {
   if (oObj.vInit == 0) {
    this.InitCal(oObj.name);
   }
   dOK = true;
   if (oObj.form.value + "" != "") {
    dOK = this.CheckDate(oObj.form.value, oObj.fmt);
   } else {
    this.dDay = this.d.getDate();
    this.dMonth = this.d.getMonth();
    this.dYear = this.d.getFullYear();
    this.sDay = "";
    this.sMonth = "";
    this.sYear = "";
   }
   if (dOK) {
    this.DeleteAll();
    for (i = 0; i < oObj.l.a[1].length; i++) {
     oTd = document.getElementById(this.uID + "_s" + i);
     indx = i + oObj.fdw;
     if (indx > 6) {
      indx = indx - 7;
     }
     oTd.innerHTML = "<b>" + oObj.l.a[1][indx] + "</b>";
    };
    oTd = document.getElementById(this.uID + "_p");
    if (oObj.prev + "" == "") {
     oTd.innerHTML = "<a href=\"javascript:DCArray['" + this.uID + "'].Previous('" + oObj.name + "');\" title=\"" + oObj.l.a[3][0] + "\" onMouseOver=\"window.status='" + oObj.l.a[3][0] + "';return true;\" onMouseOut=\"window.status=' ';return true;\">&lt;</a>";
    } else {
     oTd.innerHTML = "<a href=\"javascript:DCArray['" + this.uID + "'].Previous('" + oObj.name + "');\" title=\"" + oObj.l.a[3][0] + "\" onMouseOver=\"window.status='" + oObj.l.a[3][0] + "';return true;\" onMouseOut=\"window.status=' ';return true;\">" + oObj.prev + "</a>";
    }
    oTd = document.getElementById(this.uID + "_n");
    if (oObj.next + "" == "") {
     oTd.innerHTML = "<a href=\"javascript:DCArray['" + this.uID + "'].Next('" + oObj.name + "');\" title=\"" + oObj.l.a[3][1] + "\" onMouseOver=\"window.status='" + oObj.l.a[3][1] + "';return true;\" onMouseOut=\"window.status=' ';return true;\">&gt;</a>";
    } else {
     oTd.innerHTML = "<a href=\"javascript:DCArray['" + this.uID + "'].Next('" + oObj.name + "');\" title=\"" + oObj.l.a[3][1] + "\" onMouseOver=\"window.status='" + oObj.l.a[3][1] + "';return true;\" onMouseOut=\"window.status=' ';return true;\">" + oObj.next + "</a>";
    }
    oTd = document.getElementById(this.uID + "_c");
    if (oObj.close + "" == "") {
     oTd.innerHTML = "<a href=\"javascript:DCArray['" + this.uID + "'].Hide('" + oObj.name + "');\" title=\"" + oObj.l.a[3][3] + "\" onMouseOver=\"window.status='" + oObj.l.a[3][3] + "';return true;\" onMouseOut=\"window.status=' ';return true;\"><b>X</b></a>";
    } else {
     oTd.innerHTML = "<a href=\"javascript:DCArray['" + this.uID + "'].Hide('" + oObj.name + "');\" title=\"" + oObj.l.a[3][3] + "\" onMouseOver=\"window.status='" + oObj.l.a[3][3] + "';return true;\" onMouseOut=\"window.status=' ';return true;\">" + oObj.close + "</a>";
    }
    this.GetCal(oObj);
    var gTbl = document.getElementById("TBL_" + this.uID);
    gTbl.style.backgroundColor = oObj.cBg;
    gTbl.style.borderTop = "1px solid " + oObj.cBo;
    gTbl.style.borderBottom = "1px solid " + oObj.cBo;
    gTbl.style.borderLeft = "1px solid " + oObj.cBo;
    gTbl.style.borderRight = "1px solid " + oObj.cBo;
    gTbl.style.position = "absolute";
    gTbl.style.top = oObj.top;
    gTbl.style.left = oObj.left;
    gTbl.style.visibility = "visible";
    gTbl.style.display = "";
   } else {
    oObj.form.value = "";
    alert(oObj.l.a[4]);
   }
  }
 };
 this.IsValidDate = function( inNAME )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   if (oObj.form.value + "" != "") {
    if (this.CheckDate(oObj.form.value, oObj.fmt)) {
     return true;
    } else {
     oObj.form.value = "";
     alert(oObj.l.a[4]);
     return false;
    }
   }
  }
 };
 this.GetStatus = function( inNAME )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   return oObj.l.a[3][2];
  }
 };
 this.CheckDate = function( inDATE, inFMT )
 {
  var myArrayDate, myDay, myMonth, myYear, myString, myYearDigit;
  myString = inDATE;
  myArrayDate = myString.split(inFMT.sDel);
  myDay = Math.round(parseFloat(myArrayDate[inFMT.sFmt[1]]));
  myMonth = Math.round(parseFloat(myArrayDate[inFMT.sFmt[2]])) - 1;
  myYear = Math.round(parseFloat(myArrayDate[inFMT.sFmt[3]]));
  myString = myYear + "";
  myYearDigit = myString.length;
  if (isNaN(myDay) || isNaN(myMonth) || isNaN(myYear) || (myYear < 1) || (myDay < 1) || (myMonth < 0) || (myMonth > 11) || (myYearDigit != 4) || (myDay > this.DaysIn(myMonth, myYear))){
   return false;
  } else{
   if (myMonth == 1) {
    if (!this.IsLeap(myYear) && myDay == 29) {
     return false;
    } else {
     this.dDay = this.sDay = myDay;
     this.dMonth = this.sMonth = myMonth;
     this.dYear = this.sYear = myYear;
     return true;
    }
   } else {
    this.dDay = this.sDay = myDay;
    this.dMonth = this.sMonth = myMonth;
    this.dYear = this.sYear = myYear;
    return true;
   }
  }
 };
 this.Next = function( inNAME )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   this.dMonth++;
   if (this.dMonth > 11) {
    this.dYear++;
    this.dMonth = 0;
   }
   this.DeleteAll();
   this.GetCal(oObj);
  }
 }
 this.Previous = function( inNAME )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   this.dMonth--;
   if (this.dMonth < 0) {
    this.dYear--;
    this.dMonth = 11;
   }
   this.DeleteAll();
   this.GetCal(oObj);
  }
 };
 this.Hide = function()
 {
  this.DeleteAll();
  var gTbl = document.getElementById("TBL_" + this.uID);
  gTbl.style.visibility = "hidden";
  gTbl.style.display = "none";
 };
 this.GetCal = function( inOBJ )
 {
  var fdm = (new Date(this.dYear, this.dMonth, "01")).getDay();
  fdm = fdm - inOBJ.fdw;
  if (fdm < 0) {
   fdm = fdm + 7;
  }
  var indx = 1;
  var end = fdm + this.DaysIn(this.dMonth, this.dYear);
  for (var i = fdm; i < end; i++) {
   if (indx == this.sDay && this.dMonth == this.sMonth && this.dYear == this.sYear) {
    var css = " style=\"color:#FF0000;\"";
   } else {
    var css = "";
   }
   var dF = inOBJ.fmt.GetDateFormatted(indx, this.dMonth, this.dYear);
   document.getElementById(this.uID + "_d" + i).innerHTML = "<a href=\"javascript:DCArray['" + this.uID + "'].SetDateBack('" + inOBJ.name + "', '" + dF + "');\"" + css + " title=\"" + dF + "\" onMouseOver=\"window.status='" + dF + "';return true;\" onMouseOut=\"window.status=' ';return true;\">" + indx + "</a>";
   indx = indx + 1;
  };
  document.getElementById(this.uID + "_a").innerHTML = "<b>" + this.dYear + "</b>";
  document.getElementById(this.uID + "_m").innerHTML = "<b>" + inOBJ.l.a[2][this.dMonth] + "</b>";
 };
 this.SetDateBack = function( inNAME, inDATE )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   oObj.form.value = inDATE;
   window.status = "";
   this.Hide();
  }
 };
 this.InitCal = function( inNAME )
 {
  var oObj = this.GetObject(inNAME);
  if (oObj) {
   oObj.Init();
   oObj.vInit = 1;
  }
 };
 this.Add = function( inNAME, inFORM )
 {
  this.dc[this.dc.length] = new DC_Calendar(inNAME, inFORM);
 }
 this.Init = function()
 {
  this.gStr += "<table id=\"TBL_" + this.uID + "\" border=\"0\" style=\"table-layout:fixed;visibility:hidden;z-index:1000;display:none;\">";
  this.gStr += "<tr height=\"20\">";
  this.gStr += "<td id=\"" + this.uID + "_p\" align=\"center\" width=\"20\"></td>";
  this.gStr += "<td id=\"" + this.uID + "_m\" colspan=\"5\" align=\"center\" width=\"100\">&nbsp;</td>";
  this.gStr += "<td id=\"" + this.uID + "_n\" align=\"center\" width=\"20\"></td></tr><tr height=\"20\">";
  for (i = 0; i < 7; i++) {
   this.gStr += "<td id=\"" + this.uID + "_s" + i + "\" align=\"center\" width=\"20\"></td>";
  };
  this.gStr += "</tr>";
  this.gStr += "<tr height=\"20\">";
  for (var i = 0; i < 42; i++) {
   if (i != 0) {
    if (i % 7 == 0) {
     this.gStr += "</tr>";
     this.gStr += "<tr height=\"20\">";
    }
   }
   this.gStr += "<td id=\"" + this.uID + "_d" + i + "\" align=\"center\" width=\"20\">&nbsp;</td>";
  };
  this.gStr += "</tr><tr height=\"20\">";
  this.gStr += "<td align=\"center\" width=\"20\">&nbsp;</td>";
  this.gStr += "<td id=\"" + this.uID + "_a\" colspan=\"5\" align=\"center\" width=\"100\">&nbsp;</td>";
  this.gStr += "<td id=\"" + this.uID + "_c\" align=\"right\" width=\"20\"></td>";
  this.gStr += "</tr>";
  this.gStr += "</table>";
  document.body.insertAdjacentHTML("beforeend", this.gStr);
 };
};

DC_Calendar = function( inNAME, inFORM )
{
 this.name = inNAME;
 this.form = inFORM;
 this.cBg = "#FFFFFF";
 this.cBo = "#000000";
 this.lng = "FR";
 this.f = "dd/mm/yyyy";
 this.fdw = 0;
 this.top = 0;
 this.left = 0;
 this.next = "";
 this.prev = "";
 this.close = "";
 this.vInit = 0;
 this.Init = function()
 {
  this.l = new DC_Language(this.lng);
  this.fmt = new DC_Format(this.f);
  this.vInit = 1;
 };
};

var DC_LANG = 
[
 [
  "FR",
  ["D", "L", "M", "M", "J", "V", "S"],
  ["JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"],
  ["Mois précédent", "Mois suivant", "Afficher le calendrier", "Fermer le calendrier"],
  "Veuillez vérifier le format de la date."
 ],
 [
  "EN",
  ["S", "M", "T", "W", "T", "F", "S"],
  ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"],
  ["Previous month", "Next month", "Show calendar", "Close calendar"],
  "Please check date format."
 ]
];

DC_Language = function( inLANG )
{
 this.a = DC_LANG[0];
 this.Init = function()
 {
  if (inLANG + "" != "") {
   var i = 0, lng = inLANG.toUpperCase();
   for (i = 0; i < DC_LANG.length; i++) {
    if (DC_LANG[i][0].toUpperCase() == lng) {
     this.a = DC_LANG[i];
     break;
    }
   };
  }
 };
 this.Init();
};

DC_Format = function( inFORMAT )
{
 this.sDel = 0;
 this.sFmt = 0;
 this.aDel = new Array();
 this.aDel[0] = "/";
 this.aDel[1] = "-";
 this.aDel[2] = ":";
 this.aFmt = new Array();
 this.aFmt[0] = ["ddmmyyyy", "0", "1", "2"];
 this.aFmt[1] = ["ddyyyymm", "0", "2", "1"];
 this.aFmt[2] = ["mmddyyyy", "1", "0", "2"];
 this.aFmt[3] = ["mmyyyydd", "2", "0", "1"];
 this.aFmt[4] = ["yyyymmdd", "2", "1", "0"];
 this.aFmt[5] = ["yyyyddmm", "1", "2", "0"];

 this.Init = function()
 {
  var i, dOK = 0, a, s = "", fOK = 0;
  for (i = 0; i < this.aDel.length; i++) {
   if (inFORMAT.split(this.aDel[i]).length == 3) {
    this.sDel = this.aDel[i];
    dOK = 1;
    break;
   }
  };
  if (dOK == 0) {
   this.sDel = this.aDel[0];
  } else {
   a = inFORMAT.split(this.sDel);
   for (i = 0; i < a.length; i++) {
    s += a[i];
   };
   for (i = 0; i < this.aFmt.length; i++) {
    if (s == this.aFmt[i][0]) {
     this.sFmt = this.aFmt[i];
     fOK = 1;
     break;
    }
   };
   if (fOK == 0) {
    this.sFmt = this.aFmt[0];
   }
  }
 };
 this.GetDateFormatted = function( inDAY, inMONTH, inYEAR )
 {
  if ((inDAY + "").length < 2) {var dD = "0" + inDAY;} else {dD = inDAY;}
  var dM = inMONTH + 1;
  if ((dM + "").length < 2) {dM = "0" + dM;}
  var i, s0, s1, s2, s;
  for (i = 1; i < this.sFmt.length; i++) {
   eval("s" + this.sFmt[i] + " = '§" + i + "§';");
  };
  s = s0 + this.sDel + s1 + this.sDel + s2;
  s = s.replace("§1§", dD);
  s = s.replace("§2§", dM);
  s = s.replace("§3§", inYEAR);
  return s;
 };
 this.Init();
};