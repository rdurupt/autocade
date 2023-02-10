LMArray = new Array();

function Layout_Manager( inElements, inLayout ) {
  this.uID = document.uniqueID;
  LMArray[this.uID] = this;
  this.fromTbl = "";
  this.fromType = "";
  this.fromIndex = "";
  this.fromParent = "";
  this.insertType = "";
  this.insertParent = "";
  this.insertIndex = "";
  this.oTop = 0;
  this.oLeft = 0;
  this.GetDiv = function( inID )
  {
    if (inID + "" != "") {
      var oDiv = document.createElement("<DIV ID=\"" + inID + "\">")
    } else {
      var oDiv = document.createElement("DIV")
    }
    return oDiv;
  };
  this.GetTable = function( inID )
  {
    if (inID + "" != "") {
      var oTable = document.createElement("<TABLE ID=\"" + inID + "\">")
    } else {
      var oTable = document.createElement("TABLE")
    }
    return oTable;
  };
  this.GetTBody = function( inID )
  {
    if (inID + "" != "") {
      var oTBody = document.createElement("<TBODY ID=\"" + inID + "\">")
    } else {
      var oTBody = document.createElement("TBODY")
    }
    return oTBody;
  };
  this.GetTr = function( inID )
  {
    if (inID + "" != "") {
      var oTr = document.createElement("<TR ID=\"" + inID + "\">")
    } else {
      var oTr = document.createElement("TR")
    }
    return oTr;
  };
  this.GetTd = function( inID )
  {
    if (inID + "" != "") {
      var oTd = document.createElement("<TD ID=\"" + inID + "\">")
    } else {
      var oTd = document.createElement("TD")
    }
    return oTd;
  };
  this.layout = new Layout_Layout(inLayout, this);
  this.elements = new Layout_Elements(inElements, this);
  this.cm = new Layout_ContextMenu(this);
  this.Init = function()
  {
    var oTable, oTBody, oTr, oTd;
    oTable = this.GetTable("TBL_LM_" + this.uID);
    oTable.width = "100%";
    oTBody = this.GetTBody("TB_LM_" + this.uID);
    oTable.appendChild(oTBody);
    oTr = this.GetTr("TR_LM_" + this.uID);
    oTBody.appendChild(oTr);
    oTd = this.GetTd("TD_LM_" + this.uID);
    oTd.vAlign = "top";
    oTd.width = 220;
    oTr.appendChild(oTd);
    oTd.appendChild(this.elements.divEl);
    oTd.appendChild(this.cm.gDiv);
    oTd = this.GetTd("TD_LM_" + this.uID);
    oTd.align = "center";
    oTd.vAlign = "top";
    oTr.appendChild(oTd);
    oTd.appendChild(this.layout.layTable);
    document.getElementById(LM_DIVNAME).innerHTML = oTable.outerHTML;
  };
  this.OnMouseDown_ELE = function( inEL )
  {
    var oDiv, oTable, oTBody, oTr, oTd, oClone, iFrom, iToggle;
    this.fromTbl = "ELE";
    this.fromType = "TR";
    this.fromIndex = inEL.parentNode.rowIndex;
    oDiv = this.GetDiv("DIV_MOVE_" + this.uID);
    oDiv.style.position = "absolute";
    oDiv.style.pixelLeft = event.clientX;
    oDiv.style.pixelTop = event.clientY;
    oDiv.style.zoom = "0.8";
    oDiv.oEL = inEL;
    oTable = this.GetTable("TBL_MOVE_" + this.uID);
    oTable.style.width = inEL.offsetWidth;
    oTable.cellPadding = 0;
    oTable.cellSpacing = 2;
    oDiv.appendChild(oTable);
    oTBody = this.GetTBody("TB_MOVE_" + this.uID);
    oTable.appendChild(oTBody);
    oTr = this.GetTr("TR_MOVE_" + this.uID);
    oTBody.appendChild(oTr);
    if (inEL.parentNode.parentNode.childNodes.length == 2) {
      iFrom = inEL.id.substring(3, 6);
      iToggle = 1;
    } else {
      iToggle = 0;
    }
    if (inEL.IsUnique == "Yes") {
      inEL.parentNode.removeNode(true);
      oTr.appendChild(inEL);
    } else {
      oClone = inEL.cloneNode(true);
      oTr.appendChild(oClone);
    }
    if (iToggle == 1) {
      oTBody = document.getElementById("TB_" + iFrom + "_ELE_" + this.uID);
      oTr = this.GetTr("TR_" + iFrom + "_ELE_" + this.uID);
      oTBody.appendChild(oTr);
      oTd = this.GetTd("TD_" + iFrom + "_ELE_NONE_" + this.uID);
      oTd.align = "center";
      oTd.style.fontFamily = "Arial";
      oTd.style.fontSize = "x-small";
      oTd.innerHTML = "Pas d'éléments";
      oTr.appendChild(oTd);
    }
    return oDiv;
  };
  this.OnMouseDown_LAY = function( inEL )
  {
    var oClone, oDiv, oTable, oTBody, oTr, oTd, oClone, iFrom, iToggle;
    this.fromTbl = "LAY";
    oClone = inEL.parentNode.parentNode.parentNode.parentNode.parentNode;
    if (inEL.parentNode.childNodes.length > 1) {
      this.fromParent = inEL.parentElement;
      this.fromType = "TD";
      this.fromIndex = inEL.cellIndex;
    } else {
      this.fromParent = oClone.parentNode;
      this.fromType = "TR";
      this.fromIndex = oClone.rowIndex;
    }
    oDiv = this.GetDiv("DIV_MOVE_" + this.uID);
    oDiv.style.position = "absolute";
    oDiv.style.pixelLeft = event.clientX;
    oDiv.style.pixelTop = event.clientY;
    oDiv.style.zoom = "0.8";
    oDiv.oEL = inEL;
    oTable = this.GetTable("TBL_MOVE_" + this.uID);
    oTable.style.width = inEL.offsetWidth;
    oTable.cellPadding = 0;
    oTable.cellSpacing = 2;
    oDiv.appendChild(oTable);
    oTBody = this.GetTBody("TB_MOVE_" + this.uID);
    oTable.appendChild(oTBody);
    oTr = this.GetTr("TR_MOVE_" + this.uID);
    oTBody.appendChild(oTr);
    if (oClone.parentNode.childNodes.length == 1 && this.fromType == "TR") {
      iToggle = 1;
    } else {
      iToggle = 0;
    }
    if (this.fromType == "TR") {
      oClone.removeNode(true);
      oTr.appendChild(inEL);
    } else {
      inEL.removeNode(true);
      oTr.appendChild(inEL);
    }
    if (iToggle == 1) {
      oTBody = document.getElementById("TB_EXT_LAY_" + this.uID);
      oTr = this.GetTr("TR_EXT_LAY_" + this.uID);
      oTBody.appendChild(oTr);
      oTd = this.GetTd("TD_EXT_LAY_NONE_" + this.uID);
      oTd.align = "center";
      oTd.style.fontFamily = "Arial";
      oTd.style.fontSize = "x-small";
      oTd.style.border = "1 solid buttonshadow";
      oTd.style.height = 25;
      oTd.innerText = "Pas de layout";
      oTr.appendChild(oTd);
    }
    return oDiv;
  };
  this.OnMouseMove = function( inDragEL, inEL )
  {
    var oLine, oClone;
    oClone = inEL;
    if (inDragEL.oLine) {
      inDragEL.oLine.removeNode(true);
    }
    while (oClone && oClone.id.indexOf("_GLO_") < 0) {
      oClone = oClone.parentElement;
    };
    if (oClone) {
      if (oClone.id.indexOf("_ELE_") >= 0) {
        inEL = document.getElementById(oClone.id);
        oLine = this.GetDiv("DIV_OVER_ELE_" + this.uID);
        oLine.style.border = "2 solid #FF0000";
        oLine.style.position = "absolute";
        oLine.style.pixelLeft = this.GetAbsoluteLeft(inEL) - 2;
        oLine.style.pixelTop = this.GetAbsoluteTop(inEL);
        oLine.style.pixelWidth = inEL.offsetWidth + 4;
        oLine.style.pixelHeight = inEL.offsetHeight;
        oLine.innerText = String.fromCharCode(160);
        inDragEL.oLine = document.body.appendChild(oLine);
      } else {
        if (inEL.id.indexOf("TD_INT_LAY_") >= 0) {
          oLine = this.GetDiv("DIV_OVER_LAY_" + this.uID);
          oLine.style.position = "absolute";
          oLine.style.pixelLeft = this.GetAbsoluteLeft(inEL);
          oLine.style.pixelTop = this.GetAbsoluteTop(inEL);
          if (inDragEL.oEL.DragOnTd == "Yes" && inEL.DragOnTd == "Yes") {
            if ((event.clientX - oLine.style.pixelLeft) < (inEL.offsetWidth / 4)) {
              this.insertType = "TD";
              this.insertParent = inEL.parentElement;
              this.insertIndex = inEL.cellIndex;
              oLine.style.border = "1 solid #FF0000";
              oLine.style.pixelWidth = 1;
              oLine.style.pixelHeight = inEL.offsetHeight;
            } else if ((event.clientX - oLine.style.pixelLeft) > ((inEL.offsetWidth / 4) * 3)) {
              this.insertType = "TD";
              this.insertParent = inEL.parentElement;
              this.insertIndex = inEL.cellIndex + 1;
              if (this.insertIndex == this.insertParent.childNodes.length) {
                this.insertIndex = -1;
              }
              oLine.style.border = "1 solid #FF0000";
              oLine.style.pixelLeft += inEL.offsetWidth - 2;
              oLine.style.pixelWidth = 1;
              oLine.style.pixelHeight = inEL.offsetHeight;
            } else {
              this.insertType = "TR";
              this.insertParent = oClone.childNodes[0];
              this.insertIndex = inEL.parentNode.parentNode.parentNode.parentNode.parentNode.rowIndex;
              oLine.style.pixelTop -= 2;
              if (inEL.parentNode.childNodes.length > 1) {
                oLine.style.pixelLeft = this.GetAbsoluteLeft(inEL.parentElement);
                oLine.style.pixelWidth = inEL.parentElement.offsetWidth - 4;
              } else {
                oLine.style.pixelWidth = inEL.offsetWidth;
              }
              oLine.style.pixelHeight = 4;
              if ((event.clientY - oLine.style.pixelTop) > (inEL.offsetHeight / 2)) {
                oLine.style.pixelTop += inEL.offsetHeight;
                this.insertIndex++;
                if (this.insertIndex == this.insertParent.childNodes.length) {
                  this.insertIndex = -1;
                }
              }
              oLine.innerHTML = "<TABLE ID=\"TBL_OVER_LAY_" + this.uID + "\" WIDTH=\"100%\" STYLE=\"BORDER-TOP:4 solid #FF0000;\"><TBODY ID=\"TB_OVER_LAY_" + this.uID + "\"><TR ID=\"TR_OVER_LAY_" + this.uID + "\"><TD ID=\"TD_OVER_LAY_" + this.uID + "\"></TD></TR></TBODY></TABLE>";
            }
          } else {
            this.insertType = "TR";
            this.insertParent = oClone.childNodes[0];
            if (inEL.ForceTr == "None") {
              this.insertIndex = inEL.parentNode.parentNode.parentNode.parentNode.parentNode.rowIndex;
              oLine.style.pixelTop -= 2;
              if (inEL.parentNode.childNodes.length > 1) {
                oLine.style.pixelLeft = this.GetAbsoluteLeft(inEL.parentElement);
                oLine.style.pixelWidth = inEL.parentElement.offsetWidth - 4;
              } else {
                oLine.style.pixelWidth = inEL.offsetWidth;
              }
              oLine.style.pixelHeight = 4;
              if ((event.clientY - oLine.style.pixelTop) > (inEL.offsetHeight / 2)) {
                oLine.style.pixelTop += inEL.offsetHeight;
                this.insertIndex++;
                if (this.insertIndex == this.insertParent.childNodes.length) {
                  this.insertIndex = -1;
                }
              }
            } else {
              this.insertIndex = inEL.parentNode.parentNode.parentNode.parentNode.parentNode.rowIndex;
              oLine.style.pixelTop -= 2;
              if (inEL.parentNode.childNodes.length > 1) {
                oLine.style.pixelLeft = this.GetAbsoluteLeft(inEL.parentElement);
                oLine.style.pixelWidth = inEL.parentElement.offsetWidth - 4;
              } else {
                oLine.style.pixelWidth = inEL.offsetWidth;
              }
              oLine.style.pixelHeight = 4;
              if (inEL.ForceTr == "Top") {
                oLine.style.pixelTop += inEL.offsetHeight;
                this.insertIndex++;
                if (this.insertIndex == this.insertParent.childNodes.length) {
                  this.insertIndex = -1;
                }
              } else {
                this.insertIndex--;
                if (this.insertIndex < 0) {
                  this.insertIndex = 0;
                }
              }
            }
            oLine.innerHTML = "<TABLE ID=\"TBL_OVER_LAY_" + this.uID + "\" WIDTH=\"100%\" STYLE=\"BORDER-TOP:4 solid #FF0000;\"><TBODY ID=\"TB_OVER_LAY_" + this.uID + "\"><TR ID=\"TR_OVER_LAY_" + this.uID + "\"><TD ID=\"TD_OVER_LAY_" + this.uID + "\"></TD></TR></TBODY></TABLE>";
          }
          inDragEL.oLine = document.body.appendChild(oLine);
        } else if (inEL.id.indexOf("_NONE_") >= 0) {
          oLine = this.GetDiv("DIV_OVER_LAY_" + this.uID);
          oLine.style.position = "absolute";
          oLine.style.border = "2 solid #FF0000";
          oLine.style.pixelLeft = this.GetAbsoluteLeft(inEL);
          oLine.style.pixelTop = this.GetAbsoluteTop(inEL);
          oLine.style.pixelWidth = inEL.offsetWidth;
          oLine.style.pixelHeight = inEL.offsetHeight;
          inDragEL.oLine = document.body.appendChild(oLine);
          this.insertType = "TR";
          this.insertIndex = -1;
          this.insertParent = document.getElementById("TB_EXT_LAY_" + this.uID);
        }
      }
    } else {
      if (inDragEL.oLine) {
        inDragEL.oLine.removeNode(true);
      }
    }
    return inDragEL;
  };
  this.OnMouseUp = function( inDragEL )
  {
    var oEL, iFrom, oTBody, oTr, oTrInt, oTd, oClone, ratio, i;
    if (inDragEL.oLine) {
      oEL = window.event.srcElement;
      if (oEL.id.indexOf("_OVER_ELE_") >= 0) {
        oEL = document.getElementById("DIV_GLO_ELE_" + this.uID);
      } else if (oEL.id.indexOf("_OVER_LAY_") >= 0) {
        oEL = document.getElementById("TBL_GLO_LAY_" + this.uID);
      } else if (oEL.id.indexOf("_LAY_NONE_") >= 0) {
        oEL = document.getElementById("TBL_GLO_LAY_" + this.uID);
      } else {
        while (oEL && oEL.id.indexOf("_GLO_") < 0) {
          oEL = oEL.parentElement;
        };
      }
      if (oEL) {
        if (oEL.id.indexOf("_ELE_") >= 0) {
          if (inDragEL.oEL.IsUnique == "Yes") {
            iFrom = inDragEL.oEL.TdType.toUpperCase();
            oTBody = document.getElementById("TB_" + iFrom + "_ELE_" + this.uID);
            oTr = this.GetTr("TR_" + iFrom + "_ELE_" + this.uID);
            oTd = this.GetTd("TD_" + iFrom + "_ELE_" + this.uID);
            oTd.width = inDragEL.oEL.width;
            oTd.className = inDragEL.oEL.className;
            oTd.vAlign = inDragEL.oEL.vAlign;
            oTd.align = "left";
            oTd.setAttribute("IsUnique", inDragEL.oEL.getAttribute("IsUnique"), 1);
            oTd.setAttribute("DragOnTd", inDragEL.oEL.getAttribute("DragOnTd"), 1);
            oTd.setAttribute("TdType", inDragEL.oEL.getAttribute("TdType"), 1);
            oTd.setAttribute("RealContextMenu", inDragEL.oEL.getAttribute("RealContextMenu"), 1);
            oTd.setAttribute("RealAction", inDragEL.oEL.getAttribute("RealAction"), 1);
            oTd.setAttribute("TypeObj", inDragEL.oEL.getAttribute("TypeObj"), 1);
            oTd.setAttribute("IdObj", inDragEL.oEL.getAttribute("IdObj"), 1);
            oTd.setAttribute("ForceTr", inDragEL.oEL.getAttribute("ForceTr"), 1);
            oTd.setAttribute("TdMove", "Yes", 1);
            oTd.innerHTML = inDragEL.oEL.innerHTML;
            oTr.appendChild(oTd);
            if (oTBody.childNodes.length == 2 && oTBody.childNodes[oTBody.childNodes.length - 1].childNodes[0].id.indexOf("_NONE_") >= 0) {
              oTBody.childNodes[oTBody.childNodes.length - 1].removeNode(true);
            }
            newEL = oTBody.insertRow(-1);
            newEL.replaceNode(oTr);
          }
        } else {
          if (inDragEL.oEL.ForceTr == "None") {
            if (this.insertType == "TR") {
              oTr = this.GetTr("TR_EXT_LAY_" + this.uID);
              oTd = this.GetTd("TD_EXT_LAY_" + this.uID);
              oTr.appendChild(oTd);
              oTable = this.GetTable("TBL_INT_LAY_" + this.uID);
              oTable.width = "100%";
              oTable.cellSpacing = 2;
              oTable.cellPadding = 0;
              oTd.appendChild(oTable);
              oTBody = this.GetTBody("TB_INT_LAY_" + this.uID);
              oTable.appendChild(oTBody);
              oTrInt = this.GetTr("TR_INT_LAY_" + this.uID);
              oTBody.appendChild(oTrInt);
              if (inDragEL.oEL.IsUnique == "Yes") {
                oClone = inDragEL.oEL;
              } else {
                oClone = inDragEL.oEL.cloneNode(true);
              }
              oTd = this.GetTd("TD_INT_LAY_" + this.uID);
              oTd.width = oClone.width;
              oTd.className = oClone.className;
              oTd.vAlign = oClone.vAlign;
              oTd.align = oClone.align;
              oTd.setAttribute("IsUnique", oClone.getAttribute("IsUnique"), 1);
              oTd.setAttribute("DragOnTd", oClone.getAttribute("DragOnTd"), 1);
              oTd.setAttribute("TdType", oClone.getAttribute("TdType"), 1);
              oTd.setAttribute("RealContextMenu", oClone.getAttribute("RealContextMenu"), 1);
              oTd.setAttribute("RealAction", oClone.getAttribute("RealAction"), 1);
              oTd.setAttribute("TypeObj", oClone.getAttribute("TypeObj"), 1);
              oTd.setAttribute("IdObj", oClone.getAttribute("IdObj"), 1);
              oTd.setAttribute("ForceTr", oClone.getAttribute("ForceTr"), 1);
              oTd.setAttribute("TdMove", "Yes", 1);
              oTd.innerHTML = oClone.innerHTML;
              oTrInt.appendChild(oTd);
              if (this.insertParent.childNodes.length == 1 && this.insertParent.childNodes[0].childNodes[0].id.indexOf("_NONE_") >= 0) {
                this.insertParent.childNodes[0].removeNode(true);
              }
              newEL = this.insertParent.insertRow(this.insertIndex);
              newEL.replaceNode(oTr);
            } else {
              if (inDragEL.oEL.IsUnique == "Yes") {
                oClone = inDragEL.oEL;
              } else {
                oClone = inDragEL.oEL.cloneNode(true);
              }
              ratio = parseInt(100 / (this.insertParent.childNodes.length + 1)) + "%";
              oTd = this.GetTd("TD_INT_LAY_" + this.uID);
              oTd.width = ratio;
              oTd.className = oClone.className;
              oTd.vAlign = oClone.vAlign;
              oTd.align = oClone.align;
              oTd.setAttribute("IsUnique", oClone.getAttribute("IsUnique"), 1);
              oTd.setAttribute("DragOnTd", oClone.getAttribute("DragOnTd"), 1);
              oTd.setAttribute("TdType", oClone.getAttribute("TdType"), 1);
              oTd.setAttribute("RealContextMenu", oClone.getAttribute("RealContextMenu"), 1);
              oTd.setAttribute("RealAction", oClone.getAttribute("RealAction"), 1);
              oTd.setAttribute("TypeObj", oClone.getAttribute("TypeObj"), 1);
              oTd.setAttribute("IdObj", oClone.getAttribute("IdObj"), 1);
              oTd.setAttribute("ForceTr", oClone.getAttribute("ForceTr"), 1);
              oTd.setAttribute("TdMove", "Yes", 1);
              oTd.innerHTML = oClone.innerHTML;
              for (i = 0; i < this.insertParent.childNodes.length; i++) {
                this.insertParent.childNodes[i].width = ratio;
              };
              newEL = this.insertParent.insertCell(this.insertIndex);
              newEL.replaceNode(oTd);
            }
          } else {
              oTr = this.GetTr("TR_EXT_LAY_" + this.uID);
              oTd = this.GetTd("TD_EXT_LAY_" + this.uID);
              oTr.appendChild(oTd);
              oTable = this.GetTable("TBL_INT_LAY_" + this.uID);
              oTable.width = "100%";
              oTable.cellSpacing = 2;
              oTable.cellPadding = 0;
              oTd.appendChild(oTable);
              oTBody = this.GetTBody("TB_INT_LAY_" + this.uID);
              oTable.appendChild(oTBody);
              oTrInt = this.GetTr("TR_INT_LAY_" + this.uID);
              oTBody.appendChild(oTrInt);
              if (inDragEL.oEL.IsUnique == "Yes") {
                oClone = inDragEL.oEL;
              } else {
                oClone = inDragEL.oEL.cloneNode(true);
              }
              oTd = this.GetTd("TD_INT_LAY_" + this.uID);
              oTd.width = oClone.width;
              oTd.className = oClone.className;
              oTd.vAlign = oClone.vAlign;
              oTd.align = oClone.align;
              oTd.setAttribute("IsUnique", oClone.getAttribute("IsUnique"), 1);
              oTd.setAttribute("DragOnTd", oClone.getAttribute("DragOnTd"), 1);
              oTd.setAttribute("TdType", oClone.getAttribute("TdType"), 1);
              oTd.setAttribute("RealContextMenu", oClone.getAttribute("RealContextMenu"), 1);
              oTd.setAttribute("RealAction", oClone.getAttribute("RealAction"), 1);
              oTd.setAttribute("TypeObj", oClone.getAttribute("TypeObj"), 1);
              oTd.setAttribute("IdObj", oClone.getAttribute("IdObj"), 1);
              oTd.setAttribute("ForceTr", oClone.getAttribute("ForceTr"), 1);
              oTd.setAttribute("TdMove", "Yes", 1);
              oTd.innerHTML = oClone.innerHTML;
              oTrInt.appendChild(oTd);
              if (this.insertParent.childNodes.length == 1 && this.insertParent.childNodes[0].childNodes[0].id.indexOf("_NONE_") >= 0) {
                this.insertParent.childNodes[0].removeNode(true);
              }
              if (inDragEL.oEL.ForceTr == "Top") {
                newEL = this.insertParent.insertRow(0);
              } else {
                newEL = this.insertParent.insertRow(-1);
              }
              newEL.replaceNode(oTr);
          }
        }
      } else {
        this.ReplaceNodeBack(inDragEL);
      }
      inDragEL.oLine.removeNode(true);
    } else {
      this.ReplaceNodeBack(inDragEL);
    }
    return null;
  };
  this.ReplaceNodeBack = function( inEL )
  {
    var iFrom, oTBody, oTr, oTrInt, oTd, newEL;
    if (this.fromTbl == "ELE") {
      if (inEL.oEL.IsUnique == "Yes") {
        iFrom = inEL.oEL.id.substring(3, 6);
        oTBody = document.getElementById("TB_" + iFrom + "_ELE_" + this.uID);
        oTr = this.GetTr("TR_" + iFrom + "_ELE_" + this.uID);
        oTd = this.GetTd("TD_" + iFrom + "_ELE_" + this.uID);
        oTd.width = inEL.oEL.width;
        oTd.className = inEL.oEL.className;
        oTd.vAlign = inEL.oEL.vAlign;
        oTd.align = inEL.oEL.align;
        oTd.setAttribute("IsUnique", inEL.oEL.getAttribute("IsUnique"), 1);
        oTd.setAttribute("DragOnTd", inEL.oEL.getAttribute("DragOnTd"), 1);
        oTd.setAttribute("TdType", inEL.oEL.getAttribute("TdType"), 1);
        oTd.setAttribute("RealContextMenu", inEL.oEL.getAttribute("RealContextMenu"), 1);
        oTd.setAttribute("RealAction", inEL.oEL.getAttribute("RealAction"), 1);
        oTd.setAttribute("TypeObj", inEL.oEL.getAttribute("TypeObj"), 1);
        oTd.setAttribute("IdObj", inEL.oEL.getAttribute("IdObj"), 1);
        oTd.setAttribute("ForceTr", inEL.oEL.getAttribute("ForceTr"), 1);
        oTd.setAttribute("TdMove", "Yes", 1);
        oTd.innerHTML = inEL.oEL.innerHTML;
        oTr.appendChild(oTd);
        if (oTBody.childNodes.length == 2 && oTBody.childNodes[oTBody.childNodes.length - 1].childNodes[0].id.indexOf("_NONE_") >= 0) {
          oTBody.childNodes[oTBody.childNodes.length - 1].removeNode(true);
        }
        newEL = oTBody.insertRow(this.fromIndex);
        newEL.replaceNode(oTr);
      }
    } else {
      if (this.fromType == "TD") {
        oTd = this.GetTd("TD_INT_LAY_" + this.uID);
        oTd.width = inEL.oEL.width;
        oTd.className = inEL.oEL.className;
        oTd.vAlign = inEL.oEL.vAlign;
        oTd.align = inEL.oEL.align;
        oTd.setAttribute("IsUnique", inEL.oEL.getAttribute("IsUnique"), 1);
        oTd.setAttribute("DragOnTd", inEL.oEL.getAttribute("DragOnTd"), 1);
        oTd.setAttribute("TdType", inEL.oEL.getAttribute("TdType"), 1);
        oTd.setAttribute("RealContextMenu", inEL.oEL.getAttribute("RealContextMenu"), 1);
        oTd.setAttribute("RealAction", inEL.oEL.getAttribute("RealAction"), 1);
        oTd.setAttribute("TypeObj", inEL.oEL.getAttribute("TypeObj"), 1);
        oTd.setAttribute("IdObj", inEL.oEL.getAttribute("IdObj"), 1);
        oTd.setAttribute("ForceTr", inEL.oEL.getAttribute("ForceTr"), 1);
        oTd.setAttribute("TdMove", "Yes", 1);
        oTd.innerHTML = inEL.oEL.innerHTML;
        newEL = this.fromParent.insertCell(this.fromIndex);
        newEL.replaceNode(oTd);
      } else {
        oTr = this.GetTr("TR_EXT_LAY_" + this.uID);
        oTd = this.GetTd("TD_EXT_LAY_" + this.uID);
        oTr.appendChild(oTd);
        oTable = this.GetTable("TBL_INT_LAY_" + this.uID);
        oTable.width = "100%";
        oTable.cellSpacing = 2;
        oTable.cellPadding = 0;
        oTd.appendChild(oTable);
        oTBody = this.GetTBody("TB_INT_LAY_" + this.uID);
        oTable.appendChild(oTBody);
        oTrInt = this.GetTr("TR_INT_LAY_" + this.uID);
        oTBody.appendChild(oTrInt);
        oTd = this.GetTd("TD_INT_LAY_" + this.uID);
        oTd.width = inEL.oEL.width;
        oTd.className = inEL.oEL.className;
        oTd.vAlign = inEL.oEL.vAlign;
        oTd.align = inEL.oEL.align;
        oTd.setAttribute("IsUnique", inEL.oEL.getAttribute("IsUnique"), 1);
        oTd.setAttribute("DragOnTd", inEL.oEL.getAttribute("DragOnTd"), 1);
        oTd.setAttribute("TdType", inEL.oEL.getAttribute("TdType"), 1);
        oTd.setAttribute("RealContextMenu", inEL.oEL.getAttribute("RealContextMenu"), 1);
        oTd.setAttribute("RealAction", inEL.oEL.getAttribute("RealAction"), 1);
        oTd.setAttribute("TypeObj", inEL.oEL.getAttribute("TypeObj"), 1);
        oTd.setAttribute("IdObj", inEL.oEL.getAttribute("IdObj"), 1);
        oTd.setAttribute("ForceTr", inEL.oEL.getAttribute("ForceTr"), 1);
        oTd.setAttribute("TdMove", "Yes", 1);
        oTd.innerHTML = inEL.oEL.innerHTML;
        oTrInt.appendChild(oTd);
        if (this.fromParent.childNodes.length == 1 && this.fromParent.childNodes[0].childNodes[0].id.indexOf("_NONE_") >= 0) {
          this.fromParent.childNodes[0].removeNode(true);
        }
        newEL = this.fromParent.insertRow(this.fromIndex);
        newEL.replaceNode(oTr);
      }
    }
  };
  this.GetAbsoluteLeft = function( inNode )
  {
    var iLeft = 0;
    while (inNode.tagName.toUpperCase() != "BODY") {
      iLeft += inNode.offsetLeft;
      inNode = inNode.offsetParent;
    };
    return iLeft;
  };
  this.GetAbsoluteTop = function( inNode )
  {
    var iTop = 0;
    while (inNode.tagName.toUpperCase() != "BODY") {
      iTop += inNode.offsetTop;
      inNode = inNode.offsetParent;
    };
    return iTop;
  };
  this.GetOptimizedHTML = function()
  {
    var str = "", oEL, i, j, oTrInt, oHidden;
    oEL = document.getElementById("TB_EXT_LAY_" + this.uID);
    if (oEL.childNodes[0].childNodes[0].id.indexOf("_NONE_") < 0) {
      for (i = 0; i < oEL.childNodes.length; i++) {
        oTrInt = oEL.childNodes[i].childNodes[0].childNodes[0].childNodes[0].childNodes[0];
        if (oTrInt.childNodes.length == 1) {
          str += oTrInt.childNodes[0].RealAction + "||" + oTrInt.childNodes[0].TypeObj + "||" + oTrInt.childNodes[0].IdObj + "||" + oTrInt.childNodes[0].innerHTML + "||" + oTrInt.childNodes[0].align;
        } else {
          col = 1;
          for (j = 0; j < oTrInt.childNodes.length; j++) {
            str += oTrInt.childNodes[j].RealAction + "||" + oTrInt.childNodes[j].TypeObj + "||" + oTrInt.childNodes[j].IdObj + "||" + oTrInt.childNodes[j].innerHTML + "||" + oTrInt.childNodes[j].align;
            str += "§||§";
          };
          str = str.substring(0, str.length - 4);
        }
        str += "§§";
      };
    }
    str = str.substring(0, str.length - 2);
    oHidden = eval("document." + LM_FORM + "." + LM_HIDDEN);
    oHidden.value = str;
  };
  this.ShowContextMenu = function( inEL )
  {
    var oMenu, aMenu, i, height = 0, oClone, oObj;
    for (i = 0; i < this.cm.aDiv.length; i++) {
      oMenu = document.getElementById(this.cm.aDiv[i]);
      oMenu.style.display = "none";
      oMenu.style.visibility = "hidden";
    };
    if (!(inEL.RealContextMenu == "None")) {
      this.cm.sDiv = inEL;
      this.oTop = event.screenY;
      this.oLeft = event.screenX + 10;
      if (inEL.RealContextMenu.indexOf(",") >= 0) {
        aMenu = inEL.RealContextMenu.split(",");
        for (i = 0; i < aMenu.length; i++) {
          oMenu = document.getElementById("CM_" + aMenu[i] + "_ELE_" + this.uID);
          if (aMenu[i] == "ALIGN") {
            oObj = document.getElementById("CM_ALIGN_IMG");
            oObj.src = LM_IMGPATH + "lm_" + this.cm.sDiv.align + ".gif";
          } else if (aMenu[i] == "HRSIZE") {
            oObj = document.getElementById("CM_HRSIZE");
            oObj.vAlign = "absmiddle";
            if (this.cm.sDiv.childNodes[0].style.width == "100%") {
              oObj.innerHTML = this.cm.sDiv.childNodes[0].style.width;
            } else {
              oObj.innerHTML = "&nbsp;&nbsp;" + this.cm.sDiv.childNodes[0].style.width;
            }
          } else if (aMenu[i] == "HRCOLOR") {
            oObj = document.getElementById("CM_HRCOLOR");
            oObj.vAlign = "absmiddle";
            if (this.cm.sDiv.childNodes[0].style.color == "" || this.cm.sDiv.childNodes[0].style.color == "buttonshadow") {
              oObj.style.backgroundColor = "buttonface";
              oObj.innerHTML = "";
            } else {
              oObj.style.backgroundColor = this.cm.sDiv.childNodes[0].style.color;
              oObj.style.color = this.cm.sDiv.childNodes[0].style.color;
              oObj.innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
            }
          } else if (aMenu[i] == "HRSTYLE") {
            oObj = document.getElementById("CM_HRSTYLE");
            if (this.cm.sDiv.childNodes[0].style.border + "" == "") {
              oObj.src = LM_IMGPATH + "lm_none.gif";
            } else {
              oObj.src = LM_IMGPATH + "lm_" + this.cm.sDiv.childNodes[0].style.border + ".gif";
            }
          }
          height += (oMenu.childNodes.length * 22) + 2;
          oMenu.style.visibility = "visible";
          oMenu.style.display = "";
        };
      } else {
        oMenu = document.getElementById("CM_" + inEL.RealContextMenu + "_ELE_" + this.uID);
        if (inEL.RealContextMenu == "ALIGN") {
          oImg = document.getElementById("CM_ALIGN_IMG");
          oImg.src = LM_IMGPATH + "lm_" + this.cm.sDiv.align + ".gif";
        } else if (inEL.RealContextMenu == "HRSIZE") {
          oObj = document.getElementById("CM_HRSIZE");
          oObj.vAlign = "absmiddle";
          if (this.cm.sDiv.childNodes[0].style.width == "100%") {
            oObj.innerHTML = this.cm.sDiv.childNodes[0].style.width;
          } else {
            oObj.innerHTML = "&nbsp;&nbsp;" + this.cm.sDiv.childNodes[0].style.width;
          }
        } else if (inEL.RealContextMenu == "HRCOLOR") {
          oObj = document.getElementById("CM_HRCOLOR");
          oObj.vAlign = "absmiddle";
          if (this.cm.sDiv.childNodes[0].style.color == "" || this.cm.sDiv.childNodes[0].style.color == "buttonshadow") {
            oObj.style.backgroundColor = "buttonface";
            oObj.innerHTML = "";
          } else {
            oObj.style.backgroundColor = this.cm.sDiv.childNodes[0].style.color;
            oObj.style.color = this.cm.sDiv.childNodes[0].style.color;
            oObj.innerHTML = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
          }
        } else if (inEL.RealContextMenu == "HRSTYLE") {
          oObj = document.getElementById("CM_HRSTYLE");
          if (this.cm.sDiv.childNodes[0].style.border + "" == "") {
            oObj.src = LM_IMGPATH + "lm_none.gif";
          } else {
            oObj.src = LM_IMGPATH + "lm_" + this.cm.sDiv.childNodes[0].style.border + ".gif";
          }
        }
        height += (oMenu.childNodes.length * 22) + 2;
        oMenu.style.visibility = "visible";
        oMenu.style.display = "";
      }
      oClone = document.getElementById("CM_" + this.uID);
      oClone.style.visibility = "visible";
      this.cm.oPopup.document.body.innerHTML = oClone.innerHTML;
      this.cm.oPopup.show(this.oLeft, this.oTop, 140, height);
    }
  };
  this.Init();
};

function Layout_Layout(inLayout, inLM) {
  this.lm = inLM;
  this.Init = function()
  {
    var i, j, oTable, oTr, oTd, oTBodyInt, oTrInt, oTdInt, ratio;
    this.layTable = this.lm.GetTable("TBL_GLO_LAY_" + this.lm.uID);
    this.layTable.width = 400;
    this.layTable.cellSpacing = 0;
    this.layTable.cellPadding = 0;
    var oTBody = this.lm.GetTBody("TB_EXT_LAY_" + this.lm.uID);
    this.layTable.appendChild(oTBody);
    if (inLayout.length > 0) {
      for (i = 0; i < inLayout.length; i++) {
        oTr = this.lm.GetTr("TR_EXT_LAY_" + this.lm.uID);
        oTBody.appendChild(oTr);
        oTd = this.lm.GetTd("TD_EXT_LAY_" + this.lm.uID);
        oTr.appendChild(oTd);
        oTable = this.lm.GetTable("TBL_INT_LAY_" + this.lm.uID);
        oTable.width = "100%";
        oTable.cellSpacing = 2;
        oTable.cellPadding = 0;
        oTd.appendChild(oTable);
        oTBodyInt = this.lm.GetTBody("TB_INT_LAY_" + this.lm.uID);
        oTable.appendChild(oTBodyInt);
        oTrInt = this.lm.GetTr("TR_INT_LAY_" + this.lm.uID);
        oTBodyInt.appendChild(oTrInt);
        for (j = 0; j < inLayout[i].length; j++) {
          ratio = parseInt((100 / inLayout[i].length), 10) + "%";
          oTdInt = this.lm.GetTd("TD_INT_LAY_" + this.lm.uID);
          oTdInt.width = ratio;
          oTdInt.vAlign = "middle";
          if (("" + inLayout[i][j][1]) != "") {
            oTdInt.align = inLayout[i][j][1];
          } else {
            oTdInt.align = "left";
          }
          if (("" + inLayout[i][j][2]) != "") {
            oTdInt.setAttribute("IsUnique", inLayout[i][j][2], 1);
          } else {
            oTdInt.setAttribute("IsUnique", "Yes", 1);
          }
          if (("" + inLayout[i][j][3]) != "") {
            oTdInt.setAttribute("DragOnTd", inLayout[i][j][3], 1);
          } else {
            oTdInt.setAttribute("DragOnTd", "Yes", 1);
          }
          if (("" + inLayout[i][j][4]) != "") {
            oTdInt.setAttribute("TdType", inLayout[i][j][4], 1);
          } else {
            oTdInt.setAttribute("TdType", "Cat", 1);
          }
          if (("" + inLayout[i][j][5]) != "") {
            oTdInt.setAttribute("RealContextMenu", inLayout[i][j][5], 1);
          } else {
            oTdInt.setAttribute("RealContextMenu", "None", 1);
          }
          if (("" + inLayout[i][j][6]) != "") {
            oTdInt.className = inLayout[i][j][6];
          } else {
            oTdInt.className = LM_CSS;
          }
          if (("" + inLayout[i][j][7]) != "") {
            oTdInt.setAttribute("RealAction", inLayout[i][j][7], 1);
          } else {
            oTdInt.setAttribute("RealAction", "", 1);
          }
          if (("" + inLayout[i][j][8]) != "") {
            oTdInt.setAttribute("TypeObj", inLayout[i][j][8], 1);
          } else {
            oTdInt.setAttribute("TypeObj", "0", 1);
          }
          if (("" + inLayout[i][j][9]) != "") {
            oTdInt.setAttribute("IdObj", inLayout[i][j][9], 1);
          } else {
            oTdInt.setAttribute("IdObj", 0, 1);
          }
          oTdInt.setAttribute("ForceTr", inLayout[i][j][10], 1);
          oTdInt.innerHTML = inLayout[i][j][0];
          oTdInt.setAttribute("TdMove", "Yes", 1);
          oTrInt.appendChild(oTdInt);
        };
      };
    } else {
      oTr = this.lm.GetTr("TR_EXT_LAY_" + this.lm.uID);
      oTBody.appendChild(oTr);
      oTd = this.lm.GetTd("TD_EXT_LAY_NONE_" + this.lm.uID);
      oTd.align = "center";
      oTd.style.fontFamily = "Arial";
      oTd.style.fontSize = "x-small";
      oTd.style.border = "1 solid buttonshadow";
      oTd.style.height = 25;
      oTd.innerText = "Pas de layout";
      oTr.appendChild(oTd);
    }
  };
  this.Init();
};

function Layout_Elements(inElements, inLM) {
  this.lm = inLM;
  this.Init = function()
  {
    var i, j, oTable, oTBody, oTr, oTd, oTxt;
    this.divEl = this.lm.GetDiv("DIV_GLO_ELE_" + this.lm.uID);
    this.divEl.style.width = 200;
    for (i = 0; i < inElements.length; i++) {
      oTable = this.lm.GetTable("TBL_" + inElements[i][1].toUpperCase() + "_ELE_" + this.lm.uID);
      oTable.width = 200;
      oTable.cellSpacing = 2;
      this.divEl.appendChild(oTable);
      oTBody = this.lm.GetTBody("TB_" + inElements[i][1].toUpperCase() + "_ELE_" + this.lm.uID);
      oTable.appendChild(oTBody);
      oTr = this.lm.GetTr("TR_" + inElements[i][1].toUpperCase() + "_ELE_" + this.lm.uID);
      oTBody.appendChild(oTr);
      oTd = this.lm.GetTd("TD_" + inElements[i][1].toUpperCase() + "_ELE_" + this.lm.uID);
      oTd.align = "left";
      oTd.style.fontFamily = "Arial";
      oTd.style.fontWeight = "bold";
      oTd.style.fontSize = "x-small";
      oTr.appendChild(oTd);
      oTxt = this.lm.GetDiv("TX_" + inElements[i][1].toUpperCase() + "_ELE_" + this.lm.uID);
      oTxt.width = "100%";
      oTxt.height = "100%";
      oTxt.innerHTML = inElements[i][0];
      oTd.appendChild(oTxt);
      if (inElements[i][2].length > 0) {
        for (j = 0; j < inElements[i][2].length; j++) {
          oTr = this.lm.GetTr("TR_" + inElements[i][1].toUpperCase() + "_ELE_" + this.lm.uID);
          oTBody.appendChild(oTr);
          oTd = this.lm.GetTd("TD_" + inElements[i][1].toUpperCase() + "_ELE_" + this.lm.uID);
          oTd.width = "100%";
          oTd.vAlign = "middle";
          oTd.className = "oEl";
          if (("" + inElements[i][2][j][1]) != "") {
            oTd.align = inElements[i][2][j][1];
          } else {
            oTd.align = "left";
          }
          if (("" + inElements[i][2][j][2]) != "") {
            oTd.setAttribute("IsUnique", inElements[i][2][j][2], 1);
          } else {
            oTd.setAttribute("IsUnique", "Yes", 1);
          }
          if (("" + inElements[i][2][j][3]) != "") {
            oTd.setAttribute("DragOnTd", inElements[i][2][j][3], 1);
          } else {
            oTd.setAttribute("DragOnTd", "Yes", 1);
          }
          if (("" + inElements[i][2][j][4]) != "") {
            oTd.setAttribute("TdType", inElements[i][2][j][4], 1);
          } else {
            oTd.setAttribute("TdType", "Cat", 1);
          }
          if (("" + inElements[i][2][j][5]) != "") {
            oTd.setAttribute("RealContextMenu", inElements[i][2][j][5], 1);
          } else {
            oTd.setAttribute("RealContextMenu", "None", 1);
          }
          if (("" + inElements[i][2][j][6]) != "") {
            oTd.className = inElements[i][2][j][6];
          } else {
            oTd.className = LM_CSS;
          }
          if (("" + inElements[i][2][j][7]) != "") {
            oTd.setAttribute("RealAction", inElements[i][2][j][7], 1);
          } else {
            oTd.setAttribute("RealAction", "", 1);
          }
          if (("" + inElements[i][2][j][8]) != "") {
            oTd.setAttribute("TypeObj", inElements[i][2][j][8], 1);
          } else {
            oTd.setAttribute("TypeObj", "0", 1);
          }
          if (("" + inElements[i][2][j][9]) != "") {
            oTd.setAttribute("IdObj", inElements[i][2][j][9], 1);
          } else {
            oTd.setAttribute("IdObj", 0, 1);
          }
          oTd.setAttribute("ForceTr", inElements[i][2][j][10], 1);
          oTd.setAttribute("TdMove", "Yes", 1);
          oTd.innerHTML = inElements[i][2][j][0];
          oTr.appendChild(oTd);
        };
      } else {
        oTr = this.lm.GetTr("TR_" + inElements[i][1].toUpperCase() + "_ELE_" + this.lm.uID);
        oTBody.appendChild(oTr);
        oTd = this.lm.GetTd("TD_" + inElements[i][1].toUpperCase() + "_ELE_NONE_" + this.lm.uID);
        oTd.align = "center";
        oTd.style.fontFamily = "Arial";
        oTd.style.fontSize = "x-small";
        oTd.innerText = "Pas d'éléments";
        oTr.appendChild(oTd);	  
      }
      this.divEl.insertAdjacentHTML("beforeEnd", "<DIV ID=\"SP_BR_ELE_" + this.lm.uID + "\"><br></DIV>");
    };
  };
  this.Init();
};