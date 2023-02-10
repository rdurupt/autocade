if (document.images) {
  var img_1 = new Image();
  img_1.src = LM_IMGPATH + "lm_left.gif";
  var img_2 = new Image();
  img_2.src = LM_IMGPATH + "lm_center.gif";
  var img_3 = new Image();
  img_3.src = LM_IMGPATH + "lm_right.gif";
  var img_4 = new Image();
  img_4.src = LM_IMGPATH + "lm_color.gif";
  var img_5 = new Image();
  img_5.src = LM_IMGPATH + "lm_none.gif";
  var img_6 = new Image();
  img_6.src = LM_IMGPATH + "lm_dotted.gif";
  var img_7 = new Image();
  img_7.src = LM_IMGPATH + "lm_dashed.gif";
  var img_8 = new Image();
  img_8.src = LM_IMGPATH + "lm_solid.gif";
}

var dragEL = null;

document.onselectstart = Document_OnSelectStart;
function Document_OnSelectStart() {
  return false;
};

document.onmousedown = Document_OnMouseDown;
function Document_OnMouseDown() {
  if (window.event.button == 1) {
    var oEl = window.event.srcElement;
    if (oEl.TdMove == "Yes") {
      var oID = oEl.id.substring(11, oEl.id.length);
      if (oEl.id.indexOf("_ELE_") >= 0) {
        dragEL = LMArray[oID].OnMouseDown_ELE(oEl);
      } else {
        dragEL = LMArray[oID].OnMouseDown_LAY(oEl);
      }
      document.body.appendChild(dragEL);
    }
  }
};

document.onmousemove = Document_OnMouseMove;
function Document_OnMouseMove() {
  if (dragEL) {
    var oEl = window.event.srcElement;
    dragEL.style.pixelLeft = event.clientX + document.body.scrollLeft + 5;
    dragEL.style.pixelTop = event.clientY + document.body.scrollTop + 5;
    var oID = dragEL.oEL.id.substring(11, dragEL.oEL.id.length);
    dragEL = LMArray[oID].OnMouseMove(dragEL, oEl);
    return false;
  }
};

document.onmouseup = Document_OnMouseUp;
function Document_OnMouseUp() {
  if (dragEL && window.event.button == 1) {
    dragEL.removeNode(true);
    var oID = dragEL.id.substring(9, dragEL.id.length);
    dragEL = LMArray[oID].OnMouseUp(dragEL);
  }
};

document.oncontextmenu = Document_OnContextMenu;
function Document_OnContextMenu() {
  if (!dragEL) {
    var oEl = window.event.srcElement;
    if (oEl.id.indexOf("TD_INT_LAY_") >= 0) {
      var oID = oEl.id.substring(11, oEl.id.length);
      LMArray[oID].ShowContextMenu(oEl);
    }
  }
  return false;
};