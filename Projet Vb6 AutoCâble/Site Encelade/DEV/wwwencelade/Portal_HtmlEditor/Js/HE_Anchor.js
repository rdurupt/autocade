function getQuery(inINDX){
 var queryString = location.search;
 var data = queryString.slice(1,queryString.length);
 var aData = data.split("&");
 var aOut = aData[inINDX].split("=");
 return aOut[1];
};

var mode = getQuery(0);

var uID = getQuery(1);

var foo = eval("window.opener.des_" + uID);

document.onkeydown = function () { 
 if (event.keyCode == 13) {
  if (mode == "Add") {
   InsertAnchor();
  } else {
   EditAnchor();
  }
 }
};

document.onkeypress = onkeyup = function () {
 if (event.keyCode == 13) {
  event.cancelBubble = true;
  event.returnValue = false;
  return false;			
 }
};

function InsertAnchor() {
 error = 0;
 var sel = foo.document.selection;
 if (sel!=null) {
  var rng = sel.createRange();
  if (rng!=null) {
   name = document.anchorForm.anchor_name.value
   if (error != 1) {
    if (name == "") {
     alert("Vous devez préciser au moins le nom");
     document.anchorForm.anchor_name.focus();
     error = 1;
    } else {
     if (window.opener.JS_HE.borderShown == "yes") {
      style = ' style="BORDER-RIGHT: #000000 1px dashed; BORDER-TOP: #000000 1px dashed; BORDER-LEFT: #000000 1px dashed; WIDTH: 20px; COLOR: #FFFFCC; BORDER-BOTTOM: #000000 1px dashed; HEIGHT: 16px; BACKGROUND-COLOR: #FFFFCC"';
     } else {
      style = "";
     }
     rng.pasteHTML("<a name=" + anchorForm.anchor_name.value + style + ">");
    }
   }
  }
 }
 if (error != 1) {
  foo.focus();
  self.close();
 }
};

function EditAnchor() {
 error = 0;
 var sel = foo.document.selection;
 if (sel!=null) {
  var rng = sel.createRange()(0);
  if (rng!=null) {
   name = document.anchorForm.anchor_name.value;
   if (error != 1) {
    if (name == "") {
     alert("Vous devez préciser au moins le nom");
     document.anchorForm.anchor_name.focus();
     error = 1;
    } else {
     var oldName = rng.name;
     rng.name = name;
    }
   }
  }
 }
 if (error != 1) {
  foo.focus();
  var allLinks = foo.document.body.getElementsByTagName("A");
  for (a=0; a < allLinks.length; a++) {
   if (allLinks[a].href.toUpperCase() != "") {
    if (allLinks[a].href.indexOf("#" + oldName) >= 0) {
     allLinks[a].href = allLinks[a].href.replace("#" + oldName, "#" + name);
    }
   }
  };
  self.close();
 }
};