function getQuery(inINDX){
 var queryString = location.search;
 var data = queryString.slice(1,queryString.length);
 var aData = data.split("&");
 var aOut = aData[inINDX].split("=");
 return aOut[1];
};

var uID = getQuery(0);

var foo = eval("window.opener.des_" + uID);

document.onkeydown = function () { 
 if (event.keyCode == 13) {
  InsertLink();
 }
};

document.onkeypress = onkeyup = function () {
 if (event.keyCode == 13) {
  event.cancelBubble = true;
  event.returnValue = false;
  return false;			
 }
};

function getLink() {
 if (foo.document.selection.type == "Control") {
  var oControlRange = foo.document.selection.createRange();
  if (oControlRange(0).tagName.toUpperCase() == "IMG") {
   var oSel = oControlRange(0).parentNode;
  }
 } else {
  oSel = foo.document.selection.createRange().parentElement();
 }
 if (oSel.tagName.toUpperCase() == "A") {
  document.linkForm.targetWindow.value = oSel.target;
  document.linkForm.link.value = oSel.href;
  document.linkForm.insertLink.value = "Modifier le lien";
 } else {
  document.linkForm.link.value = "http://";
 }
};

function SetLink(inLink) {
 document.linkForm.link.value = 'javascript:' + inLink;
};

function InsertLink() {
 targetWindow = document.linkForm.targetWindow.value;
 var linkSource = document.linkForm.link.value;
 if (linkSource != "") {
  var oNewLink = foo.document.createElement("<A border=\"0\">");
  oNewSelection = foo.document.selection.createRange();
  if (foo.document.selection.type == "Control") {
   selectedImage = foo.document.selection.createRange()(0);
   selectedImage.width = selectedImage.width;
   selectedImage.height = selectedImage.height;
   selectedImage.border = 0;
  }
  oNewSelection.execCommand("CreateLink",false,linkSource);
  if (foo.document.selection.type == "Control") {
   oLink = oNewSelection(0).parentNode;
  } else {
   oLink = oNewSelection.parentElement();
  }
  if (targetWindow != "") {
   oLink.target = targetWindow;
  } else {
   oLink.removeAttribute("target");
  }
  foo.focus();
  self.close();
 } else {
  alert("Vous devez précisez au moins l\'URL");
  document.linkForm.link.focus();
 }
};

function RemoveLink() {
 if (foo.document.selection.type == "Control") {
  selectedImage = foo.document.selection.createRange()(0);
  selectedImage.width = selectedImage.width;
  selectedImage.height = selectedImage.height;
 }
 foo.document.execCommand("Unlink");
 foo.focus();
 self.close();
};

function getAnchors() {
 var allLinks = foo.document.body.getElementsByTagName("A");
 for (a=0; a < allLinks.length; a++) {
  if (allLinks[a].href.toUpperCase() == "") {
   document.write("<option value=#" + allLinks[a].name + ">" + allLinks[a].name + "</option>");
  }
 };
};