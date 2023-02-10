function getQuery(){
 var queryString = location.search;
 var data = queryString.slice(1,queryString.length);
 var aData = data.split("=");
 return aData[1];
};

var uID = getQuery();

var foo = eval("window.opener.des_" + uID);

document.onkeydown = function () { 
 if (event.keyCode == 13) {
  InsertEmail();
 }
};

document.onkeypress = onkeyup = function () {
 if (event.keyCode == 13) {
  event.cancelBubble = true;
  event.returnValue = false;
  return false;			
 }
};

function getEmail() {
 if (foo.document.selection.type == "Control") {
  var oControlRange = foo.document.selection.createRange();
  if (oControlRange(0).tagName.toUpperCase() == "IMG") {
   var oSel = oControlRange(0).parentNode;
  }
 } else {
  oSel = foo.document.selection.createRange().parentElement();
 }
 if (oSel.tagName.toUpperCase() == "A" && oSel.href.toUpperCase().indexOf("MAILTO:") >= 0) {
  var mailto = oSel.href.replace("mailto:", "");
  var email = mailto.substring(0, mailto.indexOf("?subject=", ""));
  var subject = mailto.substring(mailto.indexOf("?subject=", "") + ("?subject=").length, mailto.length);
  document.emailForm.email.value = email;
  document.emailForm.subject.value = subject;
  document.emailForm.insertEmail.value = "Modifier l'email";
 }
};

function InsertEmail() {
 error = 0;
 var sel = foo.document.selection;
 if (sel!=null) {
  var rng = sel.createRange();
  if (rng!=null) {
   if (foo.document.selection.type == "Control") {
    var selectedImage = foo.document.selection.createRange()(0);
    selectedImage.width = selectedImage.width;
    selectedImage.height = selectedImage.height;
   }
   var email = document.emailForm.email.value;
   var subject = document.emailForm.subject.value;
   if (error != 1) {
    if (email == "") {
     alert("Vous devez préciser au moins l\'adresse email");
     document.emailForm.email.focus();
     error = 1;
    } else {
     var mailto = "mailto:" + email;
     if (subject != "") {
      mailto = mailto + "?subject=" + subject;
     }
     rng.execCommand("CreateLink",false,mailto);
    }
   }
  }
 }
 if (error != 1) {
  foo.focus()
  self.close();
 }
};

function RemoveEmail() {
 if (foo.document.selection.type == "Control") {
  selectedImage = foo.document.selection.createRange()(0);
  selectedImage.width = selectedImage.width;
  selectedImage.height = selectedImage.height;
 }
 foo.document.execCommand("Unlink");
 foo.focus();
 self.close();
};