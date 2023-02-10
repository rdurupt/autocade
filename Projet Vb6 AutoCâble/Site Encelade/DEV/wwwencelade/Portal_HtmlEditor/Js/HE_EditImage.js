function getQuery(inINDX){
 var queryString = location.search;
 var data = queryString.slice(1,queryString.length);
 var aData = data.split("&");
 var aOut = aData[inINDX].split("=");
 return aOut[1];
};

var uID = getQuery(0);

var foo = eval("window.opener.des_" + uID);

var myImage = eval("window.opener.JS_HE");

var imageWidth = myImage.selectedImage.width;
var imageHeight = myImage.selectedImage.height;
var imageAlign = myImage.selectedImage.align;
var imageBorder = myImage.selectedImage.border;
var imageAltTag = myImage.selectedImage.alt;
var imageHspace = myImage.selectedImage.hspace;
var imageVspace = myImage.selectedImage.vspace;

function setValues() {
 document.imageForm.image_width.value = imageWidth;
 document.imageForm.image_height.value = imageHeight;
 if (imageBorder == "") {imageBorder = "0";}
 document.imageForm.border.value = imageBorder;
 document.imageForm.alt_tag.value = imageAltTag;
 document.imageForm.hspace.value = imageHspace;
 document.imageForm.vspace.value = imageVspace;
 for (var i = 0; i < document.imageForm.align.length; i++) {
  if (imageAlign.toUpperCase() == document.imageForm.align[i].value.toUpperCase()) {
   document.imageForm.align[i].selected = true;
   break;
  }
 };
 this.focus();
};

document.onkeydown = function () { 
 if (event.keyCode == 13) {
  EditImg();
 }
};

document.onkeypress = onkeyup = function () {
 if (event.keyCode == 13) {
  event.cancelBubble = true;
  event.returnValue = false;
  return false;			
 }
};

function EditImg() {
 var error = 0;
 if (isNaN(document.imageForm.image_width.value) || document.imageForm.image_width.value < 0) {
  alert("Veuillez vérifier la largeur");
  error = 1;
  document.imageForm.image_width.select();
  document.imageForm.image_width.focus();
 } else if (isNaN(document.imageForm.image_height.value) || document.imageForm.image_height.value < 0) {
  alert("Veuillez vérifier la hauteur");
  error = 1;
  document.imageForm.image_height.select();
  document.imageForm.image_height.focus();
 } else if (isNaN(document.imageForm.border.value) || document.imageForm.border.value < 0 || document.imageForm.border.value == "") {
  alert("Veuillez vérifier la bordure");
  error = 1;
  document.imageForm.border.select();
  document.imageForm.border.focus();
 } else if (isNaN(document.imageForm.hspace.value) || document.imageForm.hspace.value < 0) {
  alert("Veuillez vérifier l\'espacement horizontal")
  error = 1;
  document.imageForm.hspace.select();
  document.imageForm.hspace.focus();
 } else if (isNaN(document.imageForm.vspace.value) || document.imageForm.vspace.value < 0) {
  alert("Veuillez vérifier l\'espacement vertical")
  error = 1;
  document.imageForm.vspace.select();
  document.imageForm.vspace.focus();
 }
 if (error != 1) {
  myImage.selectedImage.width = document.imageForm.image_width.value;
  myImage.selectedImage.height = document.imageForm.image_height.value;
  myImage.selectedImage.alt = document.imageForm.alt_tag.value;
  myImage.selectedImage.border = document.imageForm.border.value;
  if (document.imageForm.hspace.value != "") {
   myImage.selectedImage.hspace = document.imageForm.hspace.value;
  } else {
   myImage.selectedImage.removeAttribute('hspace', 0);
  }
  if (document.imageForm.vspace.value != "") {
   myImage.selectedImage.vspace = document.imageForm.vspace.value;
  } else {
   myImage.selectedImage.removeAttribute('vspace', 0);
  }
  if (document.imageForm.align[document.imageForm.align.selectedIndex].text != "") {
   myImage.selectedImage.align = document.imageForm.align[document.imageForm.align.selectedIndex].value;
  } else {
   myImage.selectedImage.removeAttribute('align', 0);
  }
  self.close()
 }
};