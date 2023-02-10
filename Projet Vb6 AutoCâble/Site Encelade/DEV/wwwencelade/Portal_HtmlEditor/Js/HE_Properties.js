function getQuery(inINDX){
 var queryString = location.search;
 var data = queryString.slice(1,queryString.length);
 var aData = data.split("&");
 var aOut = aData[inINDX].split("=");
 return aOut[1];
};

var uID = getQuery(0);

var foo = eval("window.opener.des_" + uID);

var metaKeywords = "", metaDescription = "", oDescription, oKeywords;

var pageTitle = foo.document.title;
var pageBgColor = foo.document.body.bgColor;
var pageLinkColor = foo.document.body.link;
var pageTextColor = foo.document.body.text;
var backgroundImage = foo.document.body.background;

var metaData = foo.document.getElementsByTagName('META');
for (var m = 0; m < metaData.length; m++) {
 if (metaData[m].name.toUpperCase() == "KEYWORDS") {
  metaKeywords = metaData[m].content;
  oKeywords = metaData[m];
 }
 if (metaData[m].name.toUpperCase() == 'DESCRIPTION') {
  metaDescription = metaData[m].content
  oDescription = metaData[m]
 }
};

function setValues() {
 document.pageForm.pagetitle.value = pageTitle;
 document.pageForm.description.value = metaDescription;
 document.pageForm.keywords.value = metaKeywords;
 document.pageForm.background.value = backgroundImage;
 document.pageForm.bgcolor.value = pageBgColor;
 document.pageForm.textcolor.value = pageTextColor;
 document.pageForm.linkcolor.value = pageLinkColor;
 this.focus();
};

document.onkeydown = function () { 
 if (event.keyCode == 13) {
  EditPage();
 }
};

document.onkeypress = onkeyup = function () {
 if (event.keyCode == 13) {
  event.cancelBubble = true;
  event.returnValue = false;
  return false;			
 }
};

function DisplayColor(inTYPE) {
 if (inTYPE == 'bg') {
  window.opener.DoPageColorBg();
 } else if (inTYPE == 'tx') {
  window.opener.DoPageColorText();
 } else if (inTYPE == 'lk') {
  window.opener.DoPageColorLink();
 }
};

function DisplayImage( inFORM ) {
 window.opener.DoImagePage( inFORM );
};

function EditPage() {
 var bgImage = document.pageForm.background.value;
 var bgcolor = document.pageForm.bgcolor.value;
 var linkcolor = document.pageForm.linkcolor.value;
 var textcolor = document.pageForm.textcolor.value;
 if (bgImage != "") {
  foo.document.body.background = bgImage;
 } else {
  foo.document.body.removeAttribute("background",0);
 }
 if (bgcolor != "") {
  foo.document.body.bgColor = bgcolor;
 } else {
  foo.document.body.removeAttribute("bgColor",0);
 }
 if (linkcolor != "") {
  foo.document.body.link = linkcolor;
 } else {
  foo.document.body.removeAttribute("link",0);
 }
 if (textcolor != "") {
  foo.document.body.text = textcolor;
 } else {
  foo.document.body.removeAttribute("text",0);
 }
 foo.document.title = document.pageForm.pagetitle.value;
 var oHead = foo.document.getElementsByTagName('HEAD');
 if (oKeywords != null) {
  oKeywords.content = document.pageForm.keywords.value;
 } else {
  var oMetaKeywords = foo.document.createElement("META");
  oMetaKeywords.name = "Keywords";
  oMetaKeywords.content = document.pageForm.keywords.value;
  oHead(0).appendChild(oMetaKeywords);
 }
 if (oDescription != null){
  oDescription.content = document.pageForm.description.value;
 } else {
  var oMetaDesc= foo.document.createElement("META");
  oMetaDesc.name = "Description";
  oMetaDesc.content = document.pageForm.description.value;
  oHead(0).appendChild(oMetaDesc);
 }
 if (window.opener.JS_HE.imageWin) {
  window.opener.JS_HE.imageWin.close();
  window.opener.JS_HE.imageWin = null;
 }
 self.close();
};