<html>
<head>
  <title>PORTAIL EUXIA</title>

<script language="Javascript" src="<%Response.Write Session("PortalPath")%>Portal_HtmlEditor/Js/HE_Class.js"></script>
<script language="Javascript" src="<%Response.Write Session("PortalPath")%>Portal_HtmlEditor/Js/HE_BarreBouttons.js"></script>
<script language="Javascript" src="<%Response.Write Session("PortalPath")%>Portal_HtmlEditor/Js/HE_ContextMenu.js"></script>

<script language="Javascript">

var HE_IMGPATH = "<%Response.Write Session("PortalPath")%>Portal_HtmlEditor/Img/";
var HE_PATH = "<%Response.Write Session("PortalPath")%>Portal_HtmlEditor/";

var oHE = null;
function InitHE() {
 oHE = new Html_Editor(window, window.document, "body");
};

</script>
</head>

<body>
&nbsp;
</body>
</html>