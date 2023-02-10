<%
Set Portal = Server.CreateObject(Session("PortalComObject") & ".Portal")
%>

    <head>
    <link rel=STYLESHEET href='http://flyway.euxia.net/portal_styles/flyway.css' type='text/css'>
    <script language='javascript' src='../Portal_Java/Portal_Generic.js'></script>
    <!--
    <style>
    body         {
    font-family: verdana;
    font-size:xx-small;
    font-weight: normal;
    background-color: White;
    margin: 0px;
    background : url(http://flyway.euxia.net/flyway/img/bgciel.jpg);
    }
    </style>
    -->    
</head>
<BODY><br><br>
<%=portal.ViewHautPage%>
<table width="80%" border="0" cellpadding='0' cellspacing='0' align='center'>

<tr class='smallerheader'><td colspan='3'>
<IMG src="http://flyway.euxia.net/flyway/img/etape5.gif" border="0" alt="" & Text & "" usemap="#etapes"><br>
<map name="etapes">
 <area shape="rect" coords="560,143,659,160" href="http://flyway.euxia.net/portal_asp/portal.asp?mode=validercommande&etape=5" target="_self">
 <area shape="rect" coords="476,143,558,160" href="http://flyway.euxia.net/portal_asp/portal.asp?mode=validercommande&etape=4" target="_self">
 <area shape="rect" coords="387,143,475,160" href="http://flyway.euxia.net/portal_asp/portal.asp?mode=validercommande&etape=3" target="_self">
 <area shape="rect" coords="294,143,385,160" href="http://flyway.euxia.net/portal_asp/portal.asp?mode=validercommande&etape=2" target="_self">
 <area shape="rect" coords="201,144,292,159" href="http://flyway.euxia.net/portal_asp/portal.asp?mode=validercommande&etape=1" target="_self">
</map>
<br><br>
<%=portal.ImgTitreEcommerce("Paiement par carte")%>
</td></tr>
<tr><td>

[SIPS]

    </td></tr></table>
<%=portal.ViewBasPage("80%", "center")%>
<br><br></BODY></HTML>
<%
Set Portal = nothing
%>
