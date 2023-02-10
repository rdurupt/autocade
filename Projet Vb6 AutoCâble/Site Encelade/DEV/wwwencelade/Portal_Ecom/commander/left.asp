<HTML>
<HEAD>
<TITLE>Untitled Document</TITLE>
<style>  
  <!--
  .text  {
  	font-size : 10px;
  	font-family : Verdana;
  	font-weight : normal;
  	font-style : normal;
  	color : #000000;
  }
  -->
  </style>

<script language="javascript">
<!--


	var download = false;
	var browser4 = false;
	var tempName = null;
	var layer_visible = top.window.menu_order;
	if(layer_visible==-1 || layer_visible>4)
		layer_visible = null;

	
	if(document.all || document.layers){
		browser4 = true;
	}

	if(document.images){
		var defaultImage = new Image();
		defaultImage.src = "img/cale.gif";
	
		var tabLayer0ImgOn = new Array(4);
		var layer0_roll_on = null;
		for(i=0;i<4;i++){
			layer0_roll_on = new Image();
			layer0_roll_on.src = "img/layer0_roll"+i+"_on.gif";
			tabLayer0ImgOn[i] = layer0_roll_on.src;
		}

		var tabLayer0ImgOff = new Array(4);
		var layer0_roll_off = null;
		for(i=0;i<4;i++){
			layer0_roll_off = new Image();
			layer0_roll_off.src = "img/layer0_roll"+i+"_off.gif";
			tabLayer0ImgOff[i] = layer0_roll_off.src;
		}

		var tabLayer1ImgOn = new Array(3);
		var layer1_roll_on = null;
		for(i=0;i<3;i++){
			layer1_roll_on = new Image();
			layer1_roll_on.src = "img/layer1_roll"+i+"_on.gif";
			tabLayer1ImgOn[i] = layer1_roll_on.src;
		}

		var tabLayer1ImgOff = new Array(3);
		var layer1_roll_off = null;
		for(i=0;i<3;i++){
			layer1_roll_off = new Image();
			layer1_roll_off.src = "img/layer1_roll"+i+"_off.gif";
			tabLayer1ImgOff[i] = layer1_roll_off.src;
		}

		var tabLayer2ImgOn = new Array(3);
		var layer2_roll_on = null;
		for(i=0;i<3;i++){
			layer2_roll_on = new Image();
			layer2_roll_on.src = "img/layer2_roll"+i+"_on.gif";
			tabLayer2ImgOn[i] = layer2_roll_on.src;
		}

		var tabLayer2ImgOff = new Array(3);
		var layer2_roll_off = null;
		for(i=0;i<3;i++){
			layer2_roll_off = new Image();
			layer2_roll_off.src = "img/layer2_roll"+i+"_off.gif";
			tabLayer2ImgOff[i] = layer2_roll_off.src;
		}

		var tabLayer3ImgOn = new Array(4);
		var layer3_roll_on = null;
		for(i=0;i<4;i++){
			layer3_roll_on = new Image();
			layer3_roll_on.src = "img/layer3_roll"+i+"_on.gif";
			tabLayer3ImgOn[i] = layer3_roll_on.src;
		}

		var tabLayer3ImgOff = new Array(4);
		var layer3_roll_off = null;
		for(i=0;i<4;i++){
			layer3_roll_off = new Image();
			layer3_roll_off.src = "img/layer3_roll"+i+"_off.gif";
			tabLayer3ImgOff[i] = layer3_roll_off.src;
		}

		var tabSousChoixOn = new Array(5);
		var schoixon = null;
		for(i=0;i<5;i++){
			schoixon = new Image();
			schoixon.src = "img/schoix"+i+"_on.gif";
			tabSousChoixOn[i] = schoixon.src;
		}

		var tabSousChoixOff = new Array(5);
		var schoixoff = null;
		for(i=0;i<5;i++){
			schoixoff = new Image();
			schoixoff.src = "img/schoix"+i+"_off.gif";
			tabSousChoixOff[i] = schoixoff.src;
		}

		var tabImgOn = new Array(8);
		var roll_on = null;
		for(i=0;i<8;i++){
			roll_on = new Image();
			roll_on.src = "img/roll"+i+"_on.gif";
			tabImgOn[i] = roll_on.src;
		}

		var tabImgOff = new Array(8);
		var roll_off = null;
		for(i=0;i<8;i++){
			roll_off = new Image();
			roll_off.src = "img/roll"+i+"_off.gif";
			tabImgOff[i] = roll_off.src;
		}

		var tabTitle = new Array(5);
		var title = null;
		for(i=0;i<5;i++){
			title = new Image();
			title.src = "img/title"+i+".gif";
			tabTitle[i] = title.src;
		}

	}
	
	function rollOver(name,i){
		if(document.images && download){
			window.document.images[name].src = tabImgOn[i];
		}
	}
	
	function rollOut(name,i){
		if(document.images && download){
			window.document.images[name].src = tabImgOff[i];
		}
	}

	function schoixOver(name,i){
		if(document.images && download){
			window.document.images[name].src = tabSousChoixOn[i];
			parent.title.window.document.images["title"].src = tabTitle[i];
			if(i != 3){
				if(i!=layer_visible){
					if(layer_visible!=null){
					    MM_showHideLayers('document.layers[\'Layer'+i+'\']','document.all[\'Layer'+i+'\']','show','document.layers[\'Layer'+layer_visible+'\']','document.all[\'Layer'+layer_visible+'\']','hide');
					    MM_showHideLayers('document.layers[\'bg'+i+'\']','document.all[\'bg'+i+'\']','show','document.layers[\'bg'+layer_visible+'\']','document.all[\'bg'+layer_visible+'\']','hide');
					    window.document.images[tempName].src = tabSousChoixOff[layer_visible];
					}
					else
						MM_showHideLayers('document.layers[\'Layer'+i+'\']','document.all[\'Layer'+i+'\']','show');
						MM_showHideLayers('document.layers[\'bg'+i+'\']','document.all[\'bg'+i+'\']','show');
					layer_visible = i;
					tempName = name;
				}
			}
			else{
				if(layer_visible!=null){
					    MM_showHideLayers('document.layers[\'Layer'+layer_visible+'\']','document.all[\'Layer'+layer_visible+'\']','hide');
					    MM_showHideLayers('document.layers[\'bg'+layer_visible+'\']','document.all[\'bg'+layer_visible+'\']','hide');
					    window.document.images[tempName].src = tabSousChoixOff[layer_visible];
					    layer_visible = null;
					    tempName = null;
				}
			}
		}
	}
	
	function schoixOut(name,i){
		if(document.images && download && i!=layer_visible){
			if(layer_visible!=null){
				window.document.images[name].src = tabSousChoixOff[i];
				parent.title.window.document.images["title"].src = tabTitle[layer_visible];
			}
			else{
				window.document.images[name].src = tabSousChoixOff[i];
				parent.title.window.document.images["title"].src = defaultImage.src;
			}
		}
	}

	function init(){
		if(layer_visible!=null){
			window.document.images["schoix"+layer_visible].src = tabSousChoixOn[layer_visible];
			parent.title.window.document.images["title"].src = tabTitle[layer_visible];
			MM_showHideLayers('document.layers[\'Layer'+layer_visible+'\']','document.all[\'Layer'+layer_visible+'\']','show');
			MM_showHideLayers('document.layers[\'bg'+layer_visible+'\']','document.all[\'bg'+layer_visible+'\']','show');
			layer_visible = layer_visible;
			tempName = "schoix"+layer_visible;
		}
	}

	function MM_showHideLayers() {
	  var i, visStr, args, theObj;
	  args = MM_showHideLayers.arguments;
	  for (i=0; i<(args.length-2); i+=3) {
	    visStr   = args[i+2];
	    if (navigator.appName == 'Netscape' && document.layers != null) {
	      theObj = eval(args[i]);
	      if (theObj)
		   theObj.visibility = visStr;
	    }
		else if (document.all != null) {
	      if (visStr == 'show')
		   visStr = 'visible';
	      if (visStr == 'hide')
		   visStr = 'hidden';
	      theObj = eval(args[i+1]);
	      if (theObj)
		theObj.style.visibility = visStr;
	  }
	}
}

//-->
</script>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</head>

<body bgcolor="#FFFFFF" onLoad="window.download=true;window.init();" marginheight=0 leftmargin=0 marginwidth=0 topmargin=0 onUnload="top.window.menu_order=window.layer_visible">

<div align="" id="bg0" style="position: absolute; left: 176px; top: 4px; width: 127px; height: 89px; z-index: 1; visibility: hidden"> 
  <img src="img/layer0_bg.gif" width="127" height="89"> </div>

<div align="" id="Layer0" style="position: absolute; left: 176px; top: 4px; width: 127px; height: 89px; z-index: 1; visibility: hidden"> 
  <table width="116" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="16"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="91" height="16"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="24" height="16"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="12"></td>
      <td align="left" valign="top"><img name="layer0_roll0" src="img/layer0_roll0_off.gif" border="0" width="91" height="12" usemap="#layer0_roll0"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="24" height="12"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="13"></td>
      <td align="left" valign="top"><img name="layer0_roll1" src="img/layer0_roll1_off.gif" border="0" width="91" height="13" usemap="#layer0_roll1"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="24" height="13"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="13"></td>
      <td align="left" valign="top"><img name="layer0_roll2" src="img/layer0_roll2_off.gif" border="0" width="91" height="13" usemap="#layer0_roll2"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="24" height="13"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="11"></td>
      <td align="left" valign="top"><img name="layer0_roll3" src="img/layer0_roll3_off.gif" border="0" width="91" height="11" usemap="#layer0_roll3"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="24" height="11"></td>
    </tr>
  </table>
<script language="javascript">
<!--
	function layer0RollOver(name,i){
		if(document.images && download){
			document.images[name].src = tabLayer0ImgOn[i];
		}
	}
	
	function layer0RollOut(name,i){
		if(document.images && download){
			document.images[name].src = tabLayer0ImgOff[i];
		}
	}

//-->
</script>
  <map name="layer0_roll0">
    <area shape="rect" coords="8,1,86,12" href="famille.asp?id_categorie=1" onMouseOver="layer0RollOver('layer0_roll0', 0);window.defaultStatus='Airway Manual'" onMouseOut="layer0RollOut('layer0_roll0', 0);window.defaultStatus=''" target="center">
  </map>
  <map name="layer0_roll1">
    <area shape="rect" coords="9,2,83,11" href="famille.asp?id_categorie=2" onMouseOver="layer0RollOver('layer0_roll1', 1);window.defaultStatus='Bottlang'" onMouseOut="layer0RollOut('layer0_roll1', 1);window.defaultStatus=''" target="center">
  </map>
  <map name="layer0_roll2">
    <area shape="rect" coords="1,-1,90,13" href="famille.asp?id_categorie=3" onMouseOver="layer0RollOver('layer0_roll2', 2);window.defaultStatus='Cartes VFR/GPS'" onMouseOut="layer0RollOut('layer0_roll2', 2);window.defaultStatus=''" target="center">
  </map>
  <map name="layer0_roll3">
    <area shape="rect" coords="-14,0,102,9" href="famille.asp?id_categorie=6" onMouseOver="layer0RollOver('layer0_roll3', 3);window.defaultStatus='Accessoires'" onMouseOut="layer0RollOut('layer0_roll3', 3);window.defaultStatus=''" target="center">
  </map>

</div>

<div align="" id="bg1" style="position: absolute; left: 176px; top: 34px; width: 116px; height: 85px; z-index: 1; visibility: hidden"> 
  <img src="img/layer1_bg.gif" width="116" height="85"> </div>

 
<div align="" id="Layer1" style="position: absolute; left: 176px; top: 34px; width: 116px; height: 85px; z-index: 1; visibility: hidden"> 
  <table width="116" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="9" height="20"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="94" height="20"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="15" height="20"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="9" height="12"></td>
      <td align="left" valign="top"><img name="layer1_roll0" src="img/layer1_roll0_off.gif" border="0" width="94" height="12" usemap="#layer1_roll0"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="15" height="12"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="9" height="13"></td>
      <td align="left" valign="top"><img name="layer1_roll1" src="img/layer1_roll1_off.gif" border="0" width="94" height="13" usemap="#layer1_roll1"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="15" height="13"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="9" height="12"></td>
      <td align="left" valign="top"><img name="layer1_roll2" src="img/layer1_roll2_off.gif" border="0" width="94" height="12" usemap="#layer1_roll2"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="15" height="12"></td>
    </tr>
  </table>
<script language="javascript">
<!--
	function layer1RollOver(name,i){
		if(document.images && download){
			document.images[name].src = tabLayer1ImgOn[i];
		}
	}
	
	function layer1RollOut(name,i){
		if(document.images && download){
			document.images[name].src = tabLayer1ImgOff[i];
		}
	}

//-->
</script>
  <map name="layer1_roll0">
    <area shape="rect" coords="18,-1,75,13" href="famille.asp?id_categorie=7" onMouseOver="layer1RollOver('layer1_roll0', 0);window.defaultStatus='Simulateurs'" onMouseOut="layer1RollOut('layer1_roll0', 0);window.defaultStatus=''" target="center">
  </map>
  <map name="layer1_roll1">
    <area shape="rect" coords="3,1,92,12" href="famille.asp?id_categorie=8" onMouseOver="layer1RollOver('layer1_roll1', 1);window.defaultStatus='Préparation de vol'" onMouseOut="layer1RollOut('layer1_roll1', 1);window.defaultStatus=''" target="center">
  </map>
  <map name="layer1_roll2">
    <area shape="rect" coords="2,-2,91,12" href="famille.asp?id_categorie=5" onMouseOver="layer1RollOver('layer1_roll2', 2);window.defaultStatus='Formation'" onMouseOut="layer1RollOut('layer1_roll2', 2);window.defaultStatus=''" target="center">
  </map>
</div>

<div align="" id="bg2" style="position: absolute; left: 172px; top: 74px; width: 94px; height: 63px; z-index: 1; visibility: hidden"> 
  <img src="img/layer2_bg.gif" width="94" height="63"> </div>

<div align="" id="Layer2" style="position: absolute; left: 172px; top: 74px; width: 94px; height: 63px; z-index: 1; visibility: hidden"> 
  <table width="90" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="11"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="63" height="11"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="15" height="11"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="11"></td>
      <td align="left" valign="top"><img name="layer2_roll0" src="img/layer2_roll0_off.gif" border="0" width="63" height="11" usemap="#layer2_roll0"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="15" height="11"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="12"></td>
      <td align="left" valign="top"><img name="layer2_roll1" src="img/layer2_roll1_off.gif" border="0" width="63" height="12" usemap="#layer2_roll1"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="15" height="12"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="12"></td>
      <td align="left" valign="top"><img name="layer2_roll2" src="img/layer2_roll2_off.gif" border="0" width="63" height="12" usemap="#layer2_roll2"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="15" height="12"></td>
    </tr>
  </table>
<script language="javascript">
<!--
	function layer2RollOver(name,i){
		if(document.images && download){
			document.images[name].src = tabLayer2ImgOn[i];
		}
	}
	
	function layer2RollOut(name,i){
		if(document.images && download){
			document.images[name].src = tabLayer2ImgOff[i];
		}
	}

//-->
</script>
  <map name="layer2_roll0">
    <area shape="rect" coords="4,-6,58,16" href="famille.asp?id_categorie=9" onMouseOver="layer2RollOver('layer2_roll0', 0);window.defaultStatus='Librairie'" onMouseOut="layer2RollOut('layer2_roll0', 0);window.defaultStatus=''" target="center">
  </map>
  <map name="layer2_roll1">
    <area shape="rect" coords="1,0,59,11" href="famille.asp?id_categorie=10" onMouseOver="layer2RollOver('layer2_roll1', 1);window.defaultStatus='Audiovisuel'" onMouseOut="layer2RollOut('layer2_roll1', 1);window.defaultStatus=''" target="center">
  </map>
  <map name="layer2_roll2">
    <area shape="rect" coords="8,0,54,12" href="famille.asp?id_categorie=11" onMouseOver="layer2RollOver('layer2_roll2', 2);window.defaultStatus='Stages'" onMouseOut="layer2RollOut('layer2_roll2', 2);window.defaultStatus=''" target="center">
  </map>
</div>

<div align="" id="bg3" style="position: absolute; left: 106px; top: 190px; width: 150px; height: 104px; z-index: 1; visibility: hidden"> 
  <img src="img/cale.gif" width="150" height="104"> </div>

<div align="" id="Layer3" style="position: absolute; left: 106px; top: 190px; width: 150px; height: 104px; z-index: 1; visibility: hidden"> 
</div>

<div align="" id="bg4" style="position: absolute; left: 102px; top: 159px; width: 153px; height: 100px; z-index: 1; visibility: hidden"> 
  <img src="img/layer3_bg.gif" width="153" height="100"> </div>

<div align="" id="Layer4" style="position: absolute; left: 102px; top: 159px; width: 153px; height: 100px; z-index: 1; visibility: hidden"> 
  <table border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="25"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="120" height="32"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="18" height="25"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="12"></td>
      <td align="left" valign="top"><img name="layer3_roll1" src="img/layer3_roll1_off.gif" border="0" width="120" height="12" usemap="#layer3_roll1"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="18" height="12"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="12"></td>
      <td align="left" valign="top"><img name="layer3_roll2" src="img/layer3_roll2_off.gif" border="0" width="120" height="12" usemap="#layer3_roll2"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="18" height="12"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="12"></td>
      <td align="left" valign="top"><img name="layer3_roll3" src="img/layer3_roll3_off.gif" border="0" width="120" height="12" usemap="#layer3_roll3"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="18" height="12"></td>
    </tr>
    <tr> 
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="12" height="12"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="120" height="12"></td>
      <td align="left" valign="top"><img src="img/cale.gif" border="0" width="18" height="12"></td>
    </tr>
  </table>

<script language="javascript">
<!--
	function layer3RollOver(name,i){
		if(document.images && download){
			document.images[name].src = tabLayer3ImgOn[i];
		}
	}
	
	function layer3RollOut(name,i){
		if(document.images && download){
			document.images[name].src = tabLayer3ImgOff[i];
		}
	}

//-->
</script>
  <map name="layer3_roll0">
    <area shape="rect" coords="10,-1,117,12" href="javascript:void(0)" onMouseOver="layer3RollOver('layer3_roll0', 0);window.defaultStatus='Payer en toute sécurité'" onMouseOut="layer3RollOut('layer3_roll0', 0);window.defaultStatus=''" target="center">
  </map> <map name="layer3_roll1">
    <area shape="rect" coords="7,1,114,12" href="conditions/comment.html" onMouseOver="layer3RollOver('layer3_roll1', 1);window.defaultStatus='Comment commander'" onMouseOut="layer3RollOut('layer3_roll1', 1);window.defaultStatus=''" target="center">
  </map> <map name="layer3_roll2">
    <area shape="rect" coords="8,0,115,11" href="conditions/conditions.html" onMouseOver="layer3RollOver('layer3_roll2', 2);window.defaultStatus='Frais de port et livraison'" onMouseOut="layer3RollOut('layer3_roll2', 2);window.defaultStatus=''" target="center">
  </map> <map name="layer3_roll3">
    <area shape="rect" coords="2,1,118,10" href="conditions/informations.html" onMouseOver="layer3RollOver('layer3_roll3', 3);window.defaultStatus='Infos CNIL'" onMouseOut="layer3RollOut('layer3_roll3', 3);window.defaultStatus=''" target="center">
  </map>

</div>


<table width="190" border="0" cellspacing="0" cellpadding="0" style="margin-left: 0; margin-top: 0; padding-left: 0; padding-top: 0; margin: 0; padding: 0;">

		<tr>
			<td align="right" valign="top"><a onMouseOver="window.schoixOver('schoix0', 0)" onMouseOut="window.schoixOut('schoix0', 0)" href="javascript:void(0)" ><img name="schoix0" src="img/schoix0_off.gif" border="0" width="190" height="46"></a></td>
		</tr>

		<tr>
			<td align="right" valign="top"><a onMouseOver="window.schoixOver('schoix1', 1)" onMouseOut="window.schoixOut('schoix1', 1)" href="javascript:void(0)" ><img name="schoix1" src="img/schoix1_off.gif" border="0" width="190" height="38"></a></td>
		</tr>

		<tr>
			<td align="right" valign="top"><a onMouseOver="window.schoixOver('schoix2', 2)" onMouseOut="window.schoixOut('schoix2', 2)" href="javascript:void(0)" ><img name="schoix2" src="img/schoix2_off.gif" border="0" width="190" height="25"></a></td>
		</tr>

		<tr>
			<td align="right" valign="top"><a onMouseOver="window.schoixOver('schoix3', 3)" onMouseOut="window.schoixOut('schoix3', 3)" href="famille.asp?id_categorie=12" target="center" ><img name="schoix3" src="img/schoix3_off.gif" border="0" width="190" height="45"></a></td>
		</tr>

		<tr>
			<td align="right" valign="top"><a onMouseOver="window.schoixOver('schoix4', 4)" onMouseOut="window.schoixOut('schoix4', 4)" href="javascript:void(0)" ><img name="schoix4" src="img/schoix4_off.gif" border="0" width="190" height="55"></a></td>
		</tr>

		<tr>
			<td align="right" valign="top"><img src="img/schoix_bottom.gif" border="0" width="190" height="40"></td>
		</tr>
		<tr>
			<td align="left" valign="top"><a href="caddie.asp" target="center"><img name="icone_caddie" id="icone_caddie" src="../img/caddie_<%=session("caddie_rempli")%>.gif" border="0" hspace="0" alt="Cliquez ici pour consulter votre commande"><br><img src="../img/edit_facture.gif" border="0" hspace="0" alt="Cliquez ici pour consulter votre commande"></a></td>
		</tr>

	</table>
<script language="JavaScript">
function set_caddie(on_off)
{
	document.images["icone_caddie"].src = "../img/caddie_" + on_off + ".gif";
}

</script>
</BODY>

</HTML>
