var newwin;
function launchwin(winurl,winname,winfeatures){
  newwin = window.open(winurl,winname,winfeatures);
  if(javascript_version > 1.0){
    setTimeout('newwin.focus();',250);
  }
};

function message(txt) {
	top.status=txt;
};

function goToPage(leMode, laPage, laTarget) {
  document.forms[0].h_mode.value = leMode;
  document.forms[0].action = laPage;
  if(laTarget!=''){
    document.forms[0].target = laTarget.substring(laTarget.lastIndexOf(".") + 1, laTarget.length);
  }
  document.forms[0].submit();
};

function goToCat(leCatId, laTarget) {
  document.forms[0].h_catid.value = leCatId;
};

function goToTopic(leTopId, laTarget) {
  document.forms[0].h_topicid.value = leTopId;
};

function goToTemplate(leTmplId, laTarget) {
  document.forms[0].h_tmplid.value = leTmplId;
};

function goToProfil(leProfil, laTarget) {
  document.forms[0].h_profil.value = leProfil;
};

function goToGroupe(leGroupId, laTarget) {
  document.forms[0].h_groupid.value = leGroupId;
};

function goToModele(leModId, laTarget) {
  document.forms[0].h_modid.value = leModId;
};

function jsTrim(monItem) {
  var monTexte = new String("");
  monTexte = monItem.value;
  while (monTexte.charAt(0) == ' ') {
    monTexte = monTexte.substring(1,monTexte.length);
  }
  while (monTexte.charAt(monTexte.length - 1) == ' ') {
    monTexte = monTexte.substring(0, (monTexte.length - 1));
  }
  monItem.value = monTexte;
};

function jsUCase(monItem) {
  var monTexte = "" + monItem.value;
  monItem.value = monTexte.toUpperCase();
};

function switchRadio(rdo, pos, val, rdoStr, parStr, listStr, rdoType) {
  // rdo     : Identifiant radio
  // pos     : Position radio racine
  // val     : Valeur radio racine
  // rdoStr  : Name Input Radio
  // parStr  : Name Hidden Parent
  // listStr : Name List ID
  // rdoType : Type du message Alert
  var myForm = document.forms[0], leRdo, laList, lePar, leTemp, leStr = "", i = 0, leCompteur = 0;
  leRdo = eval("myForm." + rdoStr + rdo);
  if (leRdo[pos].checked == true) {
    lePar = eval("myForm." + parStr + rdo);
    laList = eval("myForm." + listStr);
    leStr = ""
    while (lePar) {
      leStr = leStr + "§" + lePar.value + "§";
      lePar = eval("myForm." + parStr + lePar.value);
    };
    for (i = 0; i < laList.length; i++) {
      if (laList[i].value == rdo) {
        leCompteur = i + 1;
        break;
      }
    };
    for (i = leCompteur; i < laList.length; i++) {
      lePar = eval("myForm." + parStr + laList[i].value);
      if (leStr.indexOf("§" + lePar.value + "§") != -1) {
        break;
      } else {
        leTemp = eval("myForm." + rdoStr + laList[i].value);
        leTemp[pos].checked = true;
      }
    };
  } else {
    lePar = eval("myForm." + parStr + rdo);
    leTemp = eval("myForm." + rdoStr + lePar.value);
    if (leTemp) {
      if (leTemp[pos].checked == true) {
        leRdo[pos].checked = true;

        // GESTION ALERTE //
        if (rdoType == "Grp") {
          alert("Vous ne pouvez modifier les droits d\'un groupe si l'un de ses pères est interdit");
        } else {
          alert("Vous ne pouvez modifier les droits d\'une catégorie si l'une de ses pères est interdite");
        }

      }
    }
  }
};

function autoSelect(fld, slt) {
  // fld : Field de saisie
  // slt : Select de destination
  var myForm = this.document.forms[0], leFld, leSlt, leTxt = "", laTaille = 0, i = 0, j = 0, leTest = "";
  leFld = eval("myForm." + fld)
  leSlt = eval("myForm." + slt)
  leTxt = ("" + leFld.value).toUpperCase();
  laTaille = leTxt.length;
  if (leTxt != "") {
    for (i = 0; i < leSlt.length; i++) {
      leTest = leSlt.options[i].text;
      leTest = (leTest.substring(0, laTaille)).toUpperCase();
      if (leTest == leTxt) {
        for (j = 0; j < leSlt.length; j++) {
          leSlt.options[j].selected = false;
        };
	leSlt.options[i].selected = true;
        break;
      }
    };
  } else {
    for (i = 0; i < leSlt.length; i++) {
      leSlt.options[i].selected = false;
    };
  }
};

function switchSelect(slt_1, slt_2, type) {
  // s_1   : le select de depart
  // s_2   : le select arrivee
  // type  : type du switch - 'one' / 'all' - determine le passage
  var myForm = this.document.forms[0], sltFrom, sltTo, uneOption, i = 0, leCompteur = 0;
  sltFrom = eval("myForm." + slt_1);
  sltTo = eval("myForm." + slt_2);
  if (sltFrom.length > 0) {
    if (type == "one") {
      for (i = 0; i < sltFrom.length; i++) {
	if (sltFrom.options[i].selected == true) {
	  leCompteur = 1;
	  break;
	}
      };
      if (leCompteur == 1) {
	leCompteur = 0;
	for (i = 0; i < sltFrom.length; i++) {
	  if (sltFrom.options[i].selected == true) {
	    leCompteur++;
	  }
	};
	for (i = 0; i < leCompteur; i++) {
	  uneOption = new Option(sltFrom.options[sltFrom.selectedIndex].text, sltFrom.options[sltFrom.selectedIndex].value, false, false);
	  sltTo.options[sltTo.length] = uneOption;
	  sltFrom.options[sltFrom.selectedIndex] = null;
	};
      } else {
	alert("Il n\'y a pas d\'option(s) sélectionnée(s).");
      }
    } else {
      leCompteur = sltFrom.length;
      for (i = 0; i < leCompteur; i++) {
	uneOption = new Option(sltFrom.options[0].text, sltFrom.options[0].value, false, false);
	sltTo.options[sltTo.length] = uneOption;
	sltFrom.options[0] = null;
      };
    }
  } else {
    alert("Il n\'y a plus d\'options à deplacer dans ce sens.");
  }
};

function resetSwitch(hid_1, hid_2, slt_1, slt_2, fld_1, fld_2) {
  // hid_1 : Hidden name pour remplir le premier select
  // hid_2 : Hidden name pour remplir le deuxieme select
  // slt_1 : Nom du premier select
  // slt_2 : Nom du deuxieme select
  // fld_1 : Nom du Recherche 1
  // fld_2 : Nom du Recherche 2
  var myForm = this.document.forms[0], i = 0, leCompteur = 0, hid1, hid2, slt1, slt2, tot1, tot2, fld1, fld2;
  
  hid1 = eval("myForm." + hid_1);
  hid2 = eval("myForm." + hid_2);
  slt1 = eval("myForm." + slt_1);
  slt2 = eval("myForm." + slt_2);
  tot1 = eval("myForm." + hid_1 + "_TOT");
  tot2 = eval("myForm." + hid_2 + "_TOT");
  
  leCompteur = slt1.length;
  for (i = 0; i < leCompteur; i++) {
    slt1.options[0] = null;
  };
  if (tot1.value > 0) {
    leCompteur = 0;
    for (i = 0; i < hid1.length; i = i + 2) {
      uneOption = new Option(hid1[i + 1].value, hid1[i].value, false, false);
      slt1.options[leCompteur] = uneOption;
      leCompteur++;
    };
  }

  leCompteur = slt2.length;
  for (i = 0; i < leCompteur; i++) {
    slt2.options[0] = null;
  };
  if (tot2.value > 0) {
    leCompteur = 0;
    for (i = 0; i < hid2.length; i = i + 2) {
      uneOption = new Option(hid2[i + 1].value, hid2[i].value, false, false);
      slt2.options[leCompteur] = uneOption;
      leCompteur++;
    };
  }

  if (fld_1 != "") { 
    fld1 = eval("myForm." + fld_1);
    fld1.value = "";
  }
  if (fld_1 != "") { 
    fld2 = eval("myForm." + fld_2);
    fld2.value = "";
  }

};

function validSwitch(slt) {
  // slt : Select valide
  var myForm = this.document.forms[0], leSlt, i = 0;
  leSlt = eval("myForm." + slt)
  for (i = 0; i < leSlt.length; i++) {
    leSlt.options[i].selected = true;
  };
};

function swapAPI(slt, fld) {
  // slt : Select API
  // fld : Field de destination
  var myForm = this.document.forms[0], leSelect, leField, leTemp;
  leSelect = eval("myForm." + slt);
  leField = eval("myForm." + fld);
  if (leSelect.options[leSelect.selectedIndex].value == -1) {
    leField.value = "";
  } else {
    leTemp = eval("myForm.h_" + leSelect.options[leSelect.selectedIndex].value);
    leField.value = leTemp.value;
  }
};

function FUCase(monItem) {
  // Passage de la premiere lettre en majuscule
  leText = "" + monItem.value;
  if (leText.value + "" != "") {
    monItem.value = (leText.substring(0,1)).toUpperCase() + leText.substring(1,leText.length);
  }
};

function getOrder(prt, ord) {
  // prt : Select contenant le Parent
  // ord : Select de destination
  var myForm = this.document.forms[0], leParent, leOrder, leTemp, i = 0, leCompteur, uneOption;
  leParent = eval("myForm." + prt);
  leOrder = eval("myForm." + ord);
  leCompteur = leOrder.length
  for (i = 1; i < leCompteur; i++) {
    leOrder.options[1] = null;
  };
  leTemp = eval("myForm.h_" + leParent.options[leParent.selectedIndex].value);
  if (leTemp) {
    if (leTemp.length > 1) {
      leCompteur = leOrder.length;
      for (i = 0; i < leTemp.length; i++) {
        uneOption = new Option('Après ' + leTemp[i].value, leCompteur, false, false);
        leOrder.options[leOrder.length] = uneOption;
        leCompteur++;
      };
    } else {
      uneOption = new Option('Après ' + leTemp.value, leOrder.length, false, false);
      leOrder.options[leOrder.length] = uneOption;
    }
  }
  leOrder.focus();
};

function AfficheMaxi(inHTML, inTITLE) {
 html = '<HTML><HEAD><TITLE>' + inTITLE + '</TITLE></HEAD><BODY LEFTMARGIN=20 MARGINWIDTH=20 TOPMARGIN=20 MARGINHEIGHT=20><CENTER>' + inHTML + '</CENTER></BODY></HTML>';
 var popupImage = window.open('','_blank','toolbar=0,location=0,directories=0,menuBar=0,scrollbars=1,resizable=1,top=100,left=100');
 popupImage.document.open();
 popupImage.document.write(html);
 popupImage.document.close()
};
