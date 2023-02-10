function modifie(id)
{
	var f = document.form_caddie;
	var qte = 0;
	eval ("qte = f.qte_"+id+".value;");
	if (Math.abs(parseInt(qte)) == qte) 
	{
		if (qte>0)
		{
			f.qte_produit.value = qte;
			f.id_produit.value = id;
			f.act.value="modif";
			f.submit();
		}
		else
		{
			suppr(id);
		}
	}
	else
	{
		alert ("La quantité saisie est erronée...");
	}
	
}

function suppr(id)
{
	if (confirm("Êtes-vous sûr de vouloir supprimer cet article ?"))
	{
		var f = document.form_caddie;
		f.act.value="suppr";
		f.id_produit.value = id;
		f.submit();	
	}
}

function plus_un(id)
{
		alert ("toto");
	var f = document.form_caddie;
	var qte = 0;
	eval ("qte = f.qte_"+id+".value;");
	if (Math.abs(parseInt(qte)) == qte) 
	{
		qte++;
		f.qte_produit.value = qte;
		f.id_produit.value = id;
		f.act.value = "modif";
		f.submit();
	}
	else
	{
		alert ("La quantité saisie est erronée...");
	}
}

function moins_un(id)
{
	var f = document.form_caddie;
	var qte = 0;
	eval ("qte = f.qte_"+id+".value;");
	if (Math.abs(parseInt(qte)) == qte) 
	{
		qte--;
		if (qte>0)
		{
			f.qte_produit.value = qte;
			f.id_produit.value = id;
			f.act.value="modif";
			f.submit();
		}
		else
		{
			suppr(id);
		}
	}
	else
	{
		alert ("La quantité saisie est erronée...");
	}
}


function recopie()
{
	var f = document.form_etape4;
	
	var id_pays = f.c_id_pays.options[f.c_id_pays.selectedIndex].value;
	var existe = false;
	var index = -1;
	f.l_beneficiaire.value = f.c_nom.value + " " + f.c_prenom.value;
	f.l_adresse.value = f.c_adresse.value;
	f.l_cp.value = f.c_codepostal.value;
	f.l_ville.value = f.c_ville.value;
}

function valide_1()
{
	var f = document.form_etape1;
	if (f.id_pays.options.selectedIndex<1)
	{
		alert("Vous devez sélectionner un pays de livraison");
	}
	else
	{
//		alert ("ok");
		f.submit();
	}
}


function valide_2()
{
	var f = document.form_etape2;
	f.submit();
}


function valide_3()
{
	var f = document.form_etape3;
	var ret = true;
	var err = "Les erreurs suivantes sont apparues:";

	if (f.c_login.value.length<1)
	{
		ret = false;
		err += "\nVotre login est vide";
	}
	if (f.c_password.value.length<1)
	{
		ret = false;
		err += "\nVotre mot de passe est vide";
	}

	if (ret)
	{
//		alert ("ok");
		f.submit();
	}
	else
	{
		alert (err);
	}

}

function valide_3b()
{
	var f = document.form_etape3b;
	f.submit();
}

function valide_3c()
{
	var f = document.form_etape3c;
	f.submit();
}

function valide_4()
{
	var f = document.form_etape4;
	var ret = true;
	var err = "Les erreurs suivantes sont apparues:";
	
	if (f.c_nom.value.length<1)
	{
		ret = false;
		err += "\nVotre nom est vide";
	}
	if (f.c_prenom.value.length<1)
	{
		ret = false;
		err += "\nVotre prénom est vide";
	}
	if (f.c_adresse.value.length<1)
	{
		ret = false;
		err += "\nVotre adresse personnelle est vide";
	}
	if (f.c_codepostal.value.length<1)
	{
		ret = false;
		err += "\nVotre code postal personnel est vide";
	}
	if (f.c_ville.value.length<1)
	{
		ret = false;
		err += "\nLe nom de votre ville est vide";
	}
	if (f.c_id_pays.options.selectedIndex<1)
	{
		ret = false;
		err += "\nLe Pays de votre adresse personnelle est incorrect";
	}
	if (f.c_codepostal.value.length<1)
	{
		ret = false;
		err += "\nVotre code postal personnel est vide";
	}
	
	if (f.l_beneficiaire.value.length<1)
	{
		ret = false;
		err += "\nLe nom du bénéficiaire est vide";
	}
	if (f.l_cp.value.length<1)
	{
		ret = false;
		err += "\nLe code postal de livraision est vide";
	}
	if (f.l_ville.value.length<1)
	{
		ret = false;
		err += "\nLe nom de la ville de livraison est vide";
	}
	if (f.c_id_pays.options.selectedIndex<1)
	{
		ret = false;
		err += "\nLe Pays de l'adresse de livraison est incorrect";
	}
	if (!f.accepte.checked)
	{
		ret = false;
		err += "\nVous n'avez pas accepté les conditions générales de vente"
	}
	if (f.c_passwd1.value.length>1)
	{
    	if (f.c_login.value.length<1)
    	{
    		ret = false;
    		err += "\nVotre login est vide";
    	}
		if (f.c_passwd1.value!=f.c_passwd2.value)
		{
		ret = false;
		err += "\nVos mots de passe ne sont pas identiques";
		}
	}
	else
	{
		err += "\nVotre mot de passe est vide";
	}
	if (ret)
	{
//		alert ("ok");
		f.submit();
	}
	else
	{
		alert (err);
	}

}

function valide_5(mode_de_paiement)
{
	var f=document.form_etape5;
	f.mode_paiement.value = mode_de_paiement;
	alert (f.mode_paiement.value);
}


