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

/*function valide_2(normal, express)
{
	var f = document.form_etape2;
	if (f.type_tarif.options.selectedIndex<1)
	{
		alert("Vous devez sélectionner un tarif de livraison");
	}
	else
	{
//		alert ("ok");
		if (f.type_tarif.options.selectedIndex == 1)
		{
			f.tarif.value = normal;
		}
		else
		{
			f.tarif.value = express;
		}
		
		f.submit();
	}
}*/

function valide_3()
{
	var f = document.form_etape3;
/*	var ret = true;
	var err = "Les erreurs suivantes sont apparues:";
	if (f.client[0].checked)
	{
		f.is_client.value = 1;
		if (f.compte_client.value.length<1)
		{
			ret = false;
			err += "\nVotre identifiant est vide";
		}
		if (f.passwd_client.value.length<1)
		{
			ret = false;
			err += "\nVotre mot de passe est vide";
		}
	}
	else
	{
		f.is_client.value = 0;
	}
	if (ret)
	{*/
//		alert ("is_client=" + f.is_client.value + "\nlogin=" + f.login.value + "\npasswd=" + f.passwd.value);
		f.submit();
	/*}
	else
	{
		alert(err);
	}*/
}

function set_modif()
{
	document.form_etape4.modifie.value = 1;
//	alert ("modifié");
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
	if (f.c_email.value.length<1)
	{
		ret = false;
		err += "\nVotre email est vide";
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
	if (f.is_client.value == 0)
	{
		if ((f.compte_client.value.length>0) && ((f.passwd1.value != f.passwd2.value) || (f.passwd1.value.length<1)))
		{
			ret = false;
			err += "\nVous avez tapé un identifiant de connexion, mais votre mot de passe est incorrect.";
		}
		/*else
		{	
			f.passwd_client.value = f.passwd1.value;
		}*/
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


