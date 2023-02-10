//////////////////////////////////////////////////////////////////////////////////////////////////
//	PROJET	:	TreeView.js																		//
//	AUTEUR	:	GOLINSKI Ludwig - Ludinski														//
//	DATE	:	Mercredi 21 juillet 04															//
//																								//
//	DESCRIPTION	:	Impl�mente toutes les functions n�cessaires � la cr�ation, � l'affichage et	//
//					au fonctionnement d'une treeview											//
//					Ce fichier doit se trouver dans un dossier nomm� 'TreeView' et le code qui	//
//					y fait appel, � la racine de ce dossier										//
//																								//
//	CLASSES	:	TreeView ->	G�re l'apparence de la racine, ainsi que celle du cadre situ� autour//
//							de la treeview ( DIV ).												//
//				Noeud ---->	G�re l'apparence de chacun des noeuds de la treeview ainsi que leur	//
//							Position dans la treeview.											//
//////////////////////////////////////////////////////////////////////////////////////////////////

// VARIABLE GLOBALE CONTENANT L'INSTANCE SUR LA CLASSE 'TREEVIEW'
// IL FAUT UTILISER CETTE VARIABLE
var treeView
// CONSTRUCTEUR DE LA CLASSE 'TREEVIEW'
function TreeView( styleBorder, icoRacine, txtRacine, style, styleOnOver, funOnClick )
{
	// chaine contenant l'icone de la racine
	this.icone = icoRacine

	// chaine contenant le texte situ� � c�t� de l'icone
	this.texte = txtRacine

	// chaine contenant la classe de style du texte
	this.style = style

	// chaine contenant la classe de style lorsque le curseur est dessus
	this.styleOnOver = styleOnOver

	// chaine contenant la fonction � appeler lors d'un clique sur le texte
	this.onClick = funOnClick

	// Chaine contenant l'index (unique) du noeud
	this.index = "0"

	// entier contenant la taille horizontale (en pixel) du cadre de la treeview
	this.styleBorder = styleBorder

	// Nom de la table qui contient le noeud
	this.table = "TABLE_" + this.index

	// tableau contenant les noeuds fils de la racine
	this.tableauEnfants = new Array

	// m�thodes de la classe
	this.Start = TreeView_Start
	this.Add = TreeView_Add
	this.Contient = TreeView_Contient
	this.Noeuds = TreeView_Noeuds
}

// PERMET D'AJOUTER UN NOEUD A L'ARBORESCENCE
// PARAMETRES	: icoNoeud	 - Chaine contenant l'image symbolisant le noeud
//				  txtNoeud	 - Chaine contenant le texte affich� � c�t� de l'icone
//				  funOnClick - Chaine contenant la fonction (sans les parenth�ses) � appeler lors d'un click sur le noeud
//							   Cette fonction doit prendre deux chaines en param�tres ( l'index et le texte du noeud )
function TreeView_Add( icoNoeud, txtNoeud, style, styleOnOver, funOnClick )
{
	// Recherche le prochain emplacement libre du tableau de noeud
	var indice = 0
	while( this.tableauEnfants[ indice ] != null )
		indice ++

	// Cr�e l'index unique du noeud
	var indexNoeud = "0_" + indice

	// Cr�e le noeud et l'ajoute dans le tableau
	this.tableauEnfants[ indice ] = new Noeud( icoNoeud, txtNoeud, style, styleOnOver, indexNoeud, funOnClick, this )

	// retourne le noeud pour pouvoir y ajouter des sous-noeuds
	return this.tableauEnfants[ indice ]
}

// PERMET DE LANCER L'AFFICHAGE DE LA TREEVIEW
// A APPELER APRES AVOIR AJOUTER TOUS LES NOEUDS VOULUS
function TreeView_Start()
{
	// Cr�e le DIV o� sera int�gr� la treeview ( qui se muniera de scroll barres si n�cessaire )
	document.write( "<DIV ID = 'CONTOUR_TREEVIEW' CLASS = '" + this.styleBorder + "'>" )

	// Cr�e un tableau d'une ligne et sans bord
	document.write( "<TABLE ID = '" + this.table + "' BORDER = 0 CELLSPACING = 0 CELLPADDING = 0>" )
	document.write( "<TR><TD VALIGN = middle nowrap>" )

	// Ajoute l'icone
	document.write( "<IMG ID = 'ICONE_0' SRC = '" + this.icone + "'" )
	document.write( "ONCLICK = '" + this.onClick + "( &quot;" + this.index + "&quot; , &quot;" + this.texte + "&quot; )' " )
	document.write( "ONMOUSEOVER = 'OnOver( &quot;" + this.index + "&quot; )' " )
	document.write( "ONMOUSEOUT = 'OnOut( &quot;" + this.index + "&quot; )'>" )

	// Change de colonne
	document.write( "</TD><TD VALIGN = middle nowrap>" )

	// Ajoute le texte
	document.write( "<FONT ID = 'TEXTE_0' " )
	document.write( "CLASS = '" + this.style + "'" )
	document.write( "COLOR = '#000000' " )
	document.write( "ONCLICK = '" + this.onClick + "( &quot;" + this.index + "&quot; , &quot;" + this.texte + "&quot; )' " )
	document.write( "ONMOUSEOVER = 'OnOver( &quot;" + this.index + "&quot; )' " )
	document.write( "ONMOUSEOUT = 'OnOut( &quot;" + this.index + "&quot; )'>" )
	document.write( this.texte + "</FONT>" )

	// Referme la colonne et le tableau
	document.write( "</TD></TR></TABLE>" )

	// Parcourt le tableau de noeuds
	var indice = 0
	while( this.tableauEnfants[ indice ] != null )
	{
		// Lance l'affichage du noeud
		this.tableauEnfants[ indice ].Draw()

		// Passe au noeud suivant
		indice ++
	}

	// Referme le DIV
	document.write( "</DIV>" )
}

// INDIQUE SI UN NOEUD EST PRESENT DANS L'ARBORESCENCE
// PARAMETRE	: indexNoeud - Chaine contenant l'index du noeud recherch�
// RETOUR		: bool�en, indiquant si le noeud est pr�sent ou non
function TreeView_Contient( indexNoeud )
{
	// Parcourt le tableau de noeuds fils
	var indice = 0
	while( this.tableauEnfants[ indice ] != null )
	{
		// Fait appel � la m�thode du noeud et v�rifie s'il contient ou s'il est le noeud recherch�
		if( this.tableauEnfants[ indice ].Contient( indexNoeud ) )
		{
			// Retourne comme quoi le noeud est pr�sent
			return true
		}

		// Passe au noeud suivant
		indice ++
	}

	// Retourne comme quoi le noeud n'est pas pr�sent
	return false
}

// PERMET DE RECUPERER L'INSTANCE D'UN NOEUD D'APRES SON INDEX
// PARAMETRE	: indexNoeud - Chaine contenant l'index du noeud recherch�
// RETOUR		: l'instance sur le noeud ou 'null' s'il n'est pas pr�sent dans l'arbre
function TreeView_Noeuds( indexNoeud )
{
	// Il s'agit de la racine
	if( indexNoeud == "0" )
		return this

	// Parcourt le tableau de noeuds fils
	var indice = 0
	while( this.tableauEnfants[ indice ] != null )
	{
		// R�cup�re le noeud
		var noeudRecherche = this.tableauEnfants[ indice ].Noeuds( indexNoeud )

		// V�rifie s'il a �t� trouv�
		if( noeudRecherche != null )
		{
			// Retourne le noeud
			return noeudRecherche
		}

		// Passe au noeud suivant
		indice ++
	}

	// Retourne comme quoi le noeud n'est pas pr�sent
	return null
}










// CONSTRUCTEUR DE LA CLASSE 'NOEUD'
function Noeud( icoNoeud, txtNoeud, style, styleOnOver, indexNoeud, funOnClick )
{
	// Chaine contenant l'icone du noeud
	this.icone = icoNoeud

	// Chaine contenant le texte situ� � c�t� de l'icone
	this.texte = txtNoeud

	// Chaine contenant la classe de style du texte
	this.style = style

	// chaine contenant la classe de style lorsque le curseur est dessus
	this.styleOnOver = styleOnOver

	// Chaine contenant l'index (unique) du noeud
	this.index = indexNoeud

	// Chaine contenant la m�thode (sans parenth�ses) � appeler lors d'un click sur le noeud
	this.onClick = funOnClick

	// Tableau contenant les noeud fils
	this.tableauEnfants = new Array

	// Indique si le noeud est compact� ou non
	this.isExpand = true

	// Nom de la table qui contient le noeud
	this.table = "TABLE_" + this.index

	// M�thodes de la classe
	this.Add = Noeud_Add
	this.Draw = Noeud_Draw
	this.IsEnd = Noeud_IsEnd
	this.HaveAChild = Noeud_HaveAChild
	this.Contient = Noeud_Contient
	this.Noeuds = Noeud_Noeuds
	this.Cacher = Noeud_Cacher
	this.Montrer = Noeud_Montrer
	this.ChangerExpand = Noeud_ChangerExpand
}

// PERMET D'AJOUTER UN NOEUD FILS
// PARAMETRES	: icoNoeud   - Chaine contenant l'image symbolisant le noeud
//				  txtNoeud   - Chaine contenant le texte affich� � c�t� de l'icone
//				  funOnClick - Chaine contenant la fonction � appeler lors d'un click sur le noeud
// RETOUR		: l'instance sur le noeud
function Noeud_Add( icoNoeud, txtNoeud, style, styleOnOver, funOnClick )
{
	// Recherche le prochain emplacement libre du tableau
    var indice = 0
    while( this.tableauEnfants[ indice ] != null )
        indice ++

	// Cr�e l'index unique du noeud
    var indexNoeud = this.index + "_" + indice

	// Cr�e le noeud et le stocke dans le tableau
    this.tableauEnfants[ indice ] = new Noeud( icoNoeud, txtNoeud, style, styleOnOver, indexNoeud, funOnClick )

	// Retourne le noeud, pour pouvoir y ajouter un noeud fils
    return this.tableauEnfants[ indice ]
}

// PERMET D' AFFICHER LE NOEUD SUR LA PAGE
function Noeud_Draw()
{
	// Cr�e la table servant � contenir chaqu'une des icones et le texte du noeud
	document.write( "<TABLE ID = '" + this.table + "' BORDER = 0 CELLSPACING = 0 CELLPADDING = 0>" )
	document.write( "<TR><TD VALIGN = middle nowrap>" )

	// R�cup�re le bool�en, servant � indiquer si le noeud poss�de ou non un noeud sous lui
	var isEnd = this.IsEnd()

	// R�cup�re le bool�en, servant � indiquer si le noeud poss�de au moin un noeud fils
	var haveAChild = this.HaveAChild()

	// Parcourt l'index du noeud pour d�terminer les icones � placer devant le noeud
	var indice = 2
	while( indice < this.index.length - 1 )
	{
		// Recherche le nombre de chiffres du nombre suivant
		var tailleNombre = 1
		while( indice + tailleNombre < this.index.length && this.index.substring( indice + tailleNombre, indice + tailleNombre + 1 ) != "_" )
		{
			tailleNombre ++
		}

		// R�cup�re l'indice suivant
		var indiceSuivant = parseInt( this.index.substring( indice, indice + tailleNombre ), 10 )

		// V�rifie qu'il y ait un indice et qu'il ne sagisse pas du noeud courant
		if( ! isNaN( indiceSuivant ) && this.index != this.index.substring( 0, indice ) + indiceSuivant )
		{
			// Incr�mente l'indice
			indiceSuivant ++

			// Cr�e l'index du noeud devant se situer directement en dessous
			var indexSuivant = this.index.substring( 0, indice ) + indiceSuivant

			// Buffer o� sera stock� le nom de l'icone � ajouter
			var icone

			// La treeview contient l'index cr��
			if( treeView.Contient( indexSuivant ) )
			{
				// Met l'icone de la ligne pointill�e
				icone = "TreeView/PointillesLigne.gif"
			}
			// La treeview ne contient pas l'index cr��
			else
			{
				// Met une icone vide
				icone = "TreeView/Vide.gif"
			}

			// Ajoute l'icone dans la table, puis passe � la colonne suivante
			document.write( "<IMG SRC = '" + icone + "' >" )
			document.write( "</TD><TD VALIGN = middle nowrap>" )
		}

		// Continue � parcourir l'index du noeud
		indice = indice + tailleNombre
		indice ++
	}

	// Buffer o� sera stock� le nom de l'icone � ajouter
	var icone

	// Le noeud contient au moin un noeud fils
	if( haveAChild )
	{
		// Le noeud est d�velopp�
		if( this.isExpand )
		{
			// Met l'icone indiquant que le noeud peut �tre compact�
			icone = "Moin"
		}
		// Le noeud est compact�
		else
		{
			// Met l'icone indiquant que le noeud peut �tre d�velopp�
			icone = "Plus"
		}
	}
	// Le noeud ne contient aucun noeud fils
	else
	{
		// Met l'icone contenant les pointill�s
		icone = "Pointilles"
	}

	// Le noeud ne poss�de aucun autre noeud en dessous de lui
	if( isEnd )
	{
		// Modifie l'icone pour qu'elle finisse la branche
		icone += "Fin.gif"
	}
	// Le noeud poss�de au moin un noeud en dessous de lui
	else
	{
		// Ajoute l'extension � l'icone
		icone += ".gif"
	}

	// Le noeud poss�de au moin un noeud fils
	if( haveAChild )
	{
		// Cr�e le nom que poss�dera l'icone servant � d�velopper/compacter le noeud
		var nomIcone = "EXPAND_" + this.index

		// Ajoute l'icone en lui indiquant la m�thode � appeler lors d'un clique dessus
		document.write( "<IMG ID = '" + nomIcone + "' STYLE = {cursor:hand;} SRC = 'TreeView/" + icone + "' " )
		document.write( "ONCLICK = 'OnExpand( " + nomIcone + " )'>" )
	}
	// Le noeud ne poss�de aucun fils
	else
	{
		// Ajoute l'icone
		document.write( "<IMG SRC = 'TreeView/" + icone + "'>" )
	}

	// Passe � la colonne suivante et lui met l'icone du noeud dedans
	document.write( "</TD><TD VALIGN = middle nowrap>" )
	document.write( "<IMG ID = 'ICONE_" + this.index + "' SRC = '" + this.icone + "'" )
	document.write( "ONCLICK = '" + this.onClick + "( &quot;" + this.index + "&quot;, &quot;" + this.texte + "&quot; )'" )
	document.write( "ONMOUSEOVER = 'OnOver(&quot;" + this.index + "&quot;)' " )
	document.write( "ONMOUSEOUT = 'OnOut(&quot;" + this.index + "&quot;)'>" )

	// Passe � la colonne suivante o� sera affich� le texte du noeud
	document.write( "</TD><TD VALIGN = middle nowrap>" )

	// Cr�e le nom unique du FONT
	var nomFont = "TEXTE_" + this.index

	// Ajoute le texte dans la table
	document.write( "<FONT ID = '" + nomFont + "'" )
	document.write( "CLASS = '" + this.style + "'" )	// Le curseur prendra la forme d'une main lorsqu'il passera par dessus
	document.write( "COLOR = '#000000'" )
	document.write( "ONCLICK = '" + this.onClick + "( &quot;" + this.index + "&quot;, &quot;" + this.texte + "&quot; )'" )
	document.write( "ONMOUSEOVER = 'OnOver(&quot;" + this.index + "&quot;)' " )
	document.write( "ONMOUSEOUT = 'OnOut(&quot;" + this.index + "&quot;)'>" )
	document.write( this.texte + "</FONT>" )

	// Referme la colonne et la table
	document.write( "</TD></TR></TABLE>" )

	// V�rifie si le noeud poss�de des fils et s'il est d�velopper
	if( haveAChild )
	{
		// Parcourt le tableau de noeud fils
		var indiceNoeud = 0
		while( this.tableauEnfants[ indiceNoeud ] != null )
		{
			// Lance l'affichage du noeud
			this.tableauEnfants[ indiceNoeud ].Draw()

			// Passe au noeud suivant
			indiceNoeud ++
		}
	}
}

// PERMET DE RECUPERER L'INSTANCE D'UN NOEUD D'APRES SON INDEX
// PARAMETRE	: indexNoeud - Chaine contenant l'index du noeud recherch�
// RETOUR		: l'instance sur le noeud, ou 'null' si le noeud n'est pas pr�sent
function Noeud_Noeuds( indexNoeud )
{
	// S'il s'agit du noeud courant
	if( this.index == indexNoeud )
	{
		// Retourne l'instance du noeud
		return this
	}

	// Parcourt le tableau de noeuds fils
	var indice = 0
	while( this.tableauEnfants[ indice ] != null )
	{
		// R�cup�re l'instance sur le noeud
		var noeudRecherche = this.tableauEnfants[ indice ].Noeuds( indexNoeud )

		// V�rifie qu'il ait �t� trouv�
		if( noeudRecherche != null )
		{
			// Retourne l'instance du noeud
			return noeudRecherche
		}

		// Passe au noeud suivant
		indice ++
	}

	// Retourne comme quoi le noeud n'a pas �t� trouv�
	return null
}

// PERMET DE DETERMINER S'IL S'AGIT DU DERNIER NOEUD DE SA BRANCHE
// RETOUR		: bool�en indiquant si le noeud est le dernier de sa branche, ou non
function Noeud_IsEnd()
{
	// Recherche le dernier nombre de l'index du noeud
	var indice = this.index.length - 1
	while(  this.index.substring( indice, indice + 1 ) != "_" )
		indice -- 

	// Cr�e l'index du noeud qui devrait se trouver juste apr�s celui-ci
	var indiceSuivant = parseInt( this.index.substring( indice + 1, this.index.length ), 10 )
	var indexSuivant = this.index.substring( 0, indice + 1 ) + ( ++ indiceSuivant )

	// V�rifie s'il existe dans la treeview
	if( treeView.Contient( indexSuivant ) )
	{
		// Le noeud existe, il ne s'agit donc pas de la fin de la branche
		return false
	}
	else
	{
		// Le noeud n'existe pas, il s'agit donc bien de la fin de la branche
		return true
	}
}

// PERMET DE DETERMINER SI LE NOEUD POSSEDE DES FILS
// RETOUR		: bool�en indiquant si le noeud poss�de ou non des enfants
function Noeud_HaveAChild()
{
	// V�rifie s'il existe un �l�ment dans la premi�re case du tableau de noeuds fils
	if( this.tableauEnfants[ 0 ] != null )
	{
		// Retourne comme quoi le noeud poss�de des fils
		return true
	}
	else
	{
		// Retourne comme quoi le noeud ne poss�de aucun fils
		return false
	}
}

// PERMET DE DETERMINER SI UN NOEUD EST PRESENT DANS LA TREEVIEW, D'APRES SON INDEX
// PARAMETRE	: indexNoeud - Chaine contenant l'index du noeud recherch�
// RETOUR		: bool�en indiquant si le noeud est pr�sent ou non
function Noeud_Contient( indexNoeud )
{
	// Il s'agit du noeud courant
	if( this.index == indexNoeud )
	{
		// Retourne comme quoi le noeud est pr�sent
		return true
	}

	// Parcourt le tableau de noeuds fils
	var indice = 0
	while( this.tableauEnfants[ indice ] != null )
	{
		// V�rifie s'il contient ou s'il est le noeud recherch�
		if( this.tableauEnfants[ indice ].Contient( indexNoeud ) )
		{
			// Retourne comme quoi le noeud est pr�sent
			return true
		}

		// passe au noeud suivant
		indice ++
	}

	// Retourne comme quoi le noeud n'est pas pr�sent dans la treeview
	return false
}

// PERMET DE RENDRE INVISIBLES, LES NOEUDS FILS
function Noeud_Cacher()
{
	// Parcourt le tableau de noeuds fils
	var indice = 0
	while( this.tableauEnfants[ indice ] != null )
	{
		// Cache la table contenant le noeud
		GetElement( this.tableauEnfants[ indice ].table ).style.display = "none"

		// Cache les fils du noeud
		this.tableauEnfants[ indice ].Cacher()

		// Passe au noeud suivant
		indice ++
	}
}

// PERMET DE RENDRE VISIBLES, LES NOEUDS FILS
function Noeud_Montrer()
{
	// V�rifie s'il faut d�velopper ce noeud
	if( ! this.isExpand )
		return

	// Parcourt le tableau de noeuds fils
	var indice = 0
	while( this.tableauEnfants[ indice ] != null )
	{
		// Rend visible la table contenant le noeud
		GetElement( this.tableauEnfants[ indice ].table ).style.display = "block"

		// Rend visibles les fils du noeud
		this.tableauEnfants[ indice ].Montrer()

		// Passe au noeud suivant
		indice ++
	}
}

// PERMET DE CHANGER L'APPARENCE D'UN NOEUD
// A APPELER UNIQUEMENT APR�S LA M�THODE TREEVIEW.START()
// PARAMETRE	: bool�en, indiquant s'il faut d�ployer le noeud ou le compacter
function Noeud_ChangerExpand( isExpand )
{
	// V�rifie si un changement est n�cessaire
	if( this.isExpand == isExpand || ! this.HaveAChild() )
	{
		// Retourne de la fonction sans rien faire
		return
	}

	// R�cup�re l'instance sur l'icone d'extension du noeud
	var icone = GetElement( "EXPAND_" + this.index )
	OnExpand( icone )
}









// M�thode execut�e lorsque l'utilisateur d�veloppe ou compacte un noeud
// PARAMETRE	: ClickedIcone - L'instance sur l'icone concern�e
function OnExpand( ClickedIcone )
{
	// R�cup�re l'index du noeud, puis son instance
	var indexNoeud = ClickedIcone.id.substring( 7, ClickedIcone.id.length )
	var noeudConcerne = treeView.Noeuds( indexNoeud )

	// Cr�e le nom de la table du noeud
	var nomTable = "TABLE_" + indexNoeud

	// Recherche, dans le chamin complet de l'icone, l'indice � partir duquel son nom commence
	var indice = ClickedIcone.src.length - 1
	while( ClickedIcone.src.substring( indice, indice + 1 ) != "/" )
		indice --

	// R�cup�re le nom de l'icone cliqu�e
	var icone = ClickedIcone.src.substring( indice + 1, ClickedIcone.src.length )

	// Met l'icone et le bool�en indiquant si le noeud doit �tre d�velopp�, � jour
	if( icone == "Moin.gif" )
	{
		ClickedIcone.src = "TreeView/Plus.gif"
		noeudConcerne.isExpand = false
	}
	else if( icone == "MoinFin.gif" )
	{
		ClickedIcone.src = "TreeView/PlusFin.gif"
		noeudConcerne.isExpand = false
	}
	else if( icone == "Plus.gif" )
	{
		ClickedIcone.src = "TreeView/Moin.gif"
		noeudConcerne.isExpand = true
	}
	else if( icone == "PlusFin.gif" )
	{
		ClickedIcone.src = "TreeView/MoinFin.gif"
		noeudConcerne.isExpand = true
	}

	// Le noeud � �t� developp�
	if( noeudConcerne.isExpand )
	{
		// Rend les fils du noeud visible
		noeudConcerne.Montrer()
	}
	// Le noeud � �t� compact�
	else
	{
		// Cache les fils du noeud
		noeudConcerne.Cacher()
	}
}

// FONCTION APPELEE LORSQUE LE CURSEUR SE TROUVE AU DESSUS D'UN ELEMENT DE LA TREEVIEW
// PARAMETRE	: indexNoeud - L'index du noeud concern�
function OnOver( indexNoeud )
{
	// R�cup�re l'instance sur le noeud
	var noeud = treeView.Noeuds( indexNoeud )

	// Change la classe de style du texte
	GetElement( "TEXTE_" + indexNoeud ).className = noeud.styleOnOver

	// Compte le nombre de fils du noeud
	var nbrFils = 0
	while( noeud.tableauEnfants[ nbrFils ] != null )
		nbrFils ++

	// Affiche le nombre de fils dans la barre d'�tat du navigateur
	if( nbrFils == 1 )
		window.defaultStatus = "contient " + nbrFils + " �l�ment"
	else if( nbrFils > 0 )
		window.defaultStatus = "contient " + nbrFils + " �l�ments"
}

// FONCTION APPELEE LORSQUE LE CURSEUR SORT D'UN ELEMENT DE LA TREEVIEW
// PARAMETRE	: indexNoeud - L'index du noeud concern�
function OnOut( indexNoeud )
{
	// Change la classe de style du texte
	GetElement( "TEXTE_" + indexNoeud ).className = treeView.Noeuds( indexNoeud ).style

	// Efface le texte contenu dans la barre d'�tat du navigateur
	window.defaultStatus = ""
}

// FONCTION PERMETTANT D'OBTENIR UN ELEMENT DE LA PAGE D'APRES SON ID
function GetElement( idElement )
{
	// Appel la m�thode en fonction du navigateur
	if( document.all )
		return document.all[ idElement ]
	else
		return document.getElementById( idElement )
}