<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

 <HTML>
 <HEAD>
 <TITLE>TreeView en JavaScript</TITLE>
 <META NAME="Author" CONTENT="Golinski Ludwig">
 <STYLE> .Racine { font-weight: bold; font-size: 12px; cursor: default; color: gray; font-family: Tahoma; }
 .RacineOver { font-weight: bold; font-size: 12px; cursor: default; color: black; font-family: Tahoma; }
 .Poste { font-size: 12px; cursor: default; color: gray; font-family: Tahoma; }
 .PosteOver { font-size: 12px; cursor: default; color: black; font-family: Tahoma; }
 .Noeuds { font-size: 12px; cursor: default; color: gray; font-family: Tahoma; }
 .NoeudsOver { font-size: 12px; cursor: default; color: black; font-family: Tahoma; }
 .BordTreeView { border: black thin solid; position: absolute; left: 0px; top: 0px; overflow: auto; width: 250px; height: 400px; }
 .BordHidden { visibility: hidden; }
 </STYLE>
 <SCRIPT SRC="TreeView.js"></SCRIPT>
 </HEAD>
 <BODY oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
 <SCRIPT language=javascript>
 // Voici la méthode qui sera appelé lors d'un click sur un fichier
 // doit obligatoirement prendre l'index puis le texte du noeud en paramétre
 function OnClickFichier( index, texte )
 {
 alert( "click sur le fichier : " + texte )
 }

 // Voici la méthode qui sera appelé lors d'un click sur un lien
 function OnClickLien( index, texte )
 {
 if( texte == "Codes-Sources" )
 location.href = "http://www.codes-sources.com/"

 else if( texte == "Javascriptfr" )
 location.href = "http://www.Javascriptfr.com/"

 else if( texte == "Google" )
 location.href = "http://www.Google.com/"
 }

 // Celle-là sera appelée lors d'un click sur un dossier
 function OnClickDossier( index, texte )
 {
 treeView.Noeuds( index ).ChangerExpand( ! treeView.Noeuds( index ).isExpand )
 }

 // Instancie la treeView en lui indiquant les styles du contour et de la racine
 treeView = new TreeView( "BordTreeView", "Images/Réseau.gif", "Réseau", "Racine", "RacineOver", "" )

 for( var indicePoste = 1; indicePoste <= 4; indicePoste ++ )
 {
 // Ajoute un noeud
 var poste = treeView.Add( "Images/Poste.gif", "Disque n°" + indicePoste, "Poste", "PosteOver", "OnClickDossier" )

 switch( indicePoste )
 {
 case 1 :
 for( var indiceDossier = 1; indiceDossier <= 3; indiceDossier ++ )
 {
 // Ajoute un sous-noeud au noeud 'Poste'
 var dossier = poste.Add( "Images/Dossier.gif", "Dossier n°" + indiceDossier, "Poste", "PosteOver", "OnClickDossier" )

 if( indiceDossier == 3 )
 dossier.Add( "Images/Fichier.gif", "Fichier", "Noeuds", "NoeudsOver", "OnClickFichier" )
 }
 break

 case 2 :
 poste.Add( "Images/Dossier.gif", "Dossier vide", "Poste", "PosteOver", "OnClickDossier" )
 break

 case 3 :
 var dossier = poste.Add( "Images/Dossier.gif", "Dossier", "Poste", "PosteOver", "OnClickDossier" )

 for( var indiceFichier = 1; indiceFichier <= 5; indiceFichier ++ )
 dossier.Add( "Images/Fichier.gif", "Fichier n°" + indiceFichier, "Noeuds", "NoeudsOver", "OnClickFichier" )

 var autredossier = dossier.Add( "Images/Dossier.gif", "Dossier", "Poste", "PosteOver", "OnClickDossier" )

 for( var indiceFichier = 1; indiceFichier <= 5; indiceFichier ++ )
 autredossier.Add( "Images/Fichier.gif", "Fichier n°" + indiceFichier, "Noeuds", "NoeudsOver", "OnClickFichier" )

 dossier.Add( "Images/Dossier.gif", "Dossier vide", "Poste", "PosteOver", "OnClickDossier" )

 break

 case 4 :
 var dossier = poste.Add( "Images/Dossier.gif", "Liens", "Poste", "PosteOver", "OnClickDossier" )

 dossier.Add( "Images/Fichier.gif", "Javascriptfr", "Noeuds", "NoeudsOver", "OnClickLien" )

 dossier.Add( "Images/Fichier.gif", "Codes-Sources", "Noeuds", "NoeudsOver", "OnClickLien" )

 dossier.Add( "Images/Fichier.gif", "Google", "Noeuds", "NoeudsOver", "OnClickLien" )

 break
 }
 }

 // Lance la création de la treeView
 treeView.Start()

 // Referme les noeuds principaux
 treeView.Noeuds( "0_0_2" ).ChangerExpand( false )

 treeView.Noeuds( "0_2_0_5" ).ChangerExpand( false )

 treeView.Noeuds( "0_3_0" ).ChangerExpand( false )

 for( var indicePoste = 0; indicePoste < 4; indicePoste ++ )
 treeView.Noeuds( "0_" + indicePoste ).ChangerExpand( false )

 </SCRIPT>
 </BODY>
 </HTML> 