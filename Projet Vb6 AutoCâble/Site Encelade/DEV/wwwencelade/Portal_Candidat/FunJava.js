
function MyTrim(V){
 var monTexte = new String("");
  monTexte = V;
  while (monTexte.charAt(0) == ' ') {
    monTexte = monTexte.substring(1,monTexte.length);
  }
  while (monTexte.charAt(monTexte.length - 1) == ' ') {
    monTexte = monTexte.substring(0, (monTexte.length - 1));
  }
  return monTexte;
}

function SaisieDate(F,D){
var MyDate;
 MyDate = document.forms[F].elements[D];
var erreur=0;
erreur=0;
var jour=0;
var mois=0;
var annee=0;
valeurDate=MyTrim(MyDate.value);
indSep1=valeurDate.indexOf("/");
indSep2=valeurDate.lastIndexOf("/");
if ((indSep1 !=-1) && (indSep1 != indSep2)){
   jour=valeurDate.substring(0,indSep1);
    mois=valeurDate.substring(indSep1+1,indSep2);
    annee=valeurDate.substr(indSep2+1);
    if (!testDate(jour,mois,annee)) erreur=1;
    } else {
 erreur=1
}
if (erreur==1) {
    window.alert('Vous devez saisir une date au format jj/mm/aaaaa');
    MyDate.value='';
}

}

function testDate(jour, mois, annee){
var erreur=0;
erreur=0;
valeurJours=jour.toString();
valeurMois=mois.toString();
valeurAnnee=annee.toString();

if ((!testEntier(valeurJours))|| (valeurJours.length !=2)) erreur=1;
if ((!testEntier(valeurMois))|| (valeurMois.length !=2)) erreur=1;
if ((!testEntier(valeurAnnee))|| (valeurAnnee.length !=4)) erreur=1;
var bis=0;
var maxFev;
bis=valeurAnnee % 4;
if(bis==0) maxFev=29;
else maxFev=28;
var maxJours;
maxJours= new Array(31,maxFev,31,30,31,30,31,31,30,31,30,31);
if (valeurJours> maxJours[valeurMois-1]) erreur=1;
if (erreur==0) return true;
else return false;
}
function testEntier(valeur){
if (valeur==parseInt(valeur)){
return true;
} else {
return false;
}
}

function funMath(Frm,Objet_Entree1,Objet_Entree2,Objet_Sortie,Signe,noNegat,EntreReset,Entier){



}

