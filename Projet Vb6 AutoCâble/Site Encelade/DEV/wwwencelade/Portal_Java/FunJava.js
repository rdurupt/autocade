
function MyTrim(V){
 var monTexte = new String('');
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
indSep1=valeurDate.indexOf('/');
indSep2=valeurDate.lastIndexOf('/');
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
    return false;
	
}
return true;
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
if (valeur==null){
return false;
}
if (valeur==parseInt(valeur)){
return true;
} else {
return false;
}
}

function testFloat(ValeurTest){
if (ValeurTest==null){
return false;
}
if (ValeurTest==parseFloat(ValeurTest)) return true;
 else return false;

}
function Remplacer(Txt,C,R){
var a, tmp;
tmp = '';
a = ''+Txt;
if (a.length==0){
return;
}

 for(var i = 0; i < a.length; i++)
{

  
    if (a.charAt(i) ==  C)
    {
    tmp = tmp + R;
    } else{
        tmp = tmp + a.charAt(i);
    }
}
  
return  tmp;

}
function funMath(Val_Entree1,Val_Entree2,Signe,noNegat) {
 var Sortie=0;   
    switch  (Signe){
        case '+' : 
                    Sortie = Val_Entree1 + Val_Entree2;
         break;
        case '-':
                   Sortie = Val_Entree1 - Val_Entree2;
         break;
        case '/' :
                   Sortie=Val_Entree1 * ( 1/ Val_Entree2);
         break;
        case '*':
                   Sortie= Val_Entree1* Val_Entree2;
         break;
    }
    if (noNegat==true) {
        if (Sortie<0){
            alert('La valeur de sortie ne peut être in inférieur  à zéro.');
            Sortie='Err';
        }
    }
	
 return  Sortie;
}



function DateDif(Format,DateDeb,DateFin){
var Jd=0;
var Md=0;
var Ad=0;
indSep1=DateDeb.indexOf('/');
indSep2=DateDeb.lastIndexOf('/');
if ((indSep1 !=-1) && (indSep1 != indSep2)){
   Jd=DateDeb.substring(0,indSep1);
    Md=DateDeb.substring(indSep1+1,indSep2);
    Ad=DateDeb.substr(indSep2+1);
}

indSep1=DateFin.indexOf('/');
indSep2=DateFin.lastIndexOf('/');
if ((indSep1 !=-1) && (indSep1 != indSep2)){
   Jf=DateFin.substring(0,indSep1);
    Mf=DateFin.substring(indSep1+1,indSep2);
    Af=DateFin.substr(indSep2+1);
}

var bis=0;
var maxFev;
var maxJours;
var Incre=true;
var i=0;
if (parseInt(Af)<parseInt(Ad)) {
     Incre=false;
}
if (parseInt(Af)==parseInt(Ad)) {
    if (parseInt(Mf)<parseInt(Md)) {
        Incre=false;
    }
}
if (parseInt(Af)==parseInt(Ad)) {
   if (parseInt(Mf)==parseInt(Md)){
        if (parseInt(Jf)<parseInt(Jd)) {
            Incre=false;
        }
    }
}



switch  (Format){
    case 'j' :
        bis=parseInt(Ad) % 4;
        if(bis==0) maxFev=29;
        else maxFev=28;
        maxJours= new Array(31,maxFev,31,30,31,30,31,31,30,31,30,31);
      if (parseInt(Jd)<10) {
            DateDeb='0' + parseInt(Jd);
       } else {
            DateDeb=+ parseInt(Jd);
        }
       if (parseInt(Md)<10) {
          DateDeb=DateDeb +'/0'+ parseInt(Md);
        } else {
            DateDeb=DateDeb + '/'+parseInt(Md);
        }
           DateDeb=DateDeb+ '/'+parseInt(Ad);
        while (DateDeb!== DateFin)
         {
        if(Incre==true) {
                    i++;
                    Jd++;
                    if(parseInt(Jd)>maxJours[parseInt(Md)-1]) {
                        Jd=1;
                        Md++;
                    }
                    if (parseInt(Md)>12) {
                        Md=1;
                        Ad++;
                        bis=Ad % 4;
                        if(bis==0) maxFev=29;
                        else maxFev=28;
                        maxJours= new Array(31,maxFev,31,30,31,30,31,31,30,31,30,31);
                    }
        } else {
                    i--;
                    Jd--;
                    if(parseInt(Jd)<1) {
                        Md--;
                        Jd=maxJours[parseInt(Md)-1];
                    }
                    if (parseInt(Md)<1) {
                        Md=12;
                        Ad--;
                        bis=Ad % 4;
                        if(bis==0) maxFev=29;
                        else maxFev=28;
                        maxJours= new Array(31,maxFev,31,30,31,30,31,31,30,31,30,31);
                        Jd=maxJours[parseInt(Md)-1];
                    }
        }
        
            if (parseInt(Jd)<10) {
                DateDeb='0' + parseInt(Jd);
            } else {
                DateDeb=+ parseInt(Jd);
            }
            if (parseInt(Md)<10) {
                DateDeb=DateDeb +'/0'+ parseInt(Md);
            } else {
                DateDeb=DateDeb + '/'+parseInt(Md);
            }
            DateDeb=DateDeb+ '/'+parseInt(Ad);
           }
        break;
    }
 return i;
}
