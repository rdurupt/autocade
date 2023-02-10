#include <conio.h> // Fichier d'en-tête pour l'instruction "_outp"

#include <stdio.h> // Fichier d'en-tête pour E/S standards

#include <time.h> // Fichier d'en-tête pour les instruction "clock_t" et "endwait"

 

   void wait ( int seconds ) //Fonction "Wait" (attente)
 
{
 
clock_t endwait;
 
endwait = clock () + seconds * 1 ;
 
while (clock () < endwait) { }
 
}
 
 
 
int main ( ) // Fonction principale
 
{
 
int i; // Variable de boucle
 
               int valeur; // Variable saisie au clavier
 
 
 
printf("Entrez une valeur limite: ");
 
               scanf ("%d", &valeur) ;  // Saisie de la variable "valeur"
 
 
 
 
 
                // Mise à 0 du compteur
 
               _outp (0x378, 00) ;
 
               _outp (0x378, 00) ;
 
 
 
 
 
                
 
for (i=0; i<valeur+1; i++)
 
{
 
_outp (0x378, 00);
 
_outp (0x378, 01);
 
    wait (1000); // Appel de fonction
 
}
 
 
 
return 0;
 
}
 
