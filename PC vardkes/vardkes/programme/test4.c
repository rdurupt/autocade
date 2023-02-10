#include <stdio.h>
#include <ctime>
#include <stdlib.h>
#include <string.h>
#include <conio.h>

void main (void)
{

FILE *fichier ;
int i ;
float x[25], y ;

/*  ouverture du fichier 'mon_fichier.txt' pour lecture (r) en mode texte (t)  */


fichier = fopen("C:\test\donne.txt", "rt");

/*  en cas d'échec de l'ouverture, le pointeur est NULL: intercepter ce cas  */
 
if (fichier == NULL)
{


    /*  message d'alerte et fin du programme  */
    printf ("impossible de créer le fichier mon_fichier.txt\n") ;
	
    exit (0) ;
}

for (i=0 ; i<25 ; i++) /*  boucle pour lire dans le fichier */
{
    if (fscanf (fichier, "%f", &y) == 1) /* lecture d'un réel */
    {
        x[i] = y; /*  reussi => je le stocke */
    }
    else                                 /*  echec  */
    {
        printf ("erreur ligne %d\n", i) ;/*  message  */
        fclose (fichier) ;               /*  fermeture du fichier  */
        exit (0) ; /*  arret du programme  */
    }
}
fclose (fichier) ; /*  fermeture du fichier  */
printf("toto");


}