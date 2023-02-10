#include <stdlib.h>
#include <stdio.h>


void main (void)
{


	FILE *pf;

	pf =fopen ("donne.txt","wt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	fputs("Salut les Zér0s\nComment allez-vous ?", fichier);
        fclose(fichier);

	fclose(pf);





	int i;
	FILE *pf;

	pf =fopen ("donne.txt","rt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	fgets(chaine, TAILLE_MAX, fichier); // On lit maximum TAILLE_MAX caractères du fichier, on stocke le tout dans "chaine"
        printf("%s", chaine); // On affiche la chaîne

        

	fclose(pf);

}








