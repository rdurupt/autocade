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

	fputs("Salut les Z�r0s\nComment allez-vous ?", fichier);
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

	fgets(chaine, TAILLE_MAX, fichier); // On lit maximum TAILLE_MAX caract�res du fichier, on stocke le tout dans "chaine"
        printf("%s", chaine); // On affiche la cha�ne

        

	fclose(pf);

}








