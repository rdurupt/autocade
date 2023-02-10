#include <stdio.h>
#include <ctime>
#include <stdlib.h>
#include <string.h>
#include <conio.h>

#define TAILLE_MAX  256
#define N 256

void main (void)
{

	FILE *pf;
	char chaine[TAILLE_MAX];
	char fichier[N];
	char caractereActuel;
	int i = 0;
	

/*	pf =fopen ("C:\\test\\dede.txt","wt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	fputs("Salut les zero\nComment allez-vous \n?", pf);
	printf("ecriture\n");
        

	fclose(pf);*/

	
	gets(fichier);
	

	pf =fopen (fichier,"rt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	// Boucle de lecture des caractères un à un
        do
        {
            caractereActuel = fgetc(pf); // On lit le caractère
            printf("%c", caractereActuel); // On l'affiche
			chaine[i]=caractereActuel ;
			i++;
        } while (caractereActuel != EOF); // On continue tant que fgetc n'a pas retourné EOF (fin de fichier)

       	fclose(pf);
    




}

