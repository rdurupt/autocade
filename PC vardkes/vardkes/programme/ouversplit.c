#include <stdlib.h>
#include <stdio.h>
#define TAILLE_MAX    1000
#define N 1000
#define MAXSTR 256



void main (void)
{
	FILE *pf;
	char chaine[TAILLE_MAX];
	char fichier[N];
	int i=0,j=0,k=0,l=0,c=0,m=0,fin_ligne=0,nbr_colonne=0,nbr_ligne=0;
	int a=0,b=0,d=0,z=0;
	char str[256][10][256] ;
	char caractereActuel;
	
	  
	 

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
		chaine[i]='\0';
       	fclose(pf);


	for (j=0;j<=strlen(chaine);j++)
	{
		if(fin_ligne== 0)
		{
			if(chaine[j]==';')
			{
				nbr_colonne++;

			}
		}
		
		if(chaine[j]=='\n')
		{
			fin_ligne=1;
			nbr_ligne++;
		}
		

	}

	nbr_ligne=nbr_ligne-2;

	
	 

	for (k=0;k<=strlen(chaine);k++)
	{
		if (chaine[k]!=';' && chaine[k]!='\n')
		{
			str[l][c][m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
				str[l][c][m]='\0';
				c++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				str[l][c][m]='\0';
				l++;
				c=0;
				m=0;
				
			}
		}
	}
	for (a=0;a<=nbr_ligne;a++)	/* faire un do*/
	{
		for(b=0;b<=nbr_colonne;b++)
		{
			do
			{
				printf("%c",str[a][b][d]);
				d++;
			}
				while (str[a][b][d]!='\0');
			
				d=0;
				printf("|");
			 	
		}
		printf("\n");

	}
				
				
	  
}














	


