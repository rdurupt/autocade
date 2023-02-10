#include <stdio.h>

charger(FILE *);

void main (void)
{
	FILE *pf;

	pf =fopen ("c:\\test\\donne.txt","wt");	
	
	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}



	fputs("manger",pf);

	charger (&*pf);
	
	fclose(pf);

}

charger(FILE *pf)
{

	char chaine[10] ;

	

	
	fputs("bonjour my borther",pf);

	gets(chaine);

	fputs(chaine,pf);

}

