#include <stdio.h>
#include <conio.h>
#include <stdlib.h>

/**********************************************************/
/* fonction qui ouvre un fichier */
/**********************************************************/
FILE *MyFileOpen(char *name,char *mode)
{
FILE *f;

/* on ouvre le fichier */
f = fopen(name,mode);


printf("Ouverture du fichier '%s' en mode '%s' : ",name,mode);

/* s'il y a eu une erreur */
if(f == NULL)
	{
	printf("echouee !\n");
	/* on quitte */
	getch();
	exit(1);
	}

/* l'ouverture a reussi */
printf("reussie !\n");
return f;
}

/**********************************************************/
/* on ferme le fichier */
/**********************************************************/
void MyFileClose(FILE *f)
{
printf("Fermeture du fichier : ");

/* on ferme le fichier et on teste sa valeur de retour */
if(fclose(f) == EOF)
	{
	printf("echouee !\n");
	/* on quitte */
	getch();
	exit(1);
	}

/* la fermeture a reussi */
printf("reussie !\n");
}



/**********************************************************/
/**********************************************************/
/**********************************************************/
int main(int argc,char **argv)
{
FILE *f;
char feld[10];
char fichier[128] ;


printf("------------------------------------\n");

/* TOUT MARCHE BIEN */
f = MyFileOpen("C:\\Documents and Settings\\vardkes.aslanian\\Bureau\\vardkes\\coucou.txt","wt");
printf("Ici vous ecrivez ce que vous voulez ...\n");
MyFileClose(f);

/* NE MARCHE PAS */
printf( "\nEntrez le fichier avec son chemin : " ) ;
scanf( "%s" , fichier ) ;




f = MyFileOpen(fichier,"rt");
MyFileClose(f);



getch();
return 0;
}

