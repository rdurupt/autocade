#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <conio.h>

#define NMAX      128
#define NBD       10
#define NOM_MAX	  20
#define ISBN_MAX  15



typedef struct s_BD {

	char titre[NMAX] ;
	char serie[NOM_MAX] ;
	int  nserie ;
	char  dessinateur[NOM_MAX] ;
	char  scenariste[NOM_MAX] ;
	int  annee ;
	char  editeur[NOM_MAX];
	int  pages;
	float  prix ;
	char isbn[ISBN_MAX] ;

}BD;

int menu(void);

void affiche(BD biblio [NBD],int bd_present);
void insert(BD biblio [NBD],int *bd_present);
void supprim(BD biblio [NBD],int *bd_present);
void trier(void);
void save(BD biblio [NBD],int bd_present);
void load(BD biblio [NBD],int *bd_present);





void main(void)

{
	BD biblio [NBD];
	int choix,bd_present=0;

	 do {
		 choix=menu();
		 switch (choix)

	  {
		  case 1 :

		  affiche(biblio,bd_present);

		  break ;



		  case 2 :

		  insert(biblio,&bd_present);

		  break ;


		  case 3 :

		  supprim(biblio,&bd_present);

		  break ;




		  case 4 :

		  trier();

		  break ;



		  case 5 :

		  save(biblio,bd_present);

		  break ;




		  case 6 :

		  load(biblio,&bd_present);

		  break ;



		   case 7 :

		   exit(0);

		   break ;

	}

	} while (choix !=7);
}


int menu(void)
{

  int choix;






  printf("-------------------------\n");

  printf("gestions des BDs\n");

  printf("-------------------------\n");

  printf("1:visualiser l'ensemble des elements de la base de donnée\n");

  printf("2:inserer un element\n");

  printf("3:supprimer un element\n");

  printf("4:trier les elements par ordre alphabetique des noms de serie\n");

  printf("5:sauvegarder la BD dans un fichier\n");

  printf("6:charger la BD a partir d'un fichier\n");

  printf("7:quittez\n");




   printf("choix:");



   scanf("%d",&choix);











return (choix);

}









void affiche(BD biblio [NBD],int bd_present)
{
	int i;

	for (i=0;i<bd_present;i++)

	{

	printf("\ntitre:");
	puts(biblio [i].titre);



    printf("\n numero de serie:%d",biblio [i].nserie);



  	printf("\n dessinateur:");
	puts(biblio [i].dessinateur);


	}


}





void insert(BD biblio [NBD], int *bd_present)
{
	int i,a;

	char nl;

	printf("\ncombien de bd allez vous entrer\n");

	scanf("%d%c",&a,&nl);



	for (i=*bd_present;i<(a+*bd_present);i++)
	{

		printf("entrer les informations de la %d eme bd",(i+1));


		printf("\ntitre:");

		gets(biblio[i].titre) ;

/*
		printf("\nnom de serie:");

		gets(biblio[i].serie) ;

*/
		printf("\nnumero de serie:");

		scanf("%d%c",&biblio[i].nserie,& nl);


		printf("\ndessinateur:");

		gets(biblio[i].dessinateur) ;


/*		printf("\nscenariste:");

		gets(biblio[i].scenariste) ;


		printf("\nannee:");

		scanf("%d%c",&biblio[i].annee,& nl);


		printf("\nediteur:");

		gets(biblio[i].editeur);


		printf("\nnombre de page:");

		scanf("%d%c",&biblio[i].pages,&nl);


		printf("\nprix:");

		scanf("%f%c",&biblio[i].prix,& nl);


		printf("\nnumero isbn:");

	    gets(biblio[i].isbn) ;
*/

		printf("\n");






	}

*bd_present=*bd_present+a;

}


void supprim(BD biblio [NBD],int *bd_present)
{
	int i,a=-1,j;

	char nl,effacer[NMAX];

	printf("\nquel est le titre de la BD que voulez vous effacer\n");

	scanf("%c",&nl);

	gets(effacer);



	for (i=0;i<*bd_present;i++)

	{
		if (strcmp(biblio[i].titre, effacer) == 0)
		{
			a=i;

		}
	}

	if(a!=-1)

	{
		for (j=a+1;j<*bd_present;j++)
		{
			biblio[j-1]=biblio[j];



		}
	*bd_present=*bd_present-1;
	}

}


void trier(void)
{


	printf("je suis dans trier\n");

	getch();

	}


void save(BD biblio [NBD],int bd_present)
{

	FILE *pf;

	pf =fopen ("donne.txt","wt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	fwrite (&bd_present,sizeof(int),1,pf );

	fwrite (biblio,sizeof(BD),bd_present,pf );

	fclose(pf);


}


void load(BD biblio [NBD],int *bd_present)
{
	int i;
	FILE *pf;

	pf =fopen ("donne.txt","rt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	fread (bd_present,sizeof(int),1,pf);

	for(i=0;i<*bd_present;i++)

	fread (&biblio[i],sizeof(BD),1,pf);

	fclose(pf);

}


































































