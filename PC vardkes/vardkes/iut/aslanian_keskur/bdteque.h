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
void trier(BD biblio [NBD],int bd_present);
void save(BD biblio [NBD],int bd_present);
void load(BD biblio [NBD],int *bd_present);



