#include <stdio.h>
#include <ctime>
#include <stdlib.h>
//#include <string.h>
#include <conio.h>


#define MAX	 10
#define TAILLE_MAX    8000
#define N 30   
#define LIGNEMAX  140
#define RMAX	 11
#define MAPMAX 260
#define ERREURMAX  30




typedef struct s_fil
{
	char liason[MAX] ;
	char app1[MAX] ;
	char voie1[MAX] ;
	char ref1[RMAX];
	char position1[MAX];
	char app2[MAX] ;
	char voie2[MAX] ;
	char ref2[RMAX];
	char position2[MAX];
	char section[MAX];
	char couleur[MAX];
	char resultat[MAX];


	
}fil;

typedef struct s_mapping
{	
	char ref[11] ;
	char voie[MAX] ;
	char nom[MAX];

}mapping;




void emprunte (void);
void charger (fil lstfil [LIGNEMAX],mapping map [MAPMAX],int*,char*,char* ,char*);
void testcalculo (fil lstfil[LIGNEMAX],mapping map[MAPMAX],int lsterreur[ERREURMAX],int*,int*, int* );
unsigned char lire_port(unsigned short adresse_reg);
void sleep (int nbr_seconds);
void present(char*);
void affiche(fil);
int  classement(fil lstfil[LIGNEMAX],int*, char* , char*  ,char* );
void testaffiche(fil);
void afftestok(fil*);
void afferreur(fil*);
void nomme (fil lstfil [LIGNEMAX],mapping map [MAPMAX],int* );
void alpha (fil lstfil[LIGNEMAX], int ) ;
void testfaisceau(fil lstfil [LIGNEMAX],int lsterreur[ERREURMAX], int* , int*,int*);
void afftestfaisseau(fil lstfil) ;











