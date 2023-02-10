/***************************************************************************
			Simulation de protocoles
				sous OS XP
		TOUS DROITS RESERVES A LA SOCIETE E.T.E.P.
****************************************************************************
 Nom du programme  : portpara.cpp
 CODE PROJET	   : TMU
 Appareil          : DTMUX_TO
 Librairies        : WinIo.lib
 Date Modification : 09 Février 2004
****************************************************************************
 Details           : Generation Protocole en logique TTL sur port PRN
****************************************************************************/
#pragma hdrstop 
#include <windows.h> 
#include <stdio.h> 
#include <conio.h> 
//#include <math.h>
#include "./Source/Dll/WinIo.h"
#include <time.h>
#include <sys/types.h>
#include <sys/timeb.h>
#include <string.h>

#define SUCCES 1
#define ERREUR 0

typedef struct {
				int enrg;
				int arret;
				}PTITCYCLE;

PTITCYCLE T[10];

 		

/*Fonction pour mettre à un ou zéro un bit de sortie du port parallele*/
void ecritPRN(int bit)
{
	_outp(888,bit);
}

/* Programme de création d'une copie du fichier toto */

int test_fichier (FILE *in1,FILE *in2) 
{
	char car;
	if ((in1) == NULL) 
	{
		printf( "Impossible d'ouvrir %s en lecture.\n", in1);
		return 1;
	}
	if ((in2 ) == NULL) 
	{
		printf( "Impossible d'ouvrir %s en écriture.\n", in2);
		return 2;
	}
/* Parcours du fichier en lecture, et recopie */
	car = fgetc(in1);
	while (!feof(in1)) 
	{
		fputc( car, in2);
		car = fgetc(in1);
	}
	fclose(in1);
	fclose(in2);
	return 0;
}

void tempo(int s)	//tempo de 1 seconde
{
	short i;
	for(i=0;i<=s;i++)
	{
		Sleep(900);
	}
}

void tempomn(int mn)	//tempo de 1 minute
{
	unsigned short i;
	for(i=0;i<=mn;i++)
	{
			Sleep(56000);
	}
}

void tempo5mn(int mn)
{
	unsigned short i;
	for(i=1;i<=mn;i++)
	{
	Sleep(500000);
	}
}

void tempo15mn(int m)
{
	unsigned short i;
	for(i=1;i<=m;i++)
	{
		Sleep(500000);
		Sleep(500000);
		Sleep(500000);
	}
}

void tempoh(int h)	//tempo de 1 heure
{
	short i;
	for(i=0;i<=h;i++)
	{
			tempo15mn(4);
	}
}



/*
*************************************************************************
		PROGRAMME PRINCIPAL
 *************************************************************************
 */


void menu( void )
{
	int i,j,choix,ch,nbcycles,scycle;
	char a;
	unsigned char barre;
	char tmpbuf[10],tmpbuf2[10];
	char numero[10];
	char nomessai[20],titre[20];
	char *type = ".txt";
	FILE *inpor;

	barre=0xDB;	//création du caractère pour le barre-graphe

	InitializeWinIo(); //libération de l'accès au port parallele
	
	//printf("initialisation terminee\n");

	 _strtime( tmpbuf );	//récupération de l'heure
    printf( "OS time:\t\t\t\t%s\n", tmpbuf );
    _strdate( tmpbuf2 );	//récupération de la date
    printf( "OS date:\t\t\t\t%s\n", tmpbuf2 );


	choix=0;
	do{
		printf("\nSELECTION DU PROTOCOLE A GENERER ...");
		printf("\n1 = Lancement de cycles a definir\n");
		printf("\n2 = Lancement de cycles equivalent ramene sur une heure\n");
		printf("\n0 = QUITTER");
		printf("\nVotre Selection ..................:\n");
		scanf("%d",&choix);
		getchar();

		switch( choix )
		{
			case 1:
					printf("Combien de cycles doivent-ils etre execute a la suite?\n");
					scanf("%i",&nbcycles);
					printf("\nVous allez rentrer les temps pour definir un cycle\n");
					printf("\nIl sera execute %i fois\n",nbcycles);

					printf("\nCombien de sous-cycles voulez vous?(10 max)\n");
					printf("Un sous-cycle est compose :\n");
					printf("d'un enregistrement continu et d'un arret\n");
					scanf("%i",&scycle);

					for(j=1;j<=scycle;j++)
					{
						if(j==1)
							printf("Configuration du %i er cycle\n",j);
						else
							printf("Configuration du %i eme cycle\n",j);

						printf("Entrez la duree en mn d'un enregistrement continu\n");
						scanf("%i",&T[j].enrg);
						printf("Entrez la duree en mn d'un arret\n");
						scanf("%i",&T[j].arret);
					}


					printf("Donnez un nom au fichier de compte rendu(20 caracteres max)\n");
					scanf("%s",&titre);
					printf("Entrez le nom de cet essai pour ce(s) %i cycle(s) (20 caracteres max)\n",nbcycles);
					scanf("%s",&nomessai);
					printf("Entrez le numero de serie de l'appareil (10 caracteres max)\n");
					scanf("%s",&numero);

					
					strcat(titre,type);
					inpor=fopen(titre, "w+t"); //création du fichier en écriture

					fprintf(inpor,"etep");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Centre d'AFFAIRES  GRAND VAR");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Bâtiment A");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"83130 LA GARDE");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"TEL:  04 94 08 50 26");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"FAX: 04 94 08 28 03");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Compte Rendu d'essai d'utilisation");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Démarrage le ");
					fprintf(inpor,tmpbuf2);
					fprintf(inpor," à ");
					fprintf(inpor,tmpbuf);
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Nom de cet essai : ");
					fprintf(inpor,nomessai);
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Numéro de série : ");
					fprintf(inpor,numero);
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Description d'un cycle :");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					
					for(i=1;i<=scycle;i++)
					{
						fprintf(inpor,"%i",i);
						fprintf(inpor,"\t enregistrement continu : %i minutes\n",T[i].enrg);
						fprintf(inpor,"\t arrêt : %i minutes\n",T[i].arret);
					}

					i=1;
					do{
						fprintf(inpor,"\nCycle %i:",i);

						if(i==1)
							printf("\nExecution du %i er cycle\n",i);
						else
							printf("\nExecution du %i eme cycle\n",i);
						fprintf(inpor,"\n");
						for(j=1;j<=scycle;j++)
						{
							
							_strtime( tmpbuf );		//récupération de l'heure
							_strdate( tmpbuf2 );	//récupération de la date
							fprintf(inpor,"\tLancement sous-cycle %i  à %s le %s.",j,tmpbuf2,tmpbuf);
							fprintf(inpor,"\n");

							ecritPRN(2);
							printf("%c",barre);
							tempomn(T[j].enrg);
							ecritPRN(0);
							printf("%c",barre);
							tempomn(T[j].arret);

							_strtime( tmpbuf );		//récupération de l'heure
							_strdate( tmpbuf2 );	//récupération de la date
							fprintf(inpor,"\tLe sous-cycle %i terminé à %s le %s.",j,tmpbuf2,tmpbuf);
							fprintf(inpor,"\n");
						}

						fprintf(inpor,"\nFin du cycle %i:",i);
						fprintf(inpor,"\n");
						i++;
					}while(i<=nbcycles);
					
					_strtime( tmpbuf );		//récupération de l'heure
					_strdate( tmpbuf2 );	//récupération de la date

					fprintf(inpor,"\nCet essai de %i cycle(s) d'utilisation s'est déroulé entièrement",nbcycles);
					fprintf(inpor,"\n");
					fprintf(inpor,"avec succès et a fini à ");
					fprintf(inpor,tmpbuf);
					fprintf(inpor," le ");
					fprintf(inpor,tmpbuf2);
					fprintf(inpor,".");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\t Pour le service technique :");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\t Nom :");
					fprintf(inpor,"\t\t\t Signature :");
					fprintf(inpor,"\t\t\t Date :");
					fprintf(inpor,"\r\n");
					printf("\nCet essai de %i cycle(s) d'utilisation est fini avec succes.\n",nbcycles);
					
					printf("Voulez vous ouvrir ou imprimer le compte rendu?\n");
					printf("1: Ouvrir le fichier %s \n",titre);
					printf("2: Imprimer le fichier %s \n",titre);
					printf("0: Quitter.\n");
					scanf("%d",&ch);
					if(ch==1)
						ShellExecute(NULL, "open", titre, NULL, NULL, SW_SHOWDEFAULT );
					else
						if(ch==2)
							ShellExecute(NULL, "print", titre, NULL, NULL, SW_SHOWDEFAULT );

					fclose(inpor);
					break;

			case 2:
					printf("Combien de cycles doivent-ils etre execute a la suite?\n");
					scanf("%i",&nbcycles);
					printf("\nLancement du cycle equivalent ramene sur une heure\n");
					printf("\nIl sera execute %i fois\n",nbcycles);
					printf("Donnez un nom au fichier de compte rendu(20 caracteres max)\n");
					scanf("%s",&titre);
					printf("Entrez le nom de cet essai pour ce(s) %i cycle(s) (20 caracteres max)\n",nbcycles);
					scanf("%s",&nomessai);
					printf("Entrez le numero de serie de l'appareil (10 caracteres max)\n");
					scanf("%s",&numero);

					
					strcat(titre,type);
					inpor=fopen(titre, "w+t"); //création du fichier en écriture

					fprintf(inpor,"etep");
					fprintf(inpor,"\n");
					fprintf(inpor,"Centre d'AFFAIRES  GRAND VAR");
					fprintf(inpor,"\n");
					fprintf(inpor,"Bâtiment A");
					fprintf(inpor,"\n");
					fprintf(inpor,"83130 LA GARDE");
					fprintf(inpor,"\n");
					fprintf(inpor,"TEL:  04 94 08 50 26");
					fprintf(inpor,"\n");
					fprintf(inpor,"FAX: 04 94 08 28 03");
					fprintf(inpor,"\n");
					fprintf(inpor,"\n");
					fprintf(inpor,"Compte Rendu d'essai d'utilisation");
					fprintf(inpor,"\n");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Démarrage le ");
					fprintf(inpor,tmpbuf2);
					fprintf(inpor," à ");
					fprintf(inpor,tmpbuf);
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Nom de cet essai : ");
					fprintf(inpor,nomessai);
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Numéro de série : ");
					fprintf(inpor,numero);
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"Description d'un cycle :");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					
					for(i=1;i<=4;i++)
					{
						fprintf(inpor,"%i",i);
						fprintf(inpor,"\t enregistrement continu : 5 minutes");
						fprintf(inpor,"\r\n");
						fprintf(inpor,"\t arrêt : 5 minutes");
						fprintf(inpor,"\r\n");
					}
					fprintf(inpor,"5");
					fprintf(inpor,"\t enregistrement continu : 15 minutes");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\t arrêt : 5 minutes");

					i=1;
					do{
						fprintf(inpor,"\nCycle %i:",i);

						if(i==1)
							printf("\nExecution du %i er cycle\n",i);
						else
							printf("\nExecution du %i eme cycle\n",i);
						fprintf(inpor,"\n");
						for(j=1;j<=4;j++)
						{
							
							_strtime( tmpbuf );		//récupération de l'heure
							_strdate( tmpbuf2 );	//récupération de la date
							fprintf(inpor,"\tLancement sous-cycle %i  à %s le %s.",j,tmpbuf2,tmpbuf);
							fprintf(inpor,"\n");

							ecritPRN(2);
							printf("%c",barre);
							tempo5mn(1);
							ecritPRN(0);
							printf("%c",barre);
							tempo5mn(1);

							_strtime( tmpbuf );		//récupération de l'heure
							_strdate( tmpbuf2 );	//récupération de la date
							fprintf(inpor,"\tLe sous-cycle %i terminé à %s le %s.",j,tmpbuf2,tmpbuf);
							fprintf(inpor,"\n");
						}

						_strtime( tmpbuf );		//récupération de l'heure
						_strdate( tmpbuf2 );	//récupération de la date
						fprintf(inpor,"\tLancement sous-cycle 5  à %s le %s.",tmpbuf2,tmpbuf);
						fprintf(inpor,"\n");

						ecritPRN(2);
						printf("%c",barre);
						tempo15mn(1);
						printf("%c",barre);
						ecritPRN(0);
						tempo5mn(1);

						_strtime( tmpbuf );		//récupération de l'heure
						_strdate( tmpbuf2 );	//récupération de la date

						fprintf(inpor,"\tLe sous-cycle 5 terminé à %s le %s.",tmpbuf2,tmpbuf);
						fprintf(inpor,"\n");
						fprintf(inpor,"\nFin du cycle %i:",i);
						fprintf(inpor,"\n");
						i++;
					}while(i<=nbcycles);
					
					_strtime( tmpbuf );		//récupération de l'heure
					_strdate( tmpbuf2 );	//récupération de la date

					fprintf(inpor,"\nCet essai de %i cycle(s) d'utilisation s'est déroulé entièrement",nbcycles);
					fprintf(inpor,"\n");
					fprintf(inpor,"avec succès et a fini à ");
					fprintf(inpor,tmpbuf);
					fprintf(inpor," le ");
					fprintf(inpor,tmpbuf2);
					fprintf(inpor,".");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\t Pour le service technique :");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\r\n");
					fprintf(inpor,"\t Nom :");
					fprintf(inpor,"\t\t\t Signature :");
					fprintf(inpor,"\t\t\t Date :");
					fprintf(inpor,"\r\n");
					printf("\nCet essai de %i cycle(s) d'utilisation est fini avec succes.\n",nbcycles);
					
					printf("Voulez vous ouvrir ou imprimer le compte rendu?\n");
					printf("1: Ouvrir le fichier %s \n",titre);
					printf("2: Imprimer le fichier %s \n",titre);
					printf("0: Quitter.\n");
					scanf("%d",&ch);
					if(ch==1)
						ShellExecute(NULL, "open", titre, NULL, NULL, SW_SHOWDEFAULT );
					else
						if(ch==2)
							ShellExecute(NULL, "print", titre, NULL, NULL, SW_SHOWDEFAULT );

					fclose(inpor);
					break;
		
		default : break;

		}
	if (choix==0)
	{
		printf("\nVoulez-vous reellement quitter? (o/n)\n");
		scanf("%c",&a);
		getchar();
	}
	else
		a='n';

	}while((a=='n')||(a=='N'));

	ShutdownWinIo();	//désactivation du contrôle sur le port parallèle

}

int main (void)
{
	menu();
	return EXIT_SUCCESS;
}
/*
 *************************************************************************
			FIN DU PROGRAMME portpara.cpp
 **************************************************************************
*/