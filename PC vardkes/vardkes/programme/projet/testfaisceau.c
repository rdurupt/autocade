#include "projet.h"

void testfaisceau(fil lstfil [LIGNEMAX],int lsterreur[ERREURMAX], int* nbr_ligne_total, int* nbr_ligne, int* nbr_erreur)
{
	int i,cpt=0;
	char test=0,reception=0;

	for (i=*nbr_ligne;i<=*nbr_ligne_total;i++)
	{
		if(strcmp(lstfil[i].resultat, "ok") != 0 )
		{
		
			cpt=0;
			while (cpt<3)
			{
				afftestfaisseau(lstfil[i]) ;
				
				//while(libre=='1');

				//donnee=lire_port(adresse_reg);
			//	printf("entrez");
			//	scanf("%x",&reception);
		
				if( reception == 0)
				{
					afftestok(&lstfil[i]);
					cpt=5;
				}
	
				else
				{

					afferreur(&lstfil[i]);
					cpt++;
					if (cpt==3)
					{
						lsterreur[*nbr_erreur]=i;
						*nbr_erreur=*nbr_erreur+1;
					}
				}
			
			}
		}
	}
}


void afftestfaisseau(fil lstfil)
{
	

		printf("|");	
		printf("%s",lstfil.position1);
		present(lstfil.position1);
		printf("|");

		printf("%s",lstfil.app1);
		present(lstfil.app1);
		printf("|");

		printf("%s",lstfil.voie1);
		present(lstfil.voie1);
		printf("|");

		printf("%s",lstfil.position2);
		present(lstfil.position2);
		printf("|");

		printf("%s",lstfil.app2);
		present(lstfil.app2);
		printf("|");

		printf("%s",lstfil.voie2);
		present(lstfil.voie2);
		printf("|");
		
}	   
