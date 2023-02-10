#include "projet.h"
        
     
void main (void)
{
	
	fil lstfil[LIGNEMAX];
	char ref1[RMAX]	 ;
	char ref2[RMAX];
	char ref3[RMAX];
	mapping map [MAPMAX] ;
	int nbr_ligne=0,a,nbr_ligne_total=0,nbr_erreur=0;
	int lsterreur[ERREURMAX];
	int cpt=0, erreur=0;

	_outp (0x37a, 0x20);

	charger(lstfil ,map,&nbr_ligne, ref1, ref2 , ref3);
	
	nbr_ligne_total=classement(lstfil,&nbr_ligne, ref1, ref2 , ref3);
	
	nomme(lstfil,map,&nbr_ligne_total);

	printf("-------------------------\n");

	printf("TEST\n");

	printf("-------------------------\n\n\n");

	
	
	testcalculo (lstfil, map,lsterreur,&nbr_ligne, &nbr_ligne_total,&nbr_erreur);

	testfaisceau(lstfil , lsterreur, &nbr_ligne_total,&nbr_ligne,&nbr_erreur);
	
	while (cpt<3)
	{
		erreur=0;

		for (a=0;a<=nbr_ligne_total;a++)
		{
			if(strcmp(lstfil[a].resultat, "ok") == 0 );

			else erreur++;	
		
		}

		if(erreur!=0)
		{
			printf("\n\n\n-------------------------\n");

			printf("CORRECTION\n");

			printf("-------------------------\n\n\n");
	 
			
	
			testcalculo (lstfil, map,lsterreur,&nbr_ligne, &nbr_ligne_total,&nbr_erreur);
		
			testfaisceau(lstfil , lsterreur, &nbr_ligne_total,&nbr_ligne,&nbr_erreur);

		}

		cpt++;
	}
	
	for (a=0;a<=nbr_ligne_total;a++)
	{
		if(strcmp(lstfil[a].resultat, "ok") == 0 );

		else erreur++;	
		
	}
	
	
	if(erreur==0)
		
	printf("\n\n\nFAISCEAU OK\n\n\n\n");

	else

	printf("\n\n ERREUR NON CORRIGE SUR LE FAISCEAU\n\n\n\n");

	

}


void testcalculo (fil lstfil[LIGNEMAX],mapping map[MAPMAX],int lsterreur[ERREURMAX],int* nbr_ligne,
				  int *nbr_ligne_total,int* nbr_erreur)
{
	int r=0,i,n, j,k, clk_consigne,cpt,pin=0,p;
	unsigned char d0,d1,d2,d3,d4,d5,d6,d7,libre='0';
	unsigned char donnee;
	char clk;
	char valide;





	//while (valide ='1');


	for (i=0;i<=*nbr_ligne;i++)
	{
		if(strcmp(lstfil[i].resultat, "ok") != 0 )
		{
			for(j=1;j<=257;j++)
			{
				if (strcmp(lstfil[i].ref2, map[j].ref) == 0 && strcmp(lstfil[i].voie2, map[j].voie) == 0)
				{
					clk_consigne=(j-1)%16;

					//clear

					for (k=1;k<=clk_consigne;k++)
					{
					clk=1;
					sleep(0.01);
					clk=0;
					}
					cpt=0;
				
					//while(cpt<4)
				//	{
		
			
						while(cpt<3)
						{
							testaffiche(lstfil[i]);
							

							//while(libre=='1');
	
							donnee=(_inp(0x378));
							//printf("entrez");
							//scanf("%x",&donnee);	
						
						
							d0= donnee & 0x01 ;
							d1= donnee & 0x02 ;
							d2= donnee & 0x04 ;	
							d3= donnee & 0x08 ;
							d4= donnee & 0x10 ;
							d5= donnee & 0x20 ;
							d6= donnee & 0x40 ;
							d7= donnee & 0x80 ;
				


							if (j>=1 && j<=16 && d0==0)
							{
								afftestok(&lstfil[i]);
								cpt=5;
							}
						
							else if (j>=17 && j<=32 && d1==0)
							{
								afftestok(&lstfil[i]);
								cpt=5;
							}
					
							else if (j>=33 && j<=48 && d2==0)
							{
								afftestok(&lstfil[i]);
								cpt=5;
							}
							
							else if (j>=49 && j<=64 && d3==0)
							{
								afftestok(&lstfil[i]);
								cpt=5;
							}		
		
							else if (j>=65 && j<=80 && d4==0)
							{
								afftestok(&lstfil[i]);
								cpt=5;
							}
								
							else if (j>=81 && j<=96 && d5==0)
							{
								afftestok(&lstfil[i]);
								cpt=5;
							}
						
							else if (j>=97 && j<=112 && d6==0)
							{
								afftestok(&lstfil[i]);
								cpt=5;
							}
					
							else if (j>=113 && j<=128 && d7==0)
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
						if(cpt==5)												 
							break;
			
				
						//clear
						
						for(p=0;p<=15;p++)
						{
							//printf("entrez");
							//scanf("%x",&donnee);
							donnee=(_inp(0x378));

					
							d0= donnee & 0x01 ;
							d1= donnee & 0x02 ;
							d2= donnee & 0x04 ;
							d3= donnee & 0x08 ;
							d4= donnee & 0x10 ;
							d5= donnee & 0x20 ;
							d6= donnee & 0x40 ;
							d7= donnee & 0x80 ;


	
							if(d0==0)
							{
								pin=1+p;
					
								for(n=0;n<=*nbr_ligne;n++)
								{
									if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
									{
										affiche(lstfil[n]);
										printf("\n");
										cpt=5;
										pin=0;
										
									}
									else if (n==*nbr_ligne && pin!=0)
									{
										printf("|");	
										printf("%s",map[pin].nom);
										present(map[pin].nom);
										printf("|");
								
										printf("%s",map[pin].voie);
										present(map[pin].voie);
										printf("|");
										printf("\n");
								   		pin=0;
									}
									
			
								}
					
							}
							
							if(d1==0)
							{
								pin=17+p;
				
								for(n=0;n<=*nbr_ligne;n++)
								{
									if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
									{
										affiche(lstfil[n]);
										printf("\n");
										cpt=5;
										pin=0;
									}
									else if (n==*nbr_ligne && pin!=0)
									{
										printf("|");	
										printf("%s",map[pin].nom);
										present(map[pin].nom);
										printf("|");
										
										printf("%s",map[pin].voie);
										present(map[pin].voie);
										printf("|");
										printf("\n");
									   	pin=0;
									}
									
								
								}		
		
								
							}
					
							if(d2==0)
							{
								pin=33+p;

								for(n=0;n<=*nbr_ligne;n++)
								{
									if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
									{
										affiche(lstfil[n]);
										printf("\n");
										cpt=5;
										pin=0;
									}
									else if (n==*nbr_ligne && pin!=0)
									{
										printf("|");	
										printf("%s",map[pin].nom);
										present(map[pin].nom);
										printf("|");
								
										printf("%s",map[pin].voie);
										present(map[pin].voie);
										printf("|");
										printf("\n");
									   	pin=0;
									}
									
			
								}
					
								
							}		
	
							if(d3==0)
							{
								pin=49+p;
					
								for(n=0;n<=*nbr_ligne;n++)
								{
									if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
									{
										affiche(lstfil[n]);
										printf("\n");
										cpt=5;
										pin=0;
									}
									else if (n==*nbr_ligne && pin!=0)
									{
										printf("|");	
										printf("%s",map[pin].nom);
										present(map[pin].nom);
										printf("|");
								
										printf("%s",map[pin].voie);
										present(map[pin].voie);
										printf("|");
										printf("\n");
									   	pin=0;
									}
								

								}

							
							}

							if(d4==0)
							{
								pin=65+p;

								for(n=0;n<=*nbr_ligne;n++)
								{
									if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
									{
										affiche(lstfil[n]);
										printf("\n");
										cpt=5;
										pin=0;
									}
									else if (n==*nbr_ligne && pin!=0)
									{
										printf("|");	
										printf("%s",map[pin].nom);
										present(map[pin].nom);
										printf("|");
							
										printf("%s",map[pin].voie);
										present(map[pin].voie);
										printf("|");
										printf("\n");
									   	pin=0;
									}
									

								}

							
							}

							if(d5==0)
							{
								pin=81+p;
						
								for(n=0;n<=*nbr_ligne;n++)
								{
									if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
									{
										affiche(lstfil[n]);
										printf("\n");
										cpt=5;
										pin=0;
									}
									else if (n==*nbr_ligne && pin!=0)
									{
										printf("|");	
										printf("%s",map[pin].nom);
										present(map[pin].nom);
										printf("|");
								
										printf("%s",map[pin].voie);
										present(map[pin].voie);
										printf("|");
										printf("\n");
									   	pin=0;
									}
								

								}


							
							}

							if(d6==0)
							{
								pin=97+p;
							
								for(n=0;n<=*nbr_ligne;n++)
								{		
									if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
									{
										affiche(lstfil[n]);
										printf("\n");
										cpt=5;
										pin=0;
									}
									else if (n==*nbr_ligne && pin!=0)
									{		
										printf("|");	
										printf("%s",map[pin].nom);
										present(map[pin].nom);
										printf("|");
								
										printf("%s",map[pin].voie);
										present(map[pin].voie);
										printf("|");
										printf("\n");
									   	pin=0;
									}
									
				
								}
						
								
							}		

							if(d7==0)
							{
								pin=113+p;
					
								for(n=0;n<=*nbr_ligne;n++)
								{
									if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
									{
										affiche(lstfil[n]);
										printf("\n");
										cpt=5;
										pin=0;
									}
									else if (n==*nbr_ligne && pin!=0)
									{
										printf("|");	
										printf("%s",map[pin].nom);
										present(map[pin].nom);
										printf("|");
							
										printf("%s",map[pin].voie);
										present(map[pin].voie);
										printf("|");
										printf("\n");
									   	pin=0;
									}
									
								
								}

							
							}


							clk=1;
							sleep(0.01);
							clk=0;

						/*	
							for(n=0;n<=*nbr_ligne;n++)
							{
								if(strcmp(lstfil[n].ref2, map[pin].ref) == 0 && strcmp(lstfil[n].voie2, map[pin].voie) == 0)
								{
									affiche(lstfil[n]);
									printf("\n");
									cpt=5;
									pin=0;
								}
							}  */
							

							

						

										
						}
				//	}

			  			
				}
			}	
		}
	}
}
	

	


/*unsigned char lire_port(unsigned short adresse_reg)
{
  asm
  {
    mov DX,adresse_reg
    in AL,DX
    mov Result,AL
  }
}  */

void sleep(int nbr_seconds)
{
	clock_t goal;

	goal = (nbr_seconds * CLOCKS_PER_SEC) + clock();

	while(goal > clock())
	{
		;
	}
}  
 
void testaffiche(fil lstfil)
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
		
}	   
			  
void afftestok(fil *lstfil)
{
	int b=0	;

//	printf("\a");


	strcpy(lstfil->resultat,"ok");
		
	printf("%s",lstfil->resultat);
	present(lstfil->resultat);
	printf("|");

	printf("\n");
	
		
		for (b=0;b<=77;b++)
		{
			if(b==0 || b==11 || b==22 || b==33 || b==44 || b==55 || b==66 || b==77)
			printf("|");
			else
			printf("-");
		}

	printf("\n");

}

void afferreur(fil *lstfil)
{
	int b=0	;

//	printf("\a");
//	printf("\a");
//	printf("\a");



	strcpy(lstfil->resultat,"erreur");
		
	printf("%s",lstfil->resultat);
	present(lstfil->resultat);
	printf("|");

	printf("\n");
	
		
		for (b=0;b<=77;b++)
		{
			if(b==0 || b==11 || b==22 || b==33 || b==44 || b==55 || b==66 || b==77)
			printf("|");
			else
			printf("-");
		}

	printf("\n");

}
	 

	 


