#include "projet.h"




void charger (fil lstfil [LIGNEMAX],mapping map [MAPMAX],int* nbr_ligne,char* ref1,char* ref2 ,char* ref3)
{
	FILE *pf;
	char chaine[TAILLE_MAX];
	char fichier[N]="c:\\test\\";
	int i=0,e=0,j=0,k=0,l=0,c=0,m=0,fin_ligne=0,nbr_colonne=0;
	int a=0,b=0,d=0,z=0,w=0,p=0,g=0,u=0;
	char caractereActuel;
//	fil lstfil[;
	int cpt=0,y=0,indicebloc;

	
	char ref_calculo [MAX];
	char ref_emprunte[MAX];
	char *string_ptr;

	
	
	//gets(fichier);

	pf =fopen ("c:\\test\\LI_1615_06_2089_1.csv","rt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	// Boucle de lecture des caract?res un ? un
        do
        {
            caractereActuel = fgetc(pf); // On lit le caract?re
            //printf("%c", caractereActuel); // On l'affiche
			chaine[i]=caractereActuel ;
			i++;
        } while (caractereActuel != EOF); // On continue tant que fgetc n'a pas retourn? EOF (fin de fichier)
		chaine[i]='\0';
       	fclose(pf);

		printf("\n");
		


		while (chaine[w]!='\n')
		{
			if (chaine[w]==';')
			{
				y++;
				indicebloc=w;

				if(y==2)
				{
					chaine[w]='\n';
					y=0;
					printf("\n");
				}
				 
				else 
				chaine[w]=' ';
			}
			printf("%c",chaine[w]);


			w++;
		}

	
		/*	if (chaine[w]==';')
			{
				if (p==1)
					ref_calculo[g]='\0';

				p++;
			}
			else if(p==1)
			{	ref_calculo[g]=chaine[w];
				g++;
			}

			else if(p==3)
			{
				if (chaine[w]=='\n')
					ref_emprunte[u]='\0';
				else
				{
				ref_emprunte[u]=chaine[w];
				u++;
			}	}
			 */
		w++;
		y=0;
		indicebloc++;


		while (chaine[indicebloc]!='\n')
		{
			ref_emprunte[y]=chaine[indicebloc];
			y++;
			indicebloc++;
		}

		ref_emprunte[y]='\0';
			

	
		


		printf("\n utilisez le bloc emprunte %s \n",ref_emprunte);


	for (j=w;j<=strlen(chaine);j++)
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
			*nbr_ligne=*nbr_ligne+1;
		}
		

	}

	*nbr_ligne=*nbr_ligne-1;


	
	
	for (k=w;k<=strlen(chaine);k++)
	{
		switch (cpt)

	  {
		  case 0 :


		if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].liason[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
				lstfil[l].liason[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].liason[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		  

		  break ;

		   case 1 :


		if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].section[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
				lstfil[l].section[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].section[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		  

		  break ;


		   case 2 :


		if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].couleur[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
				lstfil[l].couleur[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].couleur[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		  

		  break ;


		  case 3 :

			  if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].position1[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
			lstfil[l].position1[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].position1[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}

		break ;




		  case 4 :

			  if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].app1[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
				lstfil[l].app1[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].app1[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		

		  break ;


		  case 5 :

			  if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].voie1[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
				lstfil[l].voie1[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].voie1[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		  
		  break ;

		  case 6 :

			  if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].ref1[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
			lstfil[l].ref1[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].ref1[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		break ;




		case 7 :

			  if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].position2[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
			lstfil[l].position2[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].position2[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		break;



		  case 8 :

			  if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].app2[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
				lstfil[l].app2[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].app2[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		 

		  break ;

		  
		  
		  case 9 :

			  if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].voie2[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
			lstfil[l].voie2[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].voie2[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		 

		  break ;

		  case 10 :

			  if (chaine[k]!=';' && chaine[k]!='\n')
		{
			lstfil[l].ref2[m]=chaine[k];
			m++;
		}

		else
		{
			if(chaine[k]==';')
			{
			lstfil[l].ref2[m]='\0';
				cpt++;
				m=0;
				

			}
			if(chaine[k]=='\n')
			{
				lstfil[l].ref2[m]='\0';
				l++;
				cpt=0;
				m=0;
				
			}
		}


		 

		  break ;



		 


		}



		 }
		  memset (chaine, 0, sizeof (chaine));
		  m=0;

		  strcat(fichier,ref_emprunte);
		  strcat(fichier,".map");
		 
		 pf =fopen (fichier,"rt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	// Boucle de lecture des caract?res un ? un
        do
        {
            caractereActuel = fgetc(pf); // On lit le caract?re
            //printf("%c", caractereActuel); // On l'affiche
			chaine[e]=caractereActuel ;
			e++;
        } while (caractereActuel != EOF); // On continue tant que fgetc n'a pas retourn? EOF (fin de fichier)
		chaine[e]='\0';
       	fclose(pf);


		for (k=0;k<=strlen(chaine);k++)
		{
			switch (cpt)
				{
				case 0 :
					if (chaine[k]!=';' && chaine[k]!='\n')
						{
						ref1[m]=chaine[k];
						m++;
						}
					else
						{
						if(chaine[k]==';')
							{
							ref1[m]='\0';
							cpt++;
							m=0;
							}

						if(chaine[k]=='\n')
							{
							ref1[m]='\0';
							l++;
							cpt=0;
							m=0;
				
							}
						}


		  

						break ;

				case 1 :


					if (chaine[k]!=';' && chaine[k]!='\n')
					{
						ref2[m]=chaine[k];
						m++;
					}

					else
					{
						if(chaine[k]==';')
						{
						ref2[m]='\0';
						cpt++;
						m=0;
						}

					
					if(chaine[k]=='\n')
						{
						ref2[m]='\0';
						l++;
						cpt=0;
						m=0;
				
						}
					}


		  

					 break ;


				  case 2 :


				  if (chaine[k]!=';' && chaine[k]!='\n')
				  {
					  ref3[m]=chaine[k];
					  m++;
				  }

				  else
				  {
						if(chaine[k]==';')
						{
			   			ref3[m]='\0';
						 cpt++;
			   			 m=0;
					 
						}

						if(chaine[k]=='\n')
						{
						ref3[m]='\0';
						l++;
						cpt=4;
						m=0;
				
						}
				}
				  break ;

			}

		} cpt=0;l=1;w=0;

		while (chaine[w]!='\n')
		{
			
			w++;
		}
		w++;




	for (k=w;k<=strlen(chaine);k++)
		{
		map[0].ref[0]='\0';
		map[0].voie[0]='\0';

			switch (cpt)
				{
				case 0 :
					if (chaine[k]!=';' && chaine[k]!='\n')
						{
						
						
						}
					else
						{
						if(chaine[k]==';')
							{
							
							cpt++;
							
							}

						if(chaine[k]=='\n')
							{
						
							cpt=0;
							m=0;
							l++;
							map[l].ref[m]='\0';
							map[l].voie[m]='\0';
							}
						}

							   
		  

						break ;

				case 2 :


					if (chaine[k]!=';' && chaine[k]!='\n')
					{
						map[l].ref[m]=chaine[k];
						m++;
					}

					else
					{
						if(chaine[k]==';')
						{
						map[l].ref[m]='\0';
						map[l].voie[m]='\0';
						cpt++;
						m=0;
						}
				

					
						if(chaine[k]=='\n')
						{
						map[l].ref[m]='\0';
						cpt=0;
						m=0;
						l++;
						}

					}


		  

					 break ;


				  case 1 :


				  if (chaine[k]!=';' && chaine[k]!='\n')
				  {
					  map[l].voie[m]=chaine[k];
					  m++;
				  }

				  else
				  {
						if(chaine[k]==';')
						{
			   			map[l].voie[m]='\0';
						map[l].ref[m]='\0';
						 cpt++;
			   			 m=0;
					 
						}

						if(chaine[k]=='\n')
						{
						map[l].voie[m]='\0';
						map[l].ref[m]='\0';
						cpt=0;
						m=0;
						l++;
						}
				}
				  break ;

			}
			

		}

	  puts(ref1);
	  puts(ref2);
	  puts(ref3);

	
	
		



	/*for(a=0;a<=*nbr_ligne;a++)
	
	affiche(lstfil[a]);

	for (i=0;i<=l;i++)
	{
		printf("%d\n",i);
		puts(map[i].ref);
		puts(map[i].voie);
		printf("\n\n");
	}
	*/
}	  

void affiche(fil lstfil)

{
	int b=0	;

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
		
		printf("%s",lstfil.app2);
		present(lstfil.app2);
		printf("|");
				 
		printf("%s",lstfil.voie2);
		present(lstfil.voie2);
		printf("|");

		printf("%s",lstfil.couleur);
		present(lstfil.couleur);
		printf("|");

		printf("%s",lstfil.section);
		present(lstfil.section);
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
void present(char *liste)
{
	int i,espace;

	espace=MAX-strlen(liste);

	for(i=0;i<espace;i++)

		printf(" ");
}


int classement(fil lstfil[LIGNEMAX], int *nbr_ligne, char* ref1, char* ref2 ,char* ref3)
{
	int i,j,p=0,k,m=0,l,cpt=-1,nbr_ligne_total;
	fil liste, tampon[LIGNEMAX];
	

	for (i=0;i<*nbr_ligne;i++)
	{
		if (strcmp(lstfil[i].ref1,ref1)==0 || strcmp(lstfil[i].ref1,ref2)==0 || strcmp(lstfil[i].ref1,ref3)==0 ||
			strcmp(lstfil[i].ref2,ref1)==0 || strcmp(lstfil[i].ref2,ref2)==0 || strcmp(lstfil[i].ref2,ref3)==0);

		else
		{
			tampon [m]=lstfil[i];
			m++;


			for(j=i;j<=*nbr_ligne+1;j++)
			{
				
				lstfil[j]=lstfil[j+1];
			}

			
			*nbr_ligne=*nbr_ligne-1;
			i--;
		
		}

		if (strcmp(lstfil[i].ref1,ref1)==0 || strcmp(lstfil[i].ref1,ref2)==0 || strcmp(lstfil[i].ref1,ref3)==0)
		{
			strcpy(liste.app1,lstfil[i].app1);
			strcpy(liste.position1,lstfil[i].position1);
			strcpy(liste.ref1,lstfil[i].ref1);
			strcpy(liste.voie1,lstfil[i].voie1);
			
			strcpy(lstfil[i].app1,lstfil[i].app2);
			strcpy(lstfil[i].position1,lstfil[i].position2);
			strcpy(lstfil[i].ref1,lstfil[i].ref2);
			strcpy(lstfil[i].voie1,lstfil[i].voie2);
			
			strcpy(lstfil[i].app2,liste.app1);
			strcpy(lstfil[i].position2,liste.position1);
			strcpy(lstfil[i].ref2,liste.ref1);
			strcpy(lstfil[i].voie2,liste.voie1);

		  

		}
		   
	}

	alpha(lstfil ,nbr_ligne);
	alpha(tampon ,&m);

	

	for (i=*nbr_ligne;i<=*nbr_ligne+m;i++)
	{
		lstfil[i]=tampon[p];
		p++;
	}

	nbr_ligne_total=*nbr_ligne+m-1;

	return (nbr_ligne_total);

 	
}
	



    

void nomme (fil lstfil [LIGNEMAX],mapping map [MAPMAX],int* nbr_ligne_total)
{
	int i,j;

	for (i=0;i<*nbr_ligne_total;i++)
	{
		for(j=1;j<=257;j++)
		{
			if (strcmp(lstfil[i].ref2, map[j].ref) == 0 )
			
				strcpy(map[j].nom,lstfil[i].app2);
		}
	}
}

	
 void alpha (fil lstfil[LIGNEMAX], int *nbr_ligne)
{
	fil liste;
	int k,l;

		for (k=0;k<*nbr_ligne-1;k++)

		{

			for (l=k+1;l<*nbr_ligne;l++)
			{



				if (strcmp(lstfil[k].position1,lstfil[l].position1 ) > 0)
				{
					liste=lstfil[k];
					lstfil[k]=lstfil[l];
					lstfil[l]=liste;

				}

				else if (strcmp(lstfil[k].position1,lstfil[l].position1 ) == 0)
				{
					if (strcmp(lstfil[k].app1,lstfil[l].app1 ) > 0)
						{
						liste=lstfil[k];
						lstfil[k]=lstfil[l];
						lstfil[l]=liste;
						}
					
					else if (strcmp(lstfil[k].app1,lstfil[l].app1 ) == 0)
						{
						if (strcmp(lstfil[k].voie1,lstfil[l].voie1 ) > 0 )
							{
							liste=lstfil[k];
							lstfil[k]=lstfil[l];
							lstfil[l]=liste;
							}
						}
				}
			}
		}



}
	/*
void remplace (string)
{
	char *string_ptr;
   
	while((string_ptr=strpbrk(string,";"))!=NULL)
	{
	   string_ptr[0]=' ';
	   
	}

  
}		  */



   
	
