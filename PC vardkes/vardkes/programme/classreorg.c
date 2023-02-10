#include <stdlib.h>
#include <stdio.h>



#define TAILLE_MAX    8000
#define N 30   
#define MAX	 10
#define LIGNEMAX  140
#define RMAX	 11



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

	
}fil;


void present(char*);
void  affiche(fil  );
void classement(fil lstfil[LIGNEMAX],int*, char* , char*  ,char* );



  


  

void main (void)
{
	FILE *pf;
	char chaine[TAILLE_MAX];
	char fichier[N]="c:\\test\\";
	int i=0,e=0,j=0,k=0,l=0,c=0,m=0,fin_ligne=0,nbr_colonne=0,nbr_ligne=0;
	int a=0,b=0,d=0,z=0,w=0,p=0,g=0,u=0;
	char caractereActuel;
	fil lstfil[LIGNEMAX];
	int cpt=0;
	char ref1[RMAX]	 ;
	char	ref2[RMAX];
	char ref3[RMAX];
	char ref_calculo [MAX];
	char ref_emprunte[MAX]; 
	  
	
	//gets(fichier);

	pf =fopen ("c:\\test\\durupt.csv","rt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	// Boucle de lecture des caractères un à un
        do
        {
            caractereActuel = fgetc(pf); // On lit le caractère
            //printf("%c", caractereActuel); // On l'affiche
			chaine[i]=caractereActuel ;
			i++;
        } while (caractereActuel != EOF); // On continue tant que fgetc n'a pas retourné EOF (fin de fichier)
		chaine[i]='\0';
       	fclose(pf);

		printf("\n");


		while (chaine[w]!='\n')
		{
			if (chaine[w]==';')
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
			w++;
		}
		ref_emprunte[u]='\0';
		w++;

	printf("reference caculateur : %s",ref_calculo);
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
			nbr_ligne++;
		}
		

	}

	nbr_ligne=nbr_ligne-2;


	
	
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

	// Boucle de lecture des caractères un à un
        do
        {
            caractereActuel = fgetc(pf); // On lit le caractère
            //printf("%c", caractereActuel); // On l'affiche
			chaine[e]=caractereActuel ;
			e++;
        } while (caractereActuel != EOF); // On continue tant que fgetc n'a pas retourné EOF (fin de fichier)
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

		}

	  puts(ref1);
	  puts(ref2);
	  puts(ref3);

	  classement(lstfil,&nbr_ligne, ref1, ref2 , ref3);
	
		



	for(a=0;a<=nbr_ligne;a++)
	{
	affiche(lstfil[a]);
	
	
		
	}  

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


void classement(fil lstfil[LIGNEMAX], int *nbr_ligne, char* ref1, char* ref2 ,char* ref3)
{
	int i,j,k,l,cpt=-1;
	fil liste;
	

	for (i=0;i<*nbr_ligne+1;i++)
	{
		if (strcmp(lstfil[i].ref1,ref1)==0 || strcmp(lstfil[i].ref1,ref2)==0 || strcmp(lstfil[i].ref1,ref3)==0 ||
			strcmp(lstfil[i].ref2,ref1)==0 || strcmp(lstfil[i].ref2,ref2)==0 || strcmp(lstfil[i].ref2,ref3)==0);

		else
		{

			for(j=i;j<=*nbr_ligne+1;j++)
			lstfil[j]=lstfil[j+1];
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

	for (k=0;k<*nbr_ligne-1;k++)

		{

			for (l=k+1;l<*nbr_ligne;l++)
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
			


   