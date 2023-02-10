#include<string.h>
#include<stdio.h>


#define N    75

void decale (char chaine[N], char string[N]);
void file_rep( char *string, char rep);

int main(void)
{
	int i;
	char string[N]="C:\test";
	char *string_ptr;
	char tab[N],rep[N];


file_rep(string,rep);

  puts(rep);
	
	
  

  decale(string,string);
 
  
  

 /* while((string_ptr=strpbrk(string," "))!=NULL)
  {
	   string_ptr[0]='m';
	   string_ptr[1]='a';
	}

  printf("New string is \"%s\".\n",string);
  return 0;*/
}

void decale (char chaine[N], char string[N])
{

	int i,j=0;
	char tampon[75];
	
	for (i=0;i<75;i++)
	{
		if (chaine[i]==' ')
		{
			tampon[i+j]=chaine[i];
			j++;
			tampon[i+j]=chaine[i];
			
		}

		else 		
		
			tampon[i+j]=chaine[i];

	

	}

	 	puts(tampon);
		
	
	

}



void file_rep(char *rep, char *chemin)
{
 int i,j;
 i=strlen(chemin);
 do {
     i--;
  } while (chemin[i]!='\\');
  for(j=0;j<=i;j++){
        rep[j]=chemin[j];
  }
  rep[j]='\0';
}    

