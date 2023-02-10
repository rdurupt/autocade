#include<string.h>
#include<stdio.h>


#define N    75

void decale (char chaine[N], char string[N]);
void file_rep(char *rep,char *string);
char *FileName (char *path);

int main(void)
{
	int i;
	char string[N]="ma;maison;est;rouge";
	char *string_ptr;
	char tab[N],rep[N];
	char *path="ma maison esr rouge";



while((string_ptr=strpbrk(string,";"))!=NULL)
  {
	   string_ptr[0]=' ';
	   
	}

  
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

char *FileName (char *path)
{
    char *c = path, *ret = path;
    if(c == 0) return 0;
    while(*c)
    {
        if(*c == '\\')  ret = c+1;
        c++;
    }
    return ret;
}

