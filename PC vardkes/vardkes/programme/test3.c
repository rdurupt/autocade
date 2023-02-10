#include <stdio.h>
#include <ctime>
#include <stdlib.h>
#include <string.h>
#include <conio.h>

#define FIL_MAX  150
#define MAXSTR 256

char** split( char* str, char c );
void load(void);

void main(void)
{
	char liste_fil[FIL_MAX][6];
	char text[255];

	load();

}


void load(void)
{
	int i,j;
	char *pf;
	char tab;
	
	pf =fopen ("donne.txt","rt");

	puts(pf);		

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	

tab=split(pf,"/n");

puts (tab);




}



/* split (not split_r) */
char** split( char* str, char c )
{
    static char* tmp[ MAXSTR ] ;    /* 256 colonnes max */
    int current = 0;
    tmp[current++] = str;
    while( *str ) {
        if ( *str == c ) {
            *str = '\0';
            tmp[current++] = str+1;  /* on devrait vérifier si on dépasse pas 256 */
        }
        ++str;
    }
    tmp[ current ] = 0;
    return tmp;
}

