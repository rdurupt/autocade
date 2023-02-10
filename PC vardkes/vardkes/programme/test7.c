#include <stdlib.h>
#include <stdio.h>
#define TAILLE_MAX    1000
#define N 1000

char** split( char* str, char c );

void main (void)
{
	FILE *pf;
	char chaine[TAILLE_MAX];
	char fichier[N];
	int i = 0;
	char **explode;
    char str[][256];

/*	pf =fopen ("C:\\test\\dede.txt","wt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	fputs("Salut les zero\nComment allez-vous \n?", pf);
	printf("ecriture\n");
        

	fclose(pf);*/





	gets(fichier);
	

	pf =fopen (fichier,"rt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	while (fgets(chaine, TAILLE_MAX, pf) != NULL)// On lit maximum TAILLE_MAX caractères du fichier, on stocke le tout dans "chaine"
        printf("%s", chaine); // On affiche la chaîne
	
	fclose(pf);

	
    for( i = 0; i < sizeof(str)/sizeof(str[0]); ++i )
	{
        printf( "----- %d [%s]-----\n", i, str[i] );
      explode = split( str[i], '\n' );
        while( *explode )
            printf( "[%s]\n", *explode++ );

	


}


#define MAXSTR 256

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




