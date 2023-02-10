#include <stdio.h>
#include <stdlib.h>

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

/* ----- Programme principal pour le test ----- */
void main(void) {
    int i = 0;
	char **explode;
    char str[][256] = { "",
                    "|",
                    "||||",
                    "chat|chien|maison|souris",
                    "|chien|maison|souris",
                    "chat|chien|maison|",
                    "chat|chien||souris",
                    "chat|chien",
                    "chat" };
   
    for( i = 0; i < sizeof(str)/sizeof(str[0]); ++i )
	{
        printf( "----- %d [%s]-----\n", i, str[i] );
      explode = split( str[i], '|' );
        while( *explode )
            printf( "[%s]\n", *explode++ );
   }
}
