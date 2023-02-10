#include <stdio.h>
#include <ctime>
#include <stdlib.h>
//#include <string.h>
#include <conio.h>


void main (void)
{
	char listefil[60], rapport[60], etiquette[60];


	 gets (listefil);

	strcpy(rapport,listefil);				   
	strcat(rapport,"rap.txt");

	strcpy(etiquette,listefil);
	strcat(etiquette,"eti.txt");

	puts (rapport);
	puts (etiquette);


}

