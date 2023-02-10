#include "bdteque.h"

void save(BD biblio [NBD],int bd_present)
{

	FILE *pf;

	pf =fopen ("donne.txt","wt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	fwrite (&bd_present,sizeof(int),1,pf );

	fwrite (biblio,sizeof(BD),bd_present,pf );

	fclose(pf);


}


void load(BD biblio [NBD],int *bd_present)
{
	int i;
	FILE *pf;

	pf =fopen ("donne.txt","rt");

	if(pf==NULL)
	{
		perror ("pb ouverture donne.txt");
		exit(-1);
	}

	fread (bd_present,sizeof(int),1,pf);

	for(i=0;i<*bd_present;i++)

	fread (&biblio[i],sizeof(BD),1,pf);

	fclose(pf);

}








