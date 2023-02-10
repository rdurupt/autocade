#include "bdteque.h"

void trier(BD biblio [NBD],int bd_present)
{
	int i, j ;
	BD tampon;


	for (i=0;i<bd_present-1;i++)

		{

			for (j=i+1;j<bd_present;j++)
			{



				if (strcmp(biblio[i].serie,biblio[j].serie ) > 0)
				{
					tampon=biblio[i];
					biblio[i]=biblio[j];
					biblio[j]=tampon;

				}

				else if (strcmp(biblio[i].serie,biblio[j].serie ) == 0)

				{
					if (biblio[i].nserie == biblio[j].nserie )
					{
						tampon=biblio[i];
						biblio[i]=biblio[j];
						biblio[j]=tampon;

					}

				}


			}
		}


}
