#include "bdteque.h"

void trier(BD biblio [NBD])
{
	int i, j ;
	BD tampon;


	for (i=0;i<*bd_present;i++)

		{
			for (j=1;j<*bd_present-1;j++)
			{



				if (strcmp(biblio[i].serie,biblio[j].serie ) > 0)
				{
					biblio[i]=tampon;
					biblio[i]=biblio[j];
					biblio[j]=tampon;

				}

			}
		}

}
