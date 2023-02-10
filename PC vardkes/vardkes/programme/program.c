#include <stdio.h>
#include <math.h>


void main (void)
{

	char donnee;
	char d0,d1,d2,d3,d4,d5,d6,d7;

	donnee=0x7F;



	d0= donnee & 0x01 ;
	d1= donnee & 0x02 ;
	d2= donnee & 0x04 ;
	d3= donnee & 0x08 ;
	d4= donnee & 0x10 ;
	d5= donnee & 0x20 ;
	d6= donnee & 0x40 ;
	d7= donnee & 0x80 ;



	
	printf("%x\n",d0);
	printf("%x\n",d1);
	printf("%x\n",d2);
	printf("%x\n",d3);
	printf("%x\n",d4);
	printf("%x\n",d5);
	printf("%x\n",d6);
	printf("%x\n",d7);

	



}
