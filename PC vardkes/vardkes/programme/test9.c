#include <stdio.h>

void main (void)
{
	unsigned char a;
	unsigned char d0,d1,d2,d3,d4,d5,d6,d7;

	scanf("%x",&a);
				
		
			d0= a & 0x01 ;
			d1= a & 0x02 ;
			d2= a & 0x04 ;
			d3= a & 0x08 ;
			d4= a & 0x10 ;
			d5= a & 0x20 ;
			d6= a & 0x40 ;
			d7= a & 0x80 ;



	printf("%d\n",a);

	if (d0==0)
   	printf("d0 %d\n",d0);
	
	if (d1==0)
	printf("d1 %d\n",d1);
 	
	if (d2==0)
	printf("d2 %d\n",d2);
	
	if (d3==0)
	printf("d3 %d\n",d3);
	
	if (d4==0)
	printf("d4 %d\n",d4);
	
	if (d5==0)
	printf("d5 %d\n",d5);
	
	if (d6==0)
	printf("d6 %d\n",d6);
	
	if (d7==0)
	printf("d7 %d\n",d7);
		

}