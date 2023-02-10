#include <stdio.h>
#include <ctime>
#include <stdlib.h>
#include <string.h>
#include <conio.h>

void main (void)
{
	int i;
	int input[10];

FILE *ReadInput; 
 
ReadInput = fopen ("c:\\signal.txt", "r" ) ;  
  
 
for (i=0 ; i<100 ; i++)                           
{  
 fscanf(ReadInput, "%f", &input[i]);    
}  
 
fclose (ReadInput) ;  

} 