#include <stdio.h>
#include <stdlib.h>
#include <time.h>		        


void sleep (int nbr_seconds);

void sleep(int nbr_seconds)
{
	clock_t goal;

	goal = (nbr_seconds * CLOCKS_PER_SEC) + clock();

	while(goal > clock())
	{
		;
	}
} 


void main (void)
{
	int a ;

	printf("bonjour");

	sleep (0.010);

	printf("entrer");

	scanf("%d",&a);

}

void sleep(int nbr_seconds)
{
	clock_t goal;

	goal = (nbr_seconds * CLOCKS_PER_SEC) + clock();

	while(goal > clock())
	{
		;
	}
} 

