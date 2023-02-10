#include<stdio.h>
#include<dos.h>
#include<conio.h>
#include<math.h>

void outportb(unsigned short ADRESSE, unsigned char VALEUR);
unsigned char inportb(unsigned short ADRESSE);


void main(void)
{
    int valeur,i;

    
    outportb(0x378,0x00); //initialisation
    valeur=i=1;
    do
    {
        printf("Envoie de %d sur LPT1\n",valeur);
        outportb(0x378,valeur);
       
        valeur=pow(2,i);
        i++;
    }
    while(valeur<=0xFF);
}
