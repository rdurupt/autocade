#include <conio.h> // Fichier d'en-t�te pour l'instruction "_outp"

#include <stdio.h> // Fichier d'en-t�te pour E/S standards

#include <time.h> // Fichier d'en-t�te pour les instruction "clock_t" et "endwait"

 
  
   
void main (void) 
{
	char valeur;

  /*
	_outp(0x378, valeur);
	valeur=_inp(0x379) ;	  */
	
	_outportb(0x378, valeur);
	_inportb(0x379);
  }
