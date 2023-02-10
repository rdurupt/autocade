 #include <conio.h> // Fichier d'en-tête pour l'instruction "_outp"

#include <stdio.h> // Fichier d'en-tête pour E/S standards

#include <time.h> // Fichier d'en-tête pour les instruction "clock_t" et "endwait"

void main (void)
{valeur;

	valeur=lire_port(0.379);
}






unsigned char lire_port(unsigned short adresse_reg)
{
  asm;
  {
    mov DX,adresse_reg;
    in AL,DX		   ;
    mov Result,AL		;
  }
}
