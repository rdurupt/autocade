 //------------------------------------------------------------------------------------
// BIBLIOTHEQUES
//------------------------------------------------------------------------------------
#include <c8051f000.h>
#include<stdio.h>


//------------------------------------------------------------------------------------
// AFFECTATION DES PINS ET VARIABLE GLOBALE
//------------------------------------------------------------------------------------

sbit E = P2^5;
sbit RS = P2^6;
sbit RW = P2^7;

sbit BF = P3^7;

char caractere[16];
int donnee;

//------------------------------------------------------------------------------------
// CONSTANTES GLOBALES
//------------------------------------------------------------------------------------
#define 	TRUE			0x01			// Value representing TRUE
#define		FALSE			0x00			// Value representing FALSE
#define 	DB P3

//------------------------------------------------------------------------------------
// FONCTIONS PROTOTYPES
//------------------------------------------------------------------------------------
void delai(unsigned int);
void delai_long(unsigned int);
void code_LCD(char);
void code_LCD_init(char);
void init_aff (void);
void my_init_aff (void);
void clear_display(void);
void function_set(void);
void display_ON(void);
void display_OFF(void);
void mode_set(void);
void chaine (char* caractere, int nb_caractere);

char message[16]="Init ok        ";   //Message à Afficher

//------------------------------------------------------------------------------------
// MAIN
//------------------------------------------------------------------------------------
void main(void)
{
long int j;



/********** definition de l'horloge du système : externe 12MHz*********/
OSCXCN=0x67;
for(j=0;j<256;j++);
 while(!(OSCXCN&0x80));
OSCICN=0x0C;

/********** configuration des registres ***********/
//arret du watchdog timer
WDTCN = 0xde;
WDTCN = 0xad;

// Configure the XBRn Registers
XBR0 = 0x00;	// XBAR0: Initial Reset Value	
XBR1 = 0x00;	// XBAR1: Initial Reset Value
XBR2 = 0x40;	// XBAR2: Initial Reset Value

//P1, P2 et P3 drain ouvert
PRT1CF=0x00 ;
PRT2CF=0x00 ;
PRT3CF=0x00 ;



/********** initialisation de l'afficheur*************/

init_aff();
init_aff();
chaine(message,16);
while(TRUE);

}


//------------------------------------------------------------------------------------
// FONCTIONS
//------------------------------------------------------------------------------------

/***********fonction d'attente en microseconde *************/
void delai (unsigned int T)
{
	//configure Timer 3
	TMR3H=(unsigned char)((0xFFFF-T)>>8);
	TMR3L=(unsigned char)(0xFFFF-T);
	TMR3CN=0x04;				//TR3=1   T3M=0 (car on a une horloge de 12MHz et on veut 1MHz)
						//		  et on garde les autres valeurs
	while(!(TMR3CN&0x80)); 		//tant que TF3=0
}

/***********fonction d'attente en milliseconde *************/
void delai_long (unsigned int T)
{
	int i;
	for(i=0;i<T;i++)
		delai(1000);
	
}

/**************initialisation code LCD******************/
void code_LCD_init (unsigned char donnees)
{
	E=0;
	RS=0;					//afficheur attend une commande
	RW=0;
	E=1;					//validation de l'afficheur
	delai(20);				//attente de la validation de la commande
	DB=donnees;				//commande envoyé au LCD
	delai(10);				//attente de la validation de la commande
	E=0;					//écriture dans l'afficheur
}

/***********initialisation afficheur************/
void init_aff (void)
{
	delai_long(20);
	clear_display();
	delai_long(6);
	function_set();
	delai_long(6);
	display_ON();
	delai_long(6);
	mode_set();
}


/*************clear display************/
void clear_display(void)
{
	code_LCD_init(0x01);
}

/*************function set*************/
void function_set(void)
{
	code_LCD_init(0x38);
}

/*************display ON**********/
void display_ON(void)
{
	code_LCD_init(0x0F);
}

/*************display OFF**********/
void display_OFF(void)
{
	code_LCD_init(0x08);
}

/**************Mode set*************/
void mode_set(void)
{
	code_LCD_init(0x06);
}

/*************affichage ************/
void code_LCD (char donnees)
{
	E=0;
	RS=1;					//afficheur attend une commande
	RW=0;
	E=1;					//validation de l'afficheur
	delai(20);				//attente de la validation de la commande
	DB=donnees;				//commande envoyé au LCD
	delai(10);				//attente de la validation de la commande
	E=0;					//écriture dans l'afficheur
}


/************fonction chaine***********/
void chaine (char* caractere,int nb_caractere)
{
int i;	
for(i=0;i<(nb_caractere-1);i++)
	{
	if(i==8)
	        code_LCD_init(0xC0);
	code_LCD (caractere[i]);
	}
	
}



