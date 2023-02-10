#include 
#include 
#include 

void __cdecl main(void)
{
    unsigned char value;
    printf("IoExample for PortTalk V2.0\nCopyright 2001 Craig Peacock\nhttp://www.beyondlogic.org\n");
    OpenPortTalk();
    outportb(0x378, 0xFF);
    value = inportb(0x378);
    printf("Value returned = 0x%02X \n",value);
    outp(0x378, 0xAA);
    value = inp(0x378);
    printf("Value returned = 0x%02X \n",value);
    ClosePortTalk();
}

