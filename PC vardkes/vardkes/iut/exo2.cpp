#include <stdio.h>
#include <time.h>
#include <sys/netmgr.h>
#include <sys/neutrino.h>
#include <errno.h>
#include <stdlib.h>
#include <sys/dispatch.h>

#define ATTACH_POINT "myname"

typedef union {
        struct _pulse   pulse;
        /* your other message structures would go 
           here too */
} my_message_t;

main()
{
   struct sigevent         event;
   struct itimerspec       itime;
   timer_t                 timer_id;
   int                     chid;
   int                     rcvid;
   my_message_t            msg;
   name_attach_t *attach;
   int i;
   
   
   
   /* Create a local name (/dev/name/local/...) */
   if ((attach = name_attach(NULL, ATTACH_POINT, 0)) == NULL) {
       return EXIT_FAILURE;    
      

   for (i=0;i<10;i++) {
       rcvid = MsgReceive(attach->chid, &msg, sizeof(msg), NULL);
       if (rcvid == 0) { /* we got a pulse */
       
		      printf("%d\n",msg.pulse.code);
                printf("we got a pulse from \n");
       } /* else other messages ... */
   }
}
}