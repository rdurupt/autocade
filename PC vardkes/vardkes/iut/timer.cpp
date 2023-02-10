#include <stdio.h>
#include <time.h>
#include <sys/netmgr.h>
#include <sys/neutrino.h>
#include <errno.h>
#include <stdlib.h>
#include <sys/dispatch.h>

#define MY_PULSE_CODE  11
#define ATTACH_POINT "myname"

typedef union {
        struct _pulse   pulse;
        /* your other message structures would go 
MY_PULSE_CODE           here too */
} my_message_t;

main()
{
   struct sigevent         event;
   struct itimerspec       itime;
   timer_t                 timer_id;
   int                     chid;
   int                     rcvid;
    my_message_t            msg;

   int fd;


   chid = ChannelCreate(0);   /* cree un canal*/

   event.sigev_notify = SIGEV_PULSE;
   event.sigev_coid = ConnectAttach(ND_LOCAL_NODE, 0, 
                                    chid, 
                                    _NTO_SIDE_CHANNEL, 0);
   event.sigev_priority = getprio(0);
   event.sigev_code = MY_PULSE_CODE;
   timer_create(CLOCK_REALTIME, &event, &timer_id);

   itime.it_value.tv_sec = 1;
   /* 500 million nsecs = .5 secs */
   itime.it_value.tv_nsec = 700000000; 
   itime.it_interval.tv_sec = 1;
   /* 500 million nsecs = .5 secs */
   itime.it_interval.tv_nsec = 500000000; 
   timer_settime(timer_id, 0, &itime, NULL);

   /*
    * As of the timer_settime(), we will receive our pulse 
    * in 1.5 seconds (the itime.it_value) and every 1.5 
    * seconds thereafter (the itime.it_interval)
    */
       

    if ((fd = name_open(ATTACH_POINT, 0)) == -1) {
        return EXIT_FAILURE;

   for (;;) {
       rcvid = MsgReceive(chid, &msg, sizeof(msg), NULL);
       if (rcvid == 0) { /* we got a pulse */
       printf("%d\n",msg.pulse.code);
            if (msg.pulse.code == MY_PULSE_CODE) {
                printf("we got a pulse from our timer \n");
                // MsgSendPulse
                 if (MsgSendPulse(fd, 1 , MY_PULSE_CODE, 0) == -1) 
            break;
            } /* else other pulses ... */
       } /* else other messages ... */
   }
}
}