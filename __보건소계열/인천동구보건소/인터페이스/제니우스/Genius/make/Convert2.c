#include <stdio.h>
#include <stdlib.h>
#include <dos.h>
#include <stdio.h>
#include <conio.h>
#include <time.h>

#define	CR      0x0D
#define	SOH     0x01
#define	STX     0x02
#define	ETX     0x03
#define	EOT     0x04
#define	ENQ     0x05
#define	ACK     0x06
#define	NAK     0x15

#define	NSR		0xD0
#define	TA1		0x50
#define	TA2		0x10

#define	DS1		0x41
#define	RS1		0x30
#define	RS2		0x39

#define COM1	0x3F8
#define	COM2	0x2F8

#define	E_Mask1	0xEF		/*	COM1 IRQ4 Enable Mask	*/
#define	E_Mask2	0xf7		/*	COM2 IRQ3 Enable Mask	*/
#define	D_Mask1	0x10		/*	COM1 IRQ4 Disable Mask	*/
#define	D_Mask2	0x08		/*	COM2 IRQ3 Disable Mask	*/
#define	ICW1	0x20		/*	8259 PIC Address	*/
#define	ICW2	0x21		/*	8259 PIC Address	*/
#define V_COM1	0x0c		/*	COM1 Interrupt Vector */
#define V_COM2	0x0b		/*	COM2 Interrupt Vector */
#define	MaxBuf	4096
#define	TRUE	1
#define	FALSE	0
#define	TIMER_IntV 0x1c			
#define TimedOut() (tic==0)
#define SetTimer(m_sec) tic=m_sec*182/10+1

int	tic;

typedef unsigned char BYTE;
typedef unsigned int INT;
typedef unsigned long LONG;

static INT Buf[2][MaxBuf], BufFull=0, qHead[2], qTail[2];

static enum { NONE , ODD, DUME ,EVEN } Parity;
static enum { D_5BIT , D_6BIT , D_7BIT , D_8BIT } DataBit;
static enum { S_1BIT , S_2BIT } StopBit;
static enum { YES , NO } LogFile;

static void interrupt (*O_ComVec1)();
static void interrupt N_ComVec1();

static void interrupt (*O_ComVec2)();
static void interrupt N_ComVec2();

INT WhatCom(BYTE ComP);
BYTE Data[140];	

void BufAdd(BYTE ComP);
void OutPort(BYTE ComP, BYTE ch);
void V_ComSet();
void V_ComRst();
INT BufEmpty(BYTE ComP);
BYTE GetBuf(BYTE ComP);
void Initialize();
void setModem(BYTE ComP);

void setBaud(BYTE ComP, LONG baud);
void setLine(BYTE ComP, INT databit, INT stopbit, INT parity);
void Int_ComSet();
void clearBuffer(BYTE ComP);
void POLL(BYTE Port_No);

void setBaud(BYTE ComP, LONG baud)
{
	INT Ureg0;
	BYTE port;

	Ureg0=WhatCom(ComP);
	port=inportb(Ureg0+3);
	outportb(Ureg0+3, (Ureg0+3) | 0x80);				/* set DLAB=1 */
	outportb(Ureg0,(BYTE)(115200/baud & 0xFF));
	outportb(Ureg0+1,(BYTE)(115200/baud >> 8));
	outportb(Ureg0+3,port);						/* reset DLAB bit	*/
}

void setLine(BYTE ComP, INT databit, INT stopbit, INT parity)
{
	BYTE lcr=0;
	lcr |= databit;
	lcr |= (stopbit << 2);
	lcr |= (parity << 3);
	outportb(WhatCom(ComP)+3,lcr);
}

void setModem(BYTE ComP)
{
	outportb(WhatCom(ComP)+4,0x0B);
}

static void V_ComSet(void)
{
	O_ComVec1 = getvect(V_COM1);  /* intno=0xc */
	setvect(V_COM1, N_ComVec1);
	O_ComVec2 = getvect(V_COM2); /* intno=0xb */
	setvect(V_COM2, N_ComVec2);
}

static void V_ComRst()
{
	BYTE portstat;

	outport(COM1+1,0x00);
	portstat = inportb(ICW2);
	outportb(ICW2,D_Mask1 | portstat);
	setvect(V_COM1 , O_ComVec1);

	outport(COM2+1,0x00);
	portstat = inportb(ICW2);
	outportb(ICW2,D_Mask2 | portstat);
	setvect(V_COM2 , O_ComVec2);

	enable();
}

static void Int_ComSet(void)
{
	outportb(COM1+1,0x01);
	outportb(ICW2,inportb(ICW2) & E_Mask1);
	outportb(COM2+1,0x01);
	outportb(ICW2,inportb(ICW2) & E_Mask2);
}

static void interrupt N_ComVec1()
{
	BufAdd(1);
	outportb(ICW1,0x20);
}

static void interrupt N_ComVec2()
{
	BufAdd(2);
	outportb(ICW1,0x20);
}

/*-------------------------------------------------------------------*/

void clearBuffer(BYTE ComP)
{
	qHead[ComP-1]=qTail[ComP-1]=0;
	inportb(WhatCom(ComP));
	inportb(WhatCom(ComP)+6);
}

void BufAdd(BYTE ComP)
{
static	BYTE ch;

	ch=inportb(WhatCom(ComP));
	if ((qHead[ComP-1]+1==qTail[ComP-1]) ||
			(
			(qHead[ComP-1]+1 == MaxBuf) && (qTail[ComP-1] == 0) ))
			return;
	else {
			Buf[ComP-1][qHead[ComP-1]]=ch;
			qHead[ComP-1]++ ;
			if (qHead[ComP-1]==MaxBuf)	qHead[ComP-1]=0;
	}
}

INT BufEmpty(BYTE ComP)
{
	if(qHead[ComP-1]==qTail[ComP-1])
				return TRUE;
		else	return FALSE;
}

BYTE GetBuf(BYTE ComP)
{
	INT i=qTail[ComP-1];
	qTail[ComP-1] = ((qTail[ComP-1]+1)==MaxBuf) ? 0 : qTail[ComP-1]+1;
	return Buf[ComP-1][i];
}

INT WhatCom(BYTE ComP)
{
	if (ComP==1) return (COM1);
	return (COM2);
}

void OutPort(BYTE ComP, BYTE ch)
{
	while ((inportb(WhatCom(ComP)+5) & 0x20) == 0) ;
	outportb(WhatCom(ComP),ch);
}

void Initialize()
{
	LONG speed;
	BYTE i;

	Parity = NONE;
/*	Parity = EVEN;*/
/*	Parity = ODD;*/
	DataBit=D_8BIT;
	StopBit=S_1BIT;
	speed=4800;
/*	printf("speed=%d\n",speed);*/
	for (i=1;i<=2;i++) {
		setBaud(i,speed);
		setLine(i,DataBit,StopBit,Parity);
		setModem(i);
	}
	disable();
	V_ComSet();
	Int_ComSet();
	clearBuffer(1);
	clearBuffer(2);
	enable();
}

void SysClose()
{
	V_ComRst();
}

int BccChk(char *buf,int cnt)
{
	register i,	Bcc=0x00;

	for (i=3;i<cnt;i++)	Bcc ^= buf[i];
	return(Bcc);
}

static void (interrupt far *O_TimeVec)();
static void interrupt far N_TimeVec();

static void V_TimeSet()
{
	if (O_TimeVec==NULL) {
		O_TimeVec=getvect(TIMER_IntV);
		setvect(TIMER_IntV,N_TimeVec);
	}
}
static void V_TimeRst()
{
	if (O_TimeVec) {
		setvect(TIMER_IntV, O_TimeVec);
		O_TimeVec = NULL;
	}
}

static void interrupt far N_TimeVec()
{
	(*O_TimeVec)();
	if (tic) {
		--tic;
	}
}

extern void TimeKill(float msec)
{
	V_TimeSet();
	SetTimer(msec);
	while (TimedOut()==0);
	V_TimeRst();
}

int main(int argc, char *argv[])
{
	FILE *fp2;
	int being = TRUE;
	int datasure,sohset,timer;
	int timeout = 0;
	BYTE ch,i;
	if(argc < 2){
		 i = 2;
	}else{
		i = (*argv[1]-0x30) ;
	}
/*	printf("i = %d ",i);*/
	timer = 1;  /*time*/
	Initialize();

	fp2 = fopen("giorn.dat", "w");

	while(BufEmpty(i)==TRUE){
		OutPort(i, NAK);
/*		printf("[%d,%x]",i,NAK);*/
		TimeKill(timer);
	}

  while (being) {
        while(BufEmpty(i)!=TRUE) {
			  ch=GetBuf(i);
/*			  printf("%x ",ch);*/
			  if(ch == EOT){
					being = FALSE;
					datasure = TRUE;
			  }else{
				  if(ch==0xB0) ch=0x20;
				  fputc(ch, fp2);
			  }
		}/* //while*/
		OutPort(i, ACK);
/*		printf("ACK");*/
		if(being!= FALSE)TimeKill(timer);

		if (BufEmpty(i)==TRUE) timeout++;
		if(timeout > 10){
/*	 	 printf("Time out2");*/
		 being = FALSE;
		 datasure = FALSE;
		 break;
	   }
	}/*//while*/
	fclose(fp2);

	return datasure;
}

