/* file: simpserv.c */
#include <stdio.h>
#include <string.h>
#include <atmi.h>
#include <fml32.h>
#include "fld.tbl.h"

FTOUPPER(TPSVCINFO *rqst)
{
        FBFR32 *transf;
	int  i;
        char temp[100];
	long len;

	
        transf=(FBFR32 *)rqst->data;
	memset(temp, 0, sizeof(temp));

        Fget32(transf,buf,0,temp,0);

        for(i=0; i < (strlen(temp)+1); i++)
		temp[i] = toupper(temp[i]);

        (void)Fchg32(transf,buf,0,temp,(FLDLEN32)0);
        (void)Fchg32(transf,buf,1,temp,(FLDLEN32)0);
        (void)Fchg32(transf,buf,2,temp,(FLDLEN32)0);
        (void)Fchg32(transf,buf,3,temp,(FLDLEN32)0);

  	/* Return the transformed buffer */
	tpreturn(TPSUCCESS, 0, rqst->data, 0L, 0 );
}

