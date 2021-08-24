#include <stdio.h>
#include <stdlib.h>

void main()
{
	FILE *fp1, *fp2;
	unsigned char sBuf;

	if((fp1 = fopen("giorn.ris", "r")) == NULL)return ;
	fp2 = fopen("giorn.dat", "w");

	while((sBuf=fgetc(fp1)) !=0xFF)	{
		if(sBuf==0xB0) sBuf=0x20;
		fputc(sBuf, fp2);
	}
	
	fclose(fp1);
	fclose(fp2);

	return ;
}
