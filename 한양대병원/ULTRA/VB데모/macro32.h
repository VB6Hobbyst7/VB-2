#define	SUCC		1
#define FAIL		0

#define	SQLCODE		sqlca.sqlcode
#define	SQLERRD2	sqlca.sqlerrd[2]
#define	SQLOK		0
#define SQLERRMSG	sqlca.sqlerrm
#define	SQLNODATA	100

#define	INPUT		1
#define	TUX		2
#define INFO		3
#define APP		4
#define USR		5

#define OCCUR(a)		Foccur32(transf,a)
#define GETNUM(f,i,v)		v=Fvall32(transf, f, i)
#define GETSTR(f,i,v)		strcpy(v, Fvals32(transf, f, i))
#define GETDBL(f,i,v)		Fget32(transf, f, i, (char *) &v, (FLDLEN32)0)
#define GETVAR(x,y,z)	{ \
   Fget32((FBFR32 *)(transf), (x), (y), (char *)(z.arr), 0); \
   (z.len) = strlen((char *)(z.arr)); \
   (z.arr)[(z.len)]=0x00; }

#define PUTNUM(f,i,v)		Fchg32(transf, f, i, (char *) &v, (FLDLEN32)0)
#define PUTSTR(f,i,v)		Fchg32(transf, f, i, (char *) v, (FLDLEN32)0)
#define PUTDBL(f,i,v)		Fchg32(transf, f, i, (char *) &v, (FLDLEN32)0)
#define PUTVAR(x,y,z)	{ \
   (z.arr)[(z.len)]=0x00; \
   Fchg32((FBFR32 *)(transf), (x), (y), (char *)(z.arr), (z.len)); }

#define BUFGETNUM(b,f,i,v)	v=Fvall32(b, f, i)
#define BUFGETSTR(b,f,i,v)	strcpy(v, Fvals32(b, f, i))
#define BUFGETDBL(b,f,i,v)	Fget32(b, f, i, (char *) &v, (FLDLEN32)0)
#define BUFPUTNUM(b,f,i,v)	Fchg32(b, f, i, (char *) &v, (FLDLEN32)0)
#define BUFPUTSTR(b,f,i,v)	Fchg32(b, f, i, (char *) v, (FLDLEN32)0)
#define BUFPUTDBL(b,f,i,v)	Fchg32(b, f, i, (char *) &v, (FLDLEN32)0)
#define BFPUTNUM(b,f,v)		Fadd32(b, f, (char *) &v, (FLDLEN32)0)
#define BFPUTSTR(b,f,v)		Fadd32(b, f, (char *) v, (FLDLEN32)0)
#define BFPUTDBL(b,f,v)		Fadd32(b, f, (char *) &v, (FLDLEN32)0)


#define INITSTR(x)	memset(x, 0x00, sizeof(x))
#define INITHOST(x,y)	memset(x, 0x00, sizeof(y))
#define INITDATE(x)	rstrdate("", &x)
#define INITDATETIME(x)	dtcvasc	("", &x)

#define TPRETURN(x,y)   { \
   Fchg32((FBFR32 *)(transf), STATLIN, 0, (char *)(x),0); \
   tpreturn(TPSUCCESS, y, (char *)(transf), 0, 0); }

#define TPRETURN_ERROR(x,y)     { \
   Fchg32((FBFR32 *)(transf), STATLIN, 0, (char *)(x),0); \
   tpreturn(TPFAIL, y, (char *)(transf), 0, 0); }
