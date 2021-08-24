#사용법 : 오라클인 경우     --> nmake /f tmsbld.mk oracle
#	   MS-SQL인 경우 --> nmake /f tmsbld.mk mssql

CC           = cl
LD           = link
RM           = del
MV           = move
CP           = copy

CFLAGS=/D _DEBUG /I $(TMAXDIR) /Gd /MD /nologo
PROFLAGS=include=$(TMAXDIR)

ORALIBPATH   = "D:\oracle\ora92\precomp\lib\msvc"
MSQLIBPATH   = "C:\Program Files\Microsoft Visual Studio\VC98\Lib"
LDFLAGS      = /nologo /nodefaultlib:LIBC /nodefaultlib:LIBCMT /libpath:$(MSQLIBPATH) /libpath:$(TMAXDIR)/lib /libpath:$(ORALIBPATH)
SYSLIBS      = kernel32.lib ws2_32.lib user32.lib msvcrt.lib
DBLIBS_ORACL = oraSQL9.LIB oraSQX9.LIB
DBLIBS_MSSQL = xaSwitch.lib
TMSLIBS      = tmaxtms.lib
LIBS         = $(SYSLIBS) $(TMSLIBS)

APOBJS       = tms_main.obj
APSRCS       = $(APOBJS: .obj=.c)
DBSTUB_ORACL = ora_stub.obj
DBSTUB_MSSQL = msqlstub.obj
OBJS         = $(APOBJS)

.SUFFIXES: .c .obj

.c.obj:
	$(CC) $(CFLAGS) -c $<

#
all: oracle mssql

oracle: dbobj_ora tmsmain tms_ora.exe

mssql: dbobj_msq tmsmain tms_msq.exe

tms_ora.exe: $(APOBJS)
	$(LD) $(LDFLAGS) /OUT:$@ $(OBJS) $(DBSTUB_ORACL) $(LIBS) $(DBLIBS_ORACL)
	$(CP) $@ $(TMAXDIR)\appbin\.

tms_msq.exe: $(APOBJS)
	$(LD) $(LDFLAGS) /OUT:$@ $(OBJS) $(DBSTUB_MSSQL) $(LIBS) $(DBLIBS_MSSQL)
	$(CP) $@ $(TMAXDIR)\appbin\.

tms_main.obj: tms_main.c
	$(CC) $(CFLAGS) -c tms_main.c

tmsmain:
	$(CP) $(TMAXDIR)\usrinc\tms_main.c .

dbobj_ora:
	$(CP) $(TMAXDIR)\lib\$(DBSTUB_ORACL) .

dbobj_msq:
	$(CP) $(TMAXDIR)\lib\$(DBSTUB_MSSQL) .

clean:
	$(RM) $(OBJS) tms_ora.exe tms_msq.exe

