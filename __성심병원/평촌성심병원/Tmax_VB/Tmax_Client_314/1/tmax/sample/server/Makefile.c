#
#
#TMAXDIR = "c:\tmax"
LIBS    = /link ws2_32.lib $(TMAXDIR)\lib\tmaxsvr.lib
CFLAGS  = /D _DEBUG /I $(TMAXDIR) /Gd /MD /nologo

#
#
APOBJ   = $(TARGET:.exe=.obj)
SVCTSRC = $(TARGET:.exe=_svctab.c)
SVCTOBJ = $(TARGET:.exe=_svctab.obj)
OBJS    = $(APOBJ) $(SVCTOBJ)

#
#
$(TARGET): svct $(OBJS)
	cl $(CFLAGS) -o $@ $(OBJS) $(LIBS)
   copy $@ $(TMAXDIR)\appbin

svct:
	copy $(TMAXDIR)\svct\$(SVCTSRC)

#
clean:
	-del $(OBJS) $(TARGET)

