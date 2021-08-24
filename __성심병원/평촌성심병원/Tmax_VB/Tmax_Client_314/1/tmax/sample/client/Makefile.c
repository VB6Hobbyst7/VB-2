#
#

LIBS    = /link $(TMAXDIR)\lib\tmax.lib
CFLAGS  = /D _DEBUG /I $(TMAXDIR) /Gd /MD /nologo

TARGET  = $(COMP_TARGET)
APOBJ   = $(TARGET:.exe=.obj)
OBJS    = $(APOBJ)

$(TARGET): $(OBJS)
	cl $(CFLAGS) /o $@ $(OBJS) $(LIBS) 


clean:
	-del $(OBJS) $(TARGET)

