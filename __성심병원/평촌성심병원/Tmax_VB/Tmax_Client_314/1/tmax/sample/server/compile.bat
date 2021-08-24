@echo off
IF !%1==! GOTO EXITO
IF %1 ==? GOTO HELP
IF !%2==! GOTO EXITO
cls
@echo *****************************************
@echo [%TIME%] %2 compile start 
@echo *****************************************

SET MKFILE=Makefile.%1
SET TARGET=%2.exe
nmake /f %MKFILE%

@echo *****************************************
@echo [%TIME%] %2 compile end
@echo *****************************************
GOTO END

:HELP
@echo ***************************************
@echo 1.Usage    : compile [option] [target]
@echo 2.[option] :  c/pc/sdl/psdl
@echo 3.example  :  compile c svr2
@echo ***************************************
GOTO END

:EXITO
@echo *********************************
@echo Warning!! 
@echo Usage : compile [option] [target]
@echo *********************************

:END
