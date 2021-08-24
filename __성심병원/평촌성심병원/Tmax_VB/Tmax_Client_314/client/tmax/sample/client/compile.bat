@echo off
IF !%1==! GOTO EXITO
IF %1 ==? GOTO HELP
IF !%2==! GOTO EXITO
cls
@echo *****************************************
@echo [%TIME%] %2 compile start 
@echo *****************************************

SET MKFILE=Makefile.%1
SET COMP_TARGET=%2.exe
nmake /f %MKFILE%

@echo *****************************************
@echo [%TIME%] %2 compile end 
@echo *****************************************
GOTO END

:HELP
cls
@echo ***************************************
@echo 1.Usage    : compile [option] [target]
@echo 2.[option] :  c
@echo 3.example  :  compile c toupper
@echo ***************************************
GOTO END

:EXITO
@echo *********************************
@echo Warning!! 
@echo Usage : compile [option] [target]
@echo *********************************

:END
