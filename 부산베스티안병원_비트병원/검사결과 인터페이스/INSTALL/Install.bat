C:
cd C:\KTASP\SetupFiles
gacutil /i C:\KTASP\EXE\BITInterfaceDataController.dll
gacutil /i C:\KTASP\EXE\BITInterfaceResultPrint.dll
gacutil /i C:\KTASP\EXE\BITLabResultInterface.dll
regasm C:\KTASP\EXE\BITLabResultInterface.dll /tlb:C:\KTASP\EXE\BITLabResultInterface.tlb



