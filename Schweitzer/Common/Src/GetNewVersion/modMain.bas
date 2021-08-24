Attribute VB_Name = "modMain"
Option Explicit

Global gExeFile As String
Global gID      As String
Global gPWD     As String

Sub Main()
    '명령 인수로 실행할 화일의 Full Path를 받는다.
    gExeFile = Command$
    
'''''    aryCmd = GetCommandLine(3)
'''''    gExeFile = aryCmd(1)
''''''    ReDim aryCmd(1)
''''''    ReDim aryCmd(2)
''''''    aryCmd(1) = "BBS"
''''''    aryCmd(2) = "APS"
'''''    strProjectId = aryCmd(1)
'''''    If Trim(strProjectId) = "" Then strProjectId = "LIS"
''''''    blnDownloadMyself = IIf(aryCmd(2) = "1", True, False)
'''''
'''''    RegHdApp = App.LegalTrademarks & " " & strProjectId
'''''    RegHdSet = App.LegalTrademarks & " " & strProjectId

    frmVersionCheck.Show
End Sub
   

