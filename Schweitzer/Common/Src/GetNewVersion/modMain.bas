Attribute VB_Name = "modMain"
Option Explicit

Global gExeFile As String
Global gID      As String
Global gPWD     As String

Sub Main()
    '��� �μ��� ������ ȭ���� Full Path�� �޴´�.
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
   

