Attribute VB_Name = "modS2DSO"
Option Explicit

Global RegAppName As String

Global Const RegHdApp = "Schweitzer"
    ' File Server Path
Global Const RegSsSet = "Setup"
Global Const RegK1Set = "Server IP"
    ' ��������
Global Const RegK2Set = "Hospital"
Global Const RegK3Set = "HelpLine"
    
    
    ' App Path
Global Const RegSsApp = "App"
Global Const RegK1App = "Path"
Global Const RegK2App = "ExeName"
    
    ' Registry ��� (�������� �ɼ�)
Global Const RegSsOpt = "Options"
Global Const RegK1Opt = "ShowAtStart"
Global Const RegK2Opt = "RunSplash"
    
    ' Registry ��� (�ǹ�����)
Global Const RegSsBld = "Building"
Global Const RegK0Bld = "On/Off"
Global Const RegK1Bld = "Key"
Global Const RegK2Bld = "Name"
Global Const RegK3Bld = "No"
    
    ' Registry ��� (����Ÿ���̽�����)
Global Const RegSsSvr = "Server"
Global Const RegK1Svr = "Name"
Global Const RegK2Svr = "DB"
Global Const RegK3Svr = "UID"
Global Const RegK4Svr = "PWD"
Global Const RegK5Svr = "Type"

'# medAlwaysOn
'Private Declare Function SetWindowPos Lib "user32" _
'    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
'    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
'    ByVal wFlags As Long) As Long
'Private Const HWND_NOTOPMOST = -2   'Not Always top
'Private Const HWND_TOPMOST = -1  'Always top
'Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOSIZE = &H1

'# medGetComNm
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'# Ini�� �ٷ�� api
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



'*-----------------------------------------------------------------
'*  1. ��� : �ش����� �׻� ���� ���ְ� �Ѵ�.
'*  2. Parameter : frmForm - �ش� ��
'*                 OnOff - 0 : ����, 1 : ����
'*-----------------------------------------------------------------
'Public Sub medAlwaysOn(ByVal frmForm As Object, ByVal OnOff As Integer)
'Dim hWndMode As Integer
'
'    hWndMode = Choose(OnOff + 1, -2, -1)
'    SetWindowPos frmForm.hwnd, hWndMode, 0, 0, 10, 10, _
'                                        SWP_NOMOVE Or SWP_NOSIZE
''    SetWindowPos frmForm.hwnd, HWND_TOPMOST, 0, 0, 10, 10, _
'                                        SWP_NOMOVE Or SWP_NOSIZE
'
'End Sub
 


'*-----------------------------------------------------------------
'*  1. ��� : ��ǻ�� �̸� ��������..
'*-----------------------------------------------------------------
'Public Function medGetComNm()
'
'   Dim sBuffer$, nSize As Long, rtn As Long
'   sBuffer = String(256, Chr(0))
'   rtn = GetComputerName(sBuffer$, Len(sBuffer))
'   medGetComNm = sBuffer
'
'End Function


'Public Function medGetP(ByVal strText As String, _
'                  ByVal intPosition As Integer, ByVal Delimiter As String) As String
'
'    Dim intPos1 As Integer, intPos2 As Integer, i As Integer
'
'    intPos1 = 0: intPos2 = 0
'
'    ' intPosition �μ��� 1�� ��� For�� Skip
'    For i = 1 To intPosition - 1
'       intPos1 = intPos2 + 1
'       intPos2 = InStr(intPos2 + 1, strText, Delimiter)
'       If intPos2 = 0 Then GoTo ReturnNull
'    Next i
'
'    ' �ش� �÷�
'    intPos1 = intPos2 + 1
'    intPos2 = InStr(intPos2 + 1, strText, Delimiter)
'    If intPos2 = 0 Then intPos2 = Len(strText) + 1
'
'    medGetP = Mid$(strText, intPos1, intPos2 - intPos1)
'
'    Exit Function
'
'ReturnNull:
'    medGetP = ""
'
'End Function

'' --------------------------------------------------------
'' Registry���� ������ �д� �κ��� INI�� �����ϱ� ���� �Լ�
'' GetSetting�� SaveSetting�� ������ �����ϰ� �������.
'' --------------------------------------------------------
'
'Public Function S2SaveSetting(pAppName As String, pSection As String, pKey As String, pSetting As String) As Long
'    S2SaveSetting = WritePrivateProfileString(pSection, pKey, pSetting, App.Path & "\" & pAppName & ".s2")
'End Function
'
'Public Function S2GetSetting(pAppName As String, pSection As String, pKey As String, Optional pDefault As String = "") As String
'    Dim P As String
'
'    P = Space$(256)
'    Call GetPrivateProfileString(pSection, pKey, pDefault, P, 256, App.Path & "\" & pAppName & ".s2")
'    S2GetSetting = Mid(Trim(P), 1, Len(Trim(P)) - 1)
'End Function
