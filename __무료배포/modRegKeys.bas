Attribute VB_Name = "modRegKeys"
'-----------------------------------------------------------------------------'
'   ���ϸ�  : modRegKeys.bas
'   �ۼ���  : ������
'   ��  ��  : QCS ������Ʈ�� ���� Module
'   �ۼ���  : 2015-04-29
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit
'---------------------------------------------------------------
'- ������Ʈ�� API ����...
'---------------------------------------------------------------
' Function prototypes, constants, and type definitions for Windows 32-bit Registry API
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'---------------------------------------------------------------
'- ������Ʈ�� API ���...
'---------------------------------------------------------------
' ������Ʈ�� ������ ����...
Private Const REG_NONE = 0                       ' No value type
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings

' ������Ʈ���� ���� ���� �ۼ��մϴ�...
Private Const REG_OPTION_NON_VOLATILE = 0       ' �ý����� ����õǾ Ű�� �����˴ϴ�.
Private Const REG_OPTION_VOLATILE = 1           ' �ý����� ����õǸ� Ű�� ������������.

' ������Ʈ�� Ű ���� �ɼ�...
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Private Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Private Const KEY_EXECUTE = KEY_READ
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ������Ʈ�� Ű ROOT ����...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' ��ȯ��...
Private Const ERROR_NONE = 0
Private Const ERROR_BADKEY = 2
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_SUCCESS = 0
    
Private r           As Long
Private lValueType  As Long
'---------------------------------------------------------------
'- ������Ʈ�� ���� Ư�� ����...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type


'-- ������Ʈ�� ����
Global Const REG_COMPANY    As String = "OKSOFT"
Global Const REG_SYSTEM     As String = "INTERFACE"
Global Const REG_POSITION   As String = "Software\" & REG_COMPANY & "\" & REG_SYSTEM

Public REG_MACH             As String

'Global Const REG_SERVER     As String = REG_POSITION & "\CONECT_SERVER"             '����
'Global Const REG_SERVER2    As String = REG_POSITION & "\CONECT_SERVER2"            '����2
'Global Const REG_AREAINFO   As String = REG_POSITION & "\AREAINFO"                  ''-- ���μ���


'-- ����Ÿ���̽� �����
Global Const REG_DBTYPE     As String = "DATABASE TYPE"
Global Const REG_SERVER     As String = "SERVER"
Global Const REG_DATABASE   As String = "DATABASE"
Global Const REG_SERVICE    As String = "SERVICE"
Global Const REG_USER_ID    As String = "USERID"
Global Const REG_PASSWD     As String = "PASSWD"
Global Const REG_AREACD     As String = "AREACODE"
Global Const REG_AREANM     As String = "AREANAME"
Global Const REG_CORPNAME   As String = "USER NAME"

'-- INI ���� �б�
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


' ���ҽ� ���ڿ��� ������ ���� ��Ʈ���� �Ӽ��� �ε�˴ϴ�.
' Object      Property
' Form        Caption
' Menu        Caption
' TabStrip    Caption, ToolTipText
' Toolbar     ToolTipText
' ListView    ColumnHeader.Text

'-------------------------------------------------------------------------------------------------
'���� ��� - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
'-------------------------------------------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' �ڵ� ��ȯ
    Dim hKey As Long                                    ' ������Ʈ�� Ű ó��
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' ������Ʈ�� ���� ����

    lpAttr.nLength = 50                                 ' ���� Ư���� �⺻���� ����...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- ������Ʈ�� Ű �����/����...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' �����/���� //KeyRoot//KeyName

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ���� ó��...

    '------------------------------------------------------------
    '- Ű �� �����/����...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' RegSetValueEx()�� ����ϱ� ���� �� ĭ�� �ʿ��մϴ�...

    ' Create/Modify Key Value
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ���� ó��
    '------------------------------------------------------------
    '- ������Ʈ�� Ű �ݱ�...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Ű�� ����

    UpdateKey = True                                    ' ������ ��ȯ
    Exit Function                                       ' ����
CreateKeyError:
    UpdateKey = False                                   ' ���� ��ȯ �ڵ带 ����
    rc = RegCloseKey(hKey)                              ' Ű �ݱ⸦ �õ�
End Function

'-------------------------------------------------------------------------------------------------
'���� ���� - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef tmpVal As String) As String
    Dim i           As Long                                 ' ���� ī����
    Dim rc          As Long                                 ' �ڵ� ��ȯ
    Dim hKey        As Long                                 ' ���� ������Ʈ�� Ű�� �ڵ�
    Dim hDepth      As Long                                 '
    Dim sKeyVal     As String
    Dim lKeyValType As Long                                 ' ������Ʈ�� Ű�� ������ ����
'    Dim tmpVal      As String                               ' ������Ʈ�� Ű ���� �ӽ� ����
    Dim KeyValSize  As Long                                 ' ������Ʈ�� Ű ������ ũ��

    ' KeyRoot {HKEY_LOCAL_MACHINE...} �Ʒ��� RegKey ����
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ������Ʈ�� Ű ����

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ���� ó��...

    tmpVal = String$(1024, 0)                             ' ���� ���� �Ҵ�
    KeyValSize = 1024                                       ' ���� ũ�� ǥ��

    '------------------------------------------------------------
    ' ������Ʈ�� Ű �� �˻�...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Ű �� �˾Ƴ���/�����

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ���� ó��

    tmpVal = LEFT$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' ��ȯ�� ���� Ű �� ���� ����...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' ������ ���� �˻�...
    Case REG_SZ, REG_EXPAND_SZ                              ' ���ڿ� ������Ʈ�� Ű ������ ����
        sKeyVal = tmpVal                                     ' ���ڿ� �� ����
    Case REG_DWORD                                          ' Double Word ������Ʈ�� Ű ������ ����
        For i = Len(tmpVal) To 1 Step -1                    ' ��Ʈ�� ��ȯ
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Char ������ �� Char�� ����
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Double Word�� String�� ��ȯ
    End Select

    GetKeyValue = sKeyVal                                   ' �� ��ȯ
    rc = RegCloseKey(hKey)                                  ' ������Ʈ�� Ű �ݱ�
    Exit Function                                           ' ����

GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' ����ִ� ���ڿ��� ��ȯ ���� ����
    rc = RegCloseKey(hKey)                                  ' ������Ʈ�� Ű�� ����
End Function

'������ Ʈ���� Ű �����
Public Sub SaveKey(hKey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

'������ Ʈ���� Ű �����
Public Function DeleteKey(ByVal hKey As Long, ByVal strkey As String)
    Dim r As Long
    r = RegDeleteKey(hKey, strkey)
End Function

'������ Ʈ���� Ű�� �����
Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

'������ Ʈ���� ���ڿ��� ��������
Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim DataType As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = LEFT$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

'������ Ʈ���� ���ڿ��� ����
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)))
    r = RegCloseKey(keyhand)
End Sub

'������ Ʈ���� BINARY�� ��������
Public Function GetBINARY(hKey As Long, strPath As String, strValue As String)

    Dim keyhand As Long
    Dim DataType As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_BINARY Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetBINARY = LEFT$(strBuf, intZeroPos - 1)
            Else
                GetBINARY = strBuf
            End If
        End If
    End If
End Function

'������ Ʈ���� BINARY�� ����
Public Sub SaveBINARY(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_BINARY, ByVal strData, LenB(StrConv(strData, vbFromUnicode)))
    r = RegCloseKey(keyhand)
End Sub

'������ Ʈ���� ����Ÿ ���ڿ��� ��������
Function GetDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    
    r = RegOpenKey(hKey, strPath, keyhand)
    
     ' Get length/data type
    lDataBufSize = 4
        
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDword = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
End Function

'������ Ʈ���� ����Ÿ ���ڿ��� ����
Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function





