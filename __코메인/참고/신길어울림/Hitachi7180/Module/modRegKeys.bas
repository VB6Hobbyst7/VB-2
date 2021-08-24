Attribute VB_Name = "modRegKeys"
' �� ����� ������Ʈ�� Ű�� �а� ���ϴ�. VB�� ���� ������Ʈ��
' �׼��� ����� �޸� ���ڿ� ������ ������Ʈ�� Ű��
' �а� �� �� �ֽ��ϴ�.

Option Explicit
'---------------------------------------------------------------
'- ������Ʈ�� API ����...
'---------------------------------------------------------------
' Function prototypes, constants, and type definitions for Windows 32-bit Registry API
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
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
    Dim hkey As Long                                    ' ������Ʈ�� Ű ó��
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
                        hkey, hDepth)                   ' �����/���� //KeyRoot//KeyName

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ���� ó��...

    '------------------------------------------------------------
    '- Ű �� �����/����...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' RegSetValueEx()�� ����ϱ� ���� �� ĭ�� �ʿ��մϴ�...

    ' Create/Modify Key Value
    rc = RegSetValueEx(hkey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' ���� ó��
    '------------------------------------------------------------
    '- ������Ʈ�� Ű �ݱ�...
    '------------------------------------------------------------
    rc = RegCloseKey(hkey)                              ' Ű�� ����

    UpdateKey = True                                    ' ������ ��ȯ
    Exit Function                                       ' ����
CreateKeyError:
    UpdateKey = False                                   ' ���� ��ȯ �ڵ带 ����
    rc = RegCloseKey(hkey)                              ' Ű �ݱ⸦ �õ�
End Function

'-------------------------------------------------------------------------------------------------
'���� ���� - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef tmpVal As String) As String
    Dim i           As Long                                 ' ���� ī����
    Dim rc          As Long                                 ' �ڵ� ��ȯ
    Dim hkey        As Long                                 ' ���� ������Ʈ�� Ű�� �ڵ�
    Dim hDepth      As Long                                 '
    Dim sKeyVal     As String
    Dim lKeyValType As Long                                 ' ������Ʈ�� Ű�� ������ ����
'    Dim tmpVal      As String                               ' ������Ʈ�� Ű ���� �ӽ� ����
    Dim KeyValSize  As Long                                 ' ������Ʈ�� Ű ������ ũ��

    ' KeyRoot {HKEY_LOCAL_MACHINE...} �Ʒ��� RegKey ����
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hkey) ' ������Ʈ�� Ű ����

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ���� ó��...

    tmpVal = String$(1024, 0)                             ' ���� ���� �Ҵ�
    KeyValSize = 1024                                       ' ���� ũ�� ǥ��

    '------------------------------------------------------------
    ' ������Ʈ�� Ű �� �˻�...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hkey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Ű �� �˾Ƴ���/�����

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ���� ó��

    tmpVal = left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

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
    rc = RegCloseKey(hkey)                                  ' ������Ʈ�� Ű �ݱ�
    Exit Function                                           ' ����

GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' ����ִ� ���ڿ��� ��ȯ ���� ����
    rc = RegCloseKey(hkey)                                  ' ������Ʈ�� Ű�� ����
End Function

'������ Ʈ���� Ű �����
Public Sub SaveKey(hkey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

'������ Ʈ���� Ű �����
Public Function DeleteKey(ByVal hkey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(hkey, strKey)
End Function

'������ Ʈ���� Ű�� �����
Public Function DeleteValue(ByVal hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

'������ Ʈ���� ���ڿ��� ��������
Public Function GetString(hkey As Long, strPath As String, strValue As String)

    Dim keyhand As Long
    Dim DataType As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    
    r = RegOpenKey(hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

'������ Ʈ���� ���ڿ��� ����
Public Sub SaveString(hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, LenB(StrConv(strdata, vbFromUnicode)))
    r = RegCloseKey(keyhand)
End Sub

'������ Ʈ���� BINARY�� ��������
Public Function GetBINARY(hkey As Long, strPath As String, strValue As String)

    Dim keyhand As Long
    Dim DataType As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    
    r = RegOpenKey(hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_BINARY Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetBINARY = left$(strBuf, intZeroPos - 1)
            Else
                GetBINARY = strBuf
            End If
        End If
    End If
End Function

'������ Ʈ���� BINARY�� ����
Public Sub SaveBINARY(hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_BINARY, ByVal strdata, LenB(StrConv(strdata, vbFromUnicode)))
    r = RegCloseKey(keyhand)
End Sub

'������ Ʈ���� ����Ÿ ���ڿ��� ��������
Function GetDword(ByVal hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    
    r = RegOpenKey(hkey, strPath, keyhand)
    
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
Function SaveDword(ByVal hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    
    r = RegCreateKey(hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function

'ODBC mdb ��� ����
Public Sub UpdateODBCMDB(ByVal MDBName As String)
    Dim sSubKey As String
    Dim sODBCDriverName As String
    Dim sDSNName As String
    Dim sValue As String
    Dim lValue As Long
    
    Const typeString = 1
    Const typeNumber = 2
    Const ODBCPath = "SOFTWARE\ODBC\ODBC.INI\"
    Const ODBCDataSourcePath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
    '=========================================================
    ' ODBC Driver�� �����.
    '=========================================================
    sODBCDriverName = "InterMDB"
    sSubKey = ODBCPath & sODBCDriverName

    If InStr(GetString(HKEY_LOCAL_MACHINE, sSubKey, "DBQ"), INS_NAME) = 0 Or INS_NAME = "" Then
        CreateKeyValue HKEY_LOCAL_MACHINE, ODBCDataSourcePath, typeString, sODBCDriverName, "Microsoft Access Driver (*.mdb)"
        CreateKeyValue HKEY_LOCAL_MACHINE, sSubKey, typeString, "DBQ", MDBName
        CreateKeyValue HKEY_LOCAL_MACHINE, sSubKey, typeString, "Driver", "C:\WINDOWS\system32\odbcjt32.dll"
        CreateKeyValue HKEY_LOCAL_MACHINE, sSubKey, typeString, "FIL", "MS Access;"
        CreateKeyValue HKEY_LOCAL_MACHINE, sSubKey, typeString, "UID", ""
        CreateKeyValue HKEY_LOCAL_MACHINE, sSubKey, typeNumber, "DriverId", "", 25
        CreateKeyValue HKEY_LOCAL_MACHINE, sSubKey, typeNumber, "SafeTransactions", "", 0
    End If

End Sub

Public Sub CreateKeyValue(ByVal lRoot As Long, _
                                        ByVal sSubKey As String, _
                                        ByVal pDirect As Integer, _
                                        ByVal pValueName As String, _
                                        Optional ByVal psValue As String, _
                                        Optional ByVal plValue As Long)
    Dim lResult As Long
    Dim hKeyHandle As Long
    
    lResult = RegCreateKey(lRoot, sSubKey, hKeyHandle)
'    lResult = RegCreateKeyEx(lRoot, sSubKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, _
                                        KEY_ALL_ACCESS, 0&, hKeyHandle, lResult)
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Error"
    End If
    Select Case pDirect
    Case 1  '������
        'String type : ���� ���� SubKey�� �������� Value Name�� ����� Value�� setting�Ѵ�.
        psValue = psValue & Chr$(0)
        lResult = RegSetValueEx(hKeyHandle, pValueName, 0&, REG_SZ, ByVal psValue, lstrlen(psValue))
    Case 2  '������
        lResult = RegSetValueEx(hKeyHandle, pValueName, 0&, REG_DWORD, plValue, REG_DWORD)
    End Select
    If lResult <> ERROR_SUCCESS Then
        MsgBox "Error"
    Else
        lResult = RegCloseKey(hKeyHandle)
    End If
    'Subkey�� �ݴ´�.
    lResult = RegCloseKey(hKeyHandle)
End Sub



