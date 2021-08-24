Attribute VB_Name = "modRegKeys"
'-----------------------------------------------------------------------------'
'   파일명  : modRegKeys.bas
'   작성자  : 오세원
'   내  용  : QCS 레지스트리 관련 Module
'   작성일  : 2015-04-29
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit
'---------------------------------------------------------------
'- 레지스트리 API 선언...
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
'- 레지스트리 API 상수...
'---------------------------------------------------------------
' 레지스트리 데이터 형식...
Private Const REG_NONE = 0                       ' No value type
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings

' 레지스트리는 형식 값을 작성합니다...
Private Const REG_OPTION_NON_VOLATILE = 0       ' 시스템이 재부팅되어도 키는 보존됩니다.
Private Const REG_OPTION_VOLATILE = 1           ' 시스템이 재부팅되면 키는 보존하지않음.

' 레지스트리 키 보안 옵션...
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
                     
' 레지스트리 키 ROOT 형식...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' 반환값...
Private Const ERROR_NONE = 0
Private Const ERROR_BADKEY = 2
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_SUCCESS = 0
    
Private r           As Long
Private lValueType  As Long
'---------------------------------------------------------------
'- 레지스트리 보안 특성 형식...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type


'-- 레지스트리 설정
Global Const REG_COMPANY    As String = "OKSOFT"
Global Const REG_SYSTEM     As String = "INTERFACE"
Global Const REG_POSITION   As String = "Software\" & REG_COMPANY & "\" & REG_SYSTEM

Public REG_MACH             As String

'Global Const REG_SERVER     As String = REG_POSITION & "\CONECT_SERVER"             '서버
'Global Const REG_SERVER2    As String = REG_POSITION & "\CONECT_SERVER2"            '서버2
'Global Const REG_AREAINFO   As String = REG_POSITION & "\AREAINFO"                  ''-- 지부설정


'-- 데이타베이스 사용자
Global Const REG_DBTYPE     As String = "DATABASE TYPE"
Global Const REG_SERVER     As String = "SERVER"
Global Const REG_DATABASE   As String = "DATABASE"
Global Const REG_SERVICE    As String = "SERVICE"
Global Const REG_USER_ID    As String = "USERID"
Global Const REG_PASSWD     As String = "PASSWD"
Global Const REG_AREACD     As String = "AREACODE"
Global Const REG_AREANM     As String = "AREANAME"
Global Const REG_CORPNAME   As String = "USER NAME"

'-- INI 파일 읽기
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


' 리소스 문자열은 다음과 같이 컨트롤의 속성에 로드됩니다.
' Object      Property
' Form        Caption
' Menu        Caption
' TabStrip    Caption, ToolTipText
' Toolbar     ToolTipText
' ListView    ColumnHeader.Text

'-------------------------------------------------------------------------------------------------
'예제 사용 - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
'-------------------------------------------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' 코드 반환
    Dim hKey As Long                                    ' 레지스트리 키 처리
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' 레지스트리 보안 형식

    lpAttr.nLength = 50                                 ' 보안 특성을 기본으로 설정...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- 레지스트리 키 만들기/열기...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' 만들기/열기 //KeyRoot//KeyName

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 오류 처리...

    '------------------------------------------------------------
    '- 키 값 만들기/열기...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' RegSetValueEx()를 사용하기 위해 빈 칸이 필요합니다...

    ' Create/Modify Key Value
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' 오류 처리
    '------------------------------------------------------------
    '- 레지스트리 키 닫기...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' 키를 닫음

    UpdateKey = True                                    ' 성공을 반환
    Exit Function                                       ' 끝냄
CreateKeyError:
    UpdateKey = False                                   ' 오류 반환 코드를 설정
    rc = RegCloseKey(hKey)                              ' 키 닫기를 시도
End Function

'-------------------------------------------------------------------------------------------------
'샘플 예제 - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef tmpVal As String) As String
    Dim i           As Long                                 ' 루프 카운터
    Dim rc          As Long                                 ' 코드 반환
    Dim hKey        As Long                                 ' 열린 레지스트리 키의 핸들
    Dim hDepth      As Long                                 '
    Dim sKeyVal     As String
    Dim lKeyValType As Long                                 ' 레지스트리 키의 데이터 형식
'    Dim tmpVal      As String                               ' 레지스트리 키 값의 임시 저장
    Dim KeyValSize  As Long                                 ' 레지스트리 키 변수의 크기

    ' KeyRoot {HKEY_LOCAL_MACHINE...} 아래의 RegKey 열기
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 레지스트리 키 열기

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류 처리...

    tmpVal = String$(1024, 0)                             ' 변수 공간 할당
    KeyValSize = 1024                                       ' 변수 크기 표시

    '------------------------------------------------------------
    ' 레지스트리 키 값 검색...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' 키 값 알아내기/만들기

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 오류 처리

    tmpVal = LEFT$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' 변환을 위한 키 값 형식 결정...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' 데이터 형식 검색...
    Case REG_SZ, REG_EXPAND_SZ                              ' 문자열 레지스트리 키 데이터 형식
        sKeyVal = tmpVal                                     ' 문자열 값 복사
    Case REG_DWORD                                          ' Double Word 레지스트리 키 데이터 형식
        For i = Len(tmpVal) To 1 Step -1                    ' 비트를 변환
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Char 단위로 값 Char을 만듦
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Double Word를 String로 변환
    End Select

    GetKeyValue = sKeyVal                                   ' 값 반환
    rc = RegCloseKey(hKey)                                  ' 레지스트리 키 닫기
    Exit Function                                           ' 끝냄

GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' 비어있는 문자열로 반환 값을 설정
    rc = RegCloseKey(hKey)                                  ' 레지스트리 키를 닫음
End Function

'레지스 트리에 키 만들기
Public Sub SaveKey(hKey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

'레지스 트리에 키 지우기
Public Function DeleteKey(ByVal hKey As Long, ByVal strkey As String)
    Dim r As Long
    r = RegDeleteKey(hKey, strkey)
End Function

'레지스 트리에 키값 지우기
Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

'레지스 트리에 문자열값 가져오기
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

'레지스 트리에 문자열값 저장
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)))
    r = RegCloseKey(keyhand)
End Sub

'레지스 트리에 BINARY값 가져오기
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

'레지스 트리에 BINARY값 저장
Public Sub SaveBINARY(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_BINARY, ByVal strData, LenB(StrConv(strData, vbFromUnicode)))
    r = RegCloseKey(keyhand)
End Sub

'레지스 트리에 데이타 문자열값 가져오기
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

'레지스 트리에 데이타 문자열값 저장
Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function





