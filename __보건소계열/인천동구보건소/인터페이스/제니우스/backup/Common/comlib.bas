Attribute VB_Name = "comLIB"
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  COMLIB = 공통변수 선언 Library                            *
'*                                                              *
'*  System    : 공통정보 시스템                                 *
'*  Subsystem : 공통코드 관리                                   *
'*  Function  : 공통변수 정의                                   *
'*                                                              *
'*  Designed  :                                                 *
'*  Coded     :                                                 *
'*  Modified  :                                                 *
'*                                                              *
'*                                                              *
'*  < Attentation Please ! >                                    *
'*                                                              *
'*    본 화일 변수의 등록, 수정, 삭제는 반드시 공통 담당인에게  *
'*    의뢰하도록 할 것 !                                        *
'*    임의 수정은 모두를 위하여 삼가합시다 !                    *
'*                                                              *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Option Explicit


''/**********************************
''***    MDB관련 상수
''/**********************************
'Public DB   As Database
'Public TB_CODE_COL As Field
'Public TB_NAME_COL As Field

'*****************************************
'*** SERVER NAME
'*****************************************
'Public Const D0COM_SERVER01 = "DSN=DJLAB;DATABASES=;UID=Admin;PWD=;"
'Public Const D0COM_SERVER01 = "DSN=DJLAB;DATABASES=JCE3;UID=DJP;PWD=;"    '임상병리과 서버
Public Const D0COM_SERVER01 = "DSN=DJLAB;DATABASES=SHARP;UID=sa;PWD=;"    '임상병리과 서버

Public D0COM_SERVER As String

'*****************************************
'*** SERVER와 관련 변수
'*****************************************
Public D0COM_SYSDATE     As String   'SYSTEM DATE
Public D0COM_SYSTIME     As String   'SYSTEM TIME
Public D0COM_TERMID      As String   'TERMINAL ID

Public D0COM_USERID      As String   'user id
Public D0COM_USERNM      As String   'user Name
Public D0COM_PASSWORD    As String   'Password

'*****************************************
'*** FORM간의 값을 전달
'*****************************************
Public D0COM_FORMNAME As String      'FORM간의 값을 전달

Public D0COM_PATH        As String   'CURRENT DIRECTORY
Public D0COM_DELETEOK    As Integer  '과거 INTERFACE DATA 삭제여부 TRUE, FALSE

'*****************************************
'*** Code/Name Help Parameter
'*****************************************
Public D0COM_table       As String       'Table ID 명
Public D0COM_code_col    As String       '코드 Column 명
Public D0COM_name_col    As String       '코드명칭 Column 명
Public D0COM_name_co2    As String       '코드명칭 Column 명
Public D0COM_cd_gbn_col As String        '코드구분 Column 명
Public D0COM_code        As String       '코드 값
Public D0COM_name        As String       '코드명 값
Public D0COM_name1       As String       '코드명 값
Public D0COM_name2       As String       '코드명 값
Public D0COM_name3       As String       '코드명 값
Public D0COM_cd_gbn      As String       '코드구분 값
Public D0COM_length      As Integer      'D0COM_NAME_VAL을 받을 Label의 길이
Public D0COM_length2     As Integer      'D0COM_NAME_VAL을 받을 Label의 길이
Public D0COM_ret         As Integer      'Return 값

''*****************************************
''*** PROCESS DEFINE
''*****************************************
'Public Server        As String
'Public sStr        As String
'Public Counter       As Integer
'Public Login         As Integer
'Public pnlMsgBar     As String

''*****************************************
''*** PROCESS RETURN CODE
''*****************************************
'Public return_open   As Integer
'Public return_close As Integer
'Public ret           As Integer

''*****************************************
''***   Help engine declarations.       ***
''***  Commands to pass WinHelp()       ***
''***  const constant.txt define        ***
''*****************************************
'Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any) As Integer

Type MULTIKEYHELP
    mkSize As Integer
    mkKeylist As String * 1
    szKeyphrase As String * 253
End Type

