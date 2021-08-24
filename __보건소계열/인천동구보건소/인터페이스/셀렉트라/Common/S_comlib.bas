Attribute VB_Name = "S_COMLIB"
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  S0COMLIB = 공통변수 선언 Library                            *
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


Public ary_blditem(1 To 21)   As String      ' 혈액 성분제제리스트 배열
Public ary_week(1 To 7) As String

'*****************************************
'*** SERVER NAME
'*****************************************
'Public  Const S0COM_SERVER1 = "DSN=LAB01;DATABASES=;UID=Admin;PWD=;"
'Public Const S0COM_SERVER01 = "DSN=LAB01;DATABASES=yocsdt;UID=sa;PWD=;"    'TEST 용
Public Const S0COM_SERVER01 = "DSN=LAB01;DATABASES=yocscp;UID=cp01;PWD=cp01;"    '임상병리과 서버
Public Const S0COM_SERVER02 = "DSN=LAB02;DATABASES=yocsad;UID=ad01;PWD=ad01;"    '병동 서버(원무)
Public Const S0COM_SERVER03 = "DSN=LAB02;DATABASES=yocsad;UID=ad01;PWD=ad01;"    '병동 서버(처방)
Public Const S0COM_SERVER04 = "DSN=LAB04;DATABASES=yocsns;UID=sa;PWD=;"    '외래 서버(처방)
Public Const S0COM_SERVER05 = "DSN=LAB05;DATABASES=yocsac;UID=sa;PWD=;"    '외래 서버(원무)
Public Const S0COM_SERVER06 = "DSN=LAB06;DATABASES=yocshcl;UID=sa;PWD=;"   '심혈관서버(YOCSHCL)

'*****************************************
'*** SERVER와 관련 변수
'*****************************************
'Public Const S0COM_LOGINID = "dt01"      'LOGIN ID
'Public Const S0COM_LOGINPASS = "dt01"    'LOGIN Password

Public S0COM_CONNECT     As String   'SERVER와의 접속여부
Public S0COM_LOGIN       As Integer  'server login 값
Public S0COM_SYSDATE     As String   'SYSTEM DATE
Public S0COM_SYSTIME     As String   'SYSTEM TIME
Public S0COM_TERMID      As String   'TERMINAL ID
Public S0COM_USERID      As String   'user id
Public S0COM_PASSWORD    As String   'Password

'*****************************************
'*** 코드출력
'*****************************************
Public S0COM_TEXLENGTH   As Integer
Public S0COM_PRTCODE01   As String
Public S0COM_PRTCODE02   As String

'*****************************************
'*** FORM간의 값을 전달
'*****************************************
Public S0COM_FORMNAME As String      'FORM간의 값을 전달
Public S0COM_PATH        As String   'CURRENT DIRECTORY
Public S0COM_DELETEOK    As Integer  '과거 INTERFACE DATA 삭제여부 TRUE, FALSE

'*****************************************
'*** Code/Name Help Parameter
'*****************************************
Public S0COM_table       As String       'Table ID 명
Public S0COM_code_col    As String       '코드 Column 명
Public S0COM_name_col    As String       '코드명칭 Column 명1
Public S0COM_name_co2    As String       '코드명칭 Column 명2
Public S0COM_cd_gbn_col As String        '코드구분 Column 명
Public S0COM_code        As String       '코드 값
Public S0COM_name        As String       '코드명 값
Public S0COM_name1       As String       '코드명 값
Public S0COM_name2       As String       '코드명 값
Public S0COM_name3       As String       '코드명 값
Public S0COM_cd_gbn      As String       '코드구분 값
Public S0COM_length      As Integer      'S0COM_NAME_VAL을 받을 Label의 길이
Public S0COM_length2     As Integer      'S0COM_NAME_VAL을 받을 Label의 길이
Public S0COM_ret         As Integer      'Return 값

'*****************************************
'*** PROCESS DEFINE
'*****************************************
Public Server        As String
Public SqlStr        As String
'Public SqlDoc        As String
Public Counter       As Integer
Public Login         As Integer
Public pnlMsgBar     As String

'*****************************************
'*** PROCESS RETURN CODE
'*****************************************
Public return_open  As Integer
Public return_close As Integer
Public ret          As Integer
Public Return_cd    As Integer
'Public sql_ret          As Integer

'*****************************************
'***   Help engine declarations.       ***
'***  Commands to pass WinHelp()       ***
'***  const constant.txt define        ***
'*****************************************
Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any) As Integer

Type MULTIKEYHELP
    mkSize As Integer
    mkKeylist As String * 1
    szKeyphrase As String * 253
End Type

'*****************************************
'*** 배지 Setting                      ***
'*****************************************
Public S0COM_CLTCODE(1 To 12)  As String   '배지명 SETTING
Public Const S0COM_CLTCOUNT = 12  '배지 갯수

'*****************************************
'*** 항생제결과                        ***
'*****************************************
Type ANTI_MST
    anticd(1 To 60) As String
    CHECK(1 To 60)  As String
    rst1(1 To 60)   As String
    rst2(1 To 60)   As String
End Type

Type PRT_ANTI_MST
    bctno(1 To 10)      As String
    rstvalue(1 To 10)   As String
End Type
    
'-------------------------------------*
' 한글 Lib function declare           *
'-------------------------------------*
Declare Sub SetImeMode Lib "han.dll" (ByVal hWnd As Integer, ByVal bol As Integer)
Declare Function getimemode Lib "han.dll" (ByVal hWnd As Integer) As Integer
Declare Sub cvtToHAN Lib "c:\windows\system\cvt_ime.dll" (ByVal hWnd As Integer)
Declare Sub cvtToENG Lib "c:\windows\system\cvt_ime.dll" (ByVal hWnd As Integer)

