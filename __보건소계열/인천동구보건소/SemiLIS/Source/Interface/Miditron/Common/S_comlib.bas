Attribute VB_Name = "S_COMLIB"
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  S0COMLIB = ���뺯�� ���� Library                            *
'*                                                              *
'*  System    : �������� �ý���                                 *
'*  Subsystem : �����ڵ� ����                                   *
'*  Function  : ���뺯�� ����                                   *
'*                                                              *
'*  Designed  :                                                 *
'*  Coded     :                                                 *
'*  Modified  :                                                 *
'*                                                              *
'*                                                              *
'*  < Attentation Please ! >                                    *
'*                                                              *
'*    �� ȭ�� ������ ���, ����, ������ �ݵ�� ���� ����ο���  *
'*    �Ƿ��ϵ��� �� �� !                                        *
'*    ���� ������ ��θ� ���Ͽ� �ﰡ�սô� !                    *
'*                                                              *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
Option Explicit


Public ary_blditem(1 To 21)   As String      ' ���� ������������Ʈ �迭
Public ary_week(1 To 7) As String

'*****************************************
'*** SERVER NAME
'*****************************************
'Public  Const S0COM_SERVER1 = "DSN=LAB01;DATABASES=;UID=Admin;PWD=;"
'Public Const S0COM_SERVER01 = "DSN=LAB01;DATABASES=yocsdt;UID=sa;PWD=;"    'TEST ��
Public Const S0COM_SERVER01 = "DSN=LAB01;DATABASES=yocscp;UID=cp01;PWD=cp01;"    '�ӻ󺴸��� ����
Public Const S0COM_SERVER02 = "DSN=LAB02;DATABASES=yocsad;UID=ad01;PWD=ad01;"    '���� ����(����)
Public Const S0COM_SERVER03 = "DSN=LAB02;DATABASES=yocsad;UID=ad01;PWD=ad01;"    '���� ����(ó��)
Public Const S0COM_SERVER04 = "DSN=LAB04;DATABASES=yocsns;UID=sa;PWD=;"    '�ܷ� ����(ó��)
Public Const S0COM_SERVER05 = "DSN=LAB05;DATABASES=yocsac;UID=sa;PWD=;"    '�ܷ� ����(����)
Public Const S0COM_SERVER06 = "DSN=LAB06;DATABASES=yocshcl;UID=sa;PWD=;"   '����������(YOCSHCL)

'*****************************************
'*** SERVER�� ���� ����
'*****************************************
'Public Const S0COM_LOGINID = "dt01"      'LOGIN ID
'Public Const S0COM_LOGINPASS = "dt01"    'LOGIN Password

Public S0COM_CONNECT     As String   'SERVER���� ���ӿ���
Public S0COM_LOGIN       As Integer  'server login ��
Public S0COM_SYSDATE     As String   'SYSTEM DATE
Public S0COM_SYSTIME     As String   'SYSTEM TIME
Public S0COM_TERMID      As String   'TERMINAL ID
Public S0COM_USERID      As String   'user id
Public S0COM_PASSWORD    As String   'Password

'*****************************************
'*** �ڵ����
'*****************************************
Public S0COM_TEXLENGTH   As Integer
Public S0COM_PRTCODE01   As String
Public S0COM_PRTCODE02   As String

'*****************************************
'*** FORM���� ���� ����
'*****************************************
Public S0COM_FORMNAME As String      'FORM���� ���� ����
Public S0COM_PATH        As String   'CURRENT DIRECTORY
Public S0COM_DELETEOK    As Integer  '���� INTERFACE DATA �������� TRUE, FALSE

'*****************************************
'*** Code/Name Help Parameter
'*****************************************
Public S0COM_table       As String       'Table ID ��
Public S0COM_code_col    As String       '�ڵ� Column ��
Public S0COM_name_col    As String       '�ڵ��Ī Column ��1
Public S0COM_name_co2    As String       '�ڵ��Ī Column ��2
Public S0COM_cd_gbn_col As String        '�ڵ屸�� Column ��
Public S0COM_code        As String       '�ڵ� ��
Public S0COM_name        As String       '�ڵ�� ��
Public S0COM_name1       As String       '�ڵ�� ��
Public S0COM_name2       As String       '�ڵ�� ��
Public S0COM_name3       As String       '�ڵ�� ��
Public S0COM_cd_gbn      As String       '�ڵ屸�� ��
Public S0COM_length      As Integer      'S0COM_NAME_VAL�� ���� Label�� ����
Public S0COM_length2     As Integer      'S0COM_NAME_VAL�� ���� Label�� ����
Public S0COM_ret         As Integer      'Return ��

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
'*** ���� Setting                      ***
'*****************************************
Public S0COM_CLTCODE(1 To 12)  As String   '������ SETTING
Public Const S0COM_CLTCOUNT = 12  '���� ����

'*****************************************
'*** �׻������                        ***
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
' �ѱ� Lib function declare           *
'-------------------------------------*
Declare Sub SetImeMode Lib "han.dll" (ByVal hWnd As Integer, ByVal bol As Integer)
Declare Function getimemode Lib "han.dll" (ByVal hWnd As Integer) As Integer
Declare Sub cvtToHAN Lib "c:\windows\system\cvt_ime.dll" (ByVal hWnd As Integer)
Declare Sub cvtToENG Lib "c:\windows\system\cvt_ime.dll" (ByVal hWnd As Integer)

