Attribute VB_Name = "comLIB"
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  COMLIB = ���뺯�� ���� Library                            *
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


''/**********************************
''***    MDB���� ���
''/**********************************
'Public DB   As Database
'Public TB_CODE_COL As Field
'Public TB_NAME_COL As Field

'*****************************************
'*** SERVER NAME
'*****************************************
'Public Const D0COM_SERVER01 = "DSN=DJLAB;DATABASES=;UID=Admin;PWD=;"
'Public Const D0COM_SERVER01 = "DSN=DJLAB;DATABASES=JCE3;UID=DJP;PWD=;"    '�ӻ󺴸��� ����
Public Const D0COM_SERVER01 = "DSN=DJLAB;DATABASES=SHARP;UID=sa;PWD=;"    '�ӻ󺴸��� ����

Public D0COM_SERVER As String

'*****************************************
'*** SERVER�� ���� ����
'*****************************************
Public D0COM_SYSDATE     As String   'SYSTEM DATE
Public D0COM_SYSTIME     As String   'SYSTEM TIME
Public D0COM_TERMID      As String   'TERMINAL ID

Public D0COM_USERID      As String   'user id
Public D0COM_USERNM      As String   'user Name
Public D0COM_PASSWORD    As String   'Password

'*****************************************
'*** FORM���� ���� ����
'*****************************************
Public D0COM_FORMNAME As String      'FORM���� ���� ����

Public D0COM_PATH        As String   'CURRENT DIRECTORY
Public D0COM_DELETEOK    As Integer  '���� INTERFACE DATA �������� TRUE, FALSE

'*****************************************
'*** Code/Name Help Parameter
'*****************************************
Public D0COM_table       As String       'Table ID ��
Public D0COM_code_col    As String       '�ڵ� Column ��
Public D0COM_name_col    As String       '�ڵ��Ī Column ��
Public D0COM_name_co2    As String       '�ڵ��Ī Column ��
Public D0COM_cd_gbn_col As String        '�ڵ屸�� Column ��
Public D0COM_code        As String       '�ڵ� ��
Public D0COM_name        As String       '�ڵ�� ��
Public D0COM_name1       As String       '�ڵ�� ��
Public D0COM_name2       As String       '�ڵ�� ��
Public D0COM_name3       As String       '�ڵ�� ��
Public D0COM_cd_gbn      As String       '�ڵ屸�� ��
Public D0COM_length      As Integer      'D0COM_NAME_VAL�� ���� Label�� ����
Public D0COM_length2     As Integer      'D0COM_NAME_VAL�� ���� Label�� ����
Public D0COM_ret         As Integer      'Return ��

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

