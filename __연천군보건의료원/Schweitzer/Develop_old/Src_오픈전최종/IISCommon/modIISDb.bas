Attribute VB_Name = "modIISDb"
Option Explicit

'## ������Ʈ���� DB���� ����
Public Const cDBSERVER  As String = "DbServer"
Public Const cDBTYPE    As String = "DbType"
Public Const cSOURCE    As String = "Source"
Public Const cCATALOG   As String = "Catalog"
Public Const cUID       As String = "Uid"
Public Const cPWD       As String = "Pwd"

'## ���α׷� �������
Public mEXEPATH      As String       'EXE ���ϰ��
Public mLOGPATH      As String       'Log ���ϰ��
Public mCLIENTPATH   As String       'ClientDb ���
Public mINIPATH      As String       'INI ���ϰ��

