Attribute VB_Name = "modIISDb"
Option Explicit

'## 레지스트리의 DB설정 정보
Public Const cDBSERVER  As String = "DbServer"
Public Const cDBTYPE    As String = "DbType"
Public Const cSOURCE    As String = "Source"
Public Const cCATALOG   As String = "Catalog"
Public Const cUID       As String = "Uid"
Public Const cPWD       As String = "Pwd"

'## 프로그램 경로정보
Public mEXEPATH      As String       'EXE 파일경로
Public mLOGPATH      As String       'Log 파일경로
Public mCLIENTPATH   As String       'ClientDb 경로
Public mINIPATH      As String       'INI 파일경로

