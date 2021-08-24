Attribute VB_Name = "modBioPlex2200"
Option Explicit

'## 통신기호
Public Const mENQ As Long = &H5         'Chr(5),  ""
Public Const mACK As Long = &H6         'Chr(6),  ""
Public Const mSTX As Long = &H2         'Chr(2),  ""
Public Const mETB As Long = &H17        'Chr(23), ""
Public Const mETX As Long = &H3         'Chr(3),  ""
Public Const mEOT As Long = &H4         'Chr(4),  ""
Public Const mNAK As Long = &H15        'Chr(21), ""
Public Const mSOH As Long = &H1         'Chr(1),  ""
Public Const mDLE As Long = &H10        'Chr(16), ""
Public Const mSYN As Long = &H16        'Chr(22), ""

'## BarPos Enum
Public Enum BarPosEnum
    ccPC = 0            'PC
    ccEqp = 1           '장비
End Enum

'## Log Enum
Public Enum LogEnum
    ccPCLog = 0         'PC  에서 송신한 Log
    ccEqpLog = 1        '장비에서 송신한 Log
End Enum


Private mFileNum    As Integer          '로그파일의 File Number


'-----------------------------------------------------------------------------'
'   기능 : Datalog를 파일에 저장
'   인수 :
'       - pData  : Datalog
'       - pLogFg : Datalog 구분자
'-----------------------------------------------------------------------------'
Public Sub WriteLog(ByVal pData As String, ByVal pLogFg As LogEnum)
    If pLogFg = ccPCLog Then
        Print #mFileNum, "[P C] " & pData
    Else
        Print #mFileNum, pData;
    End If
End Sub

