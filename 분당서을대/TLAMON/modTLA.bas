Attribute VB_Name = "modTLA"
'-----------------------------------------------------------------------------'
'   파일명  : modTLA.bas
'   작성자  : 오세원
'   내  용  : TLA 검사진행 상태 모니터링
'   작성일  :
'   버  전  : 1.0.0
'   병  원  : 분당서울대병원 진단검사의학과 자동화검사실
'-----------------------------------------------------------------------------'
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
                "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                              ByVal lpKeyName As Any, _
                                              ByVal lpString As Any, _
                                              ByVal lplFileName As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpDefault As String, _
                                            ByVal lpReturnedString As String, _
                                            ByVal nSize As Long, _
                                            ByVal lpFileName As String) As Long

Public Cn_Ser   As ADODB.Connection
Public RS_Ser   As ADODB.Recordset
Public SQL      As String

Public gResresh As Long
Public gGRPCD   As String
Public gEQPCD   As String
Public gWIDTH   As Integer
Public gIP      As String
Public gDB      As String
Public gUID     As String
Public gPWD     As String
Public gLimit   As Long
Public gLimitS  As Long
Public gTatARC  As Long
Public gTatAU   As Long
Public gTatCOB  As Long
Public gTatC16  As Long

Public Const colHeader = 8

Private Type BackGroundColor
    LV1     As String
    LV2     As String
    LV3     As String
    LV4     As String
    LV5     As String
    LV6     As String
    LV7     As String
    LV8     As String
    LV9     As String
    LV10    As String
    LV11    As String
    LV12    As String
    'LV13    As String
    'LV14    As String
    'LV15    As String
    'LV16    As String
    'LV17    As String
    
    LVO     As String
    LVOO    As String
    LVR     As String
    LVRR    As String
    LVH     As String
End Type

Public BGColor As BackGroundColor

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


'-----------------------------------------------------------------------------'
'   기능 : 로그 내용을 파일로 저장
'   인수 :
'       1.strName    : 로그명
'       2.argSQL     : 로그내용
'-----------------------------------------------------------------------------'
Public Sub SetSQLData(ByVal strName As String, ByVal argSQL As String)
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
        
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = strName
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Output As FilNum
    Print #FilNum, argSQL
    Close FilNum
    
End Sub
