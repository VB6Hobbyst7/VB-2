VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISError.cls
'   작성자  : 오세원
'   내  용  : Error 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mFileNm         As String       '오류가 발생한 파일명
Private mClassNm        As String       '오류가 발생한 클래스명
Private mProcedureNm    As String       '오류가 발생한 프로시져명
Private mDescription    As String       '오류에 대한 설명
Private mErrorTime      As Date         '오류가 발생한 시간

Public Property Get FileNm() As String
    FileNm = mFileNm
End Property

Public Property Let FileNm(ByVal vData As String)
    mFileNm = vData
End Property

Public Property Get ClassNm() As String
    ClassNm = mClassNm
End Property

Public Property Let ClassNm(ByVal vData As String)
    mClassNm = vData
End Property

Public Property Get ProcedureNm() As String
    ProcedureNm = mProcedureNm
End Property

Public Property Let ProcedureNm(ByVal vData As String)
    mProcedureNm = vData
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal vData As String)
    mDescription = vData
End Property

Public Property Get ErrorTime() As Date
    ErrorTime = mErrorTime
End Property

Public Property Let ErrorTime(ByVal vData As Date)
    mErrorTime = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 오류정보를 로그파일에 기록
'   인수 :
'       - pFileNm       : 오류가 발생한 파일명
'       - pClassNm      : 오류가 발생한 클래스명
'       - pProcedureNm  : 오류가 발생한 프로시져명
'       - pDescription  : 오류에 대한 설명
'       - pErrorTime    : 오류가 발생한 시간
'-----------------------------------------------------------------------------'
Public Sub SetLog(ByVal pFileNm As String, ByVal pClassNm As String, ByVal pProcedureNm As String, _
                    ByVal pDescription As String, ByVal pErrorTime As Date)
    Dim strFileNm       As String   '파일명
    Dim intFileNo       As Integer  '파일번호
    
    mFileNm = pFileNm:              mClassNm = pClassNm
    mProcedureNm = pProcedureNm:    mDescription = pDescription
    mErrorTime = pErrorTime
    
On Error Resume Next
    '## Log 디렉토리 없으면 생성
    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir App.Path & "\Log"
    End If
    
    '## 파일이 존재하면 Append, 없으면 Create
    intFileNo = FreeFile
    strFileNm = App.Path & "\Log\Error(" & Format(mErrorTime, "YYYY-MM-DD") & ").log"
    If Dir(strFileNm) = "" Then
        Open strFileNm For Output As #intFileNo
    Else
        Open strFileNm For Append As #intFileNo
    End If
    
    '## 오류내용 기록
    Print #intFileNo, "[" & CStr(mErrorTime) & "]"
    Print #intFileNo, "Project Name   : " & mFileNm
    Print #intFileNo, "Class Name     : " & mClassNm
    Print #intFileNo, "Procedure Name : " & mProcedureNm
    Print #intFileNo, "Description    : " & mDescription & vbCrLf
    
    Close intFileNo
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 멤버변수 초기화
'-----------------------------------------------------------------------------'
Private Sub Clear()
    mFileNm = "":       mClassNm = ""
    mProcedureNm = "":  mDescription = ""
End Sub
