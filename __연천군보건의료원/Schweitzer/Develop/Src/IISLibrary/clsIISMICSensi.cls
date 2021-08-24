VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMICSensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISMICSensi.cls

'   내  용  : 미생물 감수성 결과내역 클래스

'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mSQL    As clsIISMicroSql

Private mMfySeq As Long     '수정회수
Private mSeq    As Long     'Colony Sequence
Private mMnmCd  As String   '균코드
Private mMicFg  As String   'MIC구분(S:일반, C:MIC)
Private mMqtCd  As String   '정도코드(?)
Private mSCnt   As Long     '항생제개수
Private mResult As String   '항생제결과

Public Property Get MfySeq() As Long
    MfySeq = mMfySeq
End Property

Public Property Let MfySeq(ByVal vData As Long)
    mMfySeq = vData
End Property

Public Property Get Seq() As Long
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As Long)
    mSeq = vData
End Property

Public Property Get MnmCd() As String
    MnmCd = mMnmCd
End Property

Public Property Let MnmCd(ByVal vData As String)
    mMnmCd = vData
End Property

Public Property Get MicFg() As String
    MicFg = mMicFg
End Property

Public Property Let MicFg(ByVal vData As String)
    mMicFg = vData
End Property

Public Property Get MqtCd() As String
    MqtCd = mMqtCd
End Property

Public Property Let MqtCd(ByVal vData As String)
    mMqtCd = vData
End Property

Public Property Get SCnt() As Long
    SCnt = mSCnt
End Property

Public Property Let SCnt(ByVal vData As Long)
    mSCnt = vData
End Property

Public Property Get Result() As String
    Result = mResult
End Property

Public Property Let Result(ByVal vData As String)
    mResult = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 균명(약어)를 이용해 균코드를 조회
'   인수 :
'       - pMnmNm : Vitek장비에서 수신한 균명(약어)
'   반환 : 균코드, 없으면 Null
'-----------------------------------------------------------------------------'
Public Function GetMnmCd(ByVal pMnmNm As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectMnmCd(pMnmNm), , adCmdText)
    If Rs.BOF Or Rs.EOF Then
        GetMnmCd = ""
    Else
        GetMnmCd = Rs.Fields("MNMCD").Value & ""
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    MsgBox Err.Description, vbCritical, "오류"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 미생물 감수성결과 입력쿼리 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'       - pTestCd   : 검사코드
'-----------------------------------------------------------------------------'
Public Function GetSaveMICResultSql(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long, ByVal pTestCd As String) As String
    Dim lngSeq As Long
    
    lngSeq = GetMaxSeq(pWorkarea, pAccDt, pAccSeq)
    
    GetSaveMICResultSql = mSQL.InsertMICResult(pWorkarea, pAccDt, pAccSeq, pTestCd, mMfySeq, _
                            lngSeq, mMnmCd, mMicFg, mMqtCd, mSCnt, mResult)
End Function

'-----------------------------------------------------------------------------'
'   기능 : IIS505에서 MAX(Seq)+1를 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'-----------------------------------------------------------------------------'
Public Function GetMaxSeq(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As Long
    Dim Rs      As ADODB.Recordset
    Dim lngSeq  As Long

On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectMaxSeq(pWorkarea, pAccDt, pAccSeq), , adCmdText)
    If Rs.BOF Or Rs.EOF Or IsNull(Rs.Fields("MAXSEQ").Value) Then
        GetMaxSeq = 1
    Else
        GetMaxSeq = CLng(Rs.Fields("MAXSEQ").Value)
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsIISMICSensi", "GetMaxSeq", Err.Description, Now
    MsgBox Error.Description, vbCritical, "오류"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMicroSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub
