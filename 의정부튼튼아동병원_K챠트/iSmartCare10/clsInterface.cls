VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsinterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mBuffers    As clsBuffers         '버퍼 컬렉션 클래스

Private mPhase      As Long             '통신상태(전체)
Private mSndPhase   As Long             '통신상태(송신)
Private mBufCnt     As Long             '버퍼순서
Private mState      As String           '현재 통신상태
Private mFrameNo    As Long             'ASTM 방식의 FrameNo
Private mIsETB      As Boolean          'ASTM 방식의 ETB 수신여부
Private mACKFg      As Long             'ACK,ETX 전송여부(0:ETX전송후,1:ACK전송후) -- CX 검사장비에서 사용
Private mRcvPhase   As Long             '통신상태(수신)

Public Property Get Buffers() As clsBuffers
    Set Buffers = mBuffers
End Property

Public Property Get Phase() As Long
    Phase = mPhase
End Property

Public Property Let Phase(ByVal vData As Long)
    mPhase = vData
End Property

Public Property Get SndPhase() As Long
    SndPhase = mSndPhase
End Property

Public Property Let SndPhase(ByVal vData As Long)
    mSndPhase = vData
End Property

Public Property Get BufCnt() As Long
    BufCnt = mBufCnt
End Property

Public Property Let BufCnt(ByVal vData As Long)
    mBufCnt = vData
End Property

Public Property Get State() As String
    State = mState
End Property

Public Property Let State(ByVal vData As String)
    mState = vData
End Property

Public Property Let FrameNo(ByVal vData As Long)
    mFrameNo = vData
End Property

Public Property Get IsETB() As Boolean
    IsETB = mIsETB
End Property

Public Property Let IsETB(ByVal vData As Boolean)
    mIsETB = vData
End Property

Public Property Get ACKFg() As Long
    ACKFg = mACKFg
End Property

Public Property Let ACKFg(ByVal vData As Long)
    mACKFg = vData
End Property

Public Property Get RcvPhase() As Long
    RcvPhase = mRcvPhase
End Property

Public Property Let RcvPhase(ByVal vData As Long)
    mRcvPhase = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 장비에서 송신한 데이터를 버퍼클래스에 저장
'   인수 :
'       - pBuffChar : 장비에서 송신한 데이터
'-----------------------------------------------------------------------------'
Public Sub AddBuffer(ByVal pBuffChar As String)

    Set mBuffers = New clsBuffers

    If mBuffers.Exist(mBufCnt) Then
        mBuffers(mBufCnt).Buffers = mBuffers(mBufCnt).Buffers & pBuffChar
    Else
        Call mBuffers.Add(mBufCnt, pBuffChar)
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 버퍼클래스 Clear
'-----------------------------------------------------------------------------'
Public Sub ClearBuffer()
    mBuffers.RemoveAll
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더전송시 사용되는 FrameNo를 조회
'-----------------------------------------------------------------------------'
Public Function GetFrameNo() As Long
    mFrameNo = mFrameNo + 1
    If mFrameNo = 8 Then
        mFrameNo = 0
    End If
    GetFrameNo = mFrameNo
End Function

Private Sub Class_Initialize()
    mPhase = 1
    mBufCnt = 1
    
    Set mBuffers = New clsBuffers
End Sub

Private Sub Class_Terminate()
    Set mBuffers = Nothing
End Sub


