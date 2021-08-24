VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim iphase As Integer               ' 통신상태
Dim rcvbuf(1 To 50) As String       ' 결과받는 버퍼
Dim ibufcnt As Integer              ' 결과받는 버퍼 Index
Dim rcvbufs As String               ' 버퍼 통합
Dim iSnd_Phase As Integer           ' 송신 통신상태
Dim iState As String                ' 통신 모드
Dim iIsETB As String                ' ETB
Dim iFrameN As Integer              ' ASTM Protocol에서 송신시 Frame Number
Dim iBlockCnt As Integer            ' Block Count
Dim iack_etx As Integer             ' ASTM중 ACK/ETX를 확인자로 사용하는 IF

Private Sub Class_Initialize()
    iphase = 1
    ibufcnt = 1
End Sub

Public Property Get phase() As Variant
    phase = iphase
End Property

Public Property Let phase(ByVal vNewValue As Variant)
    iphase = vNewValue
End Property

Public Property Get bufcnt() As Variant
    bufcnt = ibufcnt
End Property

Public Property Let bufcnt(ByVal vNewValue As Variant)
    ibufcnt = vNewValue
End Property

Public Sub clearRcvbuf()
    Dim i%
    For i = 1 To 50
        rcvbuf(i) = ""
    Next
End Sub

Public Sub addRcvbuf(sBufChar As String)
    If ibufcnt > 0 And ibufcnt <= 50 Then
        rcvbuf(ibufcnt) = rcvbuf(ibufcnt) & sBufChar
    End If
End Sub

Public Sub addBufs()
    Dim i%
    rcvbufs = ""
    
    For i = 1 To ibufcnt
        rcvbufs = rcvbufs & rcvbuf(i)
    Next
End Sub

Public Function getrcvbufs() As String
    getrcvbufs = rcvbufs
End Function

Public Function getrcvbuf(ByVal bufIndex As Variant) As String
    getrcvbuf = rcvbuf(bufIndex)
End Function

Public Property Get Snd_Phase() As Variant
    Snd_Phase = iSnd_Phase
End Property

Public Property Let Snd_Phase(ByVal vNewValue As Variant)
    iSnd_Phase = vNewValue
End Property

Public Property Get state() As Variant
    state = iState
End Property

Public Property Let state(ByVal vNewValue As Variant)
    iState = vNewValue
End Property

Public Property Get IsETB() As Variant
    IsETB = iIsETB
End Property

Public Property Let IsETB(ByVal vNewValue As Variant)
    iIsETB = vNewValue
End Property

Public Property Get FrameN() As Variant
    FrameN = iFrameN
End Property

Public Property Let FrameN(ByVal vNewValue As Variant)
    iFrameN = vNewValue
End Property

Public Sub addFrameN()
    iFrameN = iFrameN + 1
    If iFrameN = 8 Then
        iFrameN = 0
    End If
End Sub

Public Property Get BlockCnt() As Variant
    BlockCnt = iBlockCnt
End Property

Public Property Let BlockCnt(ByVal vNewValue As Variant)
    iBlockCnt = vNewValue
End Property

Public Sub rmRcvbufChkSum()
    rcvbuf(ibufcnt) = left$(rcvbuf(ibufcnt), Len(rcvbuf(ibufcnt)) - 4)
End Sub

Public Property Get ack_etx() As Variant
    ack_etx = iack_etx
End Property

Public Property Let ack_etx(ByVal vNewValue As Variant)
    iack_etx = vNewValue
End Property


