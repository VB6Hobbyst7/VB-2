VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISBuffer.cls
'   작성자  : 오세원
'   내  용  : 장비에서 전송된 데이터를 저장하는 버퍼클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

'## 속성
'   - 버퍼순서
'   - 버퍼내용

'## 메소드

Option Explicit

Private mSeq     As Long        '버퍼순서
Private mBuffers As String      '버퍼내용

Public Property Get Seq() As Long
    Seq = mSeq
End Property

Public Property Let Seq(ByVal vData As Long)
    mSeq = vData
End Property

Public Property Get Buffers() As String
    Buffers = mBuffers
End Property

Public Property Let Buffers(ByVal vData As String)
    mBuffers = vData
End Property

