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
'   ���ϸ�  : clsIISBuffer.cls
'   �ۼ���  : �̻��
'   ��  ��  : ��񿡼� ���۵� �����͸� �����ϴ� ����Ŭ����
'   �ۼ���  : 2004-04-28
'   ��  ��  :
'-----------------------------------------------------------------------------'

'## �Ӽ�
'   - ���ۼ���
'   - ���۳���

'## �޼ҵ�

Option Explicit

Private mSeq     As Long        '���ۼ���
Private mBuffers As String      '���۳���

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
