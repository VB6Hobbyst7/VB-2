VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTLab004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TestCd       As String               ' �˻��׸� �ڵ�
Public SpcCd        As String               ' ��ü �ڵ�
Public Seq          As Integer              ' Seq
Public ApplyDt      As String               ' ������
Public SpcGrpCd     As String               ' ��ü�� �ڵ�
Public LabelCnt     As Integer              ' ���ڵ�� ������
Public RstUnit      As String               ' ��� ����
Public RndFg        As String               ' Round ä������('0':��,'1':��)
Public StatFlags    As String               ' �� �ǹ��� Stat ����
Public StatFg       As String               ' Stat ����('0':��,'1':��)
Public AvalVal      As Integer              ' ��ȿ ����
Public PanicFg      As String               ' Panic Check����('0':��,'1':��)
Public PanicFrVal   As Double               ' Panic Range(FROM)
Public PanicToVal   As Double               ' Panic Range(To)
Public DeltaFg      As String               ' Delta Check����('0':��,'1':��)
Public DeltaVal1    As Integer              ' Delta Value
Public DeltaVal2    As Integer              ' Delta Value
Public TestCost     As String               ' �����ڵ�
Public StoreCd      As String               ' ��������
Public TatAvg       As Integer              ' TAT�ҿ�ð�
Public SpcQty       As Double               ' ��ü ü�뷮(TLA)
Public SpcUnit      As String               ' ��ü ä�� ����
Public ExpDt        As String               ' �����
Public TATS         As String               ' �˻�ҿ�ð�
Public ArletFg      As String               ' Arlet Check����('0':��,'1':��)
Public ArletFrVal   As Double               ' Arlet Range(FROM)
Public ArletToVal   As Double               ' Arlet Range(To)
