VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTLab005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TestCd       As String           ' �˻��׸� �ڵ�
Public SpcCd        As String           ' ��ü �ڵ�
Public ApplyDt      As String           ' ������
Public ApplySex     As String           ' ���뼺��
Public AgeFrom      As Long             ' ���볪��(FROM)
Public AgeTo        As Long             ' ���볪��(To)
Public RefValFrom   As Double           ' ����ġ(FROM)
Public RefValTo     As Double           ' ����ġ(To)
Public RefCd        As String           ' ����ġ(Alpha)
Public RefText      As String           ' ��������ġ(Alpha)
Public ARefValFrom  As Double           ' AutoVerify ����ġ(FROM)
Public ARefValTo    As Double           ' AutoVerify ����ġ(To)
Public ExpDt        As String           ' �����
Public PanicFrVal   As Double           ' Panic Range(FROM)
Public PanicToVal   As Double           ' Panic Range(To)
Public ArletFrVal   As Double
Public ArletToVal   As Double
Public RefDiv       As String
Public AMRFrVal     As Double           ' AMR Range(FROM)
Public AMRToVal     As Double           ' AMR Range(To)