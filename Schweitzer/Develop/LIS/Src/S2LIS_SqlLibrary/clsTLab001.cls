VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTLab001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public TestCd As String         ' �˻��׸� �ڵ�
Public ApplyDt As String        ' ������
Public TestNm As String         ' �˻��
Public AbbrNm5 As String        ' ����(5�ڸ�)
Public AbbrNm10 As String       ' ����(10�ڸ�)
Public WorkArea As String       ' Work Area
Public RstType As String        ' �������
Public TestDiv As String        ' ��Ÿ�˻�,�̻����˻� ����
Public RptSeq As Integer        ' Report Seq
Public PanelFg As String        ' Pannel(Group,Detail) ó�濩��('0':��,'1':��)
Public DetailFg As String       ' ���׸񿩺�('*')
Public RstDiv As String         ' Alternative /Required ����
Public OutLabCd As String       ' �ܺΰ˻� ���
Public GrpFg As String          ' Graph �������('0':��,'1':��)
Public TxtType As String        ' �ؽ�Ʈ�������('0':��,'1':��)
Public WorkUnit As Integer      ' Work Load Unit
Public AttrCd As String         ' �Ӽ� ���� �ڵ�
Public MatchCd As String
Public ExpDt As String          ' �����
Public ItemSeq As Integer       ' �����¼���
Public GroupCd As String        ' ���Group
