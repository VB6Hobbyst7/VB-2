VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntOrder.cls
'   �ۼ���  : �̻��
'   ��  ��  : AU1000 ��� �������� Ŭ����
'   �ۼ���  : 2004-06-17
'   ��  ��  :
'       1. 1.0.1: �̻��(2004-12-29)
'          - LDL-Calculated�� ������ �˻���� 23->28�� �����!
'-----------------------------------------------------------------------------'

Option Explicit

Private mRackNo     As String   'Rack No
Private mCupPos     As String   'Cup Position
Private mBarNo      As String   '���ڵ��ȣ
Private mSampleNo   As String   'Sample No

Public Property Get RackNo() As String
    RackNo = mRackNo
End Property

Public Property Let RackNo(ByVal vData As String)
    mRackNo = vData
End Property

Public Property Get CupPos() As String
    CupPos = mCupPos
End Property

Public Property Let CupPos(ByVal vData As String)
    mCupPos = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SampleNo() As String
    SampleNo = mSampleNo
End Property

Public Property Let SampleNo(ByVal vData As String)
    mSampleNo = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim objIntNm    As clsIISIntNm      '��� �˻��׸� Ŭ����
    Dim strOrder    As String           '�������� ���ڿ�
    
    '## 1.0.1: �̻��(2004-12-29)
    '   - LDL-Calculated���� 23->28�� ����Ǿ� ���α׷� ����
    '## TP�˻��� ��ü�� Urine�̸� �˻縦 �ϸ� �ȵ�
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            Set objIntNm = objResult.IntNm
            
            '## LDL-C(28)�׸��� ����׸��̶� ������ ������ �ȵ�(CHOL, T.G, HDL-C)
            '## T.P(01)�� ��ü�� Urine�϶� �˻縦 �ϸ� �ȵ�
            If objIntNm.IntBase <> "28" And _
                Not (Mid$(pAccInfo.SpcNm, 1, IISSPCLEN) = IISSPCURINE And objIntNm.IntBase = "01") Then
                strOrder = strOrder & objIntNm.IntBase
            End If
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            
            '## LDL-C(28)�׸��� ����׸��̶� ������ ������ �ȵ�(CHOL, T.G, HDL-C)
            '## T.P(01)�� ��ü�� Urine�϶� �˻縦 �ϸ� �ȵ�
            If objIntNm.IntBase <> "28" And _
                Not (Mid$(pAccInfo.SpcNm, 1, IISSPCLEN) = IISSPCURINE And objIntNm.IntBase = "01") Then
                strOrder = strOrder & objIntNm.IntBase
            End If
        Next
        Set objResult = Nothing
    End If
    Set objIntNm = Nothing
    
    GetOrder = STX & "S " & mRackNo & mCupPos & mBarNo & mSampleNo & strOrder & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ��������� ���ڿ� ��ȸ
'   ��ȯ : ���������� ��������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    GetNoOrder = STX & "S " & mRackNo & mCupPos & Format$(mBarNo, String$(SPCLEN, "@")) & mSampleNo & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mRackNo = ""
    mCupPos = ""
    mBarNo = ""
    mSampleNo = ""
End Sub