VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSParameters"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================
'      ���� �������� ����
'================================
Public HOSPITAL_MAIN        As String               '������
Public HOSPITAL_NAME        As String               '���׿� ��Ī
Public HOSPITAL_TEL1        As String               '��ȭ��ȣ1
Public HOSPITAL_TEL2        As String               '��ȭ��ȣ2
Public HOSPITAL_ADDR        As String               '�ּ�
Public HOSPITAL_ADDR2       As String               '�ּ�2
Public HOSPITAL_FAX         As String               '�ѽ�
Public BLOOD_LABEL_CNT      As String
Public BBS_PTID_LENGTH      As String               'ȯ��ID ����(�����ھ�:8)
Public BLOOD_DEPTCD         As String               '�������� Deptcd

'================================
'      ������ ���α׷� ��뿩��
'================================
Public BLOOD_DEL_USED       As Boolean              '�����ǥ ��뿩��
Public DonorUserFg          As Boolean              '�������� ��뿩��
Public TRANS_REQUIRE_USED   As Boolean              '������û ��뿩��
Public TransReactionUsed    As Boolean              '�������ۿ��Ͽ���
Public ABO_DoubleChk        As Boolean              '���������� üũ����
Public BloodSplitUsed       As Boolean              '���׺�ȹ��뿩��
Public BloodTransfer        As Boolean              '�����̵���뿩��
Public BloodLocalDelivery   As Boolean              '����Local �����뿩��
Public DonationPaper        As Boolean              '����������
Public AttributeBBS         As Boolean              'ó��Ӽ��� ���(BBS���̺�)
Public TEST_FOR_PHERSIS     As Boolean              '�������� Deptcd

Public BLOOD_STORE_BARCODE_USED As Boolean               '�����԰��� ���׿��ڵ� ���(���ڵ�θ� �Է�)
Public USE_DONOR_INFORM As Boolean '���� ��ı�� ���(�����ڵ��, ���׵��, ��������ȸ, ������ұ�ɸ� ����)

Private Sub Class_Initialize()
    
    If Dir(INIPath) = "" Then
        MsgBox INIPath & " ���������� Ȯ���ϼ���.", vbInformation + vbOKOnly, "Info"
        Exit Sub
    End If
    '�������
    AttributeBBS = GetConstVariable("AttributeBBS")             ' True
    DonorUserFg = GetConstVariable("DonorUserFg")               ' True
    TEST_FOR_PHERSIS = GetConstVariable("TEST_FOR_PHERSIS")     ' True
    BLOOD_DEL_USED = GetConstVariable("BLOOD_DEL_USED")         ' False
    TRANS_REQUIRE_USED = GetConstVariable("TRANS_REQUIRE_USED") ' False
    TransReactionUsed = GetConstVariable("TransReactionUsed")   ' False
    ABO_DoubleChk = GetConstVariable("ABO_DoubleChk")           ' False
    BloodSplitUsed = GetConstVariable("BloodSplitUsed")         ' False
    BloodTransfer = GetConstVariable("BloodTransfer")           ' False
    BloodLocalDelivery = GetConstVariable("BloodLocalDelivery") ' False
    DonationPaper = GetConstVariable("DonationPaper")           ' False
    BLOOD_STORE_BARCODE_USED = GetConstVariable("BLOOD_STORE_BARCODE_USED")
    USE_DONOR_INFORM = GetConstVariable("USE_DONOR_INFORM")
    'String
    HOSPITAL_MAIN = GetConstVariable("HOSPITAL_MAIN")           ' "�������б�����"
    HOSPITAL_NAME = GetConstVariable("HOSPITAL_NAME")           ' "�������б����� ��������"
    HOSPITAL_TEL1 = GetConstVariable("HOSPITAL_TEL1")           ' "            "
    HOSPITAL_TEL2 = GetConstVariable("HOSPITAL_TEL2")           ' "            "
    HOSPITAL_ADDR = GetConstVariable("HOSPITAL_ADDR")           ' "���ֽ� �������б� ���� ��������"
    HOSPITAL_ADDR2 = GetConstVariable("HOSPITAL_ADDR2")         ' "���ֽ� �������б� ���� ��������"
    HOSPITAL_FAX = GetConstVariable("HOSPITAL_FAX")             ' "             "
    BLOOD_DEPTCD = GetConstVariable("BLOOD_DEPTCD")             ' "CP"
    BBS_PTID_LENGTH = GetConstVariable("BBS_PTID_LENGTH")       ' "8"
    'Long
    BLOOD_LABEL_CNT = GetConstVariable("BLOOD_LABEL_CNT")       ' 1
End Sub

Private Function GetConstVariable(ByVal qKeyString As String) As Variant
    Dim strTmp  As String
    
    If Dir(INIPath) = "" Then Exit Function
    strTmp = medGetINI("BBS_CONST", qKeyString, INIPath)
    
    Select Case medGetP(strTmp, 2, LINE_DIV)
        Case "0": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
        Case "1": GetConstVariable = CLng(medGetP(strTmp, 1, LINE_DIV))
        Case "2": GetConstVariable = medGetP(strTmp, 1, LINE_DIV)
    End Select
    
End Function


