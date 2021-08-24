VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISEqpChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISEqpChoice.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : ���� PC���� ����� �˻���� ���� Ŭ����
'   �ۼ���  : 2004-03-08
'   ��  ��  :
'       1. 1.1.12: �̻��(2005-07-28)
'   ��  ��  : NOTE Ȯ���Ұ�!
'-----------------------------------------------------------------------------'

Option Explicit

'## EqpType Enum
Public Enum EqpCommTypeEnum
    ccBidirection = 0   '�����
    ccUnidirection = 1  '�ܹ���
End Enum

'## Barcode Enum
Public Enum BarcodeUseEnum
    ccNoBarcode = 0     '���ڵ� ������
    ccUseBarcode = 1    '���ڵ� ���
End Enum

'## AutoVfy Enum
Public Enum AutoVfyEnum
    ccNo = 0            '������
    ccYes = 1           '���
End Enum

'## HLCheck Enum
Public Enum HLCheckEnum
    ccNoCheck = 0       'H/L�� ������� �ڵ�������
    ccCheck = 1         'H/L�� �ش�Ǹ� �ڵ������� �Ұ���
End Enum

'## BarPos Enum
Public Enum BarPosEnum
    ccPC = 0            'PC
    ccEqp = 1           '���
End Enum

'## AutoRerun Enum
Public Enum AutoRerunEnum
    ccNoAutoRerun = 0
    ccAutoRerun = 1
End Enum

Private mSQL            As clsIISMasterSql

Private mEqpCd1         As String           '����ڵ�1
Private mEqpNm1         As String           '����1
Private mEqpKey1        As String           '��� ������ Key1
Private mAutoVfy1       As AutoVfyEnum      '�ڵ������� ��뿩��1
Private mBarcode1       As BarcodeUseEnum   '���1�� ���ڵ� �������
Private mHLCheck1       As HLCheckEnum      '���1�� �ڵ������Ͻ� H/L�� üũ����
Private mBarPos1        As BarPosEnum       '���1�� ���ڵ� ������ ��ġ
Private mEqpCd2         As String           '����ڵ�2
Private mEqpNm2         As String           '����2
Private mEqpKey2        As String           '��� ������ Key2
Private mAutoVfy2       As AutoVfyEnum      '�ڵ������� ��뿩��2
Private mBarcode2       As BarcodeUseEnum   '���2�� ���ڵ� �������
Private mHLCheck2       As HLCheckEnum      '���2�� �ڵ������Ͻ� H/L�� üũ����
Private mBarPos2        As BarPosEnum       '���2�� ���ڵ� ������ ��ġ

'## 1.1.12: �̻��(2005-07-28)
'   - �ڵ���� ���� ���� �߰�
Private mAutoRerun1     As AutoRerunEnum    '���1�� �ڵ���� ����
Private mAutoRerun2     As AutoRerunEnum    '���2�� �ڵ���� ����

Public Property Get EqpCd1() As String
    EqpCd1 = mEqpCd1
End Property

Public Property Let EqpCd1(ByVal vData As String)
    mEqpCd1 = vData
End Property

Public Property Get EqpNm1() As String
    EqpNm1 = mEqpNm1
End Property

Public Property Let EqpNm1(ByVal vData As String)
    mEqpNm1 = vData
End Property

Public Property Get EqpKey1() As String
    EqpKey1 = mEqpKey1
End Property

Public Property Let EqpKey1(ByVal vData As String)
    mEqpKey1 = vData
End Property

Public Property Get AutoVfy1() As String
    AutoVfy1 = mAutoVfy1
End Property

Public Property Let AutoVfy1(ByVal vData As String)
    mAutoVfy1 = vData
End Property

Public Property Get Barcode1() As BarcodeUseEnum
    Barcode1 = mBarcode1
End Property

Public Property Let Barcode1(ByVal vData As BarcodeUseEnum)
    mBarcode1 = vData
End Property

Public Property Get HLCheck1() As HLCheckEnum
    HLCheck1 = mHLCheck1
End Property

Public Property Let HLCheck1(ByVal vData As HLCheckEnum)
    mHLCheck1 = vData
End Property

Public Property Get BarPos1() As BarPosEnum
    BarPos1 = mBarPos1
End Property

Public Property Let BarPos1(ByVal vData As BarPosEnum)
    mBarPos1 = vData
End Property

Public Property Get EqpCd2() As String
    EqpCd2 = mEqpCd2
End Property

Public Property Let EqpCd2(ByVal vData As String)
    mEqpCd2 = vData
End Property

Public Property Get EqpNm2() As String
    EqpNm2 = mEqpNm2
End Property

Public Property Let EqpNm2(ByVal vData As String)
    mEqpNm2 = vData
End Property

Public Property Get EqpKey2() As String
    EqpKey2 = mEqpKey2
End Property

Public Property Let EqpKey2(ByVal vData As String)
    mEqpKey2 = vData
End Property

Public Property Get AutoVfy2() As String
    AutoVfy2 = mAutoVfy2
End Property

Public Property Let AutoVfy2(ByVal vData As String)
    mAutoVfy2 = vData
End Property

Public Property Get Barcode2() As BarcodeUseEnum
    Barcode2 = mBarcode2
End Property

Public Property Let Barcode2(ByVal vData As BarcodeUseEnum)
    mBarcode2 = vData
End Property

Public Property Get HLCheck2() As HLCheckEnum
    HLCheck2 = mHLCheck2
End Property

Public Property Let HLCheck2(ByVal vData As HLCheckEnum)
    mHLCheck2 = vData
End Property

Public Property Get BarPos2() As BarPosEnum
    BarPos2 = mBarPos2
End Property

Public Property Let BarPos2(ByVal vData As BarPosEnum)
    mBarPos2 = vData
End Property

'## 1.1.12: �̻��(2005-07-28)
'   - �ڵ���� ���� ������Ƽ �߰�
Public Property Get AutoRerun1() As AutoRerunEnum
    AutoRerun1 = mAutoRerun1
End Property

Public Property Let AutoRerun1(ByVal vData As AutoRerunEnum)
    mAutoRerun1 = vData
End Property

Public Property Get AutoRerun2() As AutoRerunEnum
    AutoRerun2 = mAutoRerun2
End Property

Public Property Let AutoRerun2(ByVal vData As AutoRerunEnum)
    mAutoRerun2 = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���� PC���� ����� ��� ����
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function SetEqp() As Boolean
    Dim strFileNm   As String   'ini���� ���+�̸�
    Dim strAppName  As String   '������ AppName
    
    '## 1.ini���Ͽ��� �ش� ����ڵ��� ������ Key�� ����
    '## 2.������ ����
On Error GoTo Errors
    strFileNm = IniPath & "\" & "IIS.ini"
    strAppName = AppName
    
    '## 1.1.12: �̻��(2005-07-28)
    '   - ��� 1, 2�� �ڵ���� �ɼ��߰�
    '## �˻����1
    If mEqpCd1 <> "" Then
        EqpKey1 = mReadINI(strFileNm, "EQPKEY", mEqpCd1)
        SaveSetting strAppName, "Machine1", "EqpCd", mEqpCd1
        SaveSetting strAppName, "Machine1", "EqpKey", mEqpKey1
        SaveSetting strAppName, "Machine1", "AutoVfy", mAutoVfy1
        SaveSetting strAppName, "Machine1", "Barcode", mBarcode1
        SaveSetting strAppName, "Machine1", "HLCheck", mHLCheck1
        SaveSetting strAppName, "Machine1", "BarPos", mBarPos1
        SaveSetting strAppName, "Machine1", "AutoRerun", mAutoRerun1
    Else
        SaveSetting strAppName, "Machine1", "EqpCd", ""
        SaveSetting strAppName, "Machine1", "EqpKey", ""
        SaveSetting strAppName, "Machine1", "AutoVfy", ""
        SaveSetting strAppName, "Machine1", "Barcode", ""
        SaveSetting strAppName, "Machine1", "HLCheck", ""
        SaveSetting strAppName, "Machine1", "BarPos", ""
        SaveSetting strAppName, "Machine1", "AutoRerun", ""
    End If
    
    '## �˻����2
    If mEqpCd2 <> "" Then
        EqpKey2 = mReadINI(strFileNm, "EQPKEY", mEqpCd2)
        SaveSetting strAppName, "Machine2", "EqpCd", mEqpCd2
        SaveSetting strAppName, "Machine2", "EqpKey", mEqpKey2
        SaveSetting strAppName, "Machine2", "AutoVfy", mAutoVfy2
        SaveSetting strAppName, "Machine2", "Barcode", mBarcode2
        SaveSetting strAppName, "Machine2", "HLCheck", mHLCheck2
        SaveSetting strAppName, "Machine2", "BarPos", mBarPos2
        SaveSetting strAppName, "Machine2", "AutoRerun", mAutoRerun2
    Else
        SaveSetting strAppName, "Machine2", "EqpCd", ""
        SaveSetting strAppName, "Machine2", "EqpKey", ""
        SaveSetting strAppName, "Machine2", "AutoVfy", ""
        SaveSetting strAppName, "Machine2", "Barcode", ""
        SaveSetting strAppName, "Machine2", "HLCheck", ""
        SaveSetting strAppName, "Machine2", "BarPos", ""
        SaveSetting strAppName, "Machine2", "AutoRerun", ""
    End If
    SetEqp = True
    Exit Function
    
Errors:
    Error.SetLog App.EXEName, "clsEqpChoice", "SetEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���� PC���� ����ϴ� ��� ��ȸ
'   ��ȯ : True(����), False(����)
'-----------------------------------------------------------------------------'
Public Function GetEqp() As Boolean
    Dim strAppName  As String   '������ AppName
    
On Error GoTo Errors
    strAppName = AppName
    
    '## NOTE: ���α׷������δ� AutoVfy, Barcode, HLCheck�� '0','1'�̿��� ����
    '   �Է��Ҽ� ������ ����ڰ� ������ ������Ʈ������ �����Ұ�쿡�� ��Ÿ�ӿ����� �߻��ϱ�
    '   ������ ���α׷��� �����ؾ� �Ұ� ����.
    '## 1.1.12: �̻��(2005-07-28)
    '   - ��� 1, 2�� �ڵ���� �ɼ��߰�
    '## �˻����1
    mEqpCd1 = GetSetting(strAppName, "Machine1", "EqpCd", "")
    If mEqpCd1 <> "" Then
        mEqpNm1 = GetEqpNm(mEqpCd1)
        mEqpKey1 = GetSetting(strAppName, "Machine1", "EqpKey", "")
        mAutoVfy1 = CLng(GetSetting(strAppName, "Machine1", "AutoVfy", "0"))
        mBarcode1 = CLng(GetSetting(strAppName, "Machine1", "Barcode", "0"))
        mHLCheck1 = CLng(GetSetting(strAppName, "Machine1", "HLCheck", "0"))
        mBarPos1 = CLng(GetSetting(strAppName, "Machine1", "BarPos", "0"))
        mAutoRerun1 = CLng(GetSetting(strAppName, "Machine1", "AutoRerun", "0"))
    End If
    
    '## �˻����2
    mEqpCd2 = GetSetting(strAppName, "Machine2", "EqpCd", "")
    If mEqpCd2 <> "" Then
        mEqpNm2 = GetEqpNm(mEqpCd2)
        mEqpKey2 = GetSetting(strAppName, "Machine2", "EqpKey", "")
        mAutoVfy2 = CLng(GetSetting(strAppName, "Machine2", "AutoVfy", "0"))
        mBarcode2 = CLng(GetSetting(strAppName, "Machine2", "Barcode", "0"))
        mHLCheck2 = CLng(GetSetting(strAppName, "Machine2", "HLCheck", "0"))
        mBarPos2 = CLng(GetSetting(strAppName, "Machine2", "BarPos", "0"))
        mAutoRerun2 = CLng(GetSetting(strAppName, "Machine2", "AutoRerun", "0"))
    End If
    GetEqp = True
    Exit Function
    
Errors:
    Error.SetLog App.EXEName, "clsEqpChoice", "GetEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : ������� ����ڵ�, ���� ��ȸ
'   ��ȯ : Recordset
'-----------------------------------------------------------------------------'
Public Function GetUsingEqp() As ADODB.Recordset
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectUsingEqp, , adCmdText)
    Set GetUsingEqp = Rs
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsEqpChoice", "GetUsingEqp", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ����ڵ��� ���� ��ȸ
'   �μ� :
'       1.pEqpCd : ����ڵ�
'   ��ȯ : �����ϴ� ����̸� ���� ��ȯ, �������� ������ Null��ȯ
'-----------------------------------------------------------------------------'
Public Function GetEqpNm(ByVal pEqpCd As String) As String
    Dim Rs As ADODB.Recordset
    
On Error GoTo Errors
    Set Rs = DbCon.Execute(mSQL.SelectEqpCd(pEqpCd), , adCmdText)
    
    If Not (Rs.BOF Or Rs.EOF) Then
        GetEqpNm = Rs.Fields("EQPNM").Value
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Error.SetLog App.EXEName, "clsEqpChoice", "GetEqpNm", Err.Description, Now
    MsgBox Error.Description, vbCritical, "����"
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISMasterSql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub