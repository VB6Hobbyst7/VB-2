VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISIntInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISIntInfo.cls
'   �ۼ���  : �̻��
'   ��  ��  : ��񿡼� �����Ѱ�� ����Ŭ����
'   �ۼ���  : 2004-05-26
'   ��  ��  :
'       1. 1.1.3: �̻��(2005-01-03)
'-----------------------------------------------------------------------------'

Option Explicit

'## �Ϲ���񿡼� ���
Private mIntResults As clsIISIntResults '�������̽� ��� �÷��� Ŭ����
Private mBarNo      As String       '���ڵ��ȣ
Private mSpcPos     As String       '��񿡼� ��ü��ġ

'## Vitek��񿡼� ���
Private mDrugs      As clsIISMICDrugs   '�׻������ �÷��� Ŭ����
Private mVitekNo    As String       'Vitek No
Private mMnmCd      As String       '���ڵ�
Private mMnmNm      As String       '�ո�(���)
Private mMnmNmFull  As String       '�ո�(��ü)

Public Property Get IntResults() As clsIISIntResults
    Set IntResults = mIntResults
End Property

Public Property Let IntResults(ByVal vData As clsIISIntResults)
    Set mIntResults = vData
End Property

Public Property Get Drugs() As clsIISMICDrugs
    Set Drugs = mDrugs
End Property

Public Property Let Drugs(ByVal vData As clsIISMICDrugs)
    Set mDrugs = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SpcPos() As String
    SpcPos = mSpcPos
End Property

Public Property Let SpcPos(ByVal vData As String)
    mSpcPos = vData
End Property

Public Property Get VitekNo() As String
    VitekNo = mVitekNo
End Property

Public Property Let VitekNo(ByVal vData As String)
    mVitekNo = vData
End Property

Public Property Get MnmCd() As String
    MnmCd = mMnmCd
End Property

Public Property Let MnmCd(ByVal vData As String)
    mMnmCd = vData
End Property

Public Property Get MnmNm() As String
    MnmNm = mMnmNm
End Property

Public Property Let MnmNm(ByVal vData As String)
    mMnmNm = vData
End Property

Public Property Get MnmNmFull() As String
    MnmNmFull = mMnmNmFull
End Property

Public Property Let MnmNmFull(ByVal vData As String)
    mMnmNmFull = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ��ȸ
'   ��ȯ : ��������
'-----------------------------------------------------------------------------'
Public Function GetAccDt() As String
    Dim Rs      As ADODB.Recordset
    Dim objSQL  As clsIISMicroSql
    
    If mVitekNo = "" Then Exit Function
    
    '## 1.1.3: �̻��(2005-01-03)
    '   - Vitek No�� �̿��� �������ڸ� ���ϴ� ��������
On Error GoTo Errors
    Set objSQL = New clsIISMicroSql
    Set Rs = DbCon.Execute(objSQL.SelectAccDt(IISMICWA, Mid$(mVitekNo, 1, 2), _
                Mid$(mVitekNo, 3)), , adCmdText)
    If Rs.BOF Or Rs.EOF Then
        GetAccDt = Format$(Now, "YYYY") & Mid$(mVitekNo, 1, 2)
    Else
        GetAccDt = Rs.Fields("accdt").Value & ""
    End If
    Rs.Close
    Set Rs = Nothing
    Set objSQL = Nothing
    Exit Function
    
Errors:
    Set Rs = Nothing
    Set objSQL = Nothing
    Error.SetLog App.EXEName, "clsIISIntInfo", "GetAccDt", Err.Description, Now
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ��ȸ
'   ��ȯ : ��������
'-----------------------------------------------------------------------------'
Public Function GetAccSeq() As Long
    If mVitekNo = "" Then Exit Function
    GetAccSeq = CLng(Mid$(mVitekNo, 3))
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ո�(���)�� �̿��� ���ڵ� ��ȸ
'-----------------------------------------------------------------------------'
Public Sub GetMnmCd()
    Dim objSensi As clsIISMICSensi
    
    Set objSensi = New clsIISMICSensi
    
    mMnmCd = objSensi.GetMnmCd(mMnmNm)
    Set objSensi = Nothing
End Sub

Private Sub Class_Initialize()
    Set mIntResults = New clsIISIntResults
    Set mDrugs = New clsIISMICDrugs
End Sub

Private Sub Class_Terminate()
    Set mDrugs = Nothing
    Set mIntResults = Nothing
End Sub