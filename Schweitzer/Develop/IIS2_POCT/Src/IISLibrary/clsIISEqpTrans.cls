VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISEqpTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISEqpTrans.cls (�ڻ� LIS�� �����Ҷ�)
'   �ۼ���  : �̻��
'   ��  ��  : ������۳��� Ŭ����
'   �ۼ���  : 2004-04-16
'   ��  ��  :
'-----------------------------------------------------------------------------'

'## �Ӽ�
'   - ����ڵ�
'   - ��������
'   - ���۽ð�
'   - ��ü����
'   - ��ü��ȣ
'   - ��񿡼� ��ü��ġ

'## �޼ҵ�
'   - ���۳��� �Է�
'   - ���۳��� ���� (����ڵ�, ��ü����+��ȣ�� ��ġ�ϴ°͸�)

Option Explicit

Private mSQL        As clsIISLibrarySql

Private mEqpCd      As String   '����ڵ�
Private mTransDt    As String   '��������
Private mTransTm    As String   '���۽ð�
Private mSpcYy      As String   '��ü����
Private mSpcNo      As String   '��ü��ȣ
Private mTransNo    As String   '��񿡼��� ��ü��ġ

Public Property Get EqpCd() As String
    EqpCd = mEqpCd
End Property

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Get TransDt() As String
    TransDt = mTransDt
End Property

Public Property Let TransDt(ByVal vData As String)
    mTransDt = vData
End Property

Public Property Get TransTm() As String
    TransTm = mTransTm
End Property

Public Property Let TransTm(ByVal vData As String)
    mTransTm = vData
End Property

Public Property Get SpcYy() As String
    SpcYy = mSpcYy
End Property

Public Property Let SpcYy(ByVal vData As String)
    mSpcYy = vData
End Property

Public Property Get SpcNo() As String
    SpcNo = mSpcNo
End Property

Public Property Let SpcNo(ByVal vData As String)
    mSpcNo = vData
End Property

Public Property Get TransNo() As String
    TransNo = mTransNo
End Property

Public Property Let TransNo(ByVal vData As String)
    mTransNo = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ��� ���۳��� �Է�
'   �μ� :
'       - pEqpCd    : ����ڵ�
'       - pTransDt  : ��񿡼� ��� ��������
'       - pTransTm  : ��񿡼� ��� ���۽ð�
'       - pSpcYy    : ��ü����
'       - pSpcNo    : ��ü��ȣ
'       - pTransNo  : ��񿡼� ��ü��ġ
'   ��ȯ : ��� ���۳��� �Է�����
'-----------------------------------------------------------------------------'
Public Function GetAddEqpTransSql(ByVal pEqpCd As String, ByVal pTransDt As String, _
                    ByVal pTransTm As String, ByVal pSpcYy As String, ByVal pSpcNo As String, _
                    ByVal pTransNo As String) As String
    Dim SQL As String
    
    SQL = mSQL.InsertEqpTrans(pEqpCd, pTransDt, pTransTm, pSpcYy, pSpcNo, pTransNo)
    
    GetAddEqpTransSql = SQL
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��� ���۳��� ����
'   �μ� :
'       - pEqpCd    : ����ڵ�
'       - pSpcYy    : ��ü����
'       - pSpcNo    : ��ü��ȣ
'   ��ȯ : ��� ���۳��� ��������
'-----------------------------------------------------------------------------'
Public Function GetDelEqpTransSql(ByVal pEqpCd As String, ByVal pSpcYy As String, _
                    ByVal pSpcNo As String) As String
    Dim SQL As String
    
    SQL = mSQL.DeleteEqpTrans(pEqpCd, pSpcYy, pSpcNo)
                
    GetDelEqpTransSql = SQL
End Function

Private Sub Class_Initialize()
    Set mSQL = New clsIISLibrarySql
End Sub

Private Sub Class_Terminate()
    Set mSQL = Nothing
End Sub