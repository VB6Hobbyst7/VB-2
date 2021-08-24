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
'   �ۼ���  : ������
'   ��  ��  : XE2100ASTM ��� �������� Ŭ����
'   �ۼ���  : 2008-02-20
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo   '�������� Ŭ����

Private mBarNo      As String   '���ڵ��ȣ
Private mSpcType    As String   '��ü����
Private mRackNo     As String   'Rack No
Private mPos        As String   'Position
Private mKind       As String   'Run(R1), Rerun(R2), QC(Q) ����
Private mPriority   As String   '���ޱ���(R:Routine, S:Stat)
Private mOrder      As String   '���� ������ ��������
Private mIsSending  As Boolean  '���� ������ ���� ���ڿ��� ��������
Private mNoOrder    As Boolean  '�������� �������� (True:������, False:����)

Public Property Get IsSending() As Boolean
    IsSending = mIsSending
End Property

Public Property Let IsSending(ByVal vData As Boolean)
    mIsSending = vData
End Property

Public Property Get AccInfo() As clsIISAccInfo
    Set AccInfo = mAccInfo
End Property

Public Property Let AccInfo(ByVal vData As clsIISAccInfo)
    Set mAccInfo = vData
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get SpcType() As String
    SpcType = mSpcType
End Property

Public Property Let SpcType(ByVal vData As String)
    mSpcType = vData
End Property

Public Property Get RackNo() As String
    RackNo = mRackNo
End Property

Public Property Let RackNo(ByVal vData As String)
    mRackNo = vData
End Property

Public Property Get Pos() As String
    Pos = mPos
End Property

Public Property Let Pos(ByVal vData As String)
    mPos = vData
End Property

Public Property Get Kind() As String
    Kind = mKind
End Property

Public Property Let Kind(ByVal vData As String)
    mKind = vData
End Property

Public Property Get Priority() As String
    Priority = mPriority
End Property

Public Property Let Priority(ByVal vData As String)
    mPriority = vData
End Property

Public Property Get Order() As String
    Order = mOrder
End Property

Public Property Let Order(ByVal vData As String)
    mOrder = vData
End Property

Public Property Get NoOrder() As Boolean
    NoOrder = mNoOrder
End Property

Public Property Let NoOrder(ByVal vData As Boolean)
    mNoOrder = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ��� ��ȸ
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim objIntNm    As clsIISIntNm      '��� �˻��׸� Ŭ����
    Dim strIntBase  As String           '������ �˻��
    Dim strItems    As String           '������ �˻��׸�
    Dim blnISE      As Boolean          'ISE��� �˻��׸� ����

    '## �������ڿ� ����: ^^^^WBC\^^^^RBC\......\^^^^HFR
    
    If mAccInfo.QcFg = "0" Then
        For Each objResult In mAccInfo.Results
            Set objIntNm = objResult.IntNm
            strIntBase = objIntNm.IntBase
            'NRBC%�� ������ ���ش�.
            If strIntBase <> "NRBC%" Then
                If strItems = "" Then
                    strItems = "^^^^" & strIntBase '& "/"
                Else
                    strItems = strItems & "\^^^^" & strIntBase '& "/"
                End If
            End If
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then
        For Each objQCResult In mAccInfo.QCResults
            Set objIntNm = objQCResult.IntNm
            strIntBase = objIntNm.IntBase
            'NRBC%�� ������ ���ش�.
            If strIntBase = "NRBC%" Then
                If strItems = "" Then
                    strItems = "^^^^" & strIntBase '& "/"
                Else
                    strItems = strItems & "\^^^^" & strIntBase '& "/"
                End If
            End If
        Next
        Set objQCResult = Nothing
    End If
    Set objIntNm = Nothing

    'strItems = Mid(strItems, 1, Len(strItems) - 1)
    GetOrder = strItems
    
End Function

'-----------------------------------------------------------------------------'
'   ��� : ��ü������ ��ȸ
'   ��ȯ : ��ü����
'-----------------------------------------------------------------------------'
'Public Function GetSpcType() As String
''
'End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ��ȸ
'   ��ȯ : ��������
'-----------------------------------------------------------------------------'
'Public Function GetStatFg() As String
''
'End Function

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� CheckSum�� ����
'   �μ� :
'       - pMsg : ���ڿ�
'   ��ȯ : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mSpcType = ""
    mRackNo = ""
    mPos = ""
    mKind = ""
    mPriority = ""
    mOrder = ""
    mNoOrder = False
    mIsSending = False
    
    Set mAccInfo = Nothing
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub