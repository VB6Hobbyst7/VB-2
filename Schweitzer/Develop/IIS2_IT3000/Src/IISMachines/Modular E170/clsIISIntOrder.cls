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
'   ��  ��  : Modular E170 ��� �������� Ŭ����
'   �ۼ���  : 2005-02-07
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mAccInfo    As clsIISAccInfo   '�������� Ŭ����

Private mBarNo      As String   '���ڵ��ȣ
Private mRackNo     As String   'Rack No
Private mPos        As String   'Position
Private mOrder      As String   '���� ������ ��������
Private mNoOrder    As Boolean  '�������� �������� (True:������, False:����)

'## V1.0
Private mSpcType    As String   '��ü����
Private mKind       As String   'Run(R1), Rerun(R2), QC(Q) ����
Private mPriority   As String   '���ޱ���(R:Routine, S:Stat)

'## V1.3
Private mRackType   As String   'Rack Type
Private mConType    As String   'Container Type


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

Public Property Get SpcType() As String
    SpcType = mSpcType
End Property

Public Property Let SpcType(ByVal vData As String)
    mSpcType = vData
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

Public Property Get RackType() As String
    RackType = mRackType
End Property

Public Property Let RackType(ByVal vData As String)
    mRackType = vData
End Property

Public Property Get ConType() As String
    ConType = mConType
End Property

Public Property Let ConType(ByVal vData As String)
    mConType = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : Instrement Specimean ID ��ȸ
'   ��ȯ :
'       - <Sampe No>^<Rack ID>^<Position No>^^<Rack Type>^<Container Type> ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetInstSpcId() As String
    Dim strTemp As String
    
    GetInstSpcId = "0^" & mRackNo & "^" & mPos & "^^" & mRackType & "^" & mConType
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ��� ��ȸ
'   ��ȯ : �������� ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetOrder() As String
    Dim objResult   As clsIISResult     '������� Ŭ����
    Dim objQCResult As clsIISQCResult   'QC������� Ŭ����
    Dim strIntBase  As String           '������ �˻��
    Dim blnISE      As Boolean          'ISE��� �˻��׸� ����
    Dim strItems    As String           '�۽��� �˻��׸� ���ڿ�

    '## �������ڿ� ����: ^^^1/\^^^2
    If mAccInfo.QcFg = "0" Then
        For Each objResult In mAccInfo.Results
            strIntBase = objResult.IntNm.IntBase

            If strItems = "" Then
                strItems = "^^^" & strIntBase & "/"
            Else
                strItems = strItems & "\^^^" & strIntBase & "/"
            End If
        Next
        Set objResult = Nothing
    ElseIf mAccInfo.QcFg = "1" Then
        For Each objQCResult In mAccInfo.QCResults
            strIntBase = objQCResult.IntNm.IntBase

            If strItems = "" Then
                strItems = "^^^" & strIntBase & "/"
            Else
                strItems = strItems & "\^^^" & strIntBase & "/"
            End If
        Next
        Set objQCResult = Nothing
    End If

    GetOrder = strItems
End Function

'-----------------------------------------------------------------------------'
'   ��� : Priority ��ȸ (����/������ ��ȸ)
'   ��ȯ : R(Routin), S(Stat)
'-----------------------------------------------------------------------------'
Public Function GetPriority() As String
    GetPriority = IIf(mAccInfo.StatFg = "1", "S", "R")
End Function

'-----------------------------------------------------------------------------'
'   ��� : Sample Type ��ȸ
'   ��ȯ : Sample Type
'-----------------------------------------------------------------------------'
Public Function GetSampleType() As String
    Select Case RackType
        Case "S1":  GetSampleType = "1"
        Case "S2":  GetSampleType = "2"
        Case "S3":  GetSampleType = "3"
        Case "S4":  GetSampleType = "4"
        Case "S5":  GetSampleType = "5"
        Case Else:  GetSampleType = "1"
    End Select
End Function

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
    mRackNo = ""
    mPos = ""
    mOrder = ""
    mNoOrder = False
    mSpcType = ""
    mKind = ""
    mPriority = ""
    mRackType = ""
    mConType = ""
    
    Set mAccInfo = Nothing
End Sub

Private Sub Class_Terminate()
    Set mAccInfo = Nothing
End Sub