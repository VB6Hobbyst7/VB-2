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
'   ��  ��  : ADVIA 120 ��� �������� Ŭ����
'   �ۼ���  : 2004-06-10
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mMT     As Long     'MT Value
Private mBarNo  As String   '���ڵ��ȣ

Public Property Get MT() As String
    MT = Chr(mMT)
End Property

Public Property Let MT(ByVal vData As String)
    mMT = Asc(vData)
End Property

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
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
    Dim objIntNm    As String           '��� �˻��׸� Ŭ����
    Dim strIntBase  As String           '������ �˻��׸�
    Dim strStatFg   As String           '��������
    Dim strOrder    As String           '�������� ���ڿ�
    
    '## ������ �˻��׸� ��ȸ
    '## ������ �˻���� 12�ڸ� ������ ������ ����
    If pAccInfo.QcFg = "0" Then         '## �Ϲݰ�ü
        For Each objResult In pAccInfo.Results
            strIntBase = strIntBase & objResult.IntNm.IntBase
        Next
        Set objResult = Nothing
    ElseIf pAccInfo.QcFg = "1" Then     '## QC��ü
        For Each objQCResult In pAccInfo.QCResults
            strIntBase = strIntBase & objQCResult.IntNm.IntBase
        Next
        Set objQCResult = Nothing
    End If
    
'    strIntBase = strIntBase & "001002003"


    '## �������� ��ȸ (�������׸��� ��������)
    strStatFg = IIf(pAccInfo.QcFg = "1", " U", "  ")
'    strStatFg = "  "
    
    '## �������� �ʴ� �׸��� ��� �������� ����
    strOrder = GetMT & "Y " & strStatFg & "  " & Format$(mBarNo, String$(14, "0")) & Space(25) & _
               Format$(pAccInfo.PtId, String$(14, "0")) & Space(75) & vbCrLf & strIntBase & vbCrLf
    
'    strOrder = GetMT & "Y " & strStatFg & "  " & Format$(mBarNo, String$(14, "0")) & Space(25) & _
               Space(14) & Space(75) & vbCrLf & strIntBase & vbCrLf
    
    GetOrder = STX & strOrder & GetLRC(strOrder) & ETX
End Function

'-----------------------------------------------------------------------------'
'   ��� : ���������� ��������� ���ڿ� ��ȸ
'   ��ȯ : No Workorder ���ڿ�
'-----------------------------------------------------------------------------'
Public Function GetNoOrder() As String
    Dim strNoOrder As String    'No Workorder ���ڿ�
    
    strNoOrder = GetMT & "N R " & Format$(mBarNo, String$(14, "0")) & vbCrLf
    strNoOrder = STX & strNoOrder & GetLRC(strNoOrder) & ETX
    GetNoOrder = strNoOrder
End Function

'------------------------------------------------------------------------------'
'   ��� : Initialization Message ��ȸ
'   ��ȯ : Initialization Message
'------------------------------------------------------------------------------'
Public Function GetInit() As String
    Dim strInit As String   'Initialization ���ڿ�
    
    mMT = &H30
    strInit = MT & "I " & vbCrLf
    strInit = STX & strInit & GetLRC(strInit) & ETX
    GetInit = strInit
End Function

'-----------------------------------------------------------------------------'
'   ��� : Token Transfer Message ��ȸ
'   ��ȯ : Token Transfer Message
'-----------------------------------------------------------------------------'
Public Function GetToken() As String
    Dim strToken As String  'Token Transfer ���ڿ�
    
    strToken = GetMT & "S" & Space(10) & vbCrLf
    strToken = STX & strToken & GetLRC(strToken) & ETX
    GetToken = strToken
End Function

'-----------------------------------------------------------------------------'
'   ��� : Result Validation Message ��ȸ
'   ��ȯ : Result Validation Message
'-----------------------------------------------------------------------------'
Public Function GetResultValid() As String
    Dim strValid As String  'Result Validation ���ڿ�
    
    strValid = GetMT & "Z" & Space(17) & " 0" & vbCrLf
    strValid = STX & strValid & GetLRC(strValid) & ETX
    GetResultValid = strValid
End Function

'-----------------------------------------------------------------------------'
'   ��� : MT ��ȸ
'   ��ȯ : MT
'-----------------------------------------------------------------------------'
Public Function GetMT() As String
    If mMT = 0 Or mMT = 90 Then '## H30=0, H5A=90
        mMT = &H30
    Else
        mMT = mMT + 1
    End If
    
    GetMT = Chr(mMT)
End Function

'-----------------------------------------------------------------------------'
'   ��� : LRC(Longgitudinal Redundancy Check) ��ȸ
'   ��ȯ : LRC
'-----------------------------------------------------------------------------'
Public Function GetLRC(ByVal pData As String) As String
    Dim lngLRC  As Long     'LRC
    Dim i       As Long
    
    For i = 1 To Len(pData)
        lngLRC = lngLRC Xor Asc(Mid$(pData, i, 1))
    Next i
    
    If lngLRC = 3 Then
        GetLRC = Chr(&H7F)
    Else
        GetLRC = Chr(lngLRC)
    End If
End Function

'-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
End Sub