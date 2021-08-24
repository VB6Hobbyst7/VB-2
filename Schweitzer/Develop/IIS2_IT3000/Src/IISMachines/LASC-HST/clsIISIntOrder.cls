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
'   ��  ��  : LASC-HST �������� Ŭ����
'   �ۼ���  : 2005-09-15
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mBarNo  As String           '���ڵ��ȣ
Private mRow    As Long             'Spread�� Row

Public Property Get BarNo() As String
    BarNo = mBarNo
End Property

Public Property Let BarNo(ByVal vData As String)
    mBarNo = vData
End Property

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal vData As Long)
    mRow = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� ���ڿ� ��ȸ
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'   ��ȯ : �������� ���ڿ�1
'-----------------------------------------------------------------------------'
Public Function GetOrder(ByVal pAccInfo As clsIISAccInfo) As String
    Dim objResult       As clsIISResult     '������� Ŭ����
    Dim strIntBase      As String           '������ �˻��
    Dim strItems        As String           '�۽��� �˻��׸� ���ڿ�
    Dim strOutput       As String           '�۽��� ������
    Dim strPrintNo1     As String
    Dim strPrintNo2     As String
    Dim strPrintNo3     As String
    Dim blnSPOrder      As Boolean

    strItems = String$(38, "0")
    For Each objResult In pAccInfo.Results
        strIntBase = mGetP(objResult.IntNm.IntBase, 1, "|")
        
        Select Case strIntBase
            Case "RET"
                Mid$(strItems, 1, 1) = 1
            Case "SP"
                Mid$(strItems, 2, 1) = 1
                blnSPOrder = True
            Case Else
                Mid$(strItems, CLng(strIntBase) + 3, 1) = 1
        End Select
    Next
    Set objResult = Nothing
    
    If blnSPOrder Then
        Mid$(strItems, 3, 1) = "0"
    Else
        Mid$(strItems, 3, 1) = "1"
    End If
    
    mBarNo = Format$(mBarNo, String$(13, "0"))
    strPrintNo1 = Format$(pAccInfo.AccDt & "-" & CStr(pAccInfo.AccSeq), String$(13, "@"))
    strPrintNo2 = Space(13)
    strPrintNo3 = Space(13)
    
    strOutput = STX & "S" & Space(8) & String$(2, "0") & mBarNo & "********" & strItems & _
                String$(19, "0") & strPrintNo1 & strPrintNo2 & strPrintNo3 & "100" & _
                String$(25, "*") & Space(15) & ETX
    
    GetOrder = strOutput
End Function

''-----------------------------------------------------------------------------'
'   ��� : Ŭ���� ������� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Public Sub ClsClear()
    mBarNo = ""
    mRow = 0
End Sub