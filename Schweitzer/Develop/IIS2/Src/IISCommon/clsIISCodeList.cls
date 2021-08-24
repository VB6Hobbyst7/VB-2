VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISCodeList.cls
'   �ۼ���  : �̻��
'   ��  ��  : Code List �����츦 �����ִ� Ŭ����
'   �ۼ���  : 2004-02-19
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Public Event SelectedItem(ByRef pSelItem As String)

Private WithEvents mCodeList As frmIISCodeList     'CodeList��
Attribute mCodeList.VB_VarHelpID = -1

Private mLeft       As Long         'CodeList���� ǥ�õǴ� Left��ǥ
Private mTop        As Long         'CodeList���� ǥ�õǴ� Top��ǥ
Private mCaption    As String       'CodeList���� Caption
Private mHeaderCd   As String       'CodeList���� Listview�� �ڵ忡 �ش��ϴ� �÷���
Private mHeaderCdNm As String       'CodeList���� Listview�� �ڵ���� �ش��ϴ� �÷���

Public Property Get Left() As Long
    Left = mLeft
End Property

Public Property Let Left(ByVal vData As Long)
    mLeft = vData
End Property

Public Property Get Top() As Long
    Top = mTop
End Property

Public Property Let Top(ByVal vData As Long)
    mTop = vData
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal vData As String)
    mCaption = "�� " & vData
End Property

Public Property Get HeaderCd() As String
    HeaderCd = mHeaderCd
End Property

Public Property Let HeaderCd(ByVal vData As String)
    mHeaderCd = vData
End Property

Public Property Get HeaderCdNm() As String
    HeaderCdNm = mHeaderCdNm
End Property

Public Property Let HeaderCdNm(ByVal vData As String)
    mHeaderCdNm = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���޵� SQL���� �����Ͽ� Rs(0), Rs(1)�� ������ Listview�� ǥ��
'   �μ� :
'       1.pSQL  : �ڵ�, �ڵ���� ������ SQL��
'       2.pLeft : CodeList���� ǥ�õǴ� Left��ǥ
'       3.pTop  : CodeList���� ǥ�õǴ� Top��ǥ
'-----------------------------------------------------------------------------'
Public Sub CodeListBySql(ByVal pSQL As String, Optional ByVal pLeft As Long = 0, _
                         Optional ByVal pTop As Long = 0)

    '## Left, Top��ǥ�� �˻�
    Call SetWinPos(pLeft, pTop)
    
    '## CodeList �Ӽ�����, ǥ��
    Set mCodeList = frmIISCodeList
    With mCodeList
        .Left = mLeft: .Top = mTop
        .Caption = mCaption
        .lvwCodeList.ColumnHeaders(1).Text = mHeaderCd
        .lvwCodeList.ColumnHeaders(2).Text = mHeaderCdNm
        .ParamType = ccSql: .SQL = pSQL
        
        .ZOrder 0
        .Show vbModal
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���޵� Recordset�� Rs(0), Rs(1)�� ������ Listview�� ǥ��
'   �μ� :
'       1.pRs   : �ڵ�, �ڵ���� ������ RecordSet
'       2.pLeft : CodeList���� ǥ�õǴ� Left��ǥ
'       3.pTop  : CodeList���� ǥ�õǴ� Top��ǥ
'-----------------------------------------------------------------------------'
Public Sub CodeListByRs(ByRef pRs As Recordset, Optional ByVal pLeft As Long = 0, _
                        Optional ByVal pTop As Long = 0)
    '## Left, Top��ǥ�� �˻�
    Call SetWinPos(pLeft, pTop)
    
    '## CodeList �Ӽ�����, ǥ��
    Set mCodeList = frmIISCodeList
    With mCodeList
        .Left = mLeft: .Top = mTop
        .Caption = mCaption
        .lvwCodeList.ColumnHeaders(1).Text = mHeaderCd
        .lvwCodeList.ColumnHeaders(2).Text = mHeaderCdNm
        .ParamType = ccRecordset: .Rs = pRs
        
        .ZOrder 0
        .Show vbModal
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���޵� �÷��� Ŭ������ .Cd, .CdNm�� ������ Listview�� ǥ��
'   �μ� :
'       1.pCol  : �÷��� Ŭ����
'       2.pLeft : CodeList���� ǥ�õǴ� Left��ǥ
'       3.pTop  : CodeList���� ǥ�õǴ� Top��ǥ
'   �ʿ���� : Ŭ������ �ݵ�� Cd, CdNm Property�� �����ؾ� �Ѵ�.
'-----------------------------------------------------------------------------'
Public Sub CodeListByCol(ByVal pCol As Object, Optional ByVal pLeft As Long = 0, _
                         Optional ByVal pTop As Long = 0)
    '## Left, Top��ǥ�� �˻�
    Call SetWinPos(pLeft, pTop)
    
    '## CodeList �Ӽ�����, ǥ��
    Set mCodeList = frmIISCodeList
    With mCodeList
        .Left = mLeft: .Top = mTop
        .Caption = mCaption
        .lvwCodeList.ColumnHeaders(1).Text = mHeaderCd
        .lvwCodeList.ColumnHeaders(2).Text = mHeaderCdNm
        .ParamType = ccCollection: .Col = pCol
        
        .ZOrder 0
        .Show vbModal
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���� ��ġ�� �������� ���� ��� ���� ���콺 �����ͷ� ���߾�
'          CodeList���� Left, Top�� ��ġ�� ����
'-----------------------------------------------------------------------------'
Private Sub SetWinPos(ByVal pLeft As Long, ByVal pTop As Long)
    Dim Point As POINTAPI
    
    Call GetCursorPos(Point)
    
    mLeft = IIf(pLeft = 0, Point.X * Screen.TwipsPerPixelX, pLeft)
    mTop = IIf(pTop = 0, Point.Y * Screen.TwipsPerPixelY, pTop)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : CodeList������ �������� ���������� �߻��ϴ� �̺�Ʈ
'   �μ�
'       1.pSelItem : CodeList������ ������ ������
'-----------------------------------------------------------------------------'
Private Sub mCodeList_SelectedItem(ByVal pSelItem As String)
    RaiseEvent SelectedItem(pSelItem)
    
    '## �̺�Ʈ �����ʿ��� pSelItem=""�� �Է��ϸ� ���� CodeList���� �״�� ����
    If pSelItem <> "" Then
        Unload mCodeList
        Set mCodeList = Nothing
    End If
End Sub

Private Sub Class_Initialize()
    mHeaderCd = "�ڵ�"
    mHeaderCdNm = "�ڵ��"
End Sub

Private Sub Class_Terminate()
    Set mCodeList = Nothing
End Sub
