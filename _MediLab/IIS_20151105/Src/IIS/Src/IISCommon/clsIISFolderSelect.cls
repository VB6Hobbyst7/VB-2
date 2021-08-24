VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISFolderSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISFolderSelect.cls
'   �ۼ���  : ������
'   ��  ��  : ������������ �����ִ� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Public Event SelectedFolder(ByVal pSelFolder As String)

Private WithEvents mFolder As frmIISFolderSelect   '�������� ��
Attribute mFolder.VB_VarHelpID = -1

Private mPath       As String           '������ ���
Private mCaption    As String           '������������ Caption
Private mLeft       As Long             '������������ Left ��ǥ
Private mTop        As Long             '������������ Top ��ǥ

Public Property Let Path(ByVal vData As String)
    mPath = vData
End Property

Public Property Let Caption(ByVal vData As String)
    mCaption = vData
End Property

Public Property Let Left(ByVal vData As Long)
    mLeft = vData
End Property

Public Property Let Top(ByVal vData As Long)
    mTop = vData
End Property

'-----------------------------------------------------------------------------'
'   ��� : ���������� ǥ��
'   �μ� :
'       - pLeft : ������������ Left ��ǥ
'       - pTop  : ������������ Top ��ǥ
'-----------------------------------------------------------------------------'
Public Sub ShowFolderSelect(Optional ByVal pLeft As Long = 0, Optional ByVal pTop As Long = 0)
    '## Left, Top��ǥ�� �˻�
    Call SetWinPos(pLeft, pTop)
    
    If Dir$(mPath, vbDirectory) = "" Then
        mPath = ""
    End If
    
    '## FolderSelect �Ӽ�����, ǥ��
    Set mFolder = frmIISFolderSelect
    With mFolder
        .Left = mLeft: .Top = mTop
        .Caption = "�� " & mCaption
        .Path = mPath
        .ZOrder 0
        .Show vbModal
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���������� ǥ��
'   �μ� :
'       - pCaption  : ������������ Caption
'       - pPath     : ���缱�õ� Path
'       - pLeft     : ������������ Left ��ǥ
'       - pTop      : ������������ Top ��ǥ
'-----------------------------------------------------------------------------'
Public Sub ShowFolderSelectX(ByVal pCaption As String, ByVal pPath As String, _
                             Optional ByVal pLeft As Long = 0, _
                             Optional ByVal pTop As Long = 0)
    '## Left, Top��ǥ�� �˻�
    Call SetWinPos(pLeft, pTop)
    
    If Dir$(pPath, vbDirectory) = "" Then
        mPath = ""
    End If
    
    '## FolderSelect �Ӽ�����, ǥ��
    Set mFolder = frmIISFolderSelect
    With mFolder
        .Left = mLeft: .Top = mTop
        .Caption = "�� " & pCaption
        .Path = pPath
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
'   ��� : �������������� ������ ���������� �߻��ϴ� �̺�Ʈ
'   �μ�
'       1.pSelFolder : �������������� ������ ������ ��ü���
'-----------------------------------------------------------------------------'
Private Sub mFolder_SelectedFolder(ByVal pSelFolder As String)
    RaiseEvent SelectedFolder(pSelFolder)
    
    Unload mFolder
    Set mFolder = Nothing
End Sub

Private Sub Class_Terminate()
    Set mFolder = Nothing
End Sub