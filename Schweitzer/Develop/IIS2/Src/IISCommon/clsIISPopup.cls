VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISPopup.cls
'   �ۼ���  : ������
'   ��  ��  : frmIISPopup ������ �������̽�
'   �ۼ���  : 2004-06-02
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

'Popup Menu Style Enum
Public Enum eMenuStyle
    ccDefault = MF_STRING
    ccDisabled = MF_GRAYED Or MF_DISABLED
    ccSeparator = MF_SEPARATOR
End Enum

Private hMenu       As Long     '������ �˾� �޴��� �ڵ鼱��
Private mvarMenuID  As Long     'Menu ID

Public Event Click(ByVal vMenuID As Long)   'Click �̺�Ʈ

Public Property Get MenuID() As Long
    MenuID = mvarMenuID
End Property

Public Sub AddMenu(ByVal vMenuID As Long, ByVal vMenuCaption As String, _
                Optional vMenuStyle As eMenuStyle = ccDefault)
    If hMenu = 0 Then hMenu = CreatePopupMenu
    If hMenu = 0 Then
        MsgBox "�˾� �޴� ������ �����Ͽ����ϴ�.", vbCritical
        Exit Sub
    End If
    
    If vMenuID = 0 Then
        MsgBox "�޴� ���̵�� ""0""�� ����� �� �����ϴ�.", vbCritical
        Exit Sub
    End If
    
    AppendMenu hMenu, vMenuStyle, vMenuID, vMenuCaption
End Sub

Public Sub PopupMenus(ByVal vhWnd As Long)
    Dim Pt  As POINTAPI
    Dim ret As Long
    
    GetCursorPos Pt

    ret = TrackPopupMenu(hMenu, TPM_LEFTALIGN & TPM_RETURNCMD, Pt.X, Pt.Y, 0, vhWnd, ByVal 0&)
    If ret <> 0 Then
        mvarMenuID = ret
        RaiseEvent Click(ret)
    End If
End Sub

Private Sub Class_Initialize()
    hMenu = 0
    mvarMenuID = 0
End Sub

Private Sub Class_Terminate()
    If hMenu <> 0 Then DestroyMenu hMenu
End Sub
