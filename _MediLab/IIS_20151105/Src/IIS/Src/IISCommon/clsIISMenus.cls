VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISMenus.cls
'   �ۼ���  : ������
'   ��  ��  : Menus �÷��� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'
Option Explicit

Private mMenus As Collection            '�޴� �÷���

Public Property Get Count() As Long
    Count = mMenus.Count
End Property

Public Property Get Item(ByVal pKey As String) As clsIISMenu
Attribute Item.VB_UserMemId = 0
    Set Item = mMenus(pKey)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mMenus.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ� �׸��߰�
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pKey As String, ByVal pVisible As Boolean) As clsIISMenu
    Dim objMenu As clsIISMenu
    
    '## �÷��ǿ� �ִ� �׸��� ��� ������ ���׸� �߰�
    Set objMenu = New clsIISMenu
    With objMenu
        .Key = pKey
        .Visible = pVisible
    End With
    
    mMenus.Add objMenu, pKey
    
    Set Add = objMenu
    Set objMenu = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �÷��ǿ��� �ش��� ����
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pKey As String)
    mMenus.Remove pKey
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �÷����� ��� ��һ���
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i   As Long
    
    For i = mMenus.Count To 1 Step -1
        mMenus.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� Ű�� ���������� ��ȯ - 1.1.7: (2005-04-13)
'   �μ� :
'       - pKey : ���ۼ���
'   ��ȯ : True(����) / Flase(������)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pKey As String) As Boolean
    Dim objMenu As clsIISMenu
    
    For Each objMenu In mMenus
        With objMenu
            If pKey = .Key Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objMenu = Nothing
End Function

Private Sub Class_Initialize()
    Set mMenus = New Collection
End Sub

Private Sub Class_Terminate()
    Set mMenus = Nothing
End Sub