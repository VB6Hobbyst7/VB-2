VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISMenu.cls
'   �ۼ���  : �̻��
'   ��  ��  : Menu Ŭ����
'   �ۼ���  : 2003-12-18
'   ��  ��  :
'-----------------------------------------------------------------------------'
Option Explicit

Private mKey        As String       '�޴��׸��� Key
Private mCaption    As String       '�޴��׸��� Caption
Private mTag        As String       '�޴��׸��� Tag
Private mVisible    As Boolean      '�޴��׸��� Visible

Public Property Get Key() As String
    Key = mKey
End Property

Public Property Let Key(ByVal vData As String)
    mKey = vData
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal vData As String)
    mCaption = vData
End Property

Public Property Get Tag() As String
    Tag = mTag
End Property

Public Property Let Tag(ByVal vData As String)
    mTag = vData
End Property

Public Property Get Visible() As Boolean
    Visible = mVisible
End Property

Public Property Let Visible(ByVal vData As Boolean)
    mVisible = vData
End Property
