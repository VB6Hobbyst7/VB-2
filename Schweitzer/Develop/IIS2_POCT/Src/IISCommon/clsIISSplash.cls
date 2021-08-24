VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISSplash.cls
'   �ۼ���  : �̻��
'   ��  ��  : Splash���� �ε�, �޽��� ���� Ŭ����
'   �ۼ���  : 2003-12-04
'   ��  ��  :
'-----------------------------------------------------------------------------'
Option Explicit

Private mSplashForm     As frmIISSplash     'Splash Form
Private mProjectNm      As String           '������Ʈ��
Private mVersion        As String           '����
Private mRegister       As String           '�����
Private mMessage        As String           '�޽���

Public Property Let ProjectNm(ByVal vData As String)
    mProjectNm = vData
    mSplashForm.lblProjectNm = mProjectNm
End Property

Public Property Let Version(ByVal vData As String)
    mVersion = vData
    mSplashForm.lblVersion = "Version " & mVersion
End Property

Public Property Let Register(ByVal vData As String)
    mRegister = vData
    mSplashForm.lblRegister.Caption = "This version is registered to " & mRegister
End Property

Public Property Let Message(ByVal vData As String)
    mMessage = vData
    mSplashForm.lblMessage = mMessage & "..."
End Property

'-----------------------------------------------------------------------------'
'   ��� : Splash�� �ε�
'   �μ� : ǥ���� �޽���
'-----------------------------------------------------------------------------'
Public Sub LoadSplash()
    Dim objCom  As clsIISCommon
    
    '## ��ǻ�� �̸� ���
    Set objCom = New clsIISCommon
    If mRegister = "" Then
        mRegister = objCom.mGetComNm
        mSplashForm.lblRegister.Caption = "This version is registered to " & mRegister
    End If
    
    Set objCom = Nothing
    mSplashForm.Show
    mSplashForm.ZOrder 0
End Sub

'-----------------------------------------------------------------------------'
'   ��� : Splash���� �޽����� ǥ��
'   �μ� : ǥ���� �޽���
'-----------------------------------------------------------------------------'
Public Sub SetMsg(ByVal pMsg As String)
    mSplashForm.lblMessage = pMsg & "..."
End Sub

Private Sub Class_Initialize()
    Set mSplashForm = New frmIISSplash
End Sub

Private Sub Class_Terminate()
    Unload mSplashForm
    Set mSplashForm = Nothing
End Sub