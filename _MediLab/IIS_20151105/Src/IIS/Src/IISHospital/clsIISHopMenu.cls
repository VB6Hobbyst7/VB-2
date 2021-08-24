VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISHopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISHopMenu.cls
'   �ۼ���  : ������
'   ��  ��  : �������� �޴����� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private mMenus      As clsIISMenus      '�޴� �÷���
Private mHopFrm     As frmIISHospital   '������ Imagelist�� ������ ��
Private mImgList    As Object           'ImageList ����

Public Property Get Menus() As clsIISMenus
    Set Menus = mMenus
End Property

Public Property Get ImgList() As Object
    Set ImgList = mImgList
End Property

'-----------------------------------------------------------------------------'
'   ��� : �������� Ǯ�ٿ�޴� ����
'-----------------------------------------------------------------------------'
Public Sub GetFullMenu()
    '## Ǯ�ٿ�޴� �׸��� �⺻������ Visible=True�� �Ǿ�����
    With MainFrm
        '## ��� �������̽�
        .mnuIIS200.Visible = mMenus("IIS200").Visible
        .mnuIIS201.Visible = mMenus("IIS201").Visible
        .mnuIIS202.Visible = mMenus("IIS202").Visible
        .mnuIISSEP02.Visible = mMenus("IISSEP02").Visible
        .mnuIIS204.Visible = mMenus("IIS204").Visible
        
        .mnuIIS201.Caption = mMenus("IIS201").Caption
        .mnuIIS201.Tag = mMenus("IIS201").Tag
        .mnuIIS202.Caption = mMenus("IIS202").Caption
        .mnuIIS202.Tag = mMenus("IIS202").Tag
        
        '## ������/��ȸ
        .mnuIIS300.Visible = mMenus("IIS300").Visible
        .mnuIIS301.Visible = mMenus("IIS301").Visible
        
        '## Manger
        .mnuIIS600.Visible = mMenus("IIS600").Visible
        .NODE3.Visible = mMenus("NODE3").Visible
        .mnuIIS609.Visible = mMenus("IIS609").Visible
        .mnuIIS610.Visible = mMenus("IIS610").Visible
        .mnuIIS611.Visible = mMenus("IIS611").Visible
        .mnuIIS612.Visible = mMenus("IIS612").Visible
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���� PC�� ������ ��� ������
'-----------------------------------------------------------------------------'
Private Sub GetMachine()
    Dim strAppName  As String
    Dim strM1EqpKey As String   '���1 EqpKey
    Dim strM1EqpCd  As String   '���1 EqpCd
    Dim strM2EqpKey As String   '���2 EqpKey
    Dim strM2EqpCd  As String   '���2 EqpCd
    
    strAppName = AppName
    
On Error Resume Next
    strM1EqpKey = GetSetting(strAppName, "Machine1", "EqpKey", "")
    strM1EqpCd = GetSetting(strAppName, "Machine1", "EqpCd", "")
    strM2EqpKey = GetSetting(strAppName, "Machine2", "EqpKey", "")
    strM2EqpCd = GetSetting(strAppName, "Machine2", "EqpCd", "")

    '## ������ ������� �ȵǾ� �ִ� ���
    If strM1EqpKey = "" And strM2EqpKey = "" Then
        mMenus("IIS200").Visible = False
        Exit Sub
    End If
    
    '## 1.4.1: (2005-04-13)
    '   - ������ ���Ű�� ��Ͽ� ���� ��� �������� �ʵ��� ����
    '## �˻����1 ����
    If strM1EqpKey = "" Then
        mMenus("IIS201").Visible = False
    Else
        mMenus("IIS201").Caption = strM1EqpKey
        mMenus("IIS201").Tag = strM1EqpCd
        If mMenus.Exist(strM1EqpKey) Then
            mMenus(strM1EqpKey).Visible = True
        Else
            mMenus("IIS201").Visible = False
        End If
    End If
    
    '## �˻����2 ����
    If strM2EqpKey = "" Then
        mMenus("IIS202").Visible = False
    Else
        mMenus("IIS202").Caption = strM2EqpKey
        mMenus("IIS202").Tag = strM2EqpCd
        If mMenus.Exist(strM2EqpKey) Then
            mMenus(strM2EqpKey).Visible = True
        Else
            mMenus("IIS202").Visible = False
        End If
    End If
End Sub

Private Sub Class_Initialize()
    Set mMenus = New clsIISMenus
    Set mHopFrm = New frmIISHospital
    Set mImgList = mHopFrm.imlHospital
    
    '## ����
    mMenus.Add "IIS000", True               '����
    
    '## ����������̽�
    mMenus.Add "IIS200", True               '����������̽�
    mMenus.Add "IIS201", True               '�˻����1
    mMenus.Add "IIS202", True               '�˻����2
    mMenus.Add "IISSEP02", True            '������02
    mMenus.Add "IIS204", True              '�˻��� ��ȸ
    
    '## ���
    mMenus.Add "Centaur", False
    mMenus.Add "Architect", False
    mMenus.Add "CobasE601", False
    mMenus.Add "DPC", False
    
    '## ������/��ȸ
    mMenus.Add "IIS300", False
    mMenus.Add "IIS301", True
    
    '## Manager
    mMenus.Add "IIS600", True               '�Ŵ���
    mMenus.Add "NODE3", True                '�˻���� ���� ����
    mMenus.Add "IIS609", True               '�˻���� ����
    mMenus.Add "IIS610", True               '�˻���� ��ż���
    mMenus.Add "IIS611", True               '�˻���� �˻��׸�
    mMenus.Add "IIS612", True               '�˻���� ����
    
    Call GetMachine
End Sub

Private Sub Class_Terminate()
    Set mMenus = Nothing
    If Not mHopFrm Is Nothing Then Unload mHopFrm
    Set mHopFrm = Nothing
    Set mImgList = Nothing
End Sub