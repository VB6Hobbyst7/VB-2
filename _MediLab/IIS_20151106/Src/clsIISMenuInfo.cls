VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMenuInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : clsIISMenuInfo.cls
'   �ۼ���  : ������
'   ��  ��  : �������� ������ �޴����� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Private WithEvents mToolbar As frmIISToolbar    '���� ������
Attribute mToolbar.VB_VarHelpID = -1

'-----------------------------------------------------------------------------'
'   ��� : ���ٸ� �����ϴ� ����ǥ��
'-----------------------------------------------------------------------------'
Public Function ConfigToolbar()
    '## ���ٱ��� ���� ǥ��
    Set mToolbar = New frmIISToolbar
    mToolbar.Show vbModal
    
    Unload mToolbar
    Set mToolbar = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ���� ���������� �о� ���ٱ���
'-----------------------------------------------------------------------------'
Public Sub GetToolbar()
    Dim objChoice   As clsIISEqpChoice  '������ ���� Ŭ����
    Dim objHop      As clsIISHopMenu    '������ �޴� Ŭ����
    Dim imlHospital As ImageList        '���� ImageList
    Dim strKey      As String           '���Ű
    Dim strCaption  As String           '���� Caption
    Dim strTooltip  As String           '���� Tooltip
    Dim strAppName  As String           '������ Application Name
    Dim strSection  As String           '������ Section
    Dim strTemp     As String
    Dim Index       As Long
    Dim Count       As Long
    Dim i           As Long             'Toolbar Index
    Dim j           As Long             'ImageList Index
    Dim k           As Long
    
    Set objChoice = New clsIISEqpChoice
    Set objHop = New clsIISHopMenu
    Set imlHospital = objHop.ImgList
    
    i = 1:  j = 1
    With mdiIISMain
        '## 1.���� �ʱ�ȭ
        .tbrToolbar.ImageList = Nothing
        .tbrToolbar.Buttons.Clear
        .imlToolbar.ListImages.Clear
        .imlToolbar.ImageWidth = 32
        .imlToolbar.ImageHeight = 32
        
        '   - ������ ���Ű�� ��ϵ��� ������� �������� �ʵ��� ����
        '## 2.��� ������ �ε�
        Call objChoice.GetEqp
        If objChoice.EqpCd1 <> "" Then
            strKey = objChoice.EqpKey1
            Index = GetIconIndex(strKey)
            If Index <> 0 Then
                strTemp = imlHospital.ListImages(Index).Tag
                strCaption = mGetP(strTemp, 1, ",")
                strTooltip = mGetP(strTemp, 2, ",")
                .imlToolbar.ListImages.Add j, strKey, imlHospital.ListImages(Index).Picture
                
                .tbrToolbar.ImageList = .imlToolbar
                .tbrToolbar.Buttons.Add i, strKey, strCaption, , j
                .tbrToolbar.Buttons(i).ToolTipText = strTooltip
                .tbrToolbar.Buttons(i).Tag = strKey & ",H," & CStr(Index)
                i = i + 1:  j = j + 1
            End If
        End If
        
        If objChoice.EqpCd2 <> "" Then
            strKey = objChoice.EqpKey2
            Index = GetIconIndex(strKey)
            If Index <> 0 Then
                strTemp = imlHospital.ListImages(Index).Tag
                strCaption = mGetP(strTemp, 1, ",")
                strTooltip = mGetP(strTemp, 2, ",")
                .imlToolbar.ListImages.Add j, strKey, imlHospital.ListImages(Index).Picture
                
                .tbrToolbar.ImageList = .imlToolbar
                .tbrToolbar.Buttons.Add i, strKey, strCaption, , j
                .tbrToolbar.Buttons(i).ToolTipText = strTooltip
                .tbrToolbar.Buttons(i).Tag = strKey & ",H," & CStr(Index)
                i = i + 1:  j = j + 1
            End If
        End If
        Set objChoice = Nothing
        
        '## 3.��������� �ε�
        strAppName = AppName
        strSection = "User\Default"
        Count = CLng(GetSetting(strAppName, strSection, "Count", "-1"))
        If Count = -1 Then GoTo EndLine
        
        '## 3.1.����������� ù��Ұ� �����ڰ� �ƴϸ� ������ �߰�
        strTemp = GetSetting(strAppName, strSection, "01", "")
        strKey = mGetP(strTemp, 1, ",")
        If strKey <> "IISSEP" And i <> 1 Then
            .tbrToolbar.Buttons.Add i, , , tbrSeparator
            i = i + 1
        End If
        
        '## 3.2.��������� �ε�
        For k = 1 To Count
            strTemp = GetSetting(strAppName, strSection, Format$(k, "00"), "")
            strKey = mGetP(strTemp, 1, ",")
            
            If strKey = "IISSEP" Then   '## ������
                .tbrToolbar.Buttons.Add i, , , tbrSeparator
                i = i + 1
            Else
                If objHop.Menus(strKey).Visible = True Then
                    Index = CLng(mGetP(strTemp, 3, ","))
                    .imlToolbar.ListImages.Add j, strKey, .imlCommon.ListImages(Index).Picture
                    strCaption = mGetP(.imlCommon.ListImages(Index).Tag, 1, ",")
                    strTooltip = mGetP(.imlCommon.ListImages(Index).Tag, 2, ",")
                    
                    .tbrToolbar.ImageList = .imlToolbar
                    .tbrToolbar.Buttons.Add i, strKey, strCaption, , j
                    .tbrToolbar.Buttons(i).ToolTipText = strTooltip
                    .tbrToolbar.Buttons(i).Tag = strKey & ",C," & CStr(Index)
                    i = i + 1:  j = j + 1
                End If
            End If
        Next k
    End With
    
EndLine:
    Set imlHospital = Nothing
    Set objHop = Nothing
    Set objChoice = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���Ű�� ������ �ε����� ��ȸ
'   �μ� :
'       - pEqpKey : ���Ű
'   ��ȯ : ������ �ε���
'-----------------------------------------------------------------------------'
Public Function GetIconIndex(ByVal pEqpKey As String) As Long
    Dim objHop      As clsIISHopMenu    '������ �޴� Ŭ����
    Dim imgImage    As ListImage        'ImageList�� Image
    
    Set objHop = New clsIISHopMenu
    
    '## ��� ������ ImageList���� �ش� ���Ű�� ��ȸ
    For Each imgImage In objHop.ImgList.ListImages
        If imgImage.Key = pEqpKey Then
            GetIconIndex = imgImage.Index
            Exit For
        End If
    Next
    
    Set imgImage = Nothing
    Set objHop = Nothing
End Function

'-----------------------------------------------------------------------------'
'   ��� : �������� ���� �̺�Ʈó��, ������ ������ �������� ����
'-----------------------------------------------------------------------------'
Private Sub mToolbar_Save(pListBox As Object)
    Dim strAppName  As String           '������ AppName
    Dim strSection  As String           '������ Section
    Dim lngCnt      As Long             '-1:�ش� ����� ������ ����
    Dim i           As Long
    
    strAppName = AppName
    strSection = "User\Default"
    
On Error Resume Next
    '## �ش� ������� ���������� ������ ����� �ٽ��Է�
    lngCnt = CLng(GetSetting(strAppName, strSection, "Count", "-1"))
    If lngCnt <> -1 Then
        Call DeleteSetting(strAppName, strSection)
    End If
    
    For i = 1 To pListBox.ListCount
        Call SaveSetting(strAppName, strSection, Format$(i, "00"), Trim(mGetP(pListBox.List(i - 1), 2, Space(50))))
    Next i
    
    Call SaveSetting(strAppName, strSection, "Count", CStr(i - 1))
    Call GetToolbar
End Sub