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
'   파일명  : clsIISMenuInfo.cls
'   작성자  :
'   내  용  : 메인폼의 툴바의 메뉴구성 클래스
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private WithEvents mToolbar As frmIISToolbar    '툴바 구성폼
Attribute mToolbar.VB_VarHelpID = -1

'-----------------------------------------------------------------------------'
'   기능 : 툴바를 구성하는 폼을표시
'-----------------------------------------------------------------------------'
Public Function ConfigToolbar()
    '## 툴바구성 폼을 표시
    Set mToolbar = New frmIISToolbar
    mToolbar.Show vbModal
    
    Unload mToolbar
    Set mToolbar = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 레지에서 툴바 구성정보를 읽어 툴바구성
'-----------------------------------------------------------------------------'
Public Sub GetToolbar()
    Dim objChoice   As clsIISEqpChoice  '사용장비 선택 클래스
    Dim objHop      As clsIISHopMenu    '병원별 메뉴 클래스
    Dim imlHospital As ImageList        '병원 ImageList
    Dim strKey      As String           '장비키
    Dim strCaption  As String           '툴바 Caption
    Dim strTooltip  As String           '툴바 Tooltip
    Dim strAppName  As String           '레지의 Application Name
    Dim strSection  As String           '레지의 Section
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
        '## 1.툴바 초기화
        .tbrToolbar.ImageList = Nothing
        .tbrToolbar.Buttons.Clear
        .imlToolbar.ListImages.Clear
        .imlToolbar.ImageWidth = 32
        .imlToolbar.ImageHeight = 32
        
        '   - 레지의 장비키가 등록되지 않은경우 보여주지 않도록 수정
        '## 2.장비 아이콘 로드
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
        
        '## 3.공통아이콘 로드
        strAppName = AppName
        strSection = "User\Default"
        Count = CLng(GetSetting(strAppName, strSection, "Count", "-1"))
        If Count = -1 Then GoTo EndLine
        
        '## 3.1.공통아이콘의 첫요소가 구분자가 아니면 구분자 추가
        strTemp = GetSetting(strAppName, strSection, "01", "")
        strKey = mGetP(strTemp, 1, ",")
        If strKey <> "IISSEP" And i <> 1 Then
            .tbrToolbar.Buttons.Add i, , , tbrSeparator
            i = i + 1
        End If
        
        '## 3.2.공통아이콘 로드
        For k = 1 To Count
            strTemp = GetSetting(strAppName, strSection, Format$(k, "00"), "")
            strKey = mGetP(strTemp, 1, ",")
            
            If strKey = "IISSEP" Then   '## 구분자
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
'   기능 : 해당 장비키의 아이콘 인덱스를 조회
'   인수 :
'       - pEqpKey : 장비키
'   반환 : 아이콘 인덱스
'-----------------------------------------------------------------------------'
Public Function GetIconIndex(ByVal pEqpKey As String) As Long
    Dim objHop      As clsIISHopMenu    '병원별 메뉴 클래스
    Dim imgImage    As ListImage        'ImageList의 Image
    
    Set objHop = New clsIISHopMenu
    
    '## 장비 아이콘 ImageList에서 해당 장비키로 조회
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
'   기능 : 툴바폼의 저장 이벤트처리, 레지에 아이콘 구성정보 저장
'-----------------------------------------------------------------------------'
Private Sub mToolbar_Save(pListBox As Object)
    Dim strAppName  As String           '레지의 AppName
    Dim strSection  As String           '레지의 Section
    Dim lngCnt      As Long             '-1:해당 사용자 정보가 없음
    Dim i           As Long
    
    strAppName = AppName
    strSection = "User\Default"
    
On Error Resume Next
    '## 해당 사용자의 레지정보가 있으면 지우고 다시입력
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
