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
'   파일명  : clsIISHopMenu.cls (전주예수병원)
'   작성자  :
'   내  용  : 병원별로 메뉴설정 클래스
'   작성일  : 2004-10-05
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Private mMenus      As clsIISMenus      '메뉴 컬렉션
Private mHopFrm     As frmIISHospital   '병원별 Imagelist를 가지는 폼
Private mImgList    As Object           'ImageList 정보

Public Property Get Menus() As clsIISMenus
    Set Menus = mMenus
End Property

Public Property Get ImgList() As Object
    Set ImgList = mImgList
End Property

'-----------------------------------------------------------------------------'
'   기능 : 메인폼의 풀다운메뉴 구성
'-----------------------------------------------------------------------------'
Public Sub GetFullMenu()
    '## 풀다운메뉴 항목은 기본적으로 Visible=True로 되었있음
    With MainFrm
        '## 장비 인터페이스
        .mnuIIS200.Visible = mMenus("IIS200").Visible
        .mnuIIS201.Visible = mMenus("IIS201").Visible
        .mnuIIS202.Visible = mMenus("IIS202").Visible
        .mnuIISSEP02.Visible = mMenus("IISSEP02").Visible
        .mnuIIS204.Visible = mMenus("IIS204").Visible
        
        .mnuIIS201.Caption = mMenus("IIS201").Caption
        .mnuIIS201.Tag = mMenus("IIS201").Tag
        .mnuIIS202.Caption = mMenus("IIS202").Caption
        .mnuIIS202.Tag = mMenus("IIS202").Tag
        
        '   - 결과등록 메뉴추가
        '## 결과등록/조회
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
'   기능 : 현재 PC에 설정된 장비를 보여줌
'-----------------------------------------------------------------------------'
Private Sub GetMachine()
    Dim strAppName  As String
    Dim strM1EqpKey As String   '장비1 EqpKey
    Dim strM1EqpCd  As String   '장비1 EqpCd
    Dim strM2EqpKey As String   '장비2 EqpKey
    Dim strM2EqpCd  As String   '장비2 EqpCd
    
    strAppName = AppName
    
On Error Resume Next
    strM1EqpKey = GetSetting(strAppName, "Machine1", "EqpKey", "")
    strM1EqpCd = GetSetting(strAppName, "Machine1", "EqpCd", "")
    strM2EqpKey = GetSetting(strAppName, "Machine2", "EqpKey", "")
    strM2EqpCd = GetSetting(strAppName, "Machine2", "EqpCd", "")

    '## 레지에 장비설정이 안되어 있는 경우
    If strM1EqpKey = "" And strM2EqpKey = "" Then
        mMenus("IIS200").Visible = False
        Exit Sub
    End If
    
    '   - 레지의 장비키가 목록에 없는 경우 보여주지 않도록 수정
    '## 검사장비1 설정
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
    
    '## 검사장비2 설정
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
    
    '## 공통
    mMenus.Add "IIS000", True               '종료
    
    '## 장비인터페이스
    mMenus.Add "IIS200", True               '장비인터페이스
    mMenus.Add "IIS201", True               '검사장비1
    mMenus.Add "IIS202", True               '검사장비2
    mMenus.Add "IISSEP02", True            '구분자02
    mMenus.Add "IIS204", True              '검사대상 조회
    
    '## 장비
    mMenus.Add "Uriscan Pro-1", False
    mMenus.Add "Uriscan Pro-2", False
    mMenus.Add "Dimension RXL", False
    mMenus.Add "RapidLab 865", False
    mMenus.Add "Stks-1", False
    mMenus.Add "Stks-2", False
    mMenus.Add "Hitachi 7600", False
    mMenus.Add "Axsym", False
    mMenus.Add "SE-9000", False
    mMenus.Add "Variant II", False
    mMenus.Add "CX3 Delta", False
    mMenus.Add "LPIA-NV7", False
    mMenus.Add "Thrombolyzer compact", False
    mMenus.Add "Vitek", False
    mMenus.Add "RapidLab 850", False
    mMenus.Add "RapidLab 860", False
    mMenus.Add "Vitek II", False
    mMenus.Add "Thrombolyzer RackRotor", False
    mMenus.Add "XT-1800i", False
    mMenus.Add "CA-500", False
    mMenus.Add "CA-1500", False
    
    '## By M.G.Choi 2006.01.06
    '   CA-1500 장비 응급검사실 추가 됨
    mMenus.Add "CA-1500ER", False
    
    '## 오세원 Add
    mMenus.Add "D-RXLM", False
    mMenus.Add "Architect", False
    mMenus.Add "ESR", False
    mMenus.Add "PCX", False
    mMenus.Add "Gem3000", False
    
    '## 오세원 Add
    mMenus.Add "D-RXLM2", False
    
    '## 오세원 Add
    mMenus.Add "XE2100", False
    
    '## 오세원 Add
    mMenus.Add "NS1000", False
    
    '## MAJESTIC ADD 2011.01.28
    '   XT-2000i 장비 응급검사실 추가 됨
    mMenus.Add "XT-2000i", False
    
    '## MAJESTIC ADD 2012.01.11
    '   DxC800 장비 추가 됨
    mMenus.Add "DXC800", False
    mMenus.Add "ChorusTRIO", False
    
    mMenus.Add "RapidLab348-1", False
    mMenus.Add "RapidLab348-2", False
    mMenus.Add "RapidLab348-3", False
    mMenus.Add "RapidLab348-4", False
    
    mMenus.Add "Hitachi7600-1", False
    mMenus.Add "Hitachi7180", False
    mMenus.Add "Hitachi7180-1", False
    
    mMenus.Add "MultiReader", False
    
    mMenus.Add "IT1000", False
    
    mMenus.Add "Gemini", False
    
    '## 2013.07.18 오세원 Add
    mMenus.Add "Vitek 3", False
    mMenus.Add "Vitek 4", False
    
    '## 2013.12.11 조남식 Add
    '## 2013.12.11 오세원 Add
    mMenus.Add "UF1000i", False
    mMenus.Add "Coa1", False
    mMenus.Add "CoaER", False
    
    '## 2014.01.21 오세원 Add
    mMenus.Add "CobasE411", False
    
    '## 2014.01.21 오세원 Add
    mMenus.Add "GEM3500_1", False
    mMenus.Add "GEM3500_2", False
    
    '## 2014.12.29 오세원 Add
    mMenus.Add "Taqman", False
    mMenus.Add "Triage", False
    mMenus.Add "Sofia", False
    
    '## 2015.01.28 오세원 Add
    mMenus.Add "Sofia2", False
    
    '## 2015.05.28 오세원 Add
    mMenus.Add "GeneXpert", False
    
    '## 2015.08.25 오세원 Add
    mMenus.Add "B4C", False
    
    '## 2016.03.10 오세원 Add
    mMenus.Add "NSPRIME", False
    
    '## 2016.09.06 오세원 Add
    mMenus.Add "Hitachi7600-2", False
    
    '## 2016.11.29 오세원 Add
    mMenus.Add "CobasC513", False
    
    '## 2016.12.19 오세원 Add
    mMenus.Add "Sofia3", False
    
    '## 2016.12.22 오세원 Add
    mMenus.Add "MEDIWISS", False
    
    '## 2016.12.22 오세원 Add
    mMenus.Add "CFX96", False
    
    '## 2017.04.05 오세원 Add
    mMenus.Add "IH500", False
    
    '## 2017.04.25 오세원 Add
    mMenus.Add "Uriscan Pro-3", False
    mMenus.Add "AlloStation", False
    
    '## 2017.09.18 오세원 Add
    mMenus.Add "CFX96_RV16", False
    
    '## 2018.01.08 오세원 Add
    mMenus.Add "LIAISON", False
    
    '## 2018.02.08 오세원 Add
    mMenus.Add "CobasE411_Disk", False
    
    '## 2018.11.15 오세원 Add
    mMenus.Add "UF5000", False
    
    '## 2019.01.07 오세원 Add
    mMenus.Add "GEM5000", False
    
    '## 2019.09.19 오세원 Add
    mMenus.Add "GEM5000_2", False
    
    '## 2019.10.24 오세원 Add
    mMenus.Add "OCSENSORM", False
    
    '## 2019.12.19 오세원 Add
    mMenus.Add "ALINITY", False
    
    '## 2020.10.19 오세원 Add
    mMenus.Add "GEM5000_OP", False
    
    '## 2020.10.28 오세원 Add
    mMenus.Add "LABOSPECT_1", False
    mMenus.Add "LABOSPECT_2", False
    mMenus.Add "LABOSPECT_ER", False
    
    '## 2021.01.12 오세원 Add
    mMenus.Add "PATHFAST", False
    
    '## 1.5.1: 이상대(2005-07-22)
    '   - 결과등록 메뉴추가
    '## 결과등록/조회
    mMenus.Add "IIS300", False
    mMenus.Add "IIS301", True
    
    '## Manager
    mMenus.Add "IIS600", True               '매니저
    mMenus.Add "NODE3", True                '검사장비 관련 설정
    mMenus.Add "IIS609", True               '검사장비 관리
    mMenus.Add "IIS610", True               '검사장비 통신설정
    mMenus.Add "IIS611", True               '검사장비 검사항목
    mMenus.Add "IIS612", True               '검사장비 선택
    
    Call GetMachine
End Sub

Private Sub Class_Terminate()
    Set mMenus = Nothing
    If Not mHopFrm Is Nothing Then Unload mHopFrm
    Set mHopFrm = Nothing
    Set mImgList = Nothing
End Sub
