VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISConst"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISConst.cls
'   내  용  : 공통코드 클래스
'   버  전  :
'          - 미생물 일반감수성(IISGENSENSI), MIC(IISMIC) 결과유형 변수추가
'          - PT(%)결과코드(IISPT) 변수추가
'          - 검사항목별 결과코드(CRESULTCD) 변수 추가
'          - CBC Workarea(mIISCBCWA) 변수추가
'          - 하한값(IISLMTLOW), 상한값(IISLMTHIGH) 결과코드 프로퍼티 추가
'-----------------------------------------------------------------------------'

Option Explicit

'## 병원별 정보
Public Property Get PROJECTCODE() As String
    PROJECTCODE = mPROJECTCODE
End Property

Public Property Get PROJECTTYPE() As String
    PROJECTTYPE = mPROJECTTYPE
End Property

Public Property Get HOSPITALNM() As String
    HOSPITALNM = mHOSPITALNM
End Property

Public Property Get SPCLEN() As Long
    SPCLEN = mSPCLEN
End Property

Public Property Get SPCYYLEN() As Long
    SPCYYLEN = mSPCYYLEN
End Property

Public Property Get SPCNOLEN() As Long
    SPCNOLEN = mSPCNOLEN
End Property

Public Property Get IISNEGATIVE() As String
    IISNEGATIVE = mIISNEGATIVE
End Property

Public Property Get IISPOSITIVE() As String
    IISPOSITIVE = mIISPOSITIVE
End Property

Public Property Get IISGRAYZONE() As String
    IISGRAYZONE = mIISGRAYZONE
End Property

Public Property Get IISREACTIVE() As String
    IISREACTIVE = mIISREACTIVE
End Property

Public Property Get IISNREACTIVE() As String
    IISNREACTIVE = mIISNREACTIVE
End Property

Public Property Get IISWPOSITIVE() As String
    IISWPOSITIVE = mIISWPOSITIVE
End Property

Public Property Get IISERROR() As String
    IISERROR = mIISERROR
End Property

Public Property Get IISSPCLEN() As Long
    IISSPCLEN = mIISSPCLEN
End Property

Public Property Get IISSPCSERUM() As String
    IISSPCSERUM = mIISSPCSERUM
End Property

Public Property Get IISSPCURINE() As String
    IISSPCURINE = mIISSPCURINE
End Property

Public Property Get IISSPCPLASMA() As String
    IISSPCPLASMA = mIISSPCPLASMA
End Property

Public Property Get IISSPCCSF() As String
    IISSPCCSF = mIISSPCCSF
End Property

Public Property Get IISSPCBLOOD() As String
    IISSPCBLOOD = mIISSPCBLOOD
End Property

Public Property Get IISSPCFLUID() As String
    IISSPCFLUID = mIISSPCFLUID
End Property

Public Property Get IISSPCCAPD() As String
    IISSPCCAPD = mIISSPCCAPD
End Property

Public Property Get IISQCLOW() As String
    IISQCLOW = mIISQCLOW
End Property

Public Property Get IISQCNORMAL() As String
    IISQCNORMAL = mIISQCNORMAL
End Property

Public Property Get IISQCHIGH() As String
    IISQCHIGH = mIISQCHIGH
End Property

Public Property Get IISPANICCHECK() As String
    IISPANICCHECK = mIISPANICCHECK
End Property

Public Property Get IISMICWA() As String
    IISMICWA = mIISMICWA
End Property

Public Property Get IISMQTCD() As String
    IISMQTCD = mIISMQTCD
End Property

Public Property Get IISGENSENSI() As String
    IISGENSENSI = mIISGENSENSI
End Property

Public Property Get IISMIC() As String
    IISMIC = mIISMIC
End Property

Public Property Get IISPT() As String
    IISPT = mIISPT
End Property

Public Property Get IISCBCWA() As String
    IISCBCWA = mIISCBCWA
End Property

Public Property Get IISLMTLOW() As String
    IISLMTLOW = mIISLMTLOW
End Property

Public Property Get IISLMTHIGH() As String
    IISLMTHIGH = mIISLMTHIGH
End Property

'## 공통코드 인덱스
Public Property Get CODE() As String
    CODE = mCODE
End Property

Public Property Get CSPCCD() As String
    CSPCCD = mCSPCCD
End Property

Public Property Get CWACD() As String
    CWACD = mCWACD
End Property

Public Property Get CDETAILCD() As String
    CDETAILCD = mCDETAILCD
End Property

Public Property Get CPANELCD() As String
    CPANELCD = mCPANELCD
End Property

Public Property Get CREPEATCD() As String
    CREPEATCD = mCREPEATCD
End Property

Public Property Get CLOCATIONCD() As String
    CLOCATIONCD = mCLOCATIONCD
End Property

Public Property Get CVANDCD() As String
    CVANDCD = mCVANDCD
End Property

Public Property Get CFOOTNOTECD() As String
    CFOOTNOTECD = mCFOOTNOTECD
End Property

Public Property Get CSPCRMKCD() As String
    CSPCRMKCD = mCSPCRMKCD
End Property

Public Property Get CACCRSNCD() As String
    CACCRSNCD = mCACCRSNCD
End Property

Public Property Get CMDYRSNCD() As String
    CMDYRSNCD = mCMDYRSNCD
End Property

Public Property Get CQCREJECTCD() As String
    CQCREJECTCD = mCQCREJECTCD
End Property

Public Property Get CMNMCD() As String
    CMNMCD = mCMnmCd
End Property

Public Property Get CRESULTCD() As String
    CRESULTCD = mCRESULTCD
End Property

'## 메인폼 정보
Public Property Get MainFrm() As Object
    Set MainFrm = mMainFrm
End Property

Public Property Let MainFrm(ByRef vData As Object)
    Set mMainFrm = vData
End Property

Public Property Get StatusBar() As Object
    Set StatusBar = mMainFrm.sbrStatus
End Property

'## 레지정보
Public Property Get AppName() As String
    AppName = mAppName
End Property

Public Property Get ExePath() As String
    ExePath = mExePath
End Property

Public Property Get LogPath() As String
    LogPath = mLogPath
End Property

Public Property Get ClientDbPath() As String
    ClientDbPath = mClientDbPath
End Property

Public Property Get IniPath() As String
    IniPath = mIniPath
End Property

'## 사용자 정보
Public Property Get EMPID() As String
    EMPID = mEmpId
End Property

Public Property Get EMPNM() As String
    EMPNM = mEmpNm
End Property

'## 통신기호
Public Property Get ENQ() As String
    ENQ = Chr(mENQ)
End Property

Public Property Get ACK() As String
    ACK = Chr(mACK)
End Property

Public Property Get STX() As String
    STX = Chr(mSTX)
End Property

Public Property Get ETB() As String
    ETB = Chr(mETB)
End Property

Public Property Get ETX() As String
    ETX = Chr(mETX)
End Property

Public Property Get EOT() As String
    EOT = Chr(mEOT)
End Property

Public Property Get NAK() As String
    NAK = Chr(mNAK)
End Property

Public Property Get SOH() As String
    SOH = Chr(mSOH)
End Property

Public Property Get DLE() As String
    DLE = Chr(mDLE)
End Property

Public Property Get SYN() As String
    SYN = Chr(mSYN)
End Property

'-----------------------------------------------------------------------------'
'   기능 : 레지스트리에서 정보를 로드
'   인수 :
'       1.pAppName : 프로그램 이름 (레지스트리의 "Schweitzer IIS")
'       2.pAppPath : 프로그램 경로
'-----------------------------------------------------------------------------'
Public Function GetRegInfo(ByVal pAppName As String, ByVal pAppPath As String)
    Dim strTemp As String
    
On Error Resume Next
    strTemp = Mid$(pAppPath, 1, InStrRev(pAppPath, "\"))
    
    '## 레지스트리에서 정보 로드, 없으면 설정
    mAppName = pAppName
    mExePath = GetSetting(mAppName, "App", "Path", "")
    
    '## 레지스트리의 프로그램 경로와 현재경로를 비교하여 틀리면 현재경로로 재설정
    If mExePath <> (pAppPath & "\") Then
        mExePath = pAppPath & "\"
        SaveSetting mAppName, "App", "Path", mExePath
    End If
    
    '## ClientDb경로+파일명 설정
    mClientDbPath = GetSetting(mAppName, "App", "ClientDb", "")
    If mClientDbPath = "" Then
        mClientDbPath = strTemp & "ClientDb\ClientDb.mdb"
        SaveSetting mAppName, "App", "ClientDb", mClientDbPath
    End If
    
    '## 로그파일, INI파일 경로 설정
    mIniPath = strTemp
    mLogPath = mExePath & "Log\"
End Function

'-----------------------------------------------------------------------------'
'   기능 : 코드정보를 로드
'-----------------------------------------------------------------------------'
Public Sub LoadCodeInfo()
    Dim objCom      As clsIISCommon
    Dim strFileNm   As String         'INI파일 경로+파일명
    
    '## INI파일에서 코드정보 로드
    strFileNm = mIniPath & "IIS.ini"
    Set objCom = New clsIISCommon
    With objCom
        mPROJECTCODE = .mReadINI(strFileNm, "HOSPITAL", "CODE")
        mPROJECTTYPE = Mid$(mPROJECTCODE, 1, 1)
        mHOSPITALNM = .mReadINI(strFileNm, "HOSPITAL", "NAME")
        mSPCYYLEN = CLng(.mReadINI(strFileNm, "HOSPITAL", "SPCYYLEN"))
        mSPCNOLEN = CLng(.mReadINI(strFileNm, "HOSPITAL", "SPCNOLEN"))
        mSPCLEN = mSPCYYLEN + mSPCNOLEN
        mIISNEGATIVE = .mReadINI(strFileNm, "HOSPITAL", "IISNEGATIVE")
        mIISPOSITIVE = .mReadINI(strFileNm, "HOSPITAL", "IISPOSITIVE")
        mIISGRAYZONE = .mReadINI(strFileNm, "HOSPITAL", "IISGRAYZONE")
        mIISERROR = .mReadINI(strFileNm, "HOSPITAL", "IISERROR")
        mIISSPCLEN = CLng(.mReadINI(strFileNm, "HOSPITAL", "IISSPCLEN"))
        mIISSPCSERUM = .mReadINI(strFileNm, "HOSPITAL", "IISSPCSERUM")
        mIISSPCURINE = .mReadINI(strFileNm, "HOSPITAL", "IISSPCURINE")
        mIISSPCPLASMA = .mReadINI(strFileNm, "HOSPITAL", "IISSPCPLASMA")
        mIISSPCCSF = .mReadINI(strFileNm, "HOSPITAL", "IISSPCCSF")
        mIISSPCBLOOD = .mReadINI(strFileNm, "HOSPITAL", "IISSPCBLOOD")
        mIISSPCFLUID = .mReadINI(strFileNm, "HOSPITAL", "IISSPCFLUID")
        mIISSPCCAPD = .mReadINI(strFileNm, "HOSPITAL", "IISSPCCAPD")
        mIISQCLOW = .mReadINI(strFileNm, "HOSPITAL", "IISQCLOW")
        mIISQCNORMAL = .mReadINI(strFileNm, "HOSPITAL", "IISQCNORMAL")
        mIISQCHIGH = .mReadINI(strFileNm, "HOSPITAL", "IISQCHIGH")
        mIISPANICCHECK = .mReadINI(strFileNm, "HOSPITAL", "IISPANICCHECK")
        mIISMICWA = .mReadINI(strFileNm, "HOSPITAL", "IISMICWA")
        mIISMQTCD = .mReadINI(strFileNm, "HOSPITAL", "IISMQTCD")
        mIISGENSENSI = .mReadINI(strFileNm, "HOSPITAL", "IISGENSENSI")
        mIISMIC = .mReadINI(strFileNm, "HOSPITAL", "IISMIC")
        mIISPT = .mReadINI(strFileNm, "HOSPITAL", "IISPT")
        mIISCBCWA = .mReadINI(strFileNm, "HOSPITAL", "IISCBCWA")
        mIISLMTLOW = .mReadINI(strFileNm, "HOSPITAL", "IISLMTLOW")
        mIISLMTHIGH = .mReadINI(strFileNm, "HOSPITAL", "IISLMTHIGH")
        
        mIISREACTIVE = .mReadINI(strFileNm, "HOSPITAL", "IISREACTIVE")
        mIISNREACTIVE = .mReadINI(strFileNm, "HOSPITAL", "IISNREACTIVE")
        mIISWPOSITIVE = .mReadINI(strFileNm, "HOSPITAL", "IISWPOSITIVE")

        mCODE = .mReadINI(strFileNm, "CODE", "CODE")
        mCSPCCD = .mReadINI(strFileNm, "CODE", "SPCCD")
        mCWACD = .mReadINI(strFileNm, "CODE", "WACD")
        mCDETAILCD = .mReadINI(strFileNm, "CODE", "DETAILCD")
        mCPANELCD = .mReadINI(strFileNm, "CODE", "PANELCD")
        mCREPEATCD = .mReadINI(strFileNm, "CODE", "REPEATCD")
        mCLOCATIONCD = .mReadINI(strFileNm, "CODE", "LOCATIONCD")
        mCVANDCD = .mReadINI(strFileNm, "CODE", "VANDCD")
        mCFOOTNOTECD = .mReadINI(strFileNm, "CODE", "FOOTNOTECD")
        mCSPCRMKCD = .mReadINI(strFileNm, "CODE", "SPCRMKCD")
        mCACCRSNCD = .mReadINI(strFileNm, "CODE", "ACCRSNCD")
        mCMDYRSNCD = .mReadINI(strFileNm, "CODE", "MDYRSNCD")
        mCQCREJECTCD = .mReadINI(strFileNm, "CODE", "QCREJECTCD")
        mCMnmCd = .mReadINI(strFileNm, "CODE", "MNMCD")
        mCRESULTCD = .mReadINI(strFileNm, "CODE", "RESULTCD")
    End With
    Set objCom = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 사용자 정보를 설정
'   인수 :
'       1.pEmpId   : 직원ID
'       2.pEmpNm   : 직원이름
'-----------------------------------------------------------------------------'
Public Sub SetUserInfo(ByVal pEmpId As String, ByVal pEmpNm As String)
    '## 아이디, 이름, 그룹정보를 설정
    mEmpId = pEmpId
    mEmpNm = pEmpNm
End Sub
