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
'   ���ϸ�  : clsIISConst.cls
'   �ۼ���  : ������
'   ��  ��  : �����ڵ� Ŭ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

'## ������ ����
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

'-- 2007.08.02 ������ �߰� -----------------
Public Property Get IISREACTIVE() As String
    IISREACTIVE = mIISREACTIVE
End Property

Public Property Get IISNREACTIVE() As String
    IISNREACTIVE = mIISNREACTIVE
End Property

Public Property Get IISWPOSITIVE() As String
    IISWPOSITIVE = mIISWPOSITIVE
End Property
'-------------------------------------------

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

'   - ���Ѱ�, ���Ѱ� ������Ƽ �߰�
Public Property Get IISLMTLOW() As String
    IISLMTLOW = mIISLMTLOW
End Property

Public Property Get IISLMTHIGH() As String
    IISLMTHIGH = mIISLMTHIGH
End Property

'## �����ڵ� �ε���
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

'## ������ ����
Public Property Get MainFrm() As Object
    Set MainFrm = mMainFrm
End Property

Public Property Let MainFrm(ByRef vData As Object)
    Set mMainFrm = vData
End Property

Public Property Get StatusBar() As Object
    Set StatusBar = mMainFrm.sbrStatus
End Property

'## ��������
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

'## ����� ����
Public Property Get EMPID() As String
    EMPID = mEmpId
End Property

Public Property Get EMPNM() As String
    EMPNM = mEmpNm
End Property

'## ��ű�ȣ
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
'   ��� : ������Ʈ������ ������ �ε�
'   �μ� :
'       1.pAppName : ���α׷� �̸� (������Ʈ���� "Schweitzer IIS")
'       2.pAppPath : ���α׷� ���
'-----------------------------------------------------------------------------'
Public Function GetRegInfo(ByVal pAppName As String, ByVal pAppPath As String)
    Dim strTemp As String
    
On Error Resume Next
    strTemp = Mid$(pAppPath, 1, InStrRev(pAppPath, "\"))
    
    '## ������Ʈ������ ���� �ε�, ������ ����
    mAppName = pAppName
    mExePath = GetSetting(mAppName, "App", "Path", "")
    
    '## ������Ʈ���� ���α׷� ��ο� �����θ� ���Ͽ� Ʋ���� �����η� �缳��
    If mExePath <> (pAppPath & "\") Then
        mExePath = pAppPath & "\"
        SaveSetting mAppName, "App", "Path", mExePath
    End If
    
    '## ClientDb���+���ϸ� ����
    mClientDbPath = GetSetting(mAppName, "App", "ClientDb", "")
    If mClientDbPath = "" Then
        mClientDbPath = strTemp & "ClientDb\ClientDb.mdb"
        SaveSetting mAppName, "App", "ClientDb", mClientDbPath
    End If
    
    '## �α�����, INI���� ��� ����
    mIniPath = strTemp
    mLogPath = mExePath & "Log\"
End Function

'-----------------------------------------------------------------------------'
'   ��� : �ڵ������� �ε�
'-----------------------------------------------------------------------------'
Public Sub LoadCodeInfo()
    Dim objCom      As clsIISCommon
    Dim strFileNm   As String         'INI���� ���+���ϸ�
    
    '## INI���Ͽ��� �ڵ����� �ε�
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
        
        '-- 2007.08.02 ������ �߰�
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
'   ��� : ����� ������ ����
'   �μ� :
'       1.pEmpId   : ����ID
'       2.pEmpNm   : �����̸�
'-----------------------------------------------------------------------------'
Public Sub SetUserInfo(ByVal pEmpId As String, ByVal pEmpNm As String)
    '## ���̵�, �̸�, �׷������� ����
    mEmpId = pEmpId
    mEmpNm = pEmpNm
End Sub