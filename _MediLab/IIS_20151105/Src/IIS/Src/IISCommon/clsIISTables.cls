VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISTables"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 '-----------------------------------------------------------------------------'
'   파일명  : clsIISTables.cls
'   작성자  : 오세원
'   내  용  : INI파일에 테이블, 필드명을 로드, 초기화하는 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'
Option Explicit

Public Property Get THIS001() As String
    THIS001 = mTHIS001
End Property

Public Property Get THIS002() As String
    THIS002 = mTHIS002
End Property

Public Property Get THIS003() As String
    THIS003 = mTHIS003
End Property

Public Property Get THIS004() As String
    THIS004 = mTHIS004
End Property

Public Property Get THIS005() As String
    THIS005 = mTHIS005
End Property

Public Property Get THIS006() As String
    THIS006 = mTHIS006
End Property

Public Property Get TCOM001() As String
    TCOM001 = mTCOM001
End Property

Public Property Get TIIS001() As String
    TIIS001 = mTIIS001
End Property

Public Property Get TIIS002() As String
    TIIS002 = mTIIS002
End Property

Public Property Get TIIS003() As String
    TIIS003 = mTIIS003
End Property

Public Property Get TIIS101() As String
    TIIS101 = mTIIS101
End Property

Public Property Get TIIS102() As String
    TIIS102 = mTIIS102
End Property

Public Property Get TIIS103() As String
    TIIS103 = mTIIS103
End Property

Public Property Get TIIS104() As String
    TIIS104 = mTIIS104
End Property

Public Property Get TIIS105() As String
    TIIS105 = mTIIS105
End Property

Public Property Get TIIS106() As String
    TIIS106 = mTIIS106
End Property

Public Property Get TIIS107() As String
    TIIS107 = mTIIS107
End Property

Public Property Get TIIS201() As String
    TIIS201 = mTIIS201
End Property

Public Property Get TIIS202() As String
    TIIS202 = mTIIS202
End Property

Public Property Get TIIS203() As String
    TIIS203 = mTIIS203
End Property

Public Property Get TIIS204() As String
    TIIS204 = mTIIS204
End Property

Public Property Get TIIS205() As String
    TIIS205 = mTIIS205
End Property

Public Property Get TIIS206() As String
    TIIS206 = mTIIS206
End Property

Public Property Get TIIS301() As String
    TIIS301 = mTIIS301
End Property

Public Property Get TIIS302() As String
    TIIS302 = mTIIS302
End Property

Public Property Get TIIS303() As String
    TIIS303 = mTIIS303
End Property

Public Property Get TIIS304() As String
    TIIS304 = mTIIS304
End Property

Public Property Get TIIS305() As String
    TIIS305 = mTIIS305
End Property

Public Property Get TIIS306() As String
    TIIS306 = mTIIS306
End Property

Public Property Get TIIS307() As String
    TIIS307 = mTIIS307
End Property

Public Property Get TIIS308() As String
    TIIS308 = mTIIS308
End Property

Public Property Get TIIS401() As String
    TIIS401 = mTIIS401
End Property

Public Property Get TIIS402() As String
    TIIS402 = mTIIS402
End Property

Public Property Get TIIS403() As String
    TIIS403 = mTIIS403
End Property

Public Property Get TIIS404() As String
    TIIS404 = mTIIS404
End Property

Public Property Get TIIS405() As String
    TIIS405 = mTIIS405
End Property

Public Property Get TIIS406() As String
    TIIS406 = mTIIS406
End Property

Public Property Get TIIS501() As String
    TIIS501 = mTIIS501
End Property

Public Property Get TIIS502() As String
    TIIS502 = mTIIS502
End Property

Public Property Get TIIS503() As String
    TIIS503 = mTIIS503
End Property

Public Property Get TIIS504() As String
    TIIS504 = mTIIS504
End Property

Public Property Get TIIS505() As String
    TIIS505 = mTIIS505
End Property

'   - 특수검사 결과내역 테이블 프로퍼티 추가
Public Property Get TIIS601() As String
    TIIS601 = mTIIS601
End Property

Public Property Get TACC203() As String
    TACC203 = mTACC203
End Property

Public Property Get TACC204() As String
    TACC204 = mTACC204
End Property

Public Property Get FPTID() As String
    FPTID = mFPTID
End Property

Public Property Get FPTNM() As String
    FPTNM = mFPTNM
End Property

Public Property Get FJUMIN() As String
    FJUMIN = mFJUMIN
End Property

Public Property Get FSEX() As String
    FSEX = mFSEX
End Property

Public Property Get FAGE() As String
    FAGE = mFAGE
End Property

Public Property Get FWARDCD() As String
    FWARDCD = mFWARDCD
End Property

Public Property Get FWARDNM() As String
    FWARDNM = mFWARDNM
End Property

Public Property Get FDEPTCD() As String
    FDEPTCD = mFDEPTCD
End Property

Public Property Get FDEPTNM() As String
    FDEPTNM = mFDEPTNM
End Property

Public Property Get FDOCTCD() As String
    FDOCTCD = mFDOCTCD
End Property

Public Property Get FDOCTNM() As String
    FDOCTNM = mFDOCTNM
End Property

Public Property Get FEMPID() As String
    FEMPID = mFEMPID
End Property

Public Property Get FEMPNM() As String
    FEMPNM = mFEMPNM
End Property

Public Property Get FSPCCD() As String
    FSPCCD = mFSPCCD
End Property

Public Property Get FSPCNM() As String
    FSPCNM = mFSPCNM
End Property

Public Property Get FCONCAT() As String
    FCONCAT = mFCONCAT
End Property

'-----------------------------------------------------------------------------'
'   기능 : 테이블, 필드정보를 로드
'-----------------------------------------------------------------------------'
Public Sub LoadTableInfo()
    Dim objCom      As clsIISCommon
    Dim strFileNm   As String         'INI파일 경로+파일명
    
    '## NOTE : ini파일이 없을경우의 에러처리 할것!
    strFileNm = mIniPath & "IIS.ini"
    Set objCom = New clsIISCommon
    With objCom
        '## TABLE
        mTHIS001 = .mReadINI(strFileNm, "TABLE", "THIS001")
        mTHIS002 = .mReadINI(strFileNm, "TABLE", "THIS002")
        mTHIS003 = .mReadINI(strFileNm, "TABLE", "THIS003")
        mTHIS004 = .mReadINI(strFileNm, "TABLE", "THIS004")
        mTHIS005 = .mReadINI(strFileNm, "TABLE", "THIS005")
        mTHIS006 = .mReadINI(strFileNm, "TABLE", "THIS006")
        mTCOM001 = .mReadINI(strFileNm, "TABLE", "TCOM001")
        mTIIS001 = .mReadINI(strFileNm, "TABLE", "TIIS001")
        mTIIS002 = .mReadINI(strFileNm, "TABLE", "TIIS002")
        mTIIS003 = .mReadINI(strFileNm, "TABLE", "TIIS003")
        mTIIS101 = .mReadINI(strFileNm, "TABLE", "TIIS101")
        mTIIS102 = .mReadINI(strFileNm, "TABLE", "TIIS102")
        mTIIS103 = .mReadINI(strFileNm, "TABLE", "TIIS103")
        mTIIS104 = .mReadINI(strFileNm, "TABLE", "TIIS104")
        mTIIS105 = .mReadINI(strFileNm, "TABLE", "TIIS105")
        mTIIS106 = .mReadINI(strFileNm, "TABLE", "TIIS106")
        mTIIS107 = .mReadINI(strFileNm, "TABLE", "TIIS107")
        mTIIS201 = .mReadINI(strFileNm, "TABLE", "TIIS201")
        mTIIS202 = .mReadINI(strFileNm, "TABLE", "TIIS202")
        mTIIS203 = .mReadINI(strFileNm, "TABLE", "TIIS203")
        mTIIS204 = .mReadINI(strFileNm, "TABLE", "TIIS204")
        mTIIS205 = .mReadINI(strFileNm, "TABLE", "TIIS205")
        mTIIS206 = .mReadINI(strFileNm, "TABLE", "TIIS206")
        mTIIS301 = .mReadINI(strFileNm, "TABLE", "TIIS301")
        mTIIS302 = .mReadINI(strFileNm, "TABLE", "TIIS302")
        mTIIS303 = .mReadINI(strFileNm, "TABLE", "TIIS303")
        mTIIS304 = .mReadINI(strFileNm, "TABLE", "TIIS304")
        mTIIS305 = .mReadINI(strFileNm, "TABLE", "TIIS305")
        mTIIS306 = .mReadINI(strFileNm, "TABLE", "TIIS306")
        mTIIS307 = .mReadINI(strFileNm, "TABLE", "TIIS307")
        mTIIS308 = .mReadINI(strFileNm, "TABLE", "TIIS308")
        mTIIS401 = .mReadINI(strFileNm, "TABLE", "TIIS401")
        mTIIS402 = .mReadINI(strFileNm, "TABLE", "TIIS402")
        mTIIS403 = .mReadINI(strFileNm, "TABLE", "TIIS403")
        mTIIS404 = .mReadINI(strFileNm, "TABLE", "TIIS404")
        mTIIS405 = .mReadINI(strFileNm, "TABLE", "TIIS405")
        mTIIS406 = .mReadINI(strFileNm, "TABLE", "TIIS406")
        mTIIS501 = .mReadINI(strFileNm, "TABLE", "TIIS501")
        mTIIS502 = .mReadINI(strFileNm, "TABLE", "TIIS502")
        mTIIS503 = .mReadINI(strFileNm, "TABLE", "TIIS503")
        mTIIS504 = .mReadINI(strFileNm, "TABLE", "TIIS504")
        mTIIS505 = .mReadINI(strFileNm, "TABLE", "TIIS505")
        mTIIS601 = .mReadINI(strFileNm, "TABLE", "TIIS601")
        mTACC203 = .mReadINI(strFileNm, "TABLE", "TACC203")
        mTACC204 = .mReadINI(strFileNm, "TABLE", "TACC204")
        
        '## FIELD
        mFPTID = .mReadINI(strFileNm, "FIELD", "FPTID")
        mFPTNM = .mReadINI(strFileNm, "FIELD", "FPTNM")
        mFJUMIN = .mReadINI(strFileNm, "FIELD", "FJUMIN")
        mFSEX = .mReadINI(strFileNm, "FIELD", "FSEX")
        mFAGE = .mReadINI(strFileNm, "FIELD", "FAGE")
        mFWARDCD = .mReadINI(strFileNm, "FIELD", "FWARDCD")
        mFWARDNM = .mReadINI(strFileNm, "FIELD", "FWARDNM")
        mFDEPTCD = .mReadINI(strFileNm, "FIELD", "FDEPTCD")
        mFDEPTNM = .mReadINI(strFileNm, "FIELD", "FDEPTNM")
        mFDOCTCD = .mReadINI(strFileNm, "FIELD", "FDOCTCD")
        mFDOCTNM = .mReadINI(strFileNm, "FIELD", "FDOCTNM")
        mFEMPID = .mReadINI(strFileNm, "FIELD", "FEMPID")
        mFEMPNM = .mReadINI(strFileNm, "FIELD", "FEMPNM")
        mFSPCCD = .mReadINI(strFileNm, "FIELD", "FSPCCD")
        mFSPCNM = .mReadINI(strFileNm, "FIELD", "FSPCNM")
        
        '## DB별 구분자
        mFCONCAT = .mReadINI(strFileNm, "FIELD", "FCONCAT")
    End With
    Set objCom = Nothing
End Sub
