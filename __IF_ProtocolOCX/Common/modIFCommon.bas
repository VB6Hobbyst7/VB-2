Attribute VB_Name = "modIFCommon"
Option Explicit

'Sample 에 대한 정보
Type SAMPLE_INFO
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    Kind    As String       '1st/Rerun 구분
    ORDCNT  As Integer
    IFCD()  As String
    SVOL()  As String
    SINDEX  As Boolean      'Serum Index 검사여부
    SPCCD   As String       '검체구분(Rack Type)
    RSTDT   As String       '2005/6/10 yk
    OTHER   As String
    CMT1    As String       '2005/8/1 yk
    INSTID  As String       '2006/2/9 yk
    INSTNM  As String       '2006/10/11 yk
    CONTROLNAME As String   '2007/01/17    by YeounJu
    CONTAINER   As String   'Container Type(for Cobas6000)...2007/6/22 yk
    
    PATNO   As String
    NAME    As String
    BIRTH   As String       '환자정보 전송(for XE-2100)...2009/9/14 yk
    SEX     As String       '
    DEPT    As String
    WARD    As String
    ROOM    As String
    DIL()   As String       '희석배수...2012/7/2
    
    CANCELORDER As String   '취소오더...2012/9/26 yk
    
    SENDER  As String
    VERSION As String
End Type

'수신된 결과에 대한 정보
Type RESULT_INFO
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    Kind    As String
    RSTCNT  As Integer
    IFCD    As String
    RST1    As String
    RST2    As String
    UNIT    As String
    FLAG    As String
    ALARMCD As String
    INSTID  As String
    INSTNM  As String       '2006/10/11 yk
    RSTDT   As String       'Date/Time Results Reported or Last Modified (2005/6/10 추가 yk)
    SPCCD   As String       '검체구분 (2006/10/11 추가 yk)
    OPERID  As String       'Operation ID (       "      )
    OTHER   As String
    STATUS  As String
    DIL     As String       '2012/7/2
End Type

'Slide 에 대한 정보...for SP-1000i
Type SLIDE_INFO
    ID      As String
    NO_FILM As Integer
    HCT     As String
    WBC     As String
    RBC     As String
    PRT1    As String
    PRT2    As String
    PRT3    As String
End Type

'Sample 에 대한 정보
Type SAMPLE_INFO_M
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    Kind    As String       '1st/Rerun 구분
    ORDCNT  As Integer
    IFCD()  As String
    SVOL()  As String
    SINDEX  As Boolean      'Serum Index 검사여부
    PATNO   As String
    SPCNM   As String
    SPCCD   As String
    RSTDT   As String       '2005/6/10 yk
    OTHER   As String
    CMT1    As String       '2005/8/1 yk
    INSTID  As String       '2006/2/9 yk
    INSTNM  As String       '2006/10/11 yk
    CONTROLNAME As String   '2007/01/17    by YeounJu
    CONTAINER   As String   'Container Type(for Cobas6000)...2007/6/22 yk
    
    BIRTH   As String       '환자정보 전송(for XE-2100)...2009/9/14 yk
    SEX     As String       '        "
    URINE   As String
End Type

'수신된 결과에 대한 정보
Type RESULT_INFO_M
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    Kind    As String
    RSTCNT  As Integer
    IFCD    As String
    RST1    As String
    RST2    As String
    UNIT    As String
    FLAG    As String
    ALARMCD As String
    INSTID  As String
    INSTNM  As String       '2006/10/11 yk
    RSTDT   As String       'Date/Time Results Reported or Last Modified (2005/6/10 추가 yk)
    OPERID  As String       'Operation ID (       "      )
    OTHER   As String
    PATNO   As String
    SPCNM   As String
    SPCCD   As String
    DEPT    As String
    WARD    As String
    ROOM    As String
    URINE   As String
    
    KITCD   As String
    KITNM   As String
    ISOL    As String
    ORGCNT  As Integer
    ORGCD   As String
    ORGNM   As String
    ANTIRST As Integer
    ANTICD  As String
    ANTINM  As String
    SRI     As String
    MIC     As String
    STATUS  As String   'Isolate Status
    
    TKITCD   As String
    TKITNM   As String
    TISOL    As String
    TORGCD   As String
    TORGNM   As String
    TANTIRST As String
    TANTICD  As String
    TANTINM  As String
    TSRI     As String
    TMIC     As String
    TSTATUS  As String   'Isolate Status
End Type

Public Type PANELTBL
    PKind As String
    Panel As String
    Key As String
End Type

Public Type MANUALTBL
    PKind As String
    Panel As String
    MKind As String
    Manual As String
    Key As String
End Type

'임시로 비밀번호 설정
Public Const pOpenPW = "ACK"
Public Const pEditPW = "MEDI@CK"
'
'   ASTM Protocol CheckSum 계산
'
Public Function ChkSum_ASTM(ByVal Para As String) As String

    Dim i   As Integer
    Dim Tmp As Integer
    Dim ChkS1   As Integer
    Dim ChkS2   As String
    
    For i = 1 To Len(Para)
        Tmp = Asc(Mid$(Para, i, 1))
        ChkS1 = ChkS1 + Tmp
    Next i
    ChkS1 = ChkS1 Mod 256
    ChkS2 = Right$("0" & Hex$(ChkS1), 2)
    
    ChkSum_ASTM = ChkS2
    
End Function



