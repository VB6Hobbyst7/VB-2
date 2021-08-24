VERSION 5.00
Begin VB.UserControl ALLOSTATION 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2100
   LockControls    =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   2100
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "AlloStation"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   1500
   End
End
Attribute VB_Name = "ALLOSTATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sFormula = "0"
'Const m_def_DBPath = 0
Const m_def_Settings = ""
Const m_def_sRstFileNm = "0"
Const m_def_sRstFilePath = "0"
Const m_def_sVersion = "0"
Const m_def_EqName = "0"
Const m_def_bUseBarcode = 0
Const m_def_iPhase = 0
Const m_def_iSendPhase = 0
Const m_def_sTestMode = "0"
Const m_def_iFrameN = 0
Const m_def_p_sID = "0"
Const m_def_p_sSeq = "0"
Const m_def_p_sRack = "0"
Const m_def_p_sPos = "0"
Const m_def_p_iOrdCnt = 0
Const m_def_p_sTIFCd = "0"
Const m_def_OpenPW = "0"
Const m_def_EditPW = "0"
'속성 변수:
Dim m_p_sFormula As String
'Dim m_DBPath As Variant
Dim m_Settings As String
Dim m_sRstFileNm As String
Dim m_sRstFilePath As String
Dim m_sVersion As String
Dim m_EqName As String
Dim m_bUseBarcode As Boolean
Dim m_iPhase As Integer
Dim m_iSendPhase As Integer
Dim m_sTestMode As String
Dim m_iFrameN As Integer
Dim m_p_sID As String
Dim m_p_sSeq As String
Dim m_p_sRack As String
Dim m_p_sPos As String
Dim m_p_iOrdCnt As Integer
Dim m_p_sTIFCd As String
Dim m_OpenPW As String
Dim m_EditPW As String
'이벤트 선언:
Event GetFormula(sTIFCd$)
Event ChkCurPatID(sPatID$, sRetCd$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTInstID$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)


'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim wkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim iSpaceCnt   As Integer

'for AlloScan
Private Type PANELCODEINFO
    FieldNm(100)    As String
    Code(100)       As String
End Type
Private tPnlCdInfo()    As PANELCODEINFO

'
'   결과정보 구조체 초기화
'
Private Sub Init_pResultInfo()
    
    With pResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .KIND = ""
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
End Sub

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_OpenPW = PropBag.ReadProperty("OpenPW", m_def_OpenPW)
    m_EditPW = PropBag.ReadProperty("EditPW", m_def_EditPW)
    m_EqName = PropBag.ReadProperty("EqName", m_def_EqName)
    m_bUseBarcode = PropBag.ReadProperty("bUseBarcode", m_def_bUseBarcode)
    m_iPhase = PropBag.ReadProperty("iPhase", m_def_iPhase)
    m_iSendPhase = PropBag.ReadProperty("iSendPhase", m_def_iSendPhase)
    m_sTestMode = PropBag.ReadProperty("sTestMode", m_def_sTestMode)
    m_iFrameN = PropBag.ReadProperty("iFrameN", m_def_iFrameN)
    m_p_sID = PropBag.ReadProperty("p_sID", m_def_p_sID)
    m_p_sSeq = PropBag.ReadProperty("p_sSeq", m_def_p_sSeq)
    m_p_sRack = PropBag.ReadProperty("p_sRack", m_def_p_sRack)
    m_p_sPos = PropBag.ReadProperty("p_sPos", m_def_p_sPos)
    m_p_iOrdCnt = PropBag.ReadProperty("p_iOrdCnt", m_def_p_iOrdCnt)
    m_p_sTIFCd = PropBag.ReadProperty("p_sTIFCd", m_def_p_sTIFCd)
    m_sVersion = PropBag.ReadProperty("sVersion", m_def_sVersion)
    m_sRstFileNm = PropBag.ReadProperty("sRstFileNm", m_def_sRstFileNm)
    m_sRstFilePath = PropBag.ReadProperty("sRstFilePath", m_def_sRstFilePath)
    m_Settings = PropBag.ReadProperty("Settings", m_def_Settings)

    m_p_sFormula = PropBag.ReadProperty("p_sFormula", m_def_p_sFormula)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("OpenPW", m_OpenPW, m_def_OpenPW)
    Call PropBag.WriteProperty("EditPW", m_EditPW, m_def_EditPW)
    Call PropBag.WriteProperty("EqName", m_EqName, m_def_EqName)
    Call PropBag.WriteProperty("bUseBarcode", m_bUseBarcode, m_def_bUseBarcode)
    Call PropBag.WriteProperty("iPhase", m_iPhase, m_def_iPhase)
    Call PropBag.WriteProperty("iSendPhase", m_iSendPhase, m_def_iSendPhase)
    Call PropBag.WriteProperty("sTestMode", m_sTestMode, m_def_sTestMode)
    Call PropBag.WriteProperty("iFrameN", m_iFrameN, m_def_iFrameN)
    Call PropBag.WriteProperty("p_sID", m_p_sID, m_def_p_sID)
    Call PropBag.WriteProperty("p_sSeq", m_p_sSeq, m_def_p_sSeq)
    Call PropBag.WriteProperty("p_sRack", m_p_sRack, m_def_p_sRack)
    Call PropBag.WriteProperty("p_sPos", m_p_sPos, m_def_p_sPos)
    Call PropBag.WriteProperty("p_iOrdCnt", m_p_iOrdCnt, m_def_p_iOrdCnt)
    Call PropBag.WriteProperty("p_sTIFCd", m_p_sTIFCd, m_def_p_sTIFCd)
    Call PropBag.WriteProperty("sVersion", m_sVersion, m_def_sVersion)
    Call PropBag.WriteProperty("sRstFileNm", m_sRstFileNm, m_def_sRstFileNm)
    Call PropBag.WriteProperty("sRstFilePath", m_sRstFilePath, m_def_sRstFilePath)
    Call PropBag.WriteProperty("Settings", m_Settings, m_def_Settings)

    Call PropBag.WriteProperty("p_sFormula", m_p_sFormula, m_def_p_sFormula)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get OpenPW() As String
    OpenPW = m_OpenPW
End Property

Public Property Let OpenPW(ByVal New_OpenPW As String)
    m_OpenPW = New_OpenPW
    PropertyChanged "OpenPW"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EditPW() As String
    EditPW = m_EditPW
End Property

Public Property Let EditPW(ByVal New_EditPW As String)
    m_EditPW = New_EditPW
    PropertyChanged "EditPW"
End Property

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    m_OpenPW = m_def_OpenPW
    m_EditPW = m_def_EditPW
    m_EqName = m_def_EqName
    m_bUseBarcode = m_def_bUseBarcode
    m_iPhase = m_def_iPhase
    m_iSendPhase = m_def_iSendPhase
    m_sTestMode = m_def_sTestMode
    m_iFrameN = m_def_iFrameN
    m_p_sID = m_def_p_sID
    m_p_sSeq = m_def_p_sSeq
    m_p_sRack = m_def_p_sRack
    m_p_sPos = m_def_p_sPos
    m_p_iOrdCnt = m_def_p_iOrdCnt
    m_p_sTIFCd = m_def_p_sTIFCd
    m_sVersion = m_def_sVersion
    m_sRstFileNm = m_def_sRstFileNm
    m_sRstFilePath = m_def_sRstFilePath
    m_Settings = m_def_Settings

    m_p_sFormula = m_def_p_sFormula
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EqName() As String
    EqName = m_EqName
End Property

Public Property Let EqName(ByVal New_EqName As String)
    m_EqName = New_EqName
    PropertyChanged "EqName"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get bUseBarcode() As Boolean
    bUseBarcode = m_bUseBarcode
End Property

Public Property Let bUseBarcode(ByVal New_bUseBarcode As Boolean)
    m_bUseBarcode = New_bUseBarcode
    PropertyChanged "bUseBarcode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iPhase() As Integer
    iPhase = m_iPhase
End Property

Public Property Let iPhase(ByVal New_iPhase As Integer)
    m_iPhase = New_iPhase
    PropertyChanged "iPhase"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iSendPhase() As Integer
    iSendPhase = m_iSendPhase
End Property

Public Property Let iSendPhase(ByVal New_iSendPhase As Integer)
    m_iSendPhase = New_iSendPhase
    PropertyChanged "iSendPhase"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sTestMode() As String
    sTestMode = m_sTestMode
End Property

Public Property Let sTestMode(ByVal New_sTestMode As String)
    m_sTestMode = New_sTestMode
    PropertyChanged "sTestMode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iFrameN() As Integer
    iFrameN = m_iFrameN
End Property

Public Property Let iFrameN(ByVal New_iFrameN As Integer)
    m_iFrameN = New_iFrameN
    PropertyChanged "iFrameN"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sID() As String
    p_sID = m_p_sID
End Property

Public Property Let p_sID(ByVal New_p_sID As String)
    m_p_sID = New_p_sID
    PropertyChanged "p_sID"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sSeq() As String
    p_sSeq = m_p_sSeq
End Property

Public Property Let p_sSeq(ByVal New_p_sSeq As String)
    m_p_sSeq = New_p_sSeq
    PropertyChanged "p_sSeq"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sRack() As String
    p_sRack = m_p_sRack
End Property

Public Property Let p_sRack(ByVal New_p_sRack As String)
    m_p_sRack = New_p_sRack
    PropertyChanged "p_sRack"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sPos() As String
    p_sPos = m_p_sPos
End Property

Public Property Let p_sPos(ByVal New_p_sPos As String)
    m_p_sPos = New_p_sPos
    PropertyChanged "p_sPos"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get p_iOrdCnt() As Integer
    p_iOrdCnt = m_p_iOrdCnt
End Property

Public Property Let p_iOrdCnt(ByVal New_p_iOrdCnt As Integer)
    m_p_iOrdCnt = New_p_iOrdCnt
    PropertyChanged "p_iOrdCnt"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sTIFCd() As String
    p_sTIFCd = m_p_sTIFCd
End Property

Public Property Let p_sTIFCd(ByVal New_p_sTIFCd As String)
    m_p_sTIFCd = New_p_sTIFCd
    PropertyChanged "p_sTIFCd"
End Property


'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sVersion() As String
    sVersion = m_sVersion
End Property

Public Property Let sVersion(ByVal New_sVersion As String)
    m_sVersion = New_sVersion
    PropertyChanged "sVersion"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sRstFileNm() As String
    sRstFileNm = m_sRstFileNm
End Property

Public Property Let sRstFileNm(ByVal New_sRstFileNm As String)
    m_sRstFileNm = New_sRstFileNm
    PropertyChanged "sRstFileNm"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sRstFilePath() As String
    sRstFilePath = m_sRstFilePath
End Property

Public Property Let sRstFilePath(ByVal New_sRstFilePath As String)
    m_sRstFilePath = New_sRstFilePath
    PropertyChanged "sRstFilePath"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get Settings() As String
Attribute Settings.VB_Description = "전송 속도, 패리티, 데이터 비트, 중단 비트 매개 변수를 반환하거나 설정합니다."
    Settings = m_Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    m_Settings = New_Settings
    PropertyChanged "Settings"
End Property

'
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function RcvRstData(ByVal sDBPath As String, Optional ByVal sSDate As String, Optional ByVal sEDate As String) As Variant

    '--- 사용자 확인
    If m_EditPW <> pEditPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Function
    End If
    '---------------

    If m_EqName = "0" Or m_EqName = "" Then
        RaiseEvent DispMsg("검사장비명을 지정해 주십시오.!!!")
        Exit Function
    End If

    '--- 암호 확인
    If m_OpenPW <> pOpenPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Function
    End If
    '-----------------------

    If Trim(sDBPath) = "" Then
        Exit Function
    End If

    '결과조회
    Call GetResultData_AlloScan(sDBPath, sSDate, sEDate)
    
End Function

Private Sub GetResultData_AlloScan(ByVal sPath As String, Optional ByVal sSDate As String, Optional ByVal sEDate As String)
    On Error GoTo ErrRtn

    Dim ADORS   As New ADODB.Recordset
    Dim SqlStr  As String
    
    Dim ii%, kk%
    Dim sPnlA$, sPnlB$
    Dim aID()   As String

    Dim iRstCnt%
    Dim sID$, sTIFCd$, sTRst1$, sIFCd$, sRst1$, sTmp$
    Dim sRetCd$
    
    'StripBand 정보 조회
    Call GetPanelCodeINFO(sPath)

    'RESULT 정보 조회
    SqlStr = ""
    SqlStr = SqlStr & " select * "
    SqlStr = SqlStr & "   from RESULTS "
    SqlStr = SqlStr & "  Where PatientID <> null "
    
    If sSDate <> "" And sEDate <> "" Then
        ''SqlStr = SqlStr & "  and date between '" & sSDate & "' and '" & sEDate & "' "
        SqlStr = SqlStr & "  and examdate between '" & sSDate & "' and '" & sEDate & "' "
    End If

    ADORS.Open SqlStr, fGetCurConn(sPath, "DBresults.mdb"), adOpenForwardOnly

    If ADORS.EOF = True Then
        ADORS.Close: Set ADORS = Nothing
        Exit Sub
    End If

    Do Until ADORS.EOF
        iRstCnt = 0: sID = "": sTIFCd = "": sTRst1 = ""

        sID = Trim(ADORS.Fields("PatientID"))
        sPnlA = Trim(ADORS.Fields("StripPanel_A")) & " A"
        sPnlB = Trim(ADORS.Fields("StripPanel_B")) & " B"

        If sID <> "" Then
            sRetCd = ""
            'worklist에 작성된 환자인지 체크
'            If InStr(sID, "-") = 0 Then
'                sID = sID
'            Else
'                aID() = Split(sID, "-")
'                sID = Trim(aID(0))
'            End If
            RaiseEvent ChkCurPatID(sID, sRetCd)

            If sRetCd = "OK" Then
                'panel A
                For kk = 1 To 22
                    sTmp = "Band" & Trim(kk)
                    If kk = 1 Then
                        sIFCd = "PCA"
                    Else
                        sIFCd = FindPanelCode(sPnlA, sTmp)
                    End If
                    
                    RaiseEvent GetFormula(sIFCd & "-" & Left(sPnlA, 1))
                    sTmp = "BandVal_A" & Trim(kk)
                    sRst1 = Trim(ADORS.Fields(sTmp) & "")
                    
                    If sRst1 > 0 Then
                        If p_sFormula <> "" Then
                            sRst1 = AllergenCalc(p_sFormula, sRst1)
                        End If
                    Else
                        sRst1 = "0"
                    End If
                              
                    iRstCnt = iRstCnt + 1
                    sTIFCd = sTIFCd & sIFCd & "-" & Left(sPnlA, 1) & Chr(124)
                    sTRst1 = sTRst1 & sRst1 & Chr(124)
                Next kk
                
                'panel B
                For kk = 1 To 22
                    sTmp = "Band" & Trim(kk)
                    If kk = 1 Then
                        sIFCd = "PCB"
                    Else
                        sIFCd = FindPanelCode(sPnlB, sTmp)
                    End If
                    
                    RaiseEvent GetFormula(sIFCd & "-" & Left(sPnlB, 1))
                    sTmp = "BandVal_B" & Trim(kk)
                    sRst1 = Trim(ADORS.Fields(sTmp) & "")
                    
                    If sRst1 > 0 Then
                        If p_sFormula <> "" Then
                            sRst1 = AllergenCalc(p_sFormula, sRst1)
                        End If
                    Else
                        sRst1 = "0"
                    End If
                    
                    iRstCnt = iRstCnt + 1
                    sTIFCd = sTIFCd & sIFCd & "-" & Left(sPnlB, 1) & Chr(124)
                    sTRst1 = sTRst1 & sRst1 & Chr(124)
                Next kk
                
                '결과값 등록/화면 표시 처리...
                 RaiseEvent AppendData(sID, "", "", "", iRstCnt, sTIFCd, sTRst1, String(iRstCnt, Chr(124)), String(iRstCnt, Chr(124)), String(iRstCnt, Chr(124)), _
                                    "", "", "", "", "")
            End If
        End If

        ADORS.MoveNext
    Loop

    ADORS.Close: Set ADORS = Nothing

ErrRtn:
    If Err <> 0 Then
        If ADORS.State = 1 Then
            ADORS.Close: Set ADORS = Nothing
        End If
        RaiseEvent DispMsg("GetResultData_AlloScan - " & Err.Description)
        MsgBox "GetResultData_AlloScan - " & Err.Description, vbExclamation
    End If
End Sub

Public Function AllergenCalc(ByVal rsType As String, ByVal rsValues As String) As String
    On Error GoTo ErrRtn
      
    Dim a As Double
    Dim d As Double
    Dim M As Double
    Dim N As Double
    Dim y As Double
    Dim Tmp As Double
        
    a = 0
    d = 150
    
    Select Case rsType
        Case "A"
            M = -0.696669715055768
            N = 3.57268581287044

        Case "B"
            M = -0.58804147563697
            N = 3.68672614195763

        Case "C"
            M = -0.424608277908711
            N = 3.85830192881032
            
    End Select

    y = Val(rsValues)

    y = (y - d) / (a - d)
    y = y / (1 - y)
    y = Log(y)
    y = (y - N) / M
    y = Exp(y)
    
    Tmp = Format(y, "0.00")
    
'    If Tmp - y > 0 Then
'        Tmp = Tmp - 0.01
'    End If
    
    y = Tmp
    
    AllergenCalc = CStr(y)
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("AllergenCalc - " & Err.Description)
    End If
End Function
    
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function SendWorkList(sDBPath$, sWorkList$) As Variant
    On Error GoTo ErrRtn

    Dim ADOCN   As New ADODB.Connection
    Dim SqlStr  As String
    Dim aRow()  As String
    Dim aData() As String
    Dim ii      As Integer
    Dim sID$, sLNm$, sFNm$, sDOB$, sGrpNm$
    Dim lRet    As Long
    
    If Trim(sWorkList) = "" Then Exit Function
    
    With ADOCN
        .ConnectionString = fGetCurConn(sDBPath, "dbPATIENT.db")
        .Open
    End With
    
    aRow() = Split(sWorkList, Chr(3))
    
    For ii = 0 To UBound(aRow())
        If Trim(aRow(ii)) = "" Then Exit For
    
        Erase aData()
        aData() = Split(aRow(ii), Chr(124))
        
'        sSndList = sSndList & sBarCd & "-" & sDate & Chr(124) & sPatNm & Chr(124) & Trim(Format(iSeq, "000")) & Chr(124) & Format(Now, "yyyy-MM-dd") & Chr(3)
                                        
        sID = Trim(aData(0))
        sLNm = Trim(aData(1))
        sFNm = Trim(aData(2))
        sDOB = Trim(aData(3))
        sGrpNm = Trim(aData(4))
        
        SqlStr = " INSERT INTO PATIENT " _
                & "(UsrID, ID, LastName, FirstName, DOB, [Group], Country) VALUES ("
        SqlStr = SqlStr & "'Administrator', "
        SqlStr = SqlStr & "'" & sID & "', "
        SqlStr = SqlStr & "'" & sFNm & sLNm & "', "
        SqlStr = SqlStr & "'" & Split(sID, "-")(0) & "', "
        SqlStr = SqlStr & "'" & sDOB & "', "
        SqlStr = SqlStr & "'" & sGrpNm & "', "
        SqlStr = SqlStr & "'대한민국') "
        
        ADOCN.Execute SqlStr, lRet
    Next ii

    ADOCN.Close: Set ADOCN = Nothing

    SendWorkList = "OK"
    
ErrRtn:
    If Err <> 0 Then
        If ADOCN.State = 1 Then
            ADOCN.Close: Set ADOCN = Nothing
        End If
        RaiseEvent DispMsg("SendWorkList - " & Err.Description)
        MsgBox "SendWorkList - " & Err.Description, vbExclamation
    End If
End Function
Private Function fGetCurConn(ByVal sPath As String, ByVal sDBNm As String) As String

    fGetCurConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & "\" & sDBNm & ";User ID=admin;Jet OLEDB:Database Password=reader_admin;"
        
End Function
Private Function FindPanelCode(ByVal sPnl As String, ByVal sCd As String) As String
    On Error GoTo ErrFind
    
    Dim ix1%
    
    FindPanelCode = ""
    
    For ix1 = 1 To UBound(tPnlCdInfo())
        With tPnlCdInfo(ix1)
            If ConvertFieldCd(ix1, "StripPanel") = sPnl Then
                FindPanelCode = ConvertFieldNm(ix1, sCd)
                
                Exit For
            End If
        End With
    Next ix1
        
ErrFind:
    If Err <> 0 Then
        RaiseEvent DispMsg("FindPanelCode - " & Err.Description)
        MsgBox "FindPanelCode - " & Err.Description, vbExclamation
    End If
End Function
Private Sub GetPanelCodeINFO(ByVal sTmpPath As String)
    On Error GoTo ErrGetCode

    Dim tmpADORS    As New ADODB.Recordset
    Dim tmpStr  As String
    Dim i%, k%
    
    'StripBand 정보 조회
    tmpStr = " select * from STRIPBAND "

    tmpADORS.Open tmpStr, fGetCurConn(sTmpPath, "DBstripband.mdb"), adOpenForwardOnly

    If tmpADORS.EOF = True Then
        tmpADORS.Close: Set tmpADORS = Nothing
        RaiseEvent DispMsg("STRIPBAND 정보가 존재하지 않습니다.")
        Exit Sub
    End If
    
    Erase tPnlCdInfo()
    
    i = 0
    Do Until tmpADORS.EOF
        i = i + 1
        
        With tmpADORS
            ReDim Preserve tPnlCdInfo(i)
        
            For k = 0 To Val(.Fields.Count) - 1
                tPnlCdInfo(i).FieldNm(k + 1) = Trim(.Fields(k).Name)
                tPnlCdInfo(i).Code(k + 1) = Trim(.Fields(k) & "")
            Next k
            
            .MoveNext
        End With
    Loop
    
    tmpADORS.Close: Set tmpADORS = Nothing
    
ErrGetCode:
    If Err <> 0 Then
        If tmpADORS.State = 1 Then
            tmpADORS.Close: Set tmpADORS = Nothing
        End If
        RaiseEvent DispMsg("GetPanelCodeINFO - " & Err.Description)
        MsgBox "GetPanelCodeINFO - " & Err.Description, vbExclamation
    End If
End Sub
Private Function ConvertFieldCd(ByVal iIndex As Integer, ByVal sNm As String) As String
    
    Dim ix9 As Integer
    
    ConvertFieldCd = ""
    
    For ix9 = 1 To UBound(tPnlCdInfo())
        With tPnlCdInfo(iIndex)
            If .FieldNm(ix9) = Trim(sNm) Then
                ConvertFieldCd = .Code(ix9)
                Exit For
            End If
        End With
    Next ix9

End Function

Private Function ConvertFieldNm(ByVal iIndex As Integer, ByVal sNm As String) As String
    
    Dim ix9 As Integer
    
    ConvertFieldNm = ""
    
    For ix9 = 1 To UBound(tPnlCdInfo(iIndex).FieldNm)
        With tPnlCdInfo(iIndex)
            If .FieldNm(ix9) = Trim(sNm) Then
                ConvertFieldNm = .Code(ix9)
                Exit For
            End If
        End With
    Next ix9

End Function
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sFormula() As String
    p_sFormula = m_p_sFormula
End Property

Public Property Let p_sFormula(ByVal New_p_sFormula As String)
    m_p_sFormula = New_p_sFormula
    PropertyChanged "p_sFormula"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function SendWorkListFile(sFilePath$, sWorkList$) As Variant
    On Error GoTo ErrRtn
    
    Dim aRow()  As String
    Dim aData() As String
    Dim ii      As Integer
    Dim sID$, sNm$, sExamDate$, sGrpNm$
    Dim lRet    As Long
    
    Dim iFile%
    Dim sList$
    
    If Trim(sWorkList) = "" Then Exit Function
    
    '<S--- Header
    sList = "-----------------------------------------------------------------------------" & vbCrLf
    sList = sList & "NAME: AdvanSure" & vbCrLf
    sList = sList & "TYPE: Patient List Files" & vbCrLf
    sList = sList & "-----------------------------------------------------------------------------" & vbCrLf
    sList = sList & "ID" & vbTab & "NAME" & vbTab & "PANEL" & vbTab & "A" & vbTab & "B" & vbTab & "AGE" & vbTab & "GENDER" & vbTab & "ADDR" _
                & vbTab & "CONTACT" & vbTab & "BIRTH" & vbTab & "RRN" & vbTab & "CLIENT" & vbTab & "INSPECTOR" & vbTab & "HOSPITAL" & vbTab & "EXAMDATE" & vbCrLf
    '>E----------
    
    aRow() = Split(sWorkList, Chr(3))
    
    For ii = 0 To UBound(aRow())
        If Trim(aRow(ii)) = "" Then Exit For
    
        Erase aData()
        aData() = Split(aRow(ii), Chr(124))
        
        '           0                               1                   2                                      3
        'sSndList = sBarCd & "-" & "F" & Chr(124) & sPatNm & Chr(124) & Format(Now, "yyyy-MM-dd") & Chr(124) & sOrdGrp & Chr(3)
                                        
        sID = Trim(aData(0))
        sNm = Trim(aData(1))
        sExamDate = Trim(aData(2))
        sGrpNm = Trim(aData(3))
        
        'WorkList 편집
        sList = sList & sID & vbTab         'ID
        sList = sList & sNm & vbTab         'NAME
        sList = sList & sGrpNm & vbTab      'PANEL
        sList = sList & "1" & vbTab         'A
        sList = sList & "1" & vbTab         'B
        sList = sList & "" & vbTab          'AGE
        sList = sList & "" & vbTab          'GENDER
        sList = sList & "" & vbTab          'ADDR
        sList = sList & "" & vbTab          'CONTACT
        sList = sList & "" & vbTab          'BIRTH
        sList = sList & "" & vbTab          'RRN
        sList = sList & "" & vbTab          'CLIENT
        sList = sList & "" & vbTab          'INSPECTOR
        sList = sList & "" & vbTab          'HOSPITAL
        sList = sList & sExamDate & vbTab   'EXAMDATE
        sList = sList & vbCrLf
    Next ii

    '<S--- File 생성
    iFile = FreeFile
    
    Open sFilePath For Output As #iFile
    Print #iFile, sList
    Close #iFile
    '>E-------------
    
    SendWorkListFile = "OK"
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendWorkListFile - " & Err.Description)
        MsgBox "SendWorkListFile - " & Err.Description, vbExclamation
    End If
End Function

