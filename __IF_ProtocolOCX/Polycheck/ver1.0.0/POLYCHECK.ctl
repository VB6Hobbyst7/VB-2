VERSION 5.00
Begin VB.UserControl POLYCHECK 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2100
   LockControls    =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   2100
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "POLYCHECK"
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
Attribute VB_Name = "POLYCHECK"
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
Event GetFormula(sTifcd$)
Event ChkCurPatID(sPatID$, sRetCd$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTifcd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTInstID$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
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
Public Function RcvRstData(sDBPath$) As Variant

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
    Call GetResultData_PolyCheck(sDBPath)
    
End Function

Private Sub GetResultData_PolyCheck(ByVal sResult As String)
    On Error GoTo ErrRtn

    Dim ii%, kk%

    Dim iRstCnt%
    Dim sID$, sTifcd$, sRst1$, sTRst1$, sIFCd$, sTmp$
    Dim sTotal$
    
    Dim tmpData()   As String
    Dim tmpBuf()    As String
    Dim sTestBuf()  As String
    Dim sFood()     As String
    Dim sStandard() As String
    Dim sInhalat()  As String
    Dim sOneRow     As String
    
    Dim sClass      As String
    Dim sTestNm     As String

    tmpData = Split(sResult, vbCrLf)
    
    For ii = 0 To UBound(tmpData)
        sTestBuf = Split(tmpData(ii), vbTab)
        
        If IsNumeric(sTestBuf(1)) = False Then
            Select Case sTestBuf(7)
                Case "Standard-KOR", "Korea I-Standard"
                    sStandard = sTestBuf
                Case "Food-KOR", "Korea III-Food"
                    sFood = sTestBuf
                Case "Inhalation-KOR", "Korea II-Inhalation"
                    sInhalat = sTestBuf
            End Select
        Else
            Exit For
        End If
    Next ii
    
    For ii = 0 To UBound(tmpData)
        
        sOneRow = tmpData(ii)
        tmpBuf = Split(sOneRow, vbTab)
        
        If IsNumeric(tmpBuf(1)) = True Then
            sID = Trim(tmpBuf(2))
            sTestNm = Trim(tmpBuf(7))
            sTestNm = Split(sTestNm, "-")(1)
            
            For kk = 8 To UBound(tmpBuf)
                sRst1 = Replace(tmpBuf(kk), ",", ".")
                sTotal = sRst1
                
                If Val(sRst1) < 0.35 Then
                    sClass = "0"
                ElseIf Val(sRst1) >= 0.35 And Val(sRst1) < 0.7 Then
                    sClass = "1"
                ElseIf Val(sRst1) >= 0.7 And Val(sRst1) < 3.5 Then
                    sClass = "2"
                ElseIf Val(sRst1) >= 3.5 And Val(sRst1) < 17.5 Then
                    sClass = "3"
                ElseIf Val(sRst1) >= 17.5 And Val(sRst1) < 50 Then
                    sClass = "4"
                ElseIf Val(sRst1) >= 50 And Val(sRst1) < 100 Then
                    sClass = "5"
                ElseIf Val(sRst1) >= 100 Then
                    sClass = "6"
                End If
                
                If Val(sRst1) < 0.15 Then
                    sRst1 = "<0.15"
                ElseIf Val(sRst1) > 100 Then
                    sRst1 = ">100"
                End If
                
                Select Case sTestNm
                    Case "Food"
                        iRstCnt = iRstCnt + 1
                        sTifcd = sTifcd & sFood(kk) & Chr(124)
                        sTRst1 = sTRst1 & sClass & Chr(124)
                        
                        iRstCnt = iRstCnt + 1
                        sTifcd = sTifcd & sFood(kk) & "-C" & Chr(124)
                        sTRst1 = sTRst1 & sRst1 & Chr(124)
                    
                    Case "Standard"
                        If kk = 27 Then
                            If Val(sTotal) > 100 Then
                                iRstCnt = iRstCnt + 1
                                sTifcd = sTifcd & sStandard(kk) & Chr(124)
                                sTRst1 = sTRst1 & ">100" & Chr(124)
                            Else
                                iRstCnt = iRstCnt + 1
                                sTifcd = sTifcd & sStandard(kk) & Chr(124)
                                sTRst1 = sTRst1 & sTotal & Chr(124)
                            End If
    
                        Else
                            iRstCnt = iRstCnt + 1
                            sTifcd = sTifcd & sStandard(kk) & Chr(124)
                            sTRst1 = sTRst1 & sClass & Chr(124)
                            
                            iRstCnt = iRstCnt + 1
                            sTifcd = sTifcd & sStandard(kk) & "-C" & Chr(124)
                            sTRst1 = sTRst1 & sRst1 & Chr(124)
                        End If
                        
                    Case "Inhalation"
                        iRstCnt = iRstCnt + 1
                        sTifcd = sTifcd & sInhalat(kk) & Chr(124)
                        sTRst1 = sTRst1 & sClass & Chr(124)
                        
                        iRstCnt = iRstCnt + 1
                        sTifcd = sTifcd & sInhalat(kk) & "-C" & Chr(124)
                        sTRst1 = sTRst1 & sRst1 & Chr(124)
                End Select
            Next kk
            
            '결과값 등록/화면 표시 처리...
            RaiseEvent AppendData(sID, "", "", "", iRstCnt, sTifcd, sTRst1, String(iRstCnt, Chr(124)), String(iRstCnt, Chr(124)), String(iRstCnt, Chr(124)), _
                            "", "", "", "", "")
            iRstCnt = 0: sID = "": sTifcd = "": sTRst1 = ""
        End If
    Next ii


ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("GetResultData_PolyCheck - " & Err.Description)
        MsgBox "GetResultData_PolyCheck - " & Err.Description, vbExclamation
    End If
End Sub

Public Function AllergenCalc(ByVal rsType As String, ByVal rsValues As String) As String
    On Error GoTo ErrRtn
      
    Dim a As Double
    Dim d As Double
    Dim M As Double
    Dim N As Double
    Dim y As Double

    a = 0
    d = 300

    Select Case rsType
        Case "A"
            M = -0.74006186
            N = 3.30060413

        Case "B"
            M = -0.68208301
            N = 3.36147160884595

        Case "C"
            M = -0.617902504
            N = 3.42884972400335

        Case "D"
            M = -0.542387562
            N = 3.508126981

        Case "E"
            M = -0.444539321
            N = 3.610850229
    End Select

    y = Val(rsValues)

    y = (y - d) / (a - d)
    y = y / (1 - y)
    y = Log(y)
    y = (y - N) / M
    y = Exp(y)
    y = Round(y, 2)

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
        SqlStr = SqlStr & "'" & sLNm & "', "
        SqlStr = SqlStr & "'" & sFNm & "', "
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
                FindPanelCode = ConvertFieldCd(ix1, sCd)
                
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
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sFormula() As String
    p_sFormula = m_p_sFormula
End Property

Public Property Let p_sFormula(ByVal New_p_sFormula As String)
    m_p_sFormula = New_p_sFormula
    PropertyChanged "p_sFormula"
End Property

