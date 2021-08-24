VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl FileLink 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   1665
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "TEST"
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Top             =   1725
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   1395
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   135
      Width           =   1365
   End
End
Attribute VB_Name = "FileLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sPatInfo = 0
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
Const m_def_PortOpen = 0
Const m_def_OpenPW = "0"
Const m_def_EditPW = "0"
Const m_def_FilePath = "c:\"
Const m_def_FileFilter = "*.*"

'속성 변수:
Dim m_p_sPatInfo As Variant
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
Dim m_PortOpen As Boolean
Dim m_OpenPW As String
Dim m_EditPW As String
Dim m_FilePath As String
Dim m_FileFilter As String

'이벤트 선언:
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sQCGbn$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sRack$, sPos$)
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

Private Sub PhaseCfg_Protocol(ByVal strJobGbn As String)

    Dim strFile As String
    
    '--- 사용자 확인
    If m_EditPW <> pEditPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Sub
    End If
    '---------------
    
    If m_EqName = "0" Or m_EqName = "" Then
        RaiseEvent DispMsg("검사장비명을 지정해 주십시오.!!!")
        Exit Sub
    End If
    
    Select Case UCase(m_EqName)
        Case "ROTORGENE"
        
            With cdgFile
                .CancelError = False
                
                .InitDir = m_FilePath
                .FileName = "" 'm_FileFilter
                .Filter = m_FileFilter
                .Action = 1
                
                strFile = Trim(.FileName)
            End With
     
            If strFile = "" Or strFile = "*.*" Then Exit Sub
            
            If strJobGbn = "WK" Then
                Call DataEditResponse_RotorGeneWK(strFile)
            Else
                Call DataEditResponse_RotorGene(strFile)
            End If
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_RotorGene(ByVal strFileNm$)
    On Error GoTo ErrRtn
    
    Dim blnFlag As Boolean
    Dim strGetVal   As String
    Dim strRstVal   As String
    Dim strTmp  As String
    Dim intIdx  As Integer
    Dim sFlag   As String
    Dim objExcell   As Object
    
    Set objExcell = CreateObject("Excel.Application")
    
    objExcell.Workbooks.Open strFileNm
    
    objExcell.Visible = False
    
    blnFlag = False
    
    intIdx = 0
    Do
        If iSpaceCnt = 30 Then
            iSpaceCnt = 0
        End If
        iSpaceCnt = iSpaceCnt + 2
    
        RaiseEvent DispMsg(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
        intIdx = intIdx + 1
        
        strGetVal = objExcell.Worksheets(1).range("B" & CStr(intIdx) & ":B" & CStr(intIdx)).Value
        strRstVal = objExcell.Worksheets(1).range("F" & CStr(intIdx) & ":F" & CStr(intIdx)).Value
        
        If sTestMode = "77" Then
            RaiseEvent PrintRcvLog("PID: " & strGetVal & ", RESULT: " & strRstVal)
        End If
        
        sFlag = ""
        If strRstVal = "" Then
'            strRstVal = "Negative"
            strRstVal = "< 50"
        ElseIf strRstVal = "< 50" Or strRstVal = "<50" Then     '2007/7/10 yk
            '그대로 표시
        Else
'            strRstVal = Format$(Mid$(strRstVal, 1, 1) & "." & Mid$(strRstVal, 2, 2), "0.0") & " X 10(" & CStr(Len(strRstVal) - 1) & ")"
            '<S--- Bug 수정...2007/7/10 yk
            If Left(Trim(strRstVal), 1) = ">" Or Left(Trim(strRstVal), 1) = "<" Then
                sFlag = Left(Trim(strRstVal), 1)
                strRstVal = Trim(Mid(Trim(strRstVal), 2))
            End If
            strRstVal = Format$(Mid$(strRstVal, 1, 1) & "." & Mid$(strRstVal, 2, 2), "0.0") & " X 10(" & CStr(Len(strRstVal) - 1) & ")"
            strRstVal = Trim(sFlag & " " & strRstVal)
            '>E------------
        End If
        
        If UCase(strGetVal) = "NAME" Then blnFlag = True
        
        If blnFlag Or strGetVal = "" Then Exit Do
        
        If strGetVal <> "" And UCase(strGetVal) <> "NAME" Then
            With pResultInfo
                .ID = objExcell.Worksheets(1).range("B" & CStr(intIdx) & ":B" & CStr(intIdx)).Value
                .SEQNO = objExcell.Worksheets(1).range("A" & CStr(intIdx) & ":A" & CStr(intIdx)).Value
                .RACK = ""
                .POS = ""
                .QCGBN = ""
                .RSTCNT = 1
                .IFCD = "1" & Chr(124)
                .RST1 = strRstVal & Chr(124)
                .RST2 = "" & Chr(124)
                .UNIT = "" & Chr(124)
                .FLAG = "" & Chr(124)
                    
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .QCGBN)
            End With

            Call Init_pResultInfo
            
        End If
    Loop
    objExcell.Workbooks.Close
    Set objExcell = Nothing

    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_RotorGeneWK(ByVal strFileNm$)
    On Error GoTo ErrRtn
    
    Dim blnFlag As Boolean
    Dim strSId$, strPos$
    Dim strTmp  As String
    Dim intIdx  As Integer
    
    Dim objExcell   As Object
    
    Set objExcell = CreateObject("Excel.Application")
    
    objExcell.Workbooks.Open strFileNm
    
    objExcell.Visible = False
    
    blnFlag = False
    
    intIdx = 0
    Do
        
        pSampleInfo.ID = ""
        pSampleInfo.RACK = ""
        pSampleInfo.POS = ""
        
        If iSpaceCnt = 30 Then
            iSpaceCnt = 0
        End If
        iSpaceCnt = iSpaceCnt + 2
    
        RaiseEvent DispMsg(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
        intIdx = intIdx + 1
        
        strPos = objExcell.Worksheets(1).range("A" & CStr(intIdx) & ":A" & CStr(intIdx)).Value
        strSId = objExcell.Worksheets(1).range("B" & CStr(intIdx) & ":B" & CStr(intIdx)).Value
        
        If sTestMode = "77" Then
            RaiseEvent PrintSendLog("POS: " & strPos & ", PID: " & strSId)
        End If
        
        pSampleInfo.ID = strSId
        pSampleInfo.RACK = ""
        pSampleInfo.POS = strPos
        
        If UCase(strPos) = "POSITION" Then blnFlag = True
        
        If blnFlag Or strPos = "" Then Exit Do
        
        If pSampleInfo.POS <> "" And UCase(pSampleInfo.POS) <> "POSITION" Then
            RaiseEvent RequestCurOrder(pSampleInfo.ID, "", pSampleInfo.POS)
            Call Sleep(1000)
            RaiseEvent SendOrderOK(pSampleInfo.ID, "", pSampleInfo.POS)

        End If
    Loop
    objExcell.Workbooks.Close
    Set objExcell = Nothing

    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
            Erase .IFCD
        End With
        
        Exit Sub
    End If
    
    ReDim tmpData(m_p_iOrdCnt) As String
    tmpData() = Split(m_p_sTIFCd, Chr(124))
    
    With pSampleInfo
        .ID = m_p_sID
        .SEQNO = m_p_sSeq
        .RACK = m_p_sRack
        .POS = m_p_sPos
        .ORDCNT = m_p_iOrdCnt
        
        ReDim .IFCD(.ORDCNT)
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
        
        .OTHER = m_p_sPatInfo
    End With
        
End Sub

'
'   결과정보 구조체 초기화
'
Private Sub Init_pResultInfo()
    
    With pResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .QCGBN = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
    End With
    
End Sub
Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol("")

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
    m_p_sPatInfo = PropBag.ReadProperty("p_sPatInfo", m_def_p_sPatInfo)
    m_FilePath = PropBag.ReadProperty("FilePath", m_def_FilePath)
    m_FileFilter = PropBag.ReadProperty("FileFilter", m_def_FileFilter)
    
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
    Call PropBag.WriteProperty("p_sPatInfo", m_p_sPatInfo, m_def_p_sPatInfo)
    Call PropBag.WriteProperty("FilePath", m_FilePath, m_def_FilePath)
    Call PropBag.WriteProperty("FileFilter", m_FileFilter, m_def_FileFilter)
    
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get PortOpen() As Boolean
    PortOpen = m_PortOpen
End Property

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
    m_PortOpen = m_def_PortOpen
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
    m_p_sPatInfo = m_def_p_sPatInfo
    m_FilePath = m_def_FilePath
    m_FileFilter = m_def_FileFilter
    
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EqName() As String
    EqName = m_EqName
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get FielPath() As String
    FilePath = m_FilePath
End Property


'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get FielFilter() As String
    FileFilter = m_FileFilter
End Property

Public Property Let EqName(ByVal New_EqName As String)
    m_EqName = New_EqName
    PropertyChanged "EqName"
End Property

Public Property Let FilePath(ByVal New_FilePath As String)
    m_FilePath = New_FilePath
    PropertyChanged "FilePath"
End Property

Public Property Let FileFilter(ByVal New_FileFilter As String)
    m_FileFilter = New_FileFilter
    PropertyChanged "FileFilter"
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
'MemberInfo=14
Public Function GetData() As Variant
    On Error GoTo ErrComm
    
    Call PhaseCfg_Protocol("")
    
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("GetData 에러 - " & Err.Description)
    End If
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function GetWorkList() As Variant
    On Error GoTo ErrComm
    
    Call PhaseCfg_Protocol("WK")
    
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("GetWorkList 에러 - " & Err.Description)
    End If
End Function


'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get p_sPatInfo() As Variant
    p_sPatInfo = m_p_sPatInfo
End Property

Public Property Let p_sPatInfo(ByVal New_p_sPatInfo As Variant)
    m_p_sPatInfo = New_p_sPatInfo
    PropertyChanged "p_sPatInfo"
End Property

