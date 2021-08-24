VERSION 5.00
Begin VB.UserControl AMPLICOR 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2100
   LockControls    =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   2100
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "AMPLICOR"
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
Attribute VB_Name = "AMPLICOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTInstID$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTInstID$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


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
Private Sub DataEdit_AMPLICOR_AL241_Old(ByVal sRstData As String)
'    On Error GoTo ErrRtn
'
'    Dim tmpData()   As String
'    Dim tmpField()  As String
'    Dim ii%
'    Dim tmpBarCd$, tmpRack$, tmpPos$, tmpKind$
'    Dim tmpIFCd$, tmpRst1$, tmpRst2$, tmpUnit$, tmpFlag$, tmpInstID$
'
'    tmpData() = Split(sRstData, Chr(13))
'
'    For ii = 0 To UBound(tmpData())
'        If Left(tmpData(ii), 1) = Chr(10) Then
'            tmpData(ii) = Mid(tmpData(ii), 2)
'        End If
'        tmpData(ii) = Replace(tmpData(ii), Chr(34), "")
'
'        If Trim(tmpData(ii)) = "" Then
'            Exit Sub
'        End If
'
'        '결과구조체 초기화
'        Call Init_pResultInfo
'
'        Erase tmpField()
'
'        tmpField() = Split(tmpData(ii), Chr(9))
''---------- <ver AL2.41>
''         0  "Patient Name"
''         1  "Patient ID"
''         2  "Order/Lot Number"
''         3  "Sample ID"
''         4  "Clip#"
''         5  "Test"
''         6  "Result"
''         7  "Unit"
''         8  "Flags"
''         9  "Date/Time"
''         10 "Batch"
''         11 "Tube Position"
''         12 "Instrument ID"
''         13 "Preparation Instrument ID"
''         14 "Prep. Date/Time"
''         15 "Accepted Op"
''         16 "Accepted Date/Time"
''         17 "Comment"
''         18 "Raw Data"
''----------------------
'
'        If Trim(tmpField(5)) = "" Then
'            Exit For
'        End If
'
'        If Trim(tmpField(3)) = "Sample ID" Or Trim(tmpField(3)) = "" Then
'        Else
'            '결과 편집
'            tmpBarCd = Trim(tmpField(3))
'            tmpRack = ""
'            tmpPos = Trim(tmpField(11))     'TUBE
'
'            Select Case Left(Trim(tmpField(5)), 1)
'                Case "+", "-", "#"  'QC Result
'                    tmpKind = "QC"
'                    tmpIFCd = Mid(Trim(tmpField(5)), 2)
'                    tmpRst1 = Trim(tmpField(6))
'                    If tmpRst1 = "" Then
'                        tmpRst1 = Trim(tmpField(19))
'                    End If
'
'                Case Else
'                    tmpKind = ""
'                    tmpIFCd = Trim(tmpField(5))
'                    tmpRst1 = Trim(tmpField(6))
'            End Select
'
'            tmpRst2 = ""
'            If UBound(tmpField()) >= 19 Then
'                If Trim(tmpField(19)) <> "" Then
'                    tmpRst2 = Trim(tmpField(19))    'OD value
'                End If
'            End If
'            tmpUnit = Trim(tmpField(7))
'
'            If tmpUnit <> "-" Then
'                '정량결과(예: "1.10E+6")
'                If Left(tmpRst1, 1) = "*" Then
'                    tmpRst1 = ""
'                Else
'                    tmpRst1 = Val(tmpRst1)
'                End If
'            End If
'
'            tmpFlag = Trim(tmpField(8))
'            tmpInstID = Trim(tmpField(12))  'Inst ID
'
'            '결과정보 구조체에 저장
'            With pResultInfo
'                .ID = tmpBarCd
'                .SEQNO = ""
'                .RACK = tmpRack
'                .POS = tmpPos
'                .KIND = tmpKind
'
'                '결과값 누적
'                .RSTCNT = .RSTCNT + 1
'                .IFCD = .IFCD & tmpIFCd & Chr(124)
'                .RST1 = .RST1 & tmpRst1 & Chr(124)
'                .RST2 = .RST2 & tmpRst2 & Chr(124)
'                .UNIT = .UNIT & tmpUnit & Chr(124)
'                .FLAG = .FLAG & tmpFlag & Chr(124)
'
'                .INSTID = .INSTID & tmpInstID & Chr(124)    'Instrument ID
'
'                '결과값 등록/화면 표시 처리...
'                If .RSTCNT > 0 Then
'                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, "", .KIND, "", "")
'                End If
'
'                Call Init_pResultInfo
'            End With
'        End If
'    Next ii
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg(Err.Description)
'    End If
End Sub

Private Sub DataEdit_AMPLICOR_AL241(ByVal sRstData As String)
    On Error GoTo ErrRtn

    Dim tmpData()   As String
    Dim tmpField()  As String
    Dim ii%
    Dim tmpBarCd$, tmpRack$, tmpPos$, tmpKind$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$, tmpUnit$, tmpFlag$, tmpInstID$

    tmpData() = Split(sRstData, Chr(13))

    For ii = 0 To UBound(tmpData())
        If Left(tmpData(ii), 1) = Chr(10) Then
            tmpData(ii) = Mid(tmpData(ii), 2)
        End If
        tmpData(ii) = Replace(tmpData(ii), Chr(34), "")

        If Trim(tmpData(ii)) = "" Then
            Exit Sub
        End If

        '결과구조체 초기화
        Call Init_pResultInfo

        Erase tmpField()

        tmpField() = Split(tmpData(ii), Chr(9))
'---------- <ver AL2.41>
'         0  "Patient Name"
'         1  "Patient ID"
'         2  "Order/Lot Number"
'         3  "Sample ID"
'         4  "Clip#"
'         5  "Test"
'         6  "Result"
'         7  "Unit"
'         8  "Flags"
'         9  "Date/Time"
'         10 "Batch"
'         11 "Tube Position"
'         12 "Instrument ID"
'         13 "Preparation Instrument ID"
'         14 "Prep. Date/Time"
'         15 "Accepted Op"
'         16 "Accepted Date/Time"
'         17 "Comment"
'         18 "Raw Data"
'----------------------

        If Trim(tmpField(5)) = "" Then
            Exit For
        End If

        If Trim(tmpField(3)) = "Sample ID" Or Trim(tmpField(3)) = "" Then
            '<S--- QC 결과 별도 처리...2007/6/18 yk
            If Trim(tmpField(3)) <> "Sample ID" And _
                    (Left(Trim(tmpField(5)), 1) = "+" Or Left(Trim(tmpField(5)), 1) = "-" Or Left(Trim(tmpField(5)), 1) = "#") Then
                '결과 편집
                tmpBarCd = Trim(tmpField(3))
                tmpRack = ""
                tmpPos = Trim(tmpField(11))     'TUBE
                
                tmpKind = "QC"
                tmpIFCd = Mid(Trim(tmpField(5)), 2)
                tmpRst1 = Trim(tmpField(6))
                If tmpRst1 = "" Then
                    tmpRst1 = Trim(tmpField(19))
                End If
                
                GoTo AppendRstRtn
            End If
            '>E--------------------
        Else
            '결과 편집
            tmpBarCd = Trim(tmpField(3))
            tmpRack = ""
            tmpPos = Trim(tmpField(11))     'TUBE

            Select Case Left(Trim(tmpField(5)), 1)
                Case "+", "-", "#"  'QC Result
                    tmpKind = "QC"
                    tmpIFCd = Mid(Trim(tmpField(5)), 2)
                    tmpRst1 = Trim(tmpField(6))
                    If tmpRst1 = "" Then
                        tmpRst1 = Trim(tmpField(19))
                    End If
                    
                Case Else
                    tmpKind = ""
                    tmpIFCd = Trim(tmpField(5))
                    tmpRst1 = Trim(tmpField(6))
            End Select
            
AppendRstRtn:
            tmpRst2 = ""
            If UBound(tmpField()) >= 19 Then
                If Trim(tmpField(19)) <> "" Then
                    tmpRst2 = Trim(tmpField(19))    'OD value
                End If
            End If
            tmpUnit = Trim(tmpField(7))

            If tmpUnit <> "-" Then
                '정량결과(예: "1.10E+6")
                If Left(tmpRst1, 1) = "*" Then
                    tmpRst1 = ""
                Else
                    tmpRst1 = Val(tmpRst1)
                End If
            End If

            tmpFlag = Trim(tmpField(8))
            tmpInstID = Trim(tmpField(12))  'Inst ID

            '결과정보 구조체에 저장
            With pResultInfo
                .ID = tmpBarCd
                .SEQNO = ""
                .RACK = tmpRack
                .POS = tmpPos
                .KIND = tmpKind
                
                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst1 & Chr(124)
                .RST2 = .RST2 & tmpRst2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)

                .INSTID = .INSTID & tmpInstID & Chr(124)    'Instrument ID

                '결과값 등록/화면 표시 처리...
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, "", .KIND, "", "")
                End If

                Call Init_pResultInfo
            End With
        End If
    Next ii

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub


Private Sub OpenResultDataFile_AMPLICOR(ByVal sFileNm As String)
    On Error GoTo ErrRtn
    
    Dim sRcvBuffer  As String
    
    'FILE OPEN
    Close #1
    Open sFileNm For Input As #1
    sRcvBuffer = ""
    Do While Not EOF(1)
        sRcvBuffer = sRcvBuffer & Input(1, #1)
    Loop
    Close #1
    
    If Trim(sRcvBuffer) = "" Then
        Exit Sub
    End If
    
    '결과 편집
    If m_sVersion = "AL2.41" Then
        Call DataEdit_AMPLICOR_AL241(sRcvBuffer)
    Else
        Call DataEdit_AMPLICOR(sRcvBuffer)
    End If
    
ErrRtn:
    If Err <> 0 Then
        MsgBox Err.Description, vbExclamation
    End If
End Sub



Private Sub PhaseCfg_Protocol()

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
        Case "AMPLICOR"
            If m_sVersion = "AL2.41" Then
                Call PhaseCfg_Protocol_AMPLICOR_AL241
            ElseIf m_sVersion = "AL2.3" Then
                Call PhaseCfg_Protocol_AMPLICOR_AL23
            End If
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub DataEdit_AMPLICOR(ByVal sRstData As String)
    On Error GoTo ErrRtn

    Dim tmpData()   As String
    Dim tmpField()  As String
    Dim ii%
    Dim tmpBarCd$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$, tmpUnit$, tmpFlag$, tmpInstID$

    tmpData() = Split(sRstData, Chr(13))

    For ii = 0 To UBound(tmpData())
        If Left(tmpData(ii), 1) = Chr(10) Then
            tmpData(ii) = Mid(tmpData(ii), 2)
        End If
        tmpData(ii) = Replace(tmpData(ii), Chr(34), "")

        If Trim(tmpData(ii)) = "" Then
            Exit Sub
        End If

        '결과구조체 초기화
        Call Init_pResultInfo

        Erase tmpField()

        tmpField() = Split(tmpData(ii), Chr(9))

        If Trim(tmpField(1)) = "" Then
            Exit For
        End If

        If Trim(tmpField(0)) = "SampleId" Or Trim(tmpField(0)) = "" _
                Or Left(tmpField(1), 1) = "+" Or Left(tmpField(1), 1) = "-" Then
        Else
            '결과 편집
            tmpBarCd = Trim(tmpField(0))
            tmpRack = Trim(tmpField(6))
            tmpPos = Trim(tmpField(7))
            tmpIFCd = Trim(tmpField(1))
            tmpRst1 = Trim(tmpField(2))
            tmpRst2 = Trim(tmpField(13))
            tmpUnit = Trim(tmpField(3))
            tmpFlag = Trim(tmpField(4))
            tmpInstID = Trim(tmpField(9))   'Inst ID

            '결과정보 구조체에 저장
            With pResultInfo
                .ID = tmpBarCd
                .SEQNO = ""
                .RACK = tmpRack
                .POS = tmpPos

                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst1 & Chr(124)
                .RST2 = .RST2 & tmpRst2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)

                .INSTID = .INSTID & tmpInstID & Chr(124)    'Instrument ID

                '결과값 등록/화면 표시 처리...
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, "", "", "", "")
                End If

                Call Init_pResultInfo
            End With
        End If
    Next ii

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub


'
'   AMPLICOR AL 2.3
'
Private Sub PhaseCfg_Protocol_AMPLICOR_AL23()
'
'    Dim wkDat   As String
'    Dim ix1     As Integer
'
'    For ix1 = 1 To Len(wkBuf)
'        wkDat = Mid$(wkBuf, ix1, 1)
'
'        Select Case m_iPhase
'            Case 1
'                Select Case Asc(wkDat)
'                    Case 5      'ENQ
'                        MSComm.Output = Chr(6)
'                        m_iPhase = 2
'                    Case Else
'                        m_iPhase = 1
'                End Select
'
'            Case 2
'                Select Case Asc(wkDat)
'                    Case 2      'STX
'                        RcvBuffer = ""
'
'                    Case 10     '<LF>
'                        Call DataEditResponse_AMPLICOR
'
'                        m_iPhase = 2
'                        MSComm.Output = Chr(6)
'
'                    Case 4      'EOT
'                        m_iPhase = 1
'
'                    Case 5      'ENQ
'                        MSComm.Output = Chr(6)   'Send ACK
'
'                    Case 21     'NAK
'                        MSComm.Output = Chr(5)   'Send ENQ
'                        m_iPhase = 1
'
'                    Case Else
'                        RcvBuffer = RcvBuffer & wkDat
'                        m_iPhase = 2
'                End Select
'
'            Case 3
'                Select Case Asc(wkDat)
'                    Case 6      'ACK
'                        Call SendOrder_Elecsys2010      'Order 전송
'
'                    Case 5      'ENQ
'                        m_iPhase = 2
'                        MSComm.Output = Chr(6)
'
'                    Case 21     'NAK
'                        m_iSendPhase = m_iSendPhase - 1
'                        m_iFrameN = m_iFrameN - 1
'                        m_iPhase = 3
'
'                        Call SendOrder_Elecsys2010      'Order 전송
'
'                    Case 4      'EOT
'                        m_iPhase = 1
'
'                End Select
'        End Select
'    Next ix1
'
End Sub
'
'   AMPLICOR AL 2.41
'
Private Sub PhaseCfg_Protocol_AMPLICOR_AL241()
'
'    Dim wkDat   As String
'    Dim ix1     As Integer
'
'    For ix1 = 1 To Len(wkBuf)
'        wkDat = Mid$(wkBuf, ix1, 1)
'
'        Select Case m_iPhase
'            Case 1
'                Select Case Asc(wkDat)
'                    Case 5      'ENQ
'                        MSComm.Output = Chr(6)
'                        m_iPhase = 2
'                    Case Else
'                        m_iPhase = 1
'                End Select
'
'            Case 2
'                Select Case Asc(wkDat)
'                    Case 2      'STX
'                        RcvBuffer = ""
'
'                    Case 10     '<LF>
'                        Call DataEditResponse_AMPLICOR
'
'                        m_iPhase = 2
'                        MSComm.Output = Chr(6)
'
'                    Case 4      'EOT
'                        m_iPhase = 1
'
'                    Case 5      'ENQ
'                        MSComm.Output = Chr(6)   'Send ACK
'
'                    Case 21     'NAK
'                        MSComm.Output = Chr(5)   'Send ENQ
'                        m_iPhase = 1
'
'                    Case Else
'                        RcvBuffer = RcvBuffer & wkDat
'                        m_iPhase = 2
'                End Select
'
'            Case 3
'                Select Case Asc(wkDat)
'                    Case 6      'ACK
'                        Call SendOrder_Elecsys2010      'Order 전송
'
'                    Case 5      'ENQ
'                        m_iPhase = 2
'                        MSComm.Output = Chr(6)
'
'                    Case 21     'NAK
'                        m_iSendPhase = m_iSendPhase - 1
'                        m_iFrameN = m_iFrameN - 1
'                        m_iPhase = 3
'
'                        Call SendOrder_Elecsys2010      'Order 전송
'
'                    Case 4      'EOT
'                        m_iPhase = 1
'
'                End Select
'        End Select
'    Next ix1
'
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function RcvRstData(sFileNm$) As Variant
    
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

    If Trim(sFileNm) = "" Then
        Exit Function
    End If

    'Data File Open
    Call OpenResultDataFile_AMPLICOR(sFileNm)
        
End Function

