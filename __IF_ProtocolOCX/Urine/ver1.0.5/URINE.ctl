VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl URINE 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
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
   Begin MSCommLib.MSComm msComm 
      Left            =   255
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "URINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
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
'속성 변수:
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
'이벤트 선언:
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$)
Event DispMsg(sMsg$)
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


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
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

'For MiditronM
Dim bRstFlag    As Boolean
Dim msSampleNo  As String       '단방향에서 장비번호 체크 사용
Dim msPreSampleNo   As String   '              "

'For Urisys2400
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim RstEnd  As String

Private Sub PhaseCfg_Protocol_UriScan()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2         ' STX
                RcvBuffer = ""
                
            Case 3         ' ETX
                '--- 결과편집/등록
                Call DataEdit_UriScan
                                            
                Do
                    DoEvents
                Loop Until msComm.OutBufferCount = 0
                
            Case Else      '
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1

End Sub
Private Sub PhaseCfg_Protocol_Urometer()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 13       ' CR
                '--- 결과편집/등록
                Call DataEdit_Urometer
                
                RcvBuffer = ""
                                            
                Do
                    DoEvents
                Loop Until msComm.OutBufferCount = 0
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1

End Sub
Private Sub DataEdit_Urometer()
    On Error GoTo ErrRtn
    
    Dim ii      As Integer
    Dim tmpSeqNo    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer
    
    Dim RecType$   'Record Type
    
    RecType = Mid$(RcvBuffer, 1, 3)
    
    Select Case RecType
        Case "ID "
            '결과정보 구조체 초기화
            Call Init_pResultInfo
            
            tmpSeqNo = Trim(Mid(RcvBuffer, 8, 4))
            
        Case "BLD", "BIL", "URO", "KET", "PRO", "NIT", "GLU", "LEU"
        
            sRst1 = Trim(Mid(RcvBuffer, 7, 5))
            If UCase(sRst1) = "NEG" Then sRst1 = "-"
            
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & RecType & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
            
            If RecType = "LEU" Then
                '--- MICRO 자동입력
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 3
                    .IFCD = .IFCD & "WBC|RBC|EPI|"
                    .RST1 = .RST1 & "0-1|0-1|0-1|"
                    .RST2 = .RST2 & "|||"
                    .UNIT = .UNIT & "|||"
                    .FLAG = .FLAG & "|||"
                End With
                
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    .SEQNO = tmpSeqNo
                    
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                    End If
                End With
            
            End If
            
        Case "p.H", "S.G"
        
            sRst1 = Trim(Mid(RcvBuffer, 12, 7))
            
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & RecType & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With

        Case Else
    End Select
        
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Sub DataEdit_UriScan()
    On Error GoTo ErrRtn
    
    Dim ii      As Integer
    Dim tmpSeqNo    As String
    Dim tmpBarCd    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer

    Dim aRowData()  As String
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    '설정된 장비코드 편집
    m_p_sTIFCd = "BLD|BIL|URO|KET|PRO|NIT|GLU|p.H|S.G|LEU|VTC|"
    ReDim sTestCd(10) As String
    sTestCd() = Split(m_p_sTIFCd, Chr(124))
    
    iPos = InStr(RcvBuffer, "ID_NO:")
    If iPos <> 0 Then
        If InStr(Mid(RcvBuffer, iPos), Chr(13)) > 0 Then
            aRowData() = Split(Mid(RcvBuffer, iPos + 6), Chr(13))
                
            tmpSeqNo = Trim(Mid(aRowData(0), 1, 4))
            tmpBarCd = Trim(Mid(aRowData(0), 6))
        Else
            tmpSeqNo = Trim(Mid(RcvBuffer, iPos + 6, 4))
        End If
    End If
    
    '--- 결과값 편집
    For ii = 1 To 11
        sIFCd = Trim(sTestCd(ii - 1))
                
        iPos = InStr(RcvBuffer, sIFCd)
        
        If iPos = 0 Then
        Else
            sRstData = Mid$(RcvBuffer, iPos, 15)
            sRst1 = Trim$(Mid$(sRstData, 11, 5))
                        
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & sIFCd & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
        End If
    Next ii

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = tmpSeqNo
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
        
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,CommPort
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "통신 포트 번호를 반환하거나 설정합니다."
    CommPort = msComm.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    msComm.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

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
        Case "CLINITEK500"
            Call PhaseCfg_Protocol_Clinitek500
            
        Case "MIDITRONM"
            Call PhaseCfg_Protocol_MiditronM
        
        Case "MIDITRONJR"
            Call PhaseCfg_Protocol_MiditronJr
        
        Case "URISCAN"
            Call PhaseCfg_Protocol_UriScan
            
        Case "UROMETER"
            Call PhaseCfg_Protocol_Urometer
        
        Case "URISYS2400"
            Call PhaseCfg_Protocol_URISYS2400
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_URISYS2400()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5  'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False
                        
                        msComm.Output = Chr(6)
                        
                    Case Else
                    
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True
                        
                    Case 10     'LF
                        If bEndChk = True Then
                            Call DataEditResponse_URISYS2400
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                        
                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_URISYS2400
                            RcvBuffer = ""
                        End If
                    
                    Case 3      'ETX
                        msComm.Output = Chr(6)
                    
                    Case 4      'EOT
                        RcvBuffer = ""
                        m_iPhase = 1
                    
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                    
                    Case 21     'NAK
                        Call DataEditResponse_URISYS2400
                    
                    Case 23     'ETB
                        bEndChk = False
                        
                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkDat
                            End If
                        End If
                        
                End Select

            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
'                        Call Order_Input
                    
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    
                    Case 21     'NAK
                        msComm.Output = Chr(5)   'ENQ
'                        Call Order_Input
                        m_iPhase = 3
                        
                    Case 4      'EOT
                        m_iPhase = 1
                        
                End Select
        End Select
    Next ix1

End Sub
' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_URISYS2400()
'    On Error GoTo ErrRtn
'
'    Dim RecType     As String   'Record Type
'
'    Dim tmpData()   As String
'    Dim tmpField()  As String
'
'    Dim Tmp     As String
'    Dim i       As Integer
'    Dim ChkR    As Integer
'    Dim tData() As String
'
'    Dim sRetVal As String
'
''    Dim tmpData As String
'    Dim tmpRes  As String
'    Dim sSampNo As String
'    Dim sRack   As String
'    Dim sPos    As String
'    Dim sType   As String
'    Dim sEqCd   As String
'    Dim sResult As String
'    Dim iCurRow As Integer
'    Dim vTmp    As Variant
'    Dim sTestCd As String
'
'    i = InStr(1, RcvBuffer, "|")
'    If i <> 0 Then
'        RecType = Mid$(RcvBuffer, i - 1, 1)
'    Else
'        Exit Sub
'    End If
'
'    Select Case RecType
'        Case "H"        'Header Record
'        Case "M"
'        Case "P"        'Patient Record
'            Call Init_pResultInfo
'
'        Case "Q"        'Order Request Record
'            ' URISYS2400 장비에선 사용하지 않음...
'
'        Case "O"
'            tmpBarCd = ""
'            tmpSeqNo = "": tmpRack = "": tmpPos = "": TMPType = ""
'
'            tmpData() = Split(RcvBuffer, Chr(124))
'
'            tmpBarCd = Trim(tmpData(2))
'
'            If Trim(tmpData(3)) = "" Then Exit Sub
'            ii = InStr(tmpData(3), "^")
'            If ii <> 0 Then
'                tmpField() = Split(Trim(tmpData(3)), "^")
'
'                tmpSeqNo = Trim(tData(0))
'                tmpRack = Trim(tData(1))
'                tmpPos = Trim(tData(2))
'                TMPType = Trim(tData(4))        'SAMPLE/CONTROL
'            End If
'
'            pSampleInfo.ID = UCase(tmpBarCd)
'            pSampleInfo.RACK = tmpRack
'            pSampleInfo.POS = tmpPos
'            pSampleInfo.KIND = tmpKind
'
'        Case "R"        'Result Record
'            '--- 결과데이타 편집
'            '2: TEST ID
'            '3: Result
'            '4: UNIT
'            '6: Result Abnormal Flag
'            '8: Result Status
'
'            ' R|5|^^^5|0.25|g/L|||||(CR)
'            '*7) Test No (fixation)
'            '1 : SG,  2 : pH,  3 : LEU, 4 : NIT, 5 : PRO, 6 : GLU,7 : KET, 8 : UBG, 9 : BIL, 10: ERY,
'            '11: COL , 12: CLA
'            '*8) Measurement data (range data of item corresponding to test no.)
'            'If *10) result status is .X., this field will be unassigned.
'            '*9) Sample flags
'            '! :  ID Edit
'            'N:   Sample Short
'            'E:  Sample Empty
'            'R:  Test Strip
'            '*10) Result status: No value with .X.
'
'            Erase tmpData()
'
'            tmpData() = Split(RcvBuffer, "|")
'
'            tmpData = Trim(tData(2))
'            sResult = Trim(tData(3))
'
'            Erase tData()
'
'            tData() = Split(tmpData, "^")
'            sTestCd = Trim(tData(3))
'
'
'
'        Case "C"        'Comment Record
'
'        Case "L"
'            If Trim(sTotTestCd) = "" Then Exit Sub
''            Text2 = iTestCnt & "/ " & sTotTestCd & "/ " & sTotTestNm & "/ " & sTotRst
'            '--- 화면표시
'            sRetVal = ViewResults(0, iTestCnt, LabDate, SlipCd, LabSeq, sTotTestCd, sTotTestNm, sTotRst)
'
'            '--- 결과등록
'            ''gsTxMode="0" => Batch, gsTxMode="1" => RealTime(한 항목씩), gsTxMode="2" => RealTime(한 환자씩)
'            If gsTXMode = "0" Then
'                If sRetVal = "NONE" Then
'                Else
'                    Call SpdForeBack(spdIntList, 5, 10, gResultTable(1).iCRow, _
'                             gResultTable(1).iCRow, RGB(0, 0, 0), 연노랑)
'                End If
'            ElseIf gsTXMode = "1" Then
'                If sRetVal = "NONE" Then
'                ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
'                    Call SpdForeBack(spdIntList, 5, 10, gResultTable(1).iCRow, _
'                             gResultTable(1).iCRow, RGB(0, 0, 0), 연노랑)
'
'                    If cmdReg.Visible = False Then
'                        Call RegResult(0, LabDate, SlipCd, LabSeq)
'                        'RegResult가 성공하면 연초록으로 바뀜
'                    Else
'                        Exit Sub
'                    End If
'
'                    If sRetVal = "DONE" Then
'                        If gsAPMode = "0" Then
'                            Call AutoResultEnd(LabDate, SlipCd, LabSeq)
'                        ElseIf gsAPMode = "1" Then
'                            Call AutoPrint(LabDate, SlipCd, LabSeq)
'                        ElseIf gsAPMode = "2" Then
'                            Call AutoEmerPrint(LabDate, SlipCd, LabSeq)
'                        End If
'                    End If
'                End If
'            ElseIf gsTXMode = "2" Then
'                If sRetVal = "NONE" Then
'                ElseIf sRetVal = "MORE" Or sRetVal = "DONE" Then
'                    Call SpdForeBack(spdIntList, 5, 10, gResultTable(1).iCRow, _
'                             gResultTable(1).iCRow, RGB(0, 0, 0), 연노랑)
'
'                    '환자단위로 결과 등록 시 사용
'                    If cmdReg.Visible = False Then
'                        Call RegResult(1, LabDate, SlipCd, LabSeq)
'                        'RegResult가 성공하면 연초록으로 바뀜
'                    Else
'                        Exit Sub
'                    End If
'
'                    If gsAPMode = "0" Then
'                        Call AutoResultEnd(LabDate, SlipCd, LabSeq)
'                    ElseIf gsAPMode = "1" Then
'                        Call AutoPrint(LabDate, SlipCd, LabSeq)
'                    ElseIf gsAPMode = "2" Then
'                        Call AutoEmerPrint(LabDate, SlipCd, LabSeq)
'                    End If
'                End If
'
'            End If
'    End Select
'
'    Exit Sub
'ErrRtn:
'    ViewMsg "Edit_Data 에러 발생" & "(" & CStr(Err.Number) & ")"
End Sub



Private Sub PhaseCfg_Protocol_Clinitek500()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2         ' STX
                RcvBuffer = ""
                
            Case 3         ' ETX
                '--- 결과편집/등록
                Call DataEdit_Clinitek500
                                            
            Case 21        ' NAK
                
            Case Else      '
                RcvBuffer = RcvBuffer + wkDat
                
         End Select
    Next ix1
    
End Sub
Private Sub DataEdit_MiditronJr()
    On Error GoTo ErrRtn
    
    Dim ii      As Integer
    Dim tmpSeqNo    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer
    Dim iPos1   As Integer
    Dim iPos2   As Integer
    Dim sTmp    As String
    Dim iTmp    As Integer
    
    
    If Left$(RcvBuffer, 1) = "<" Then
        bRstFlag = True
        RcvBuffer = ""
        Exit Sub
    ElseIf Left$(RcvBuffer, 1) = ":" Then
        bRstFlag = False
        Exit Sub
    End If
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    '설정된 장비코드 편집
    m_p_sTIFCd = "SG|PH|LEU|NIT|PRO|GLU|KET|UBG|BIL|ERY|"
    ReDim sTestCd(m_p_iOrdCnt) As String
    sTestCd() = Split(m_p_sTIFCd, Chr(124))
    
    tmpSeqNo = Trim(Mid(RcvBuffer, 16, 4))
    
    '--- 같은 환자 여러번 넘어올 때의 처리
    '검사시간에 해당
    msSampleNo = Mid(RcvBuffer, 16, 19)
    
    '이전과 검사시간이 같은지 체크
    If msSampleNo = msPreSampleNo Then Exit Sub
    '-------------------------------------
    
    '--- 결과값 편집
    iPos = 1
    For ii = 1 To 10
        sIFCd = Trim(sTestCd(ii - 1))
        If Trim(sIFCd) <> "" Then
            '설정된 장비코드에 해당하는 결과값 조회
            If ii = 10 Then
                iPos1 = InStr(iPos, RcvBuffer, Trim(sTestCd(ii - 1)))
                iPos2 = InStr(iPos1 + 1, RcvBuffer, "NAG")
            Else
                iPos1 = InStr(iPos, RcvBuffer, Trim(sTestCd(ii - 1)))
                iPos2 = InStr(iPos1 + 1, RcvBuffer, Trim(sTestCd(ii)))
            End If
            
            If iPos1 = 0 And iPos2 = 0 Then
                Exit For
            End If
            
            sRstData = Mid(RcvBuffer, iPos1 + Len(sIFCd), iPos2 - iPos1 - Len(sIFCd))
            
            iTmp = InStr(Trim(sRstData), Space(1))
            If iTmp = 0 Then
                sRst1 = Trim(sRstData)
                sRst2 = ""
                sUnit = ""
            Else
                sRst1 = Trim(Mid(Trim(sRstData), 1, iTmp - 1))
                sRst2 = Trim(Mid(Trim(sRstData), iTmp + 1))
                
                iTmp = InStr(Trim(sRst2), Space(1))
                If iTmp = 0 Then
                    iTmp = InStr(Trim(sRst2), "/")
                    If iTmp = 0 Then
                        sUnit = ""
                    Else
                        sUnit = Trim(sRst2)
                        sRst2 = ""
                    End If
                Else
                    '단위가 포함된 경우 편집
                    sUnit = Trim(Mid(sRst2, 1, iTmp - 1))
                    sRst2 = Trim(Mid(sRst2, iTmp + 1))
                End If
            End If
            
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & sIFCd & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
            
            iPos = iPos1
        End If
    Next ii
    
    msPreSampleNo = msSampleNo
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .SEQNO = tmpSeqNo
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
                                            
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Sub DataEdit_MiditronM()
    On Error GoTo ErrRtn
    
    Dim ii      As Integer
    Dim tmpSeqNo    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer
    Dim iPos1   As Integer
    Dim iPos2   As Integer
    Dim sTmp    As String
    Dim iTmp    As Integer
    
    
    If Left$(RcvBuffer, 1) = "<" Then
        bRstFlag = True
        RcvBuffer = ""
        Exit Sub
    ElseIf Left$(RcvBuffer, 1) = ":" Then
        bRstFlag = False
        Exit Sub
    End If
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    '설정된 장비코드 편집
    m_p_sTIFCd = "SG|PH|LEU|NIT|PRO|GLU|KET|UBG|BIL|ERY|"
    ReDim sTestCd(m_p_iOrdCnt) As String
    sTestCd() = Split(m_p_sTIFCd, Chr(124))
    
    tmpSeqNo = Trim(Mid(RcvBuffer, 16, 4))
    
    '--- 같은 환자 여러번 넘어올 때의 처리
    '검사시간에 해당
    msSampleNo = Mid(RcvBuffer, 16, 19)
    
    '이전과 검사시간이 같은지 체크
    If msSampleNo = msPreSampleNo Then Exit Sub
    '-------------------------------------
    
    '--- 결과값 편집
    iPos = 1
    For ii = 1 To 10    'm_p_iOrdCnt
        sIFCd = Trim(sTestCd(ii - 1))
        If Trim(sIFCd) <> "" Then
            '설정된 장비코드에 해당하는 결과값 조회
            If ii = 10 Then
                iPos1 = InStr(iPos, RcvBuffer, Trim(sTestCd(ii - 1)))
                iPos2 = InStr(iPos1 + 1, RcvBuffer, "NAG")
            Else
                iPos1 = InStr(iPos, RcvBuffer, Trim(sTestCd(ii - 1)))
                iPos2 = InStr(iPos1 + 1, RcvBuffer, Trim(sTestCd(ii)))
            End If
            
            If iPos1 = 0 And iPos2 = 0 Then
                Exit For
            End If
            
            sRstData = Mid(RcvBuffer, iPos1 + Len(sIFCd), iPos2 - iPos1 - Len(sIFCd))
            
            iTmp = InStr(Trim(sRstData), Space(1))
            If iTmp = 0 Then
                sRst1 = Trim(sRstData)
                sRst2 = ""
                sUnit = ""
            Else
                sRst1 = Trim(Mid(Trim(sRstData), 1, iTmp - 1))
                sRst2 = Trim(Mid(Trim(sRstData), iTmp + 1))
                
                iTmp = InStr(Trim(sRst2), Space(1))
                If iTmp = 0 Then
                    iTmp = InStr(Trim(sRst2), "/")
                    If iTmp = 0 Then
                        sUnit = ""
                    Else
                        sUnit = Trim(sRst2)
                        sRst2 = ""
                    End If
                Else
                    '단위가 포함된 경우 편집
                    sUnit = Trim(Mid(sRst2, 1, iTmp - 1))
                    sRst2 = Trim(Mid(sRst2, iTmp + 1))
                End If
            End If
            
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & sIFCd & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
            
            iPos = iPos1
        End If
    Next ii
    
    msPreSampleNo = msSampleNo
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .SEQNO = tmpSeqNo
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
                                            
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_Clinitek500_Old()
'    On Error GoTo ErrRtn
'
'    Dim iPos    As Integer
'    Dim ii      As Integer
'    Dim tmpSeqNo$
'    Dim sTestCd()   As String
'    Dim sRstData    As String
'    Dim sIFCd   As String
'    Dim sRst    As String
'    Dim iTmp    As Integer
'
'
'    '결과정보 구조체 초기화
'    Call Init_pResultInfo
'
'    '설정된 장비코드 편집
'    ReDim sTestCd(m_p_iOrdCnt) As String
'    sTestCd() = Split(m_p_sTIFCd, Chr(124))
'
'
'    '--- Data 편집
'    'Chr(13) & chr(10)을 절삭
'    RcvBuffer = Mid$(RcvBuffer, 3)
'
'    'SerialNo, 날짜 절삭
'    tmpSeqNo = Trim(Mid(RcvBuffer, 2, 6))
'    iPos = InStr(RcvBuffer, Chr(10))
'    RcvBuffer = Mid$(RcvBuffer, iPos + 1)
'
''    'SampleNo 얻기...
''    iPos = InStr(RcvBuffer, Chr(10))
''    RcvBuffer = Mid$(RcvBuffer, iPos + 1)
'
''    'Color : 절삭...
''    cnt = InStr(RcvBuffer, Chr$(10))
''    RcvBuffer = Mid$(RcvBuffer, cnt + 1)
''
''    'Clarity : 절삭...
''    cnt = InStr(RcvBuffer, Chr$(10))
''    RcvBuffer = Mid$(RcvBuffer, cnt + 1)
'
'    pResultInfo.SEQNO = tmpSeqNo
'
'    Do While True
'        iPos = InStr(RcvBuffer, Chr(10))
'        If iPos = 0 Then
'            Exit Do
'        End If
'
'        sRstData = Left$(RcvBuffer, iPos - 2)
'        RcvBuffer = Mid$(RcvBuffer, iPos + 1)
'
'        '--- 결과값 편집
'        For ii = 1 To m_p_iOrdCnt
'            sIFCd = Trim(sTestCd(ii - 1))
'            If Trim(sIFCd) <> "" Then
'                '설정된 장비코드에 해당하는 결과값 조회
'                iPos = InStr(1, sRstData, Trim(sTestCd(ii - 1)))
'                If iPos <> 0 Then
'                    '결과값
'                    sRst = Trim(Mid(sRstData, Len(sIFCd) + 1))
'                    If Left(sRst, 1) = "*" Then
'                        sRst = Trim(Mid(sRst, 2))
'                    End If
'                    iTmp = InStr(1, sRst, Space(1))
'                    If iTmp > 0 Then
'                        sRst = Trim(Left(sRst, iTmp - 1))
'                    End If
'
'                    'Data 누적
'                    With pResultInfo
'                        .RSTCNT = .RSTCNT + 1
'                        .IFCD = .IFCD & sIFCd & Chr(124)
'                        .RST1 = .RST1 & sRst & Chr(124)
'                        .RST2 = .RST2 & Chr(124)
'                        .UNIT = .UNIT & Chr(124)
'                        .FLAG = .FLAG & Chr(124)
'                    End With
'
'                    Exit For
'                End If
'            End If
'        Next ii
'    Loop
'
'    '결과값 등록/화면 표시 처리...
'    With pResultInfo
'        If .RSTCNT > 0 Then
'            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
'        End If
'    End With
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
'    End If
End Sub

Private Sub DataEdit_Clinitek500()
    On Error GoTo ErrRtn
    
    Dim iPos    As Integer
    Dim ii      As Integer
    Dim tmpSeqNo$
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst    As String
    Dim iTmp    As Integer
    
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    '설정된 장비코드 편집
    m_p_sTIFCd = "SG|pH|LEU|NIT|PRO|GLU|KET|UBG|BIL|BLD|"
    ReDim sTestCd(m_p_iOrdCnt) As String
    sTestCd() = Split(m_p_sTIFCd, Chr(124))
    
    
    '--- Data 편집
    'Chr(13) & chr(10)을 절삭
    RcvBuffer = Mid$(RcvBuffer, 3)
    
    'SerialNo, 날짜 절삭
    tmpSeqNo = Trim(Mid(RcvBuffer, 2, 6))
    iPos = InStr(RcvBuffer, Chr(10))
    RcvBuffer = Mid$(RcvBuffer, iPos + 1)
    
'    'SampleNo 얻기...
'    iPos = InStr(RcvBuffer, Chr(10))
'    RcvBuffer = Mid$(RcvBuffer, iPos + 1)

'    'Color : 절삭...
'    cnt = InStr(RcvBuffer, Chr$(10))
'    RcvBuffer = Mid$(RcvBuffer, cnt + 1)
'
'    'Clarity : 절삭...
'    cnt = InStr(RcvBuffer, Chr$(10))
'    RcvBuffer = Mid$(RcvBuffer, cnt + 1)
    
    pResultInfo.SEQNO = tmpSeqNo
    
    Do While True
        iPos = InStr(RcvBuffer, Chr(10))
        If iPos = 0 Then
            Exit Do
        End If
        
        sRstData = Left$(RcvBuffer, iPos - 2)
        RcvBuffer = Mid$(RcvBuffer, iPos + 1)

        '--- 결과값 편집
        For ii = 1 To UBound(sTestCd())     ' m_p_iOrdCnt
            sIFCd = Trim(sTestCd(ii - 1))
            If Trim(sIFCd) <> "" Then
                '설정된 장비코드에 해당하는 결과값 조회
                iPos = InStr(1, sRstData, Trim(sTestCd(ii - 1)))
                If iPos <> 0 Then
                    '결과값
                    sRst = Trim(Mid(sRstData, Len(sIFCd) + 1))
                    If Left(sRst, 1) = "*" Then
                        sRst = Trim(Mid(sRst, 2))
                    End If
                    iTmp = InStr(1, sRst, Space(1))
                    If iTmp > 0 Then
                        sRst = Trim(Left(sRst, iTmp - 1))
                    End If
                    
                    'Data 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & sIFCd & Chr(124)
                        .RST1 = .RST1 & sRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                    
                    Exit For
                End If
            End If
        Next ii
    Loop
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub



Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
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
        For ii = 1 To .ORDCNT
            .IFCD(ii) = tmpData(ii - 1)
        Next ii
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
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
    End With
    
End Sub
Private Sub PhaseCfg_Protocol_MiditronJr()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX
                RcvBuffer = ""
                
            Case 13     'CR
                '--- 결과편집/등록
                Call DataEdit_MiditronJr
                
                If bRstFlag = True Then
                    msComm.Output = Chr(2) & Chr(62) & Chr(3) & Chr(51) & Chr(63) & Chr(13)
                End If

            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_MiditronM()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX
                RcvBuffer = ""
                
            Case 13     'CR
                '--- 결과편집/등록
                Call DataEdit_MiditronM
                
                If bRstFlag = True Then
                    msComm.Output = Chr(2) & Chr(62) & Chr(3) & Chr(51) & Chr(63) & Chr(13)
                End If

            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,RTSEnable
Public Property Get RTSEnable() As Boolean
Attribute RTSEnable.VB_Description = "전송 요청 줄이 가능한지의 여부를 결정합니다."
    RTSEnable = msComm.RTSEnable
End Property

Public Property Let RTSEnable(ByVal New_RTSEnable As Boolean)
    msComm.RTSEnable() = New_RTSEnable
    PropertyChanged "RTSEnable"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,RThreshold
Public Property Get RThreshold() As Integer
Attribute RThreshold.VB_Description = "수신할 문자의 수를 반환하거나 설정합니다."
    RThreshold = msComm.RThreshold
End Property

Public Property Let RThreshold(ByVal New_RThreshold As Integer)
    msComm.RThreshold() = New_RThreshold
    PropertyChanged "RThreshold"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,Settings
Public Property Get Settings() As String
Attribute Settings.VB_Description = "전송 속도, 패리티, 데이터 비트, 중단 비트 매개 변수를 반환하거나 설정합니다."
    Settings = msComm.Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    msComm.Settings() = New_Settings
    PropertyChanged "Settings"
End Property

Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            wkBuf = msComm.Input
            
            If sTestMode = "77" Then
                RaiseEvent PrintRcvLog(wkBuf)
            End If
                                
            If iSpaceCnt = 30 Then
                iSpaceCnt = 0
            End If
            iSpaceCnt = iSpaceCnt + 2
            
            RaiseEvent DispMsg(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
            
            Call PhaseCfg_Protocol
            
        Case MSCOMM_EV_CTS      'j
        Case MSCOMM_EV_DSR      ' Change in the DSR line.
        Case MSCOMM_EV_CD       ' Change in the CD line.
        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
        ' Errors
        Case MSCOMM_ER_BREAK    ' A Break was received.
        ' Code to handle a BREAK goes here, and so on.
        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
        Case MSCOMM_ER_FRAME    ' Framing Error.
        Case MSCOMM_ER_OVERRUN  ' Data Lost.
        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
        Case MSCOMM_ER_RXPARITY ' Parity Error.
        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
    End Select
    
End Sub
'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    msComm.CommPort = PropBag.ReadProperty("CommPort", 1)
    msComm.RTSEnable = PropBag.ReadProperty("RTSEnable", False)
    msComm.RThreshold = PropBag.ReadProperty("RThreshold", 0)
    msComm.Settings = PropBag.ReadProperty("Settings", "9600,n,8,1")
    m_PortOpen = PropBag.ReadProperty("PortOpen", m_def_PortOpen)
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

End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CommPort", msComm.CommPort, 1)
    Call PropBag.WriteProperty("RTSEnable", msComm.RTSEnable, False)
    Call PropBag.WriteProperty("RThreshold", msComm.RThreshold, 0)
    Call PropBag.WriteProperty("Settings", msComm.Settings, "9600,n,8,1")
    Call PropBag.WriteProperty("PortOpen", m_PortOpen, m_def_PortOpen)
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

End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get PortOpen() As Boolean
    PortOpen = m_PortOpen
End Property

Public Property Let PortOpen(ByVal New_PortOpen As Boolean)
    m_PortOpen = New_PortOpen
    PropertyChanged "PortOpen"
    
    '--- PortOpen시 암호 확인
    If m_OpenPW <> pOpenPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Property
    End If
    '-----------------------
    
    On Error GoTo ErrPortOpen
    If m_PortOpen = True Then
        msComm.PortOpen = True
    End If
    On Error GoTo 0
    
    '변수 초기화
    bRstFlag = False
    
ErrPortOpen:
    If Err <> 0 Then
        MsgBox "PortOpen Error!!! " & Err.Description, vbCritical
        RaiseEvent DispMsg(Err.Description)
    End If
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
'MemberInfo=14
Public Function Send_Chr(iChr%) As Variant
    On Error GoTo ErrComm
    msComm.Output = Chr(iChr)
    On Error GoTo 0
ErrComm:
    RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
End Function

