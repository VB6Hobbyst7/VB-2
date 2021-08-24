VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl ACL 
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
Attribute VB_Name = "ACL"
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
Const m_def_IFMode = "0"
'속성 변수:
Private m_p_sPatInfo As Variant
Private m_EqName As String
Private m_bUseBarcode As Boolean
Private m_iPhase As Integer
Private m_iSendPhase As Integer
Private m_sTestMode As String
Private m_iFrameN As Integer
Private m_p_sID As String
Private m_p_sSeq As String
Private m_p_sRack As String
Private m_p_sPos As String
Private m_p_iOrdCnt As Integer
Private m_p_sTIFCd As String
Private m_PortOpen As Boolean
Private m_OpenPW As String
Private m_EditPW As String
Private m_IFMode As String

'이벤트 선언:
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sSeqNo$, sRack$, sPos$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)

'===== User Define
'인터페이스에서 사용
Private strFRcvBuffer   As String
Private strFWkBuf       As String
Private strFState       As String
Private blnFSend        As Boolean
Private blnFEndChk      As Boolean
Private blnFSTXChk      As Boolean
Private strFRstEnd      As String

Private strFRcvState    As String
Private strFSndState    As String
Dim msAllBarCd   As String
Dim maAllBarCd() As String
Dim TimerFlag    As Integer
Dim SavBuffer    As String
Dim ii_SendCnt   As Integer
Dim m_aTemp()    As String
Dim miSendCnt    As Integer
Dim msSendBuff   As String

'구조체 지정
Private f_typSampleInfo As SAMPLE_INFO
Private f_typResultInfo As RESULT_INFO

Private intFSpaceCnt    As Integer

Private strFExamInfo    As String
Private intFMulti       As Integer
Private intFCurIndex    As String

Dim msMsgID    As String
Dim msSender   As String
Dim msReceiver As String
Dim msVersion  As String

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

Private Sub DataEditResponse_ACL7000()

    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer

    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpType$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$

    ii = InStr(1, strFRcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(strFRcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
            Call subFInit_ResultInfo

        Case "M"
        Case "P"        'Patient Record
'            With f_typResultInfo
'                If .RSTCNT > 0 Then
'                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, "", "")
'                End If
'            End With

            Call subFInit_ResultInfo

        Case "Q"        'Order Request Record
            Erase tmpField()
            Erase tmpData()

            tmpField() = Split(strFRcvBuffer, "|")
            tmpData() = Split(tmpField(2), "^")

            If tmpData(1) = "" Then
                strFState = ""
                f_typSampleInfo.ID = ""
                Exit Sub
            Else
                strFState = "Q"
                f_typSampleInfo.ID = tmpData(1)
            End If

        Case "O"
            Call subFInit_ResultInfo

            tmpBarCd = ""
            tmpSeqNo = "": tmpRack = "": tmpPos = "": tmpType = ""

            tmpField() = Split(strFRcvBuffer, Chr(124))

            tmpBarCd = Trim(tmpField(2))

'            If Trim(tmpField(3)) = "" Then Exit Sub
'            ii = InStr(tmpField(3), "^")
'            If ii <> 0 Then
'                tmpData() = Split(Trim(tmpField(3)), "^")
'
'                tmpSeqNo = Trim(tmpData(0))
'                tmpRack = Trim(tmpData(1))
'                tmpPos = Trim(tmpData(2))
'                tmpType = Trim(tmpData(4))      'SAMPLE/CONTROL
'            End If

            With f_typResultInfo
                .ID = UCase(tmpBarCd)
                .SEQNO = ""
                .RACK = ""
                .POS = ""
                .Kind = ""
            End With

        Case "R"        'Result Record

            Erase tmpField()
            tmpField() = Split(strFRcvBuffer, "|")

            tmpData() = Split(tmpField(2), "^")
            tmpIFCd = Trim(tmpData(3))
            Select Case tmpField(4)
                Case "s", "%", "R"
                    tmpIFCd = tmpIFCd & tmpField(4)
            End Select

            tmpRst = Trim(tmpField(3))
            tmpUnit = ""
            tmpFlag = ""

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If

            If tmpRst <> "" Then
                '결과정보 구조체에 저장
                With f_typResultInfo
'                    .ID = pSampleInfo.ID
'                    .SEQNO = pSampleInfo.SEQNO
'                    .RACK = pSampleInfo.RACK
'                    .POS = pSampleInfo.POS
'                    .KIND = pSampleInfo.KIND

                    '결과값 누적
'                    .RSTCNT = .RSTCNT + 1
'                    .IFCD = .IFCD & tmpIFCd & Chr(124)
'                    .RST1 = .RST1 & tmpRst & Chr(124)
'                    .RST2 = .RST2 & Chr(124)
'                    .UNIT = .UNIT & tmpUnit & Chr(124)
'                    .FLAG = .FLAG & tmpFlag & Chr(124)
                    .RSTCNT = 1
                    .IFCD = tmpIFCd & Chr(124)
                    .RST1 = tmpRst & Chr(124)
                    .RST2 = Chr(124)
                    .UNIT = tmpUnit & Chr(124)
                    .FLAG = tmpFlag & Chr(124)
                End With
                
                '결과값 등록/화면 표시 처리...
                With f_typResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .Kind, "", "")
                    End If
                End With
            End If

        Case "C"        'Comment Record

        Case "L"
'            '결과값 등록/화면 표시 처리...
'            With f_typResultInfo
'                If .RSTCNT > 0 Then
'                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, "", "")
'                End If
'            End With

            Call subFInit_ResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If

End Sub


Private Sub DataEditResponse_ACL9000()

    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer

    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpType$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpRstDt$

    ii = InStr(1, strFRcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(strFRcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
            Call subFInit_ResultInfo

        Case "M"
        Case "P"        'Patient Record
'            With f_typResultInfo
'                If .RSTCNT > 0 Then
'                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, "", "")
'                End If
'            End With

            Call subFInit_ResultInfo

        Case "Q"        'Order Request Record
            Erase tmpField()
            Erase tmpData()

            tmpField() = Split(strFRcvBuffer, "|")
            tmpData() = Split(tmpField(2), "^")

            If tmpData(1) = "" Then
                strFState = ""
                f_typSampleInfo.ID = ""
                Exit Sub
            Else
                strFState = "Q"
                f_typSampleInfo.ID = tmpData(1)
            End If

        Case "O"
            Call subFInit_ResultInfo

            tmpBarCd = "": tmpSeqNo = "": tmpRack = "": tmpPos = "": tmpType = ""

            tmpField() = Split(strFRcvBuffer, Chr(124))

            tmpBarCd = Trim(tmpField(2))
            tmpType = Trim(tmpField(11))

'            If Trim(tmpField(3)) = "" Then Exit Sub
'            ii = InStr(tmpField(3), "^")
'            If ii <> 0 Then
'                tmpData() = Split(Trim(tmpField(3)), "^")
'
'                tmpSeqNo = Trim(tmpData(0))
'                tmpRack = Trim(tmpData(1))
'                tmpPos = Trim(tmpData(2))
'                tmpType = Trim(tmpData(4))      'SAMPLE/CONTROL
'            End If

            With f_typResultInfo
                .ID = UCase(tmpBarCd)
                .SEQNO = ""
                .RACK = ""
                .POS = ""
                .Kind = tmpType
            End With

        Case "R"        'Result Record
'            H|\^&||||||||ACL9000||P|1|19982110134700<CR>
'            P|1||PTNT1||BLU^^^^||19391127|M|||||||||||||||||DEP 1||||||||||<CR>
'            O|1|SMP01||^^^001|S||||||||||^|DR. VERDI||||||||||O||||||<CR>
'            R|1|^^^001|12.8|||||F||||19960119114215|<CR>
'            C|1|I|31^ Invalid for QC |I<CR>
'            P|2||PTNT1||Gialli^^^^||19391127|M|||||||||||||||||DEP 1||||||||||<CR>
'            O|1|SMP10||^^^001|S||||||||||^|DR. VERDI||||||||||O||||||<CR>
'            R|1|^^^001|14.5|s||||F||||19960119114215|<CR>
'            C|1|I|31^ Invalid for QC |I<CR>
'            L|1|N<CR>
        
            Erase tmpField()
            tmpField() = Split(strFRcvBuffer, "|")

            tmpData() = Split(tmpField(2), "^")
            tmpIFCd = Trim(tmpData(3))
            Select Case tmpField(4)
                Case "s", "%", "INR", "R"
                    tmpIFCd = tmpIFCd & tmpField(4)
            End Select

            tmpRst = Trim(tmpField(3))
            tmpUnit = ""
            tmpFlag = ""

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
            
            If UBound(tmpField) >= 12 Then
                tmpRstDt = Trim(tmpField(12))
            End If
            
            If tmpRst <> "" Then
                '결과정보 구조체에 저장
                With f_typResultInfo
                    .RSTDT = tmpRstDt
                    
                    .RSTCNT = 1
                    .IFCD = tmpIFCd & Chr(124)
                    .RST1 = tmpRst & Chr(124)
                    .RST2 = Chr(124)
                    .UNIT = tmpUnit & Chr(124)
                    .FLAG = tmpFlag & Chr(124)
                End With
                
                '결과값 등록/화면 표시 처리...
                With f_typResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .Kind, .RSTDT, "")
                    End If
                End With
            End If

        Case "C"        'Comment Record

        Case "L"
'            '결과값 등록/화면 표시 처리...
'            With f_typResultInfo
'                If .RSTCNT > 0 Then
'                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, "", "")
'                End If
'            End With

            Call subFInit_ResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If

End Sub

Private Sub DataEditResponse_ACLTOP()
    On Error GoTo ErrRtn
    
    Dim RecType As String   'Record Type
    Dim ii      As Integer

    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpType$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpRstGbn$
    
    Dim i           As Integer
    Dim sRxData     As String
    Dim sBuf        As String
    Dim sRackPos    As String
    Dim ChkS    As String
    
    '### Rack Or Tray 방식과 Conflict 방지
    Call ProtectConflict("Y")
    
    sRxData = ""
    sRxData = strFRcvBuffer

   'RecType 초기화
    RecType = "S"
        
    tmpData() = Split(sRxData, Chr(13))
                
    For i = 0 To UBound(tmpData)
       
        ii = InStr(1, tmpData(i), "|")
        If ii <> 0 Then
            RecType = Mid$(tmpData(i), ii - 1, 1)
        Else
            Exit For
        End If
        
        If RecType = "" Then
           Exit For
        End If
        
        If RecType = "H" Then
            Call subFInit_ResultInfo
            
            'H|@^\|<1366275944_5952><1366275944_5953>||acltop|||||GNAH_LIS||P|1394-97|20130418180544
            tmpField = Split(tmpData(i), "|")
            
            msMsgID = Trim(tmpField(2))
            msSender = Trim(tmpField(4))
            msReceiver = Trim(tmpField(9))
            msVersion = Trim(tmpField(12))
            
        ElseIf RecType = "Q" Then
            'Q|1|^1001@^1002@^1003@^1004@^1005@^1006@^1008||||||||||O@N
                        
            'Sample ID
            tmpField = Split(tmpData(i), "|")
            msAllBarCd = Replace(tmpField(2), "^", "")
                        
            strFRcvState = "Q"
            
            If msAllBarCd <> "" Then
                m_iFrameN = 1
                miSendCnt = 0
    
                maAllBarCd = Split(msAllBarCd, "@")
                For ii = 0 To UBound(maAllBarCd)    ' - 1
                    'Order내역을 가져옴
                    If maAllBarCd(ii) <> "ALL" Then
                        RaiseEvent RequestCurOrder(maAllBarCd(ii), "", "", "")
                    Else
                        m_p_iOrdCnt = 0
                    End If
                    
                    Call Get_OrderString
            
                    'Order Packet 만들기
                    Call SendOrder_ACLTOP(ii, UBound(maAllBarCd))   ' - 1)
                Next
            End If

            Exit Sub
            
        ElseIf RecType = "P" Then
            'P|1||||^|||U|||||^
            tmpField = Split(tmpData(i), "|")
            sBuf = Trim(tmpField(1))
                     
        ElseIf RecType = "O" Then
            'O|1|10|<1176405880_874>|^^^1551|R|20070412152555|||||||||P||||||||||O@F
            tmpField = Split(tmpData(i), "|")
            f_typResultInfo.ID = Trim(tmpField(2))
            f_typResultInfo.Kind = Trim(tmpField(11))

        ElseIf RecType = "R" Then
            'R|2|^^^1551|1.09|Ratio||N||F@V||ACLTOP^ACLTOP||20070412153249|ACLTOP^12^10
            strFRcvState = "R"
            tmpField = Split(tmpData(i), "|")
            sRackPos = Mid(tmpField(13), InStr(tmpField(13), "^") + 1)
            
            tmpRack = Trim(Split(sRackPos, "^")(0))
            tmpPos = Trim(Split(sRackPos, "^")(1))
            tmpIFCd = Mid(Trim(tmpField(2)), 4)
            tmpRst = Trim(tmpField(3))
            tmpRstGbn = Trim(tmpField(4))
                        
            tmpIFCd = tmpIFCd & tmpRstGbn
            
            If tmpRst <> "" Then
                '결과정보 구조체에 저장
                With f_typResultInfo
                    .RACK = tmpRack
                    .POS = tmpPos
                    
                    '결과값 누적
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            End If
                                       
        ElseIf RecType = "C" Then
            tmpField = Split(tmpData(i), "|")
            tmpFlag = Trim(tmpField(3))
            
            f_typResultInfo.FLAG = f_typResultInfo.FLAG & tmpFlag & "@"
            
        ElseIf RecType = "L" Then
            f_typResultInfo.FLAG = f_typResultInfo.FLAG & Chr(124)
            
            If f_typResultInfo.ID <> "" Then
                With f_typResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .Kind, "", "")
                    End If
                End With
            End If
        Else
        End If
    Next
    
    If strFRcvState = "R" Then
        If (strFSndState = "E") Or (strFSndState = "H") Or (strFSndState = "P") _
                Or (strFSndState = "O") Or (strFSndState = "L") Then
            'ENQ 전송
            msComm.Output = Chr(5)
            
            If m_sTestMode = 77 Then
                RaiseEvent PrintSendLog(Chr(5))
            End If
            
            m_iPhase = 3
        Else
            m_iPhase = 1
            Call ProtectConflict("N")
        End If
    End If
    
    strFRcvState = ""
    
    Exit Sub
    
ErrRtn:
    If Err <> 0 Then
        strFRcvState = ""
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub ProtectConflict(ByVal sFlag$)
    '0=단방향
    '1=양방향(Rack Or Tray 방식 지원안함, But Rack/Pos 표시)
    '2=양방향(Rack Or Tray 방식 지원안함, But Tray/Pos 표시)
    '3=양방향(Rack Or Tray 방식 지원안함, But Tray/Cup 표시)
    '4=양방향(Rack/Pos 방식 지원),
    '5=양방향(Tray/Pos 방식 지원),
    '6=양방향(Tray/Cup 방식 지원)
    
    If UCase(sFlag) = "Y" Then
        Select Case IFMode 'gsIFMode
            Case "0", "1", "2", "3"
                TimerFlag = 0
            Case "4", "5", "6"
                TimerFlag = 1
        End Select
    ElseIf UCase(sFlag) = "N" Then
        TimerFlag = 0
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
        Case "ACL1000"
            Call PhaseCfg_Protocol_ACL1000
            
        Case "ACL7000"
            Call PhaseCfg_Protocol_ACL7000
            
        Case "ACL9000"
            Call PhaseCfg_Protocol_ACL9000
        
        Case "ACLTOP"
            Call PhaseCfg_Protocol_ACLTOP
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_ACL1000()
            
    Dim strWkDat    As String
    Dim intIdx      As Long
    
    For intIdx = 1 To Len(strFWkBuf)
        strWkDat = Mid$(strFWkBuf, intIdx, 1)
                 
        Select Case Asc(strWkDat)
            Case 10
                Call DataEditResponse_ACL1000
                strFRcvBuffer = ""
            
            Case 13
            
            Case Else      ' Data
                strFRcvBuffer = strFRcvBuffer & strWkDat
         End Select
    
    Next intIdx

End Sub


Private Sub PhaseCfg_Protocol_ACL9000()
    
    Dim wkdat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(strFWkBuf)
        wkdat = Mid$(strFWkBuf, ix1, 1)
        
        Select Case m_iPhase
            Case 1            'ENQ 대기
                Select Case Asc(wkdat)
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2      '<LF> 대기
                Select Case Asc(wkdat)
                    Case 2      'STX
                        strFRcvBuffer = ""
                        
                    Case 10     '<LF>
                        Call DataEditResponse_ACL9000   'Data 편집
                        
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                                                
                    Case 4      'EOT
                        If strFState = "Q" Then
                            msComm.Output = Chr(5)
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3
                        
                    Case 5      'ENQ
                        msComm.Output = Chr(6)   'Send ACK
                        
                    Case 21     'NAK
                        msComm.Output = Chr(5)   'Send ENQ
                        m_iPhase = 1
                        
                    Case Else
                        strFRcvBuffer = strFRcvBuffer & wkdat
                        m_iPhase = 2
                End Select
            
            Case 3      'ACK 대기
                Select Case Asc(wkdat)
                    Case 6      'ACK
                        If strFState = "Q" Then
                            Call SendOrder_ACL9000
                        End If
                    
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    
                    Case 21     'NAK
                        msComm.Output = Chr(5)
                        m_iPhase = 3
                        
                    Case 4      'EOT
                        m_iPhase = 1
                End Select
                
        End Select
    Next ix1

End Sub
Private Sub PhaseCfg_Protocol_ACL7000()
    
    Dim wkdat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(strFWkBuf)
        wkdat = Mid$(strFWkBuf, ix1, 1)
        
        Select Case m_iPhase
            Case 1            'ENQ 대기
                Select Case Asc(wkdat)
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2      '<LF> 대기
                Select Case Asc(wkdat)
                    Case 2      'STX
                        strFRcvBuffer = ""
                        
                    Case 10     '<LF>
                        Call DataEditResponse_ACL7000   'Data 편집
                        
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                                                
                    Case 4      'EOT
                        If strFState = "Q" Then
                            msComm.Output = Chr(5)
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3
                        
                    Case 5      'ENQ
                        msComm.Output = Chr(6)   'Send ACK
                        
                    Case 21     'NAK
                        msComm.Output = Chr(5)   'Send ENQ
                        m_iPhase = 1
                        
                    Case Else
                        strFRcvBuffer = strFRcvBuffer & wkdat
                        m_iPhase = 2
                End Select
            
            Case 3      'ACK 대기
                Select Case Asc(wkdat)
                    Case 6      'ACK
                        If strFState = "Q" Then
                            Call SendOrder_ACL7000
                        End If
                    
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    
                    Case 21     'NAK
                        msComm.Output = Chr(5)
                        m_iPhase = 3
                        
                    Case 4      'EOT
                        m_iPhase = 1
                End Select
                
        End Select
    Next ix1

End Sub

Private Sub PhaseCfg_Protocol_ACLTOP()
    Dim wkdat As String
    Dim ix1 As Integer
    
    For ix1 = 1 To Len(strFWkBuf)
        wkdat = Mid$(strFWkBuf, ix1, 1)
             
        Select Case m_iPhase
            'ENQ 대기 상태
            Case 1
                Select Case Asc(wkdat)
                    'ENQ
                    Case 5
                        strFRcvState = ""
                        strFSndState = ""
                        strFRcvBuffer = ""
                        
                        'ACK 전송
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = 77 Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                        m_iPhase = 2
                    Case Else
                        strFRcvState = ""
                        strFSndState = ""
                        m_iPhase = 1
                End Select
            
            'Packet 모음, Packet 분석(Edit_Data)
            Case 2
                Select Case Asc(wkdat)
                    'STX
                    Case 2
                                            
                    'EOT
                    Case 4
                        Call DataEditResponse_ACLTOP
                        
                    'ENQ
                    Case 5
                        strFRcvState = ""
                        strFSndState = ""
                        strFRcvBuffer = ""
                        
                        'ACK 전송
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = 77 Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                    'LF
                    Case 10
                        strFRcvBuffer = strFRcvBuffer & Mid(SavBuffer, 2, Len(SavBuffer) - 5)
                        SavBuffer = ""
                        'ACK 전송
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = 77 Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                    'NAK
                    Case 21
                        'ENQ 전송
                        msComm.Output = Chr(5)
                        
                        If m_sTestMode = 77 Then
                            RaiseEvent PrintSendLog(Chr(5))
                        End If
                        
                    ''Case Is < 0
                    
                    Case Else
                        SavBuffer = SavBuffer & wkdat
                End Select
                
            'SendOrder위해 ENQ후의 ACK 대기상태
            Case 3
                Select Case Asc(wkdat)
                    'EOT
                    Case 4
                        m_iPhase = 1

                    'ACK
                    Case 6
                        If strFSndState = "E" Then
                            ii_SendCnt = 0
                            
                            msComm.Output = m_aTemp(ii_SendCnt)
                                                        
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(m_aTemp(ii_SendCnt))
                            End If
                            
                            If ii_SendCnt + 1 = miSendCnt Then
                                strFSndState = "L"
                            Else
                                strFSndState = "P"
                            End If
                            
                            m_iPhase = 3
                            Exit Sub
                            
                        ElseIf strFSndState = "P" Then
                            ii_SendCnt = ii_SendCnt + 1
                            msComm.Output = m_aTemp(ii_SendCnt)
                            
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(m_aTemp(ii_SendCnt))
                            End If
                            
                            If ii_SendCnt + 1 = miSendCnt Then
                                strFSndState = "L"
                            Else
                                strFSndState = "P"
                            End If
                            
                            m_iPhase = 3
                            Exit Sub
                            
                        ElseIf strFSndState = "L" Then
                            'EOT 전송
                            msComm.Output = Chr(4)
                            
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(Chr(4))
                            End If
                            
                            m_iFrameN = 0
                            strFSndState = ""
                            ii_SendCnt = 0
                            miSendCnt = 0
                            m_iPhase = 1
                            msAllBarCd = ""
                            Erase maAllBarCd
                            Erase m_aTemp

                        End If
                    'NAK
                    Case 21
                        If strFSndState = "E" Then
                            msComm.Output = Chr(5)
                            
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(Chr(5))
                            End If
                            
                            strFSndState = "E"
                            m_iPhase = 3
                            Exit Sub
                        ElseIf strFSndState = "P" Or strFSndState = "L" Then
                            msComm.Output = m_aTemp(ii_SendCnt)
                            
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(m_aTemp(ii_SendCnt))
                            End If
                            
                            If ii_SendCnt + 1 = miSendCnt Then
                                strFSndState = "L"
                            Else
                                strFSndState = "P"
                            End If
                            
                            m_iPhase = 3
                            Exit Sub
                        End If
                    'ENQ
                    Case 5
                        strFRcvState = ""
                        strFSndState = ""
                        strFRcvBuffer = ""
                    
                        'ACK 전송
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = 77 Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                        strFRcvBuffer = ""
                        m_iPhase = 2
                End Select
        End Select
    Next
    
End Sub

Private Sub SendOrder_ACL9000()

    On Error GoTo ErrRtn

    Dim sTmp    As String
    Dim ChkS    As String
    Dim TestDat As String
    Dim i       As Integer

    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If

    Select Case m_iSendPhase
        Case 1      'Header Record
            sTmp = m_iFrameN & "H|\^&||||||||ACL9000||P|1|" & Format(Now, "YYYYMMDDHHMMSS") & Chr(13) & Chr(3)

            '----- 검사항목 조회/편집
            RaiseEvent RequestCurOrder(f_typSampleInfo.ID, "", "", "")

            Call Get_OrderString

            intFCurIndex = 1

'            'AXSYM의 경우 오더가 없는 경우는 Negative Query를 전송한다.
'            If f_typSampleInfo.ORDCNT > 0 Then
                m_iSendPhase = 2
'            Else
'                m_iSendPhase = 3
'            End If

        Case 2      'Patient Record
            sTmp = m_iFrameN & "P|1||||^^^^|||U||||||||||||||||||||||||||" & Chr(13) & Chr(3)
            m_iSendPhase = 3
  
        Case 3      'Order Record
            If f_typSampleInfo.ORDCNT > 0 Then
                sTmp = m_iFrameN & "O|" & intFCurIndex & "|" & f_typSampleInfo.ID & "|" & "|^^^" _
                                & f_typSampleInfo.IFCD(intFCurIndex) & "||||||||||||||||||||||O||||||" & Chr(13) & Chr(3)
                
                If intFCurIndex >= f_typSampleInfo.ORDCNT Then
                    m_iSendPhase = 4
                Else
                    m_iSendPhase = 3
                    
                    intFCurIndex = intFCurIndex + 1
                End If
            Else
                sTmp = m_iFrameN & "O|1|" & f_typSampleInfo.ID & "|" & "|^^^000||||||||||||||||||||||O||||||" & Chr(13) & Chr(3)
                m_iSendPhase = 4
            End If

        Case 4      'Terminator Record
            sTmp = m_iFrameN & "L|1|N" & Chr(13) & Chr(3)
            m_iSendPhase = 5

        Case 5      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1: m_iPhase = 1: m_iSendPhase = 1
            strFState = ""

            If f_typSampleInfo.ORDCNT > 0 Then
                'Barcode Mode인 경우 전송완료 이벤트 발생
                RaiseEvent SendOrderOK(f_typSampleInfo.ID, "", "", "")
            Else
                RaiseEvent SendOrderOK("", "", "", "")
            End If
            
            Exit Sub

    End Select

    'CheckSum 계산
    ChkS = ChkSum_ASTM(sTmp)

    msComm.Output = Chr(2) & sTmp & ChkS & Chr(13) & Chr(10)

    m_iFrameN = m_iFrameN + 1

    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sTmp & ChkS & Chr(13) & Chr(10))
    End If
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_ACLTOP(ByVal iCurOrd As Integer, ByVal iMaxOrd)
    On Error GoTo ErrRtn
    '환자의 Order 전송
    Dim i       As Integer
    Dim iDiv    As Integer
    Dim sBuf    As String
    Dim ChkS    As String
    Dim sPriority As String
       
    sBuf = ""
    
    If iCurOrd = 0 Then
        msSendBuff = msSendBuff & "H|@^\|" & msMsgID & "||" & msReceiver & "|||||" & msSender & "||P|" & msVersion & "|" & Format(Now, "yyyyMMddHHmmss") & Chr(13)
    End If

    If f_typSampleInfo.ORDCNT > 0 Then
        For i = 1 To f_typSampleInfo.ORDCNT
            If InStr(sBuf, "^^^" & f_typSampleInfo.IFCD(i) & "@") = 0 Then
                sBuf = sBuf & "^^^" & f_typSampleInfo.IFCD(i) & "@"
            End If
        Next i
    
        If Len(sBuf) > 0 Then
            sBuf = Mid(sBuf, 1, Len(sBuf) - 1)
        End If
    
        msSendBuff = msSendBuff & "P|" & Trim(CStr(iCurOrd + 1)) & "||||^||||||||" & Chr(13)
    
        'S (Stat)
        'R (normal)
        sPriority = f_typSampleInfo.Kind
        If sPriority = "" Then
            sPriority = "R"
        End If
        
        msSendBuff = msSendBuff & "O|" & "1" & "|" & f_typSampleInfo.ID & "|" & "|" & sBuf & "|" & sPriority & "||||||A||||P||||||||||Q" & Chr(13)
    End If
    
    If iCurOrd = iMaxOrd Then
        msSendBuff = msSendBuff & "L|1|N"
    
        iDiv = Len(msSendBuff) \ 240
        ReDim m_aTemp(iDiv)
       
        For i = 0 To iDiv
            If Len(msSendBuff) > 240 Then
                miSendCnt = miSendCnt + 1
                ChkS = ChkSum_ASTM(m_iFrameN & Mid(msSendBuff, 1, 240) & Chr(23))
                m_aTemp(i) = Chr(2) & m_iFrameN & Mid(msSendBuff, 1, 240) & Chr(23) & ChkS & Chr(13) & Chr(10)
                msSendBuff = Replace(msSendBuff, Mid(msSendBuff, 1, 240), "")
            Else
                miSendCnt = miSendCnt + 1
                ChkS = ChkSum_ASTM(m_iFrameN & msSendBuff & Chr(3))
                m_aTemp(i) = Chr(2) & m_iFrameN & msSendBuff & Chr(3) & ChkS & Chr(13) & Chr(10)
            End If
            
            m_iFrameN = m_iFrameN + 1
            
            If m_iFrameN > 7 Then      'Frame Number가 8이상이면 0으로 바꿔줌
                m_iFrameN = 0
            End If
        Next
                   
        '<ENQ> 전송
        msComm.Output = Chr(5)
                        
        If m_sTestMode = 77 Then
            RaiseEvent PrintSendLog(Chr(5))
        End If
                        
        '<ENQ>를 보낸 상태
        strFSndState = "E"
        iPhase = 3
        msSendBuff = ""
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order_Input 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_ACL7000()

    On Error GoTo ErrRtn

    Dim sTmp    As String
    Dim ChkS    As String
    Dim TestDat As String
    Dim i       As Integer

    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If

    Select Case m_iSendPhase
        Case 1      'Header Record
            sTmp = m_iFrameN & "H|\^&||||||||ACL7000||P|1|" & Format(Now, "YYYYMMDDHHMMSS") & Chr(13) & Chr(3)

            '----- 검사항목 조회/편집
            RaiseEvent RequestCurOrder(f_typSampleInfo.ID, "", "", "")

            Call Get_OrderString

            intFCurIndex = 1

            'AXSYM의 경우 오더가 없는 경우는 Negative Query를 전송한다.
            If f_typSampleInfo.ORDCNT > 0 Then
                m_iSendPhase = 2
            Else
                m_iSendPhase = 3
            End If

        Case 2      'Patient Record
            sTmp = m_iFrameN & "P|1||||^^^^|||U||||||||||||||||||||||||||" & Chr(13) & Chr(3)
            m_iSendPhase = 3

        Case 3      'Order Record
''            If intFCurIndex > f_typSampleInfo.ORDCNT Then
''                sTmp = m_iFrameN & "O|1|" & f_typSampleInfo.ID & "|" & "|^^^000||||||||||||||||||||||O||||||" & Chr(13) & Chr(3)
''                m_iSendPhase = 4
''            Else

            If f_typSampleInfo.ORDCNT > 0 Then
                sTmp = m_iFrameN & "O|" & intFCurIndex & "|" & f_typSampleInfo.ID & "|" & "|^^^" & f_typSampleInfo.IFCD(intFCurIndex) & "||||||||||||||||||||||O||||||" & Chr(13) & Chr(3)
                
                If intFCurIndex >= f_typSampleInfo.ORDCNT Then
                    m_iSendPhase = 4
                Else
                    m_iSendPhase = 3
                    
                    intFCurIndex = intFCurIndex + 1
                End If
            Else
                sTmp = m_iFrameN & "O|1|" & f_typSampleInfo.ID & "|" & "|^^^000||||||||||||||||||||||O||||||" & Chr(13) & Chr(3)
                m_iSendPhase = 4
            End If

        Case 4      'Terminator Record
            sTmp = m_iFrameN & "L|1|N" & Chr(13) & Chr(3)
            m_iSendPhase = 5

        Case 5      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1: m_iPhase = 1: m_iSendPhase = 1
            strFState = ""

            If f_typSampleInfo.ORDCNT > 0 Then
                'Barcode Mode인 경우 전송완료 이벤트 발생
                RaiseEvent SendOrderOK(f_typSampleInfo.ID, "", "", "")
            Else
                RaiseEvent SendOrderOK("", "", "", "")
            End If
            
            Exit Sub

    End Select

    'CheckSum 계산
    ChkS = ChkSum_ASTM(sTmp)

    msComm.Output = Chr(2) & sTmp & ChkS & Chr(13) & Chr(10)

    m_iFrameN = m_iFrameN + 1

    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sTmp & ChkS & Chr(13) & Chr(10))
    End If
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub


' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_ACL1000()

    On Error GoTo ErrHandler
    
    Dim strBC       As String
    Dim strLC       As String

    Dim strIFRstCd  As String   '인터페이스시 검사항목코드
    Dim strRst      As String
    Dim strRst2     As String

    Dim intPos      As Integer
    Dim intNotCoag  As Integer
    
    If strFRcvBuffer = "" Then Exit Sub
    
    strBC = Mid$(strFRcvBuffer, 23, 1)
    intNotCoag = 0
    
    strIFRstCd = "": strRst = ""
    
    '===== HEADER를 인식해서 검사항목을 찾는다
    If Trim(strBC) = ":" Then
        Call subFInit_ResultInfo
        
        intFMulti = 1
        strFExamInfo = ""
        strLC = Mid$(strFRcvBuffer, 1, 2)
        
        Select Case strLC
            Case "03"
                strFExamInfo = "PT-FIB"
            Case "06"
                strFExamInfo = "APTT"
            Case "09"
                strFExamInfo = "TT"
            Case "1B"
                strFExamInfo = "PT-FIB/APTT"
            Case "1E"
                strFExamInfo = "TT/APTT"
            Case Else
'                MsgBox strLC & " 란 숫자와 검사항목을 담당자님께 알려주세요!", vbOKOnly + vbCritical, "ACL - 검사결과받기 [오류]"
                RaiseEvent DispMsg(strLC & " 란 숫자와 검사항목을 담당자님께 알려주세요!")
        End Select
        Exit Sub
    End If
    
    If intFMulti = 1 Then Call subFInit_ResultInfo

    '===== 실제결과이면 결과를 등록한다
    intPos = Format(Mid(strFRcvBuffer, 25, 2), "#0")
    
    If intPos > 0 And intPos < 19 Then
        
        Select Case strFExamInfo
    
            Case "PT-FIB"
                    
                '====== PT(s) ======
                strIFRstCd = "PT(s)"
                strRst = Trim(Mid(strFRcvBuffer, 1, 4))
                
                If strRst <> "" Then
                    With f_typResultInfo
                        .POS = Format$(intPos, "00")
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & strIFRstCd & "|"
                        .RST1 = .RST1 & strRst & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                    End With
                End If
                
                '====== PT(%) ======
                strIFRstCd = "PT(%)"
                strRst = Trim(Mid(strFRcvBuffer, 6, 3))
                
                If strRst <> "" Then
                    With f_typResultInfo
                        .POS = Format$(intPos, "00")
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & strIFRstCd & "|"
                        .RST1 = .RST1 & strRst & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                    End With
                End If
                
                '====== PT(R) ======
                strIFRstCd = "PT(R)"
                strRst = Trim(Mid(strFRcvBuffer, 10, 4))
                
                If strRst <> "" Then
                    With f_typResultInfo
                        .POS = Format$(intPos, "00")
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & strIFRstCd & "|"
                        .RST1 = .RST1 & strRst & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                    End With
                End If
                
                '====== FIB ======
                strIFRstCd = "FIB"
                strRst = Trim(Mid(strFRcvBuffer, 15, 3))
                
                If strRst <> "" Then
                    With f_typResultInfo
                        .POS = Format$(intPos, "00")
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & strIFRstCd & "|"
                        .RST1 = .RST1 & strRst & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                    End With
                End If
                
                With f_typResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .QCGBN, "", "")
                    End If
                End With
    
                intFMulti = 1
            
            Case "APTT"
                    
                '====== APTT ======
                strIFRstCd = "APTT"
                strRst = Trim(Mid(strFRcvBuffer, 1, 4))
                
                If strRst <> "" Then
                    With f_typResultInfo
                        .POS = Format$(intPos, "00")
                        
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & strIFRstCd & "|"
                        .RST1 = .RST1 & strRst & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                    End With
                End If
                
                With f_typResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .QCGBN, "", "")
                    End If
                End With
                
                intFMulti = 1

            Case "TT"
                
                '====== TT ======
                strIFRstCd = "TT"
                strRst = Trim(Mid(strFRcvBuffer, 1, 4))
                
                If strRst <> "" Then
                    With f_typResultInfo
                        .POS = Format$(intPos, "00")
                        
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & strIFRstCd & "|"
                        .RST1 = .RST1 & strRst & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                    End With
                End If
                
                With f_typResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .QCGBN, "", "")
                    End If
                End With
            
                intFMulti = 1
            
            Case "PT-FIB/APTT"
            
                If intFMulti = 1 Then
                    
                    '====== PT(s) ======
                    strIFRstCd = "PT(s)"
                    strRst = Trim(Mid(strFRcvBuffer, 1, 4))
                    
                    If strRst <> "" Then
                        With f_typResultInfo
                            .POS = Format$(intPos, "00")
                            
                            '결과값 누적
                            .RSTCNT = .RSTCNT + 1
                            .IFCD = .IFCD & strIFRstCd & "|"
                            .RST1 = .RST1 & strRst & Chr(124)
                            .RST2 = .RST2 & "" & Chr(124)
                            .UNIT = .UNIT & "" & Chr(124)
                            .FLAG = .FLAG & "" & Chr(124)
                        End With
                    End If
                    
                    '====== PT(%) ======
                    strIFRstCd = "PT(%)"
                    strRst = Trim(Mid(strFRcvBuffer, 6, 3))
                    
                    If strRst <> "" Then
                        With f_typResultInfo
                            .POS = Format$(intPos, "00")
                            '결과값 누적
                            .RSTCNT = .RSTCNT + 1
                            .IFCD = .IFCD & strIFRstCd & "|"
                            .RST1 = .RST1 & strRst & Chr(124)
                            .RST2 = .RST2 & "" & Chr(124)
                            .UNIT = .UNIT & "" & Chr(124)
                            .FLAG = .FLAG & "" & Chr(124)
                        End With
                    End If
                    
                    '====== PT(R) ======
                    strIFRstCd = "PT(R)"
                    strRst = Trim(Mid(strFRcvBuffer, 10, 4))
                    
                    If strRst <> "" Then
                        With f_typResultInfo
                            .POS = Format$(intPos, "00")
                            '결과값 누적
                            .RSTCNT = .RSTCNT + 1
                            .IFCD = .IFCD & strIFRstCd & "|"
                            .RST1 = .RST1 & strRst & Chr(124)
                            .RST2 = .RST2 & "" & Chr(124)
                            .UNIT = .UNIT & "" & Chr(124)
                            .FLAG = .FLAG & "" & Chr(124)
                        End With
                    End If
                    
                    '====== FIB ======
                    strIFRstCd = "FIB"
                    strRst = Trim(Mid(strFRcvBuffer, 15, 3))
                    
                    If intNotCoag = 1 Then
                        strRst = "60"
                        intNotCoag = 0
                    End If
                    
                    If strRst <> "" Then
                        With f_typResultInfo
                            .POS = Format$(intPos, "00")
                            '결과값 누적
                            .RSTCNT = .RSTCNT + 1
                            .IFCD = .IFCD & strIFRstCd & "|"
                            .RST1 = .RST1 & strRst & Chr(124)
                            .RST2 = .RST2 & "" & Chr(124)
                            .UNIT = .UNIT & "" & Chr(124)
                            .FLAG = .FLAG & "" & Chr(124)
                        End With
                    End If
                    
                    intFMulti = 2
                Else
                    
                    '====== APTT ======
                    strIFRstCd = "APTT"
                    strRst = Trim(Mid(strFRcvBuffer, 1, 4))
                                        
                    If strRst <> "" Then
                        With f_typResultInfo
                            .POS = Format$(intPos, "00")
                            '결과값 누적
                            .RSTCNT = .RSTCNT + 1
                            .IFCD = .IFCD & strIFRstCd & "|"
                            .RST1 = .RST1 & strRst & Chr(124)
                            .RST2 = .RST2 & "" & Chr(124)
                            .UNIT = .UNIT & "" & Chr(124)
                            .FLAG = .FLAG & "" & Chr(124)
                        End With
                    End If
                    
                    With f_typResultInfo
                        If .RSTCNT > 0 Then
                            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .QCGBN, "", "")
                        End If
                    End With
                    
                    intFMulti = 1
                    
                End If
            
            Case "TT/APTT"
                
                '====== TT ======
                strIFRstCd = "TT"
                strRst = Trim(Mid(strFRcvBuffer, 1, 4))
                
                If strRst <> "" Then
                    With f_typResultInfo
                        .POS = Format$(intPos, "00")
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & strIFRstCd & "|"
                        .RST1 = .RST1 & strRst & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                    End With
                End If
                
                '====== APTT ======
                strIFRstCd = "APTT"
                strRst = Trim(Mid(strFRcvBuffer, 6, 4))
                
                If strRst <> "" Then
                    With f_typResultInfo
                        .POS = Format$(intPos, "00")
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & strIFRstCd & "|"
                        .RST1 = .RST1 & strRst & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                    End With
                End If
                
                With f_typResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .QCGBN, "", "")
                    End If
                End With
                
                intFMulti = 1
        
        End Select
        
    End If
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub


Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With f_typSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
            Erase .IFCD
        End With
        
        Exit Sub
    End If
    
    ReDim tmpData(m_p_iOrdCnt) As String
    tmpData() = Split(m_p_sTIFCd, Chr(124))
    
    With f_typSampleInfo
        .ID = m_p_sID
        .SEQNO = m_p_sSeq
        .RACK = m_p_sRack
        .POS = m_p_sPos
        .ORDCNT = m_p_iOrdCnt
        .Kind = m_p_sPatInfo
        
        ReDim .IFCD(.ORDCNT)
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
        
'        .ID = m_p_sID
'        .SEQNO = m_p_sSeq
'        .RACK = m_p_sRack
'        .POS = m_p_sPos
'        .ORDCNT = 1      '실제 검사 가능한 항목 갯수
'        ReDim Preserve .IFCD(1 To 1) As String
'        .IFCD(1) = ""
'        .OTHER = m_p_sPatInfo
    End With
        
End Sub

'
'   결과정보 구조체 초기화
'
Private Sub subFInit_ResultInfo()
    
    With f_typResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .QCGBN = ""
        .Kind = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .INSTID = ""
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
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

    strFWkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            strFWkBuf = msComm.Input
            
            If sTestMode = "77" Then
                RaiseEvent PrintRcvLog(strFWkBuf)
            End If
                                
            If intFSpaceCnt = 30 Then
                intFSpaceCnt = 0
            End If
            intFSpaceCnt = intFSpaceCnt + 2
            
            RaiseEvent DispMsg(Space(intFSpaceCnt) & "장비와 Interface 작업 중...")
            
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
    m_p_sPatInfo = PropBag.ReadProperty("p_sPatInfo", m_def_p_sPatInfo)
    
    m_IFMode = PropBag.ReadProperty("IFMode", m_def_IFMode)
    
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
    Call PropBag.WriteProperty("p_sPatInfo", m_p_sPatInfo, m_def_p_sPatInfo)
    
    Call PropBag.WriteProperty("IFMode", m_IFMode, m_def_IFMode)
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
    m_p_sPatInfo = m_def_p_sPatInfo
    
    m_IFMode = m_def_IFMode

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
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
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


Public Property Get IFMode() As Integer
    IFMode = m_IFMode
End Property

Public Property Let IFMode(ByVal New_IFMode As Integer)
    m_IFMode = New_IFMode
    PropertyChanged "IFMode"
End Property
