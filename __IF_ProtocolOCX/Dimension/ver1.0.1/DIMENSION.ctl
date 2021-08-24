VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl DIMENSION 
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
Attribute VB_Name = "DIMENSION"
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
Event SendOrderOK(sID$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sRack$, sPos$)
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

'For Dimension
Dim sSend_Buf   As String

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
        Case "DIMENSION"        '바코드 사용하는 양방향
            Call PhaseCfg_Protocol_Dimension
            
        Case "DIMENSION_UNI"    'DIMENSION RxL 단방향
            Call PhaseCfg_Protocol_DimensionUni
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_DimensionUni()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 1      '----- ENQ 수신
                msComm.Output = Chr$(6)
                
            Case 2      '----- STX 수신
                RcvBuffer = ""
                
            Case 3      '----- ETX 수신
                Call DataEditResponse_DimensionUni
                
            Case 6      '----- ACK 수신

            Case 21     '----- NAK 수신
'                msComm.Output = sSend_Buf
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1

End Sub


Private Sub PhaseCfg_Protocol_Dimension()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 1      '----- ENQ 수신
                msComm.Output = Chr$(6)
                
            Case 2      '----- STX 수신
                RcvBuffer = ""
                
            Case 3      '----- ETX 수신
                Call DataEditResponse_Dimension
                
            Case 6      '----- ACK 수신

            Case 21     '----- NAK 수신
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
End Sub

Private Sub DataEditResponse_DimensionUni()
    On Error GoTo ErrRtn
    
    Dim sType   As String
    Dim sPollF  As String
    Dim sPollR  As String
    Dim tmpData()   As String
    Dim ii      As Integer
    
    Dim tmpPID  As String
    Dim tmpSeq  As String
    Dim iTestCnt    As Integer
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$
    Dim sStatus As String

    
    sType = Left$(RcvBuffer, 1)       ' get Frame of RcvBuffer.
   
    Select Case sType
        Case "P"        'POLL Record
            tmpData() = Split(RcvBuffer, Chr(28))
            
            sPollF = tmpData(2)
            sPollR = tmpData(3)
                    
'            If sPollR = "1" Then
'                If m_iSendPhase = 1 Then    'Order를 눌렀을때
'                    msComm.Output = Chr$(6)
'                    'ORDER 전송
'                    Call SendOrder_Dimension
'                Else                        'New Order가 모두 전송
                    msComm.Output = Chr$(6)
''                    msComm.Output = Chr$(2) & "W" & Chr$(28) & "73" & Chr$(3)
'                    msComm.Output = Chr(2) & "N" & Chr(28) & "6A" & Chr(3)
'                End If
'            ElseIf sPollF = "1" And sPollR = "0" Then
'                msComm.Output = Chr$(6)
'                msComm.Output = Chr$(2) & "W" & Chr$(28) & "73" & Chr$(3)
''                msComm.Output = Chr(2) & "N" & Chr(28) & "6A" & Chr(3)
'            End If
    
        Case "R"        'RESULT Record
            '결과정보 구조체 초기화
            Call Init_pResultInfo
            
            tmpData() = Split(RcvBuffer, Chr(28))
                        
            tmpPID = tmpData(2)         'Patient ID
            tmpSeq = tmpData(3)         'SampleNo
            iTestCnt = Val(tmpData(10)) 'Number of Tests
            
            If Trim(tmpSeq) = "" Then
                msComm.Output = Chr$(6)
'                msComm.Output = Chr$(2) & "M" & Chr$(28) & "A" & Chr$(28) & Chr$(28) & "E2" & Chr$(3)
                Exit Sub
            End If
            
            With pResultInfo
                .ID = tmpSeq    'tmpPID
                .SEQNO = tmpSeq
                .RACK = ""
                .POS = ""
            End With
            
            For ii = 1 To iTestCnt
                tmpIFCd = tmpData(11 + ((ii - 1) * 4))
                tmpRst = tmpData(12 + ((ii - 1) * 4))
                tmpUnit = tmpData(13 + ((ii - 1) * 4))
                tmpFlag = tmpData(14 + ((ii - 1) * 4))
                If tmpRst = "" Then
                    tmpRst = "No Rst"
                End If
                                
                '결과값 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                    .FLAG = .FLAG & tmpFlag & Chr(124)
                End With
            Next ii
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With
            
            msComm.Output = Chr$(6)
'            msComm.Output = Chr$(2) & "M" & Chr$(28) & "A" & Chr$(28) & Chr$(28) & "E2" & Chr$(3)
    
        Case "M"        'REQUEST ACCEPTANCE Record
            msComm.Output = Chr$(6)
            
            tmpData() = Split(RcvBuffer, Chr(28))
            
            '또 보낼 오더있는지 확인하고 R이면 색변하지 않게...
            sStatus = tmpData(1)        '상태(A:정상,R:오류)
            
            RaiseEvent SendOrderOK(pSampleInfo.ID & Chr(124) & sStatus)
      
            If sStatus = "R" Then
                m_iSendPhase = 0
'                msComm.Output = Chr$(2) & "W" & Chr$(28) & "73" & Chr$(3)
            End If
            
        Case "I"        'QUERY Record
            msComm.Output = Chr$(6)
            
        Case "C"        'CALIBRATION RESULT MESSAGE
            msComm.Output = Chr$(6)
'            msComm.Output = Chr$(2) & "M" & Chr$(28) & "A" & Chr$(28) & Chr$(28) & "E2" & Chr$(3)
      
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub
Private Sub DataEditResponse_Dimension()
    On Error GoTo ErrRtn
    
    Dim sType   As String
    Dim sPollF  As String
    Dim sPollR  As String
    Dim tmpData()   As String
    Dim ii      As Integer
    
    Dim tmpPID  As String
    Dim tmpSeq  As String
    Dim iTestCnt    As Integer
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$
    Dim sStatus As String

    
    sType = Left$(RcvBuffer, 1)       ' get Frame of RcvBuffer.
   
    Select Case sType
        Case "P"        'POLL Record
            tmpData() = Split(RcvBuffer, Chr(28))
            
            sPollF = tmpData(2)
            sPollR = tmpData(3)
                    
            If sPollR = "1" Then
                'send a request
                msComm.Output = Chr(6)
                msComm.Output = Chr(2) & "N" & Chr(28) & "6A" & Chr(3)
            Else
                'busy send no request
                msComm.Output = Chr(6)
                msComm.Output = Chr(2) & "W" & Chr$(28) & "73" & Chr(3)
            End If
    
        Case "R"        'RESULT Record
            '결과정보 구조체 초기화
            Call Init_pResultInfo
            
            tmpData() = Split(RcvBuffer, Chr(28))
                        
            tmpPID = tmpData(3)         'Patient ID
            tmpSeq = tmpData(3)         'SampleNo
            iTestCnt = Val(tmpData(10)) 'Number of Tests
            
            If Trim(tmpSeq) = "" Then
                msComm.Output = Chr$(6)
                msComm.Output = Chr$(2) & "M" & Chr$(28) & "A" & Chr$(28) & Chr$(28) & "E2" & Chr$(3)
                Exit Sub
            End If
            
            With pResultInfo
                .ID = tmpPID
                .SEQNO = tmpSeq
                .RACK = ""
                .POS = ""
            End With
            
            For ii = 1 To iTestCnt
                tmpIFCd = tmpData(11 + ((ii - 1) * 4))
                tmpRst = tmpData(12 + ((ii - 1) * 4))
                tmpUnit = tmpData(13 + ((ii - 1) * 4))
                tmpFlag = tmpData(14 + ((ii - 1) * 4))
                If tmpRst = "" Then
                    tmpRst = "No Rst"
                End If
                                
                '결과값 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                    .FLAG = .FLAG & tmpFlag & Chr(124)
                End With
            Next ii
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With
            
            msComm.Output = Chr$(6)
            msComm.Output = Chr$(2) & "M" & Chr$(28) & "A" & Chr$(28) & Chr$(28) & "E2" & Chr$(3)
    
        Case "M"        'REQUEST ACCEPTANCE Record
            msComm.Output = Chr$(6)
            
            tmpData() = Split(RcvBuffer, Chr(28))
            
            sStatus = tmpData(1)        '상태(A:정상,R:오류)
      
            If sStatus = "R" Then
                msComm.Output = Chr$(2) & "W" & Chr$(28) & "73" & Chr$(3)
                RaiseEvent SendOrderOK("")
                RaiseEvent DispMsg(pSampleInfo.ID & " Order 전송 실퍠!!!")
            Else
                RaiseEvent SendOrderOK(pSampleInfo.ID)
            End If
            
        Case "I"        'QUERY Record
            msComm.Output = Chr$(6)
            
            tmpData() = Split(RcvBuffer, Chr(28))
            tmpPID = Trim(tmpData(1))
            
            pSampleInfo.ID = ""
            If Trim(tmpPID) = "" Then
                msComm.Output = Chr$(2) & "N" & Chr$(28) & "6A" & Chr$(3)
                Exit Sub
            End If
            
            pSampleInfo.ID = Trim(tmpPID)
            
            Call SendOrder_Dimension    'ORDER 전송
            
        Case "C"        'CALIBRATION RESULT MESSAGE
            msComm.Output = Chr$(6)
            msComm.Output = Chr$(2) & "M" & Chr$(28) & "A" & Chr$(28) & Chr$(28) & "E2" & Chr$(3)
      
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub Edit_Data()
'    On Error GoTo ErrHandler
'
'    '<---- COBAS 장비에서 주로 사용 S --->
'    Dim BC          As String
'    Dim LC          As String
'    Dim BCpos       As Integer
'    Dim LCpos       As Integer
'
'    Dim ErrCode     As Integer
'    Dim GeneralErrorCode    As String
'    '<---- COBAS 장비에서 주로 사용 E --->
'
'    '>>> Common Variable
'    Dim sLabDate$, sSlipCd$, sLabSeq$, sRack$, sPos$, sSampNo$ ' , sSampID$
'    Dim vLabDate, vSlipCd, vLabSeq, vRstCnt
'    Dim i%, j%, iCRow%
'    Dim iRstCnt%, iCmtCnt%
'    Dim sTIFCd$, sTRst$, sTComment$
'
'    Dim cnt     As Integer
'    Dim temp    As String
'
'    Dim LabDate As String
'    Dim TestCd As String
'    Dim Result As String
'    Dim LabSeq As String
'
'    Dim Frame As String
'    Dim Func As String
'    Dim TstCount As Integer
'    Dim StartPtr As Integer
'    Dim ResData  As String
'    Dim iRealCnt As Integer
'    Dim SampNo As String
'    Dim RackNo As String
'    Dim PosiNo As String
'    Dim Id_Num As String
'
'    Dim SendBuf As String
'    Dim ChekSum  As Integer             ' Check Sum
'    Dim ChkSumS  As String              ' Check Sum
'    Dim TestDat As String * 60
'    Dim i1stRow As Integer
'    Dim v1, v2, v3, v4, v5
'    Dim sTmp As String
'    Dim vTmp As Variant
'    Dim vIFCnt As Variant
'    Dim iPos As Integer
'    Dim iHDLchk As Integer
'
'    Dim sPollF  As String
'    Dim sPollR  As String
'    Dim tmpRID  As String       '장비로부터 전송된 ID 저장(12자리!!!)
'    Dim sInPhrase As String
'    Dim sBuf As String
'    Dim sOneRow As String
'    Dim iShowOk As Integer
'    Dim sSendBuff As String
'    Dim sSendBuff2 As String
'    Dim sStatus As String
'    Dim objAUIF As Object
'
'    Frame = Left$(RcvBuffer, 1)        ' get Frame of RcvBuffer.
'
'    Select Case Frame
'        Case "P"        'POLL Record
'
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            sPollF = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            sPollR = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'
'            If sPollR = 1 Then
'                'send a request
'                Comm1.Output = Chr$(6)
'                Comm1.Output = Chr$(2) & "N" & Chr$(28) & "6A" & Chr$(3)
'            Else
'                'busy send no request
'                Comm1.Output = Chr$(6)
'                Comm1.Output = Chr$(2) & "W" & Chr$(28) & "73" & Chr$(3)
'            End If
'
'        Case "R"        'RESULT Record
'
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            tmpRID = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            'LabSeq = "20" & Left(tmpRID, 6) & Mid(tmpRID, 7, 1) & "00" & Right(tmpRID, 3)
'            LabSeq = Left(UCase(tmpRID), 10)
'
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            'Number of Tests
'            TstCount = Val(GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28)))
'            iRealCnt = 0: sTIFCd = "": sTRst = ""
'            For i = 1 To TstCount
'                sTIFCd = sTIFCd & GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28)) & Chr$(124)
'                Result = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'                If Result = "" Then Result = "No Rst"
'                sTRst = sTRst & Result & Chr$(124)
'                Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'                Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'                iRealCnt = iRealCnt + 1
'            Next i
'
'            '현재의 전송과 매칭되는 Row 찾기
'            iCRow = FindCurRow(4, "", "", "", LabSeq)
'            '결과등록
'            If iCRow > 0 Then
'                Call ResultProcess(iCRow, iRealCnt, 0, sTIFCd, sTRst, "", "", LabSeq)
'            End If
'
'            Comm1.Output = Chr$(6)
'            Comm1.Output = Chr$(2) & "M" & Chr$(28) & "A" & Chr$(28) & Chr$(28) & "E2" & Chr$(3)
'
'        Case "M"        'REQUEST ACCEPTANCE Record
'
'            Comm1.Output = Chr$(6)
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            '또 보낼 오더있는지 확인하고 R이면 색변하지 않게...
'            sStatus = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))     '상태(A:정상,R:오류)
'            If Trim(sStatus) = "R" Then
'                Call SpdForeBack(spdIntList, 1, 4, gOrderTable.iCRow, gOrderTable.iCRow, RGB(0, 0, 0), 흐린빨강)
'                lblOrder = gOrderTable.sSampID & " 전송실패!!!"
'            Else
'                Call spdIntList.SetText(7, gOrderTable.iCRow, CStr(Val(gOrderTable.iOrdCnt)) & "")
'                Call SpdForeBack(spdIntList, 1, 4, gOrderTable.iCRow, gOrderTable.iCRow, RGB(0, 0, 0), 흐린파랑)
'                lblOrder = gOrderTable.sSampID
'            End If
'
'        Case "I"        'QUERY Record
'            Comm1.Output = Chr$(6)
'            SampNo = ""
'            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'            SampNo = GetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr$(28))
'
'            If Len(SampNo) <> 10 Then
'                Comm1.Output = Chr$(2) & "N" & Chr$(28) & "6A" & Chr$(3)
'            Else
'                sInPhrase = " IN ( "
'
'                For i = 1 To giIntItemCnt
'                    sBuf = sBuf & "'" & gIFItem(i).s06 & "', "
'                Next
'
'                sBuf = Left(sBuf, Len(sBuf) - 2)
'
'                sInPhrase = sInPhrase & sBuf & " ) "
'
'                Set objAUIF = CreateObject("AUIF01.DCUIF01")
'
'                Me.MousePointer = vbHourglass
'                sBuf = objAUIF.Get_TotIFList(0, gsMachineCd & "(" & gsSiteNm & ")", _
'                                               SampNo, "", "", sInPhrase)
'                Me.MousePointer = vbDefault
'
'                Set objAUIF = Nothing
'
'                If sBuf = "NONE" Then
'                    Comm1.Output = Chr$(2) & "N" & Chr$(28) & "6A" & Chr$(3)
'                    Exit Sub
'                Else
'
'                    iShowOk = 0: j = 0
'
'                    Do
'                        sOneRow = GetByOneUserSymbol(sBuf, sBuf, Chr(3))
'
'                        If sOneRow = "" Then
'                            Exit Do
'                        End If
'
'                        sOneRow = sOneRow & Chr(124)
'
'                        sLabDate = GetByOne(sOneRow, sOneRow)
'                        sSlipCd = GetByOne(sOneRow, sOneRow)
'                        sLabSeq = GetByOne(sOneRow, sOneRow)
'                        sRack = GetByOne(sOneRow, sOneRow)
'                        sPos = GetByOne(sOneRow, sOneRow)
'                        v1 = GetByOne(sOneRow, sOneRow)
'                        v2 = GetByOne(sOneRow, sOneRow)
'                        v3 = GetByOne(sOneRow, sOneRow)
'                        v4 = GetByOne(sOneRow, sOneRow)
'
'                        With spdIntList
'                            If iShowOk = 0 Then
'                                .MaxRows = .MaxRows + 1
'                                .RowHeight(.MaxRows) = 12.34
'                                Call .SetText(1, .MaxRows, Mid(sLabDate, 3, 6))
'                                Call .SetText(2, .MaxRows, sSlipCd & Mid(sLabSeq, 3, 3))
'                                Call .SetText(3, .MaxRows, sRack)
'                                Call .SetText(4, .MaxRows, sPos)
'                                iShowOk = 1
'                            End If
'
'                            Call .SetText(7, .MaxRows, "0") 'O
'                            Call .SetText(8, .MaxRows, "0") 'R
'                            j = j + 1
'                            Call .SetText(9 + j, .MaxRows, v1 & v2 & v3 & v4)
'                        End With
'                    Loop
'                    Call spdIntList.SetText(9, spdIntList.MaxRows, CVar(Str(j))) 'IF
'                End If
'                '=============================
'                gOrderTable.iCRow = spdIntList.MaxRows
'                gOrderTable.sSampID = SampNo
'                gOrderTable.iOrdCnt = j
'
'                ReDim gOrderTable.sIFTestCd(CInt(j))
'
'                For i = 1 To CInt(j)
'                    Call spdIntList.GetText(9 + i, spdIntList.MaxRows, vTmp)
'
'                    If vTmp = "" Or IsNull(vTmp) Then
'                        ViewMsg "검사항목 오류발견!! Order 전송이 취소되었습니다..."
'                        Comm1.Output = Chr$(2) & "N" & Chr$(28) & "6A" & Chr$(3)
'                        Exit Sub
'                    End If
'
'                    sTmp = CStr(vTmp) & "|"
'
'                    sTmp = GetIFTestCd_IFSpcCd(GetByOne(sTmp, sTmp))
'
'                    iPos = InStr(1, sTmp, "|")
'
'                    gOrderTable.sIFTestCd(i) = Mid$(sTmp, 1, iPos - 1)
'
'                Next
'
'                '============== ORDER 만들기 ===================
'                iHDLchk = 0: sSendBuff = "": sSendBuff2 = ""
'
'                '검사항목 Order코드 추가 (HDL제외)
'                For i = 1 To gOrderTable.iOrdCnt
'                    If gOrderTable.sIFTestCd(i) <> "HDL" Then
'                        sSendBuff2 = sSendBuff2 & gOrderTable.sIFTestCd(i) & ""
'                    Else
'                        iHDLchk = 1
'                    End If
'                Next i
'
'                sSendBuff = "D00A" & gOrderTable.sSampID & "1011" & Trim(Str(gOrderTable.iOrdCnt - iHDLchk)) & ""
'                sSendBuff = sSendBuff & sSendBuff2
'
'                ChekSum = 0
'
'                For i = 1 To Len(sSendBuff)
'                    ChekSum = ChekSum + Asc(Mid$(sSendBuff, i, 1))
'                Next
'
'                ChkSumS = Right$("00" & Hex$(ChekSum), 2)
'
'                sSendBuff = Chr$(2) & sSendBuff & ChkSumS & Chr$(3)
'                snd_buf = sSendBuff
'                Comm1.Output = sSendBuff
''                Print #302, sSendBuff;
'            End If
'
'        Case "C"        'CALIBRATION RESULT MESSAGE
'            Comm1.Output = Chr$(6)
'            Comm1.Output = Chr$(2) & "M" & Chr$(28) & "A" & Chr$(28) & Chr$(28) & "E2" & Chr$(3)
'
'    End Select
'
'    Exit Sub
'
'ErrHandler:
'    ViewMsg "Edit_Data - " & Err.Description & "(" & CStr(Val(Err.Number)) & ")"
End Sub


Private Sub SendOrder_Dimension()
    
    Dim sSend   As String
    Dim sTest   As String
    Dim ii      As Integer
    Dim CheckSum    As String
    Dim sCS     As String

        
    RaiseEvent RequestCurOrder(pSampleInfo.ID, "", "")
    
    'ORDER 편집
    Call Get_OrderString
    
    If Trim(pSampleInfo.ID) = "" Or pSampleInfo.ORDCNT = 0 Then
        msComm.Output = Chr$(2) & "N" & Chr$(28) & "6A" & Chr$(3)
        Exit Sub
    End If
    
    sSend = "": sTest = "": sSend_Buf = ""
    
    '검사항목 Order코드 추가
    For ii = 1 To pSampleInfo.ORDCNT
        sTest = sTest & pSampleInfo.IFCD(ii) & Chr(28)
    Next ii
    
    sSend = "D" & Chr(28)               'TYPE
    sSend = sSend & "0" & Chr(28)       'SAMPLE CARRIER ID
    sSend = sSend & "0" & Chr(28)       'LOADLIST ID
    sSend = sSend & "A" & Chr(28)       'TRANSACTION
    sSend = sSend & pSampleInfo.ID & Chr(28)    'PID
    sSend = sSend & pSampleInfo.ID & Chr(28)    'SAMPLE #
    sSend = sSend & "1" & Chr(28)       'SAMPLE TYPE
    sSend = sSend & "" & Chr(28)        'LOCATION
    sSend = sSend & "0" & Chr(28)       'PRIORITY
    sSend = sSend & "1" & Chr(28)       '# OF CUPS FOR SAMPLE
    sSend = sSend & "" & Chr(28)        'CUP POSITION
    sSend = sSend & "1" & Chr(28)       'DILUTION
    sSend = sSend & Trim(pSampleInfo.ORDCNT) & Chr(28)      '# OF TESTS
    sSend = sSend & sTest               'ORDER
        
    CheckSum = 0
    For ii = 1 To Len(sSend)
        CheckSum = CheckSum + Asc(Mid$(sSend, ii, 1))
    Next ii

    sCS = Right$("00" & Hex$(CheckSum), 2)

    sSend = Chr$(2) & sSend & sCS & Chr$(3)
    sSend_Buf = sSend
    
    msComm.Output = sSend
                
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    
End Sub
Private Sub SendOrder_DimensionUni()
    
    Dim sSend   As String
    Dim sTest   As String
    Dim ii      As Integer
    Dim CheckSum    As String
    Dim sCS     As String
    Dim sSamplePos  As String
    Dim sSampleNo   As String

        
    'ORDER 편집
    Call Get_OrderString
    
    If Trim(pSampleInfo.ID) = "" Or pSampleInfo.ORDCNT = 0 Then
        m_iSendPhase = 0
        msComm.Output = Chr$(2) & "W" & Chr$(28) & "73" & Chr$(3)
        Exit Sub
    End If
    
    sSend = "": sTest = "": sSend_Buf = ""
    
    'Sample # 편집(12자리)
    sSampleNo = Left(pSampleInfo.ID, 8) & Mid(pSampleInfo.ID, 13)
    
    'Sample Position 편집
'    sSamplePos = pSampleInfo.RACK & Val(pSampleInfo.POS)
    sSamplePos = Trim((Asc(pSampleInfo.RACK) - 64) * Val(pSampleInfo.POS))
    
    
    '검사항목 Order코드 추가
    For ii = 1 To pSampleInfo.ORDCNT
        sTest = sTest & pSampleInfo.IFCD(ii) & Chr(28)
    Next ii
    
    sSend = "D" & Chr(28)               'TYPE
    sSend = sSend & "0" & Chr(28)       'SAMPLE CARRIER ID
    sSend = sSend & "0" & Chr(28)       'LOADLIST ID
    sSend = sSend & "A" & Chr(28)       'TRANSACTION
    sSend = sSend & pSampleInfo.ID & Chr(28)    'PID
    sSend = sSend & sSampleNo & Chr(28)         'SAMPLE #
    sSend = sSend & "1" & Chr(28)       'SAMPLE TYPE
    sSend = sSend & "" & Chr(28)        'LOCATION
    sSend = sSend & "0" & Chr(28)       'PRIORITY
    sSend = sSend & "1" & Chr(28)       '# OF CUPS FOR SAMPLE
    sSend = sSend & sSamplePos & Chr(28)    'CUP POSITION
    sSend = sSend & "1" & Chr(28)       'DILUTION
    sSend = sSend & Trim(pSampleInfo.ORDCNT) & Chr(28)      '# OF TESTS
    sSend = sSend & sTest               'ORDER
        
    CheckSum = 0
    For ii = 1 To Len(sSend)
        CheckSum = CheckSum + Asc(Mid$(sSend, ii, 1))
    Next ii

    sCS = Right$("00" & Hex$(CheckSum), 2)

    sSend = Chr$(2) & sSend & sCS & Chr$(3)
    sSend_Buf = sSend
    
    msComm.Output = sSend
                
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    
End Sub


Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
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
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
    End If
End Function

