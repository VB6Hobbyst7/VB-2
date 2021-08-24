VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl ADVIA1650 
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
Attribute VB_Name = "ADVIA1650"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sCmt1 = ""
Const m_def_p_sSpcCd = 0
Const m_def_p_sTSVol = "0"
Const m_def_p_sRerunGbn = "0"
Const m_def_p_bSIndex = 0
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
Dim m_p_sCmt1 As String
Dim m_p_sSpcCd As Variant
Dim m_p_sTSVol As String
Dim m_p_sRerunGbn As String
Dim m_p_bSIndex As Boolean
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
Event DispMsgComm(sMsg$)
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTInstID$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sRack$, sPos$, sKind$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
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

'For E-170/Hitachi7600
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim sNextSend   As String
Dim RstEnd      As String

'for ADVIA1650
Dim iPendingFlag    As Integer
Dim iTotQueryFlag   As Integer
Dim iTmpPendingFlag As Integer
Dim iIdleFlag   As Integer
Dim iOrderFlag  As Integer
Dim iResultFlag As Integer
Dim sRcvState   As String
Dim sSndState   As String
    
Dim sSndPacket()    As String
Dim sQueryBarcd()   As String

Public piDateLen As Integer
Public piUsrCdLen As Integer

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
        Case "ADVIA1650"
            Call PhaseCfg_Protocol_ADVIA1650
            
        Case "ADVIA1800", "ADVIA2400"
            If m_bUseBarcode = True Then
                Call PhaseCfg_Protocol_ADVIA1800
            Else
                Call PhaseCfg_Protocol_ADVIA1800_Batch
            End If
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub DataEditResponse_ADVIA1650()
    On Error GoTo ErrHandler
    
'<---- COBAS 장비에서 주로 사용 S --->
    Dim sBC          As String
    Dim sLC          As String
    Dim iBCpos       As Integer
    Dim iLCpos       As Integer
    
    Dim iErrCode     As Integer
    Dim sGeneralErrCode    As String
'<---- COBAS 장비에서 주로 사용 E --->
    Dim i           As Integer
    Dim sTmp        As String
    Dim sRetVal     As String
    
    Dim TResult(100) As String
    Dim TCode(100)  As String
    Dim TFlag(100)  As String
    Dim sTotTestCd  As String
    Dim sTotRst     As String
    Dim sTotRst2    As String
    
    Dim lngRetVal   As Long
    Dim sBuf        As String
    Dim sSvrCd      As String
    Dim iTBlockNo   As Integer
    Dim iCBlockNo   As Integer
    Dim iItemNo     As Integer
    Dim iMachRstCnt As Integer
    Dim iPos        As Integer
    
    Dim tmpBarCd$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    
    iBCpos = 3
    sBC = Mid$(RcvBuffer, iBCpos, 1)
    
    Select Case sBC
        Case "q"
            sRcvState = "Q"
            sSndState = ""
            
            iPendingFlag = CStr(Val(Mid$(RcvBuffer, iBCpos + 6, 2)))
            
            For i = 1 To iPendingFlag
                tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 7 + 13 * (i - 1), 13))
    
                If Len(tmpBarCd) = 0 Then
                Else
                    pSampleInfo.ID = tmpBarCd
                                        
                    'Order 가져오는 부분
                    Call SendOrder_ADVIA1650
                End If
            Next i
            
        Case "Q"
            sRcvState = "Q"
            sSndState = ""
            
            'QueryFlag = Val(Mid$(RcvBuffer, iBCpos - 1, 1))
            'TotQueryFlag = Val(Mid$(RcvBuffer, iBCpos + 4, 2))
            
            iTmpPendingFlag = Val(Mid$(RcvBuffer, iBCpos + 6, 2))
            
            iPendingFlag = iPendingFlag + iTmpPendingFlag
            
            For i = 1 To iTmpPendingFlag
                tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 9 + 13 * (i - 1), 13))
    
                If Len(tmpBarCd) = 0 Then
                    sRcvState = ""
                Else
                    pSampleInfo.ID = tmpBarCd
                                    
                    'Order 가져오는 부분
                    Call SendOrder_ADVIA1650
                End If
            Next
        
        Case "R"
            sRcvState = "R"
            
            '결과구조체 초기화
            Call Init_pResultInfo
            
            iTBlockNo = Val(Mid$(RcvBuffer, iBCpos + 2, 2))
            iCBlockNo = Val(Mid$(RcvBuffer, iBCpos + 4, 2))
            iItemNo = Val(Mid$(RcvBuffer, iBCpos + 6, 3))
            sLC = Mid$(RcvBuffer, iBCpos + 17, 1)
            tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 19, 13))
            
            sTmp = Trim$(Mid$(RcvBuffer, iBCpos + 32, 7))
            
            iPos = InStr(sTmp, "-")
                     
            If iPos = 0 Then
                tmpRack = ""
                tmpPos = ""
            Else
                tmpRack = Mid$(sTmp, 1, iPos - 1)
                tmpPos = Mid$(sTmp, iPos + 1)
            End If
            
            'QC 결과 빠지기
            If sLC = "C" Then
                Exit Sub
            End If
            
            If iCBlockNo = 1 Then
                For i = 1 To iItemNo
                    TCode(i) = Trim$(Mid(RcvBuffer, iBCpos + 89 + 15 * (i - 1), 3))
                    TResult(i) = Trim(Mid(RcvBuffer, iBCpos + 89 + 4 + 15 * (i - 1), 8))
                    TFlag(i) = Trim(Mid(RcvBuffer, iBCpos + 89 + 8 + 4 + 15 * (i - 1), 3))
                Next
            Else
                For i = 1 To iItemNo
                    TCode(i) = Trim$(Mid(RcvBuffer, iBCpos + 39 + 15 * (i - 1), 3))
                    TResult(i) = Trim(Mid(RcvBuffer, iBCpos + 39 + 4 + 15 * (i - 1), 8))
                    TFlag(i) = Trim(Mid(RcvBuffer, iBCpos + 39 + 8 + 4 + 15 * (i - 1), 3))
                Next
            End If
            
            iMachRstCnt = iMachRstCnt + iItemNo
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = tmpBarCd
                .RACK = tmpRack
                .POS = tmpPos
            
                For i = 1 To iMachRstCnt
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & TCode(i) & Chr$(124)
                    .RST1 = .RST1 & TResult(i) & Chr$(124)
                    .RST2 = .RST2 & Chr$(124)
                    .FLAG = .FLAG & TFlag(i) & Chr(124)
                Next i
            End With
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .Kind, .RSTDT, "")
                End If
            End With

            Call Init_pResultInfo
    
    End Select
       
    RcvBuffer = ""
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
        RcvBuffer = ""
    End If
End Sub

Private Sub DataEditResponse_ADVIA1800(Optional ByVal sMode As String)
    On Error GoTo ErrHandler
    
'<---- COBAS 장비에서 주로 사용 S --->
    Dim sBC          As String
    Dim sLC          As String
    Dim iBCpos       As Integer
    Dim iLCpos       As Integer
    
    Dim iErrCode     As Integer
    Dim sGeneralErrCode    As String
'<---- COBAS 장비에서 주로 사용 E --->
    Dim i           As Integer
    Dim sTmp        As String
    Dim sRetVal     As String
    
    Dim TResult(100) As String
    Dim TCode(100)  As String
    Dim TFlag(100)  As String
    Dim sTotTestCd  As String
    Dim sTotRst     As String
    Dim sTotRst2    As String
    
    Dim lngRetVal   As Long
    Dim sBuf        As String
    Dim sSvrCd      As String
    Dim iTBlockNo   As Integer
    Dim iCBlockNo   As Integer
    Dim iItemNo     As Integer
    Dim iMachRstCnt As Integer
    Dim iPos        As Integer
    
    Dim tmpBarCd$, tmpRack$, tmpPos$, tmpKind$, tmpSpcCd$, tmpRerun$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    
    iBCpos = 2
    
    If sMode = "EOT" Then   'EOT일때 결과처리..
        '결과값 등록/화면 표시 처리...
        With pResultInfo
            If .RSTCNT > 0 Then
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .Kind, .RSTDT, "")
            End If
        End With
        
        Call Init_pResultInfo
        Exit Sub
    End If
    
    If (RcvBuffer = "" Or Len(RcvBuffer) <= 2) Then
        Exit Sub
    End If
    
    sBC = Mid$(RcvBuffer, iBCpos, 1)
    
    Select Case sBC
        Case "q"        'Batch-Test Query Text
            sRcvState = "Q"
            sSndState = ""
            
            iPendingFlag = CStr(Val(Mid$(RcvBuffer, iBCpos + 6, 2)))
            
            For i = 1 To iPendingFlag
                tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 7 + 13 * (i - 1), 13))
    
                If Len(tmpBarCd) = 0 Then
                Else
                    pSampleInfo.ID = tmpBarCd
                                        
                    'Order 가져오는 부분
                    Call SendOrder_ADVIA1800
                End If
            Next i
            
        Case "Q"        'Test Request Text
            sRcvState = "Q"
            sSndState = ""

            iTmpPendingFlag = Val(Mid$(RcvBuffer, iBCpos + 6, 2))
            
            iPendingFlag = iPendingFlag + iTmpPendingFlag
            
            For i = 1 To iTmpPendingFlag
                tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 9 + 13 * (i - 1), 13))
    
                If Len(tmpBarCd) = 0 Then
                    sRcvState = ""
                Else
                    pSampleInfo.ID = tmpBarCd
                                    
                    'Order 가져오는 부분
                    Call SendOrder_ADVIA1800
                End If
            Next
        
        Case "R"
            sRcvState = "R"
            
''            '결과구조체 초기화
''            Call Init_pResultInfo
            
            iTBlockNo = Val(Mid$(RcvBuffer, iBCpos + 2, 2))
            iCBlockNo = Val(Mid$(RcvBuffer, iBCpos + 4, 2))
            iItemNo = Val(Mid$(RcvBuffer, iBCpos + 6, 3))
            
            tmpKind = Mid$(RcvBuffer, iBCpos + 17, 1)       'N:Sample, C:Control
            
            tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 19, 13))
            
            sTmp = Trim$(Mid$(RcvBuffer, iBCpos + 32, 7))
            
            iPos = InStr(sTmp, "-")
                     
            If iPos = 0 Then
                tmpRack = ""
                tmpPos = ""
            Else
                tmpRack = Mid$(sTmp, 1, iPos - 1)
                tmpPos = Mid$(sTmp, iPos + 1)
            End If
            
            If tmpKind = "C" Then       'Control Result
                tmpKind = "QC"
'                tmpBarCd = Mid(tmpBarCd, 2)
            Else
                tmpKind = ""
            End If
            
            tmpRerun = ""
            
            If iCBlockNo = 1 Then
                For i = 1 To iItemNo
                    TCode(i) = Trim$(Mid(RcvBuffer, iBCpos + 89 + 15 * (i - 1), 3))
                    TResult(i) = Trim(Mid(RcvBuffer, iBCpos + 89 + 4 + 15 * (i - 1), 8))
                    TFlag(i) = Trim(Mid(RcvBuffer, iBCpos + 89 + 8 + 4 + 15 * (i - 1), 3))
                    
                    If InStr(TFlag(i), "R") > 0 Then
                        tmpRerun = "R"
                        TFlag(i) = Replace(TFlag(i), "R", "")
                    End If
                Next i
            Else
                For i = 1 To iItemNo
                    TCode(i) = Trim$(Mid(RcvBuffer, iBCpos + 39 + 15 * (i - 1), 3))
                    TResult(i) = Trim(Mid(RcvBuffer, iBCpos + 39 + 4 + 15 * (i - 1), 8))
                    TFlag(i) = Trim(Mid(RcvBuffer, iBCpos + 39 + 8 + 4 + 15 * (i - 1), 3))
                    
                    If InStr(TFlag(i), "R") > 0 Then
                        tmpRerun = "R"
                        TFlag(i) = Replace(TFlag(i), "R", "")
                    End If
                Next i
            End If
            
            If tmpRerun = "R" Then      'Rerun Result
                tmpKind = tmpKind & "R"
            End If
            
            iMachRstCnt = iMachRstCnt + iItemNo
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = tmpBarCd
                .RACK = tmpRack
                .POS = tmpPos
                .Kind = tmpKind
                
                For i = 1 To iMachRstCnt
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & TCode(i) & Chr$(124)
                    .RST1 = .RST1 & TResult(i) & Chr$(124)
                    .RST2 = .RST2 & Chr$(124)
                    .FLAG = .FLAG & TFlag(i) & Chr(124)
                    
                    .ALARMCD = .ALARMCD & Chr(124)
                    .RSTDT = .RSTDT & Chr(124)
                Next i
            End With
            
''            '결과값 등록/화면 표시 처리...
''            With pResultInfo
''                If .RSTCNT > 0 Then
''                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .KIND, .RSTDT, "")
''                End If
''            End With
''
''            Call Init_pResultInfo
            
    End Select
       
    RcvBuffer = ""
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
        RcvBuffer = ""
    End If
End Sub

Private Sub DataEditResponse_ADVIA1800_Batch(Optional ByVal sMode As String)
    On Error GoTo ErrHandler
    
'<---- COBAS 장비에서 주로 사용 S --->
    Dim sBC          As String
    Dim sLC          As String
    Dim iBCpos       As Integer
    Dim iLCpos       As Integer
    
    Dim iErrCode     As Integer
    Dim sGeneralErrCode    As String
'<---- COBAS 장비에서 주로 사용 E --->
    Dim i           As Integer
    Dim sTmp        As String
    Dim sRetVal     As String
    
    Dim TResult(100) As String
    Dim TCode(100)  As String
    Dim TFlag(100)  As String
    Dim sTotTestCd  As String
    Dim sTotRst     As String
    Dim sTotRst2    As String
    
    Dim lngRetVal   As Long
    Dim sBuf        As String
    Dim sSvrCd      As String
    Dim iTBlockNo   As Integer
    Dim iCBlockNo   As Integer
    Dim iItemNo     As Integer
    Dim iMachRstCnt As Integer
    Dim iPos        As Integer
    Dim iSIdx       As Integer
    
    Dim tmpBarCd$, tmpRack$, tmpPos$, tmpKind$, tmpSpcCd$, tmpRerun$, tmpRstDt$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    
    iBCpos = 2
    
    If sMode = "EOT" Then   'EOT일때 결과처리..
        '결과값 등록/화면 표시 처리...
        With pResultInfo
            If .RSTCNT > 0 Then
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .Kind, .RSTDT, "")
            End If
        End With
        
        Call Init_pResultInfo
        Exit Sub
    End If
    
    If (RcvBuffer = "" Or Len(RcvBuffer) <= 2) Then
        Exit Sub
    End If
    
    sBC = Mid$(RcvBuffer, iBCpos, 1)
    
    Select Case sBC
        Case "q"        'Batch-Test Query Text
            sRcvState = "Q"
            sSndState = ""
            
            iPendingFlag = CStr(Val(Mid$(RcvBuffer, iBCpos + 6, 2)))
            
            For i = 1 To iPendingFlag
                tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 7 + 13 * (i - 1), 13))
    
                If Len(tmpBarCd) = 0 Then
                Else
                    pSampleInfo.ID = tmpBarCd
                                        
                    'Order 가져오는 부분
                    Call SendOrder_ADVIA1800_Batch
                End If
            Next i
            
        Case "Q"        'Test Request Text
            sRcvState = "Q"
            sSndState = ""

            iTmpPendingFlag = Val(Mid$(RcvBuffer, iBCpos + 6, 2))
            
            iPendingFlag = iPendingFlag + iTmpPendingFlag
            
            For i = 1 To iTmpPendingFlag
                tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 9 + 13 * (i - 1), 13))
    
                If Len(tmpBarCd) = 0 Then
                    sRcvState = ""
                Else
                    pSampleInfo.ID = tmpBarCd
                                    
                    'Order 가져오는 부분
                    Call SendOrder_ADVIA1800_Batch
                End If
            Next
        
        Case "R"
            sRcvState = "R"
            
            '1R 010100120130523101033N144593        01-01                                  M  000000000 1.011 28M     7.1    XYZ 6E
            iTBlockNo = Val(Mid$(RcvBuffer, iBCpos + 2, 2))
            iCBlockNo = Val(Mid$(RcvBuffer, iBCpos + 4, 2))
            iItemNo = Val(Mid$(RcvBuffer, iBCpos + 6, 3))
            
            ''tmpRstDt = Mid$(RcvBuffer, iBCpos + 9, 8)
            
            '<장비 DateTime Format에 따라 달라짐..
            tmpRstDt = Mid$(RcvBuffer, iBCpos + 9, piDateLen)
            iBCpos = iBCpos + piDateLen - 8
            '>
            
            tmpKind = Mid$(RcvBuffer, iBCpos + 17, 1)       'N:Sample, C:Control
            tmpBarCd = Trim$(Mid$(RcvBuffer, iBCpos + 19, 13))
            
            sTmp = Trim$(Mid$(RcvBuffer, iBCpos + 32, 7))
            
            iPos = InStr(sTmp, "-")
                     
            If iPos = 0 Then
                tmpRack = ""
                tmpPos = ""
            Else
                tmpRack = Mid$(sTmp, 1, iPos - 1)
                tmpPos = Mid$(sTmp, iPos + 1)
            End If
            
            If tmpKind = "C" Then       'Control Result
                tmpKind = "QC"
'                tmpBarCd = Mid(tmpBarCd, 2)
            Else
                tmpKind = ""
            End If
            
            tmpRerun = ""
                        
''            If iCBlockNo = 1 Then
''                For i = 1 To iItemNo
''                    TCode(i) = Trim$(Mid(RcvBuffer, iBCpos + 89 + 15 * (i - 1), 3))
''                    TResult(i) = Trim(Mid(RcvBuffer, iBCpos + 89 + 4 + 15 * (i - 1), 8))
''                    TFlag(i) = Trim(Mid(RcvBuffer, iBCpos + 89 + 8 + 4 + 15 * (i - 1), 3))
''
''                    If InStr(TFlag(i), "R") > 0 Then
''                        tmpRerun = "R"
''                        TFlag(i) = Replace(TFlag(i), "R", "")
''                    End If
''                Next i
''            Else
''                For i = 1 To iItemNo
''                    TCode(i) = Trim$(Mid(RcvBuffer, iBCpos + 39 + 15 * (i - 1), 3))
''                    TResult(i) = Trim(Mid(RcvBuffer, iBCpos + 39 + 4 + 15 * (i - 1), 8))
''                    TFlag(i) = Trim(Mid(RcvBuffer, iBCpos + 39 + 8 + 4 + 15 * (i - 1), 3))
''
''                    If InStr(TFlag(i), "R") > 0 Then
''                        tmpRerun = "R"
''                        TFlag(i) = Replace(TFlag(i), "R", "")
''                    End If
''                Next i
''            End If

            iSIdx = 15 + piUsrCdLen '장비 User code 설정에 따라 달라짐..

            If iCBlockNo = 1 Then
                For i = 1 To iItemNo
                    TCode(i) = Trim$(Mid(RcvBuffer, iBCpos + 89 + iSIdx * (i - 1), 3))
                    TResult(i) = Trim(Mid(RcvBuffer, iBCpos + 89 + 4 + iSIdx * (i - 1), 8))
                    TFlag(i) = Trim(Mid(RcvBuffer, iBCpos + 89 + 8 + 4 + iSIdx * (i - 1), 3))

                    If InStr(TFlag(i), "R") > 0 Then
                        tmpRerun = "R"
                        TFlag(i) = Replace(TFlag(i), "R", "")
                    End If
                Next i
            Else
                For i = 1 To iItemNo
                    TCode(i) = Trim$(Mid(RcvBuffer, iBCpos + 39 + iSIdx * (i - 1), 3))
                    TResult(i) = Trim(Mid(RcvBuffer, iBCpos + 39 + 4 + iSIdx * (i - 1), 8))
                    TFlag(i) = Trim(Mid(RcvBuffer, iBCpos + 39 + 8 + 4 + iSIdx * (i - 1), 3))

                    If InStr(TFlag(i), "R") > 0 Then
                        tmpRerun = "R"
                        TFlag(i) = Replace(TFlag(i), "R", "")
                    End If
                Next i
            End If
            
            If tmpRerun = "R" Then      'Rerun Result
                tmpKind = tmpKind & "R"
            End If
            
            iMachRstCnt = iMachRstCnt + iItemNo
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = tmpBarCd
                .RACK = tmpRack
                .POS = tmpPos
                .Kind = tmpKind
                
                For i = 1 To iMachRstCnt
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & TCode(i) & Chr$(124)
                    .RST1 = .RST1 & TResult(i) & Chr$(124)
                    .RST2 = .RST2 & Chr$(124)
                    .FLAG = .FLAG & TFlag(i) & Chr(124)
                    
                    .ALARMCD = .ALARMCD & Chr(124)
                    .RSTDT = .RSTDT & Chr(124)
                Next i
            End With
            
''            '결과값 등록/화면 표시 처리...
''            With pResultInfo
''                If .RSTCNT > 0 Then
''                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .KIND, .RSTDT, "")
''                End If
''            End With
''
''            Call Init_pResultInfo
            
    End Select
       
    RcvBuffer = ""
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
        RcvBuffer = ""
    End If
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

Private Sub SendOrder_ADVIA1650()
    On Error GoTo ErrHandler
    
    '환자의 Order 전송
    Dim SendBuff As String
    Dim i%, j%, k%, iOrdCnt%
    Dim vIFCnt, vTmp
    Dim sTmp$, sTestCd$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$, sFrameNo$
    
    SendBuff = ""
    sTmp = ""
    
    Select Case sSndState
        Case ""
            '----- 검사항목 조회
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, "")
            
            Call Get_OrderString
        
            If pSampleInfo.ORDCNT = 0 Then
                iIdleFlag = CStr(Val(iIdleFlag) + 1)
            
                sFrameNo = CStr(Val(iIdleFlag) Mod 8)
                        
                SendBuff = sFrameNo & "O" & " " & "0101"
                SendBuff = SendBuff & "000"     'Format$(gOrderTable.iOrdCnt, "000")
                SendBuff = SendBuff & "N" & "2"
                SendBuff = SendBuff & Left$(pSampleInfo.ID & Space(12), 12) & " "
                SendBuff = SendBuff & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)
                SendBuff = SendBuff & Space$(8) & " 1.0" & "1" & "1"
                SendBuff = SendBuff & Space$(1) & Chr(3)
                                        
                'n개의 sSndPacket 구성
                ReDim Preserve sSndPacket(Val(iIdleFlag))
                sSndPacket(Val(iIdleFlag)) = Chr(2) & SendBuff & ChkSum_ASTM(SendBuff) & vbCr & vbLf
                                        
                pSampleInfo.ORDCNT = 0
                
                '조회된 내용이 없는 경우 환자정보 구조체 초기화
                Call Init_pResultInfo
        
                RaiseEvent SendOrderOK("", "", "", "")
                
                Exit Sub
            End If
            
            sTestCd = ""
            
            For i = 1 To pSampleInfo.ORDCNT
                sTestCd = sTestCd & Right(Space(3) & Trim$(pSampleInfo.IFCD(i)), 3) & "M"
            Next i
            
            iIdleFlag = CStr(Val(iIdleFlag) + 1)
        
            sFrameNo = CStr(Val(iIdleFlag) Mod 8)
                    
            SendBuff = sFrameNo & "O" & " " & "0101"
            SendBuff = SendBuff & Format$(pSampleInfo.ORDCNT, "000")
            SendBuff = SendBuff & "N" & "0"
            SendBuff = SendBuff & Left$(pSampleInfo.ID & Space(12), 12) & " "
            SendBuff = SendBuff & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)
            SendBuff = SendBuff & Space$(8) & " 1.0" & "1" & "1"
            SendBuff = SendBuff & sTestCd & Space$(1) & Chr(3)
            
            'n개의 sSndPacket 구성
            ReDim Preserve sSndPacket(Val(iIdleFlag))
            sSndPacket(Val(iIdleFlag)) = Chr(2) & SendBuff & ChkSum_ASTM(SendBuff) & vbCr & vbLf
                         
            RaiseEvent SendOrderOK(pSampleInfo.ID, "", "", "")
                        
        Case "E"
            '처음 Packet 전송
            msComm.Output = sSndPacket(1)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(sSndPacket(1))
            End If
            
            iOrderFlag = 1
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "P"
            'Packet 전송
            iOrderFlag = iOrderFlag + 1
            
            msComm.Output = sSndPacket(iOrderFlag)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(sSndPacket(iOrderFlag))
            End If
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "L"
            'EOT 전송
            msComm.Output = Chr(4)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If
            
            '초기화
            iOrderFlag = 0: iPendingFlag = 0: iIdleFlag = 0: iTotQueryFlag = 0
            
        Case Else
    End Select
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_ADVIA1800()
    On Error GoTo ErrHandler
    
    '환자의 Order 전송
    Dim SendBuff As String
    Dim i%, j%, k%, iOrdCnt%
    Dim vIFCnt, vTmp
    Dim sTmp$, sTestCd$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$, sFrameNo$
    
    SendBuff = ""
    sTmp = ""
    
    Select Case sSndState
        Case ""
            '----- 검사항목 조회
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, "")
            
            Call Get_OrderString
        
            If pSampleInfo.ORDCNT = 0 Then      'Order 없는 경우
                iIdleFlag = CStr(Val(iIdleFlag) + 1)
            
                sFrameNo = CStr(Val(iIdleFlag) Mod 8)
                        
                SendBuff = sFrameNo & "O" & " " & "0101"
                SendBuff = SendBuff & "000"     'Format$(gOrderTable.iOrdCnt, "000")
                SendBuff = SendBuff & "N" & "2"
                SendBuff = SendBuff & Left$(pSampleInfo.ID & Space(13), 13)
                SendBuff = SendBuff & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)
                SendBuff = SendBuff & Space$(8) & " 1.0" & "1" & "1"
                SendBuff = SendBuff & Space$(1) & Chr(3)
                                        
                'n개의 sSndPacket 구성
                ReDim Preserve sSndPacket(Val(iIdleFlag))
                sSndPacket(Val(iIdleFlag)) = Chr(2) & SendBuff & ChkSum_ASTM(SendBuff) & vbCr & vbLf
                                        
                pSampleInfo.ORDCNT = 0
                
                '조회된 내용이 없는 경우 환자정보 구조체 초기화
                Call Init_pResultInfo
        
                RaiseEvent SendOrderOK("", "", "", "")
                
                Exit Sub
            End If
            
            sTestCd = ""
            
            For i = 1 To pSampleInfo.ORDCNT
                sTestCd = sTestCd & Right(Space(3) & Trim$(pSampleInfo.IFCD(i)), 3) & "M"
            Next i
            
            iIdleFlag = CStr(Val(iIdleFlag) + 1)
        
            sFrameNo = CStr(Val(iIdleFlag) Mod 8)
                    
            SendBuff = sFrameNo & "O" & " " & "0101"
            SendBuff = SendBuff & Format$(pSampleInfo.ORDCNT, "000")
            SendBuff = SendBuff & "N"                                   'Sample classification
            SendBuff = SendBuff & "0"                                   'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
            SendBuff = SendBuff & Left$(pSampleInfo.ID & Space(13), 13)
            SendBuff = SendBuff & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)
            
'            SendBuff = SendBuff & Space$(8) & " 1.0" & "1" & "1"
            SendBuff = SendBuff & Space$(8)
            SendBuff = SendBuff & " 1.0"        'Dilution coefficient(4)
            
            If pSampleInfo.SPCCD = "2" Then
                SendBuff = SendBuff & "2"           'Sample classification(1:blood serum, 2:urine)
            Else
                SendBuff = SendBuff & "1"           'Sample classification(1:blood serum, 2:urine)
            End If
            
            SendBuff = SendBuff & "1"           'Container classification
            
            SendBuff = SendBuff & sTestCd & Space$(1) & Chr(3)
            
            'n개의 sSndPacket 구성
            ReDim Preserve sSndPacket(Val(iIdleFlag))
            sSndPacket(Val(iIdleFlag)) = Chr(2) & SendBuff & ChkSum_ASTM(SendBuff) & vbCr & vbLf
                         
            RaiseEvent SendOrderOK(pSampleInfo.ID, "", "", "")
                        
        Case "E"
            '처음 Packet 전송
            msComm.Output = sSndPacket(1)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(sSndPacket(1))
            End If
            
            iOrderFlag = 1
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "P"
            'Packet 전송
            iOrderFlag = iOrderFlag + 1
            
            msComm.Output = sSndPacket(iOrderFlag)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(sSndPacket(iOrderFlag))
            End If
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "L"
            'EOT 전송
            msComm.Output = Chr(4)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If
            
            '초기화
            iOrderFlag = 0: iPendingFlag = 0: iIdleFlag = 0: iTotQueryFlag = 0
            
        Case Else
    End Select
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_ADVIA1800_Batch()
    On Error GoTo ErrHandler
    
    '환자의 Order 전송
    Dim SendBuff As String
    Dim i%, j%, k%, iOrdCnt%
    Dim vIFCnt, vTmp
    Dim sTmp$, sTestCd$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$, sFrameNo$
    
    SendBuff = ""
    sTmp = ""
    
    Select Case sSndState
        Case ""
            '----- 검사항목 조회
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, "")
            
            Call Get_OrderString
        
            If pSampleInfo.ORDCNT = 0 Then      'Order 없는 경우
                iIdleFlag = CStr(Val(iIdleFlag) + 1)
            
                sFrameNo = CStr(Val(iIdleFlag) Mod 8)
                        
                SendBuff = sFrameNo & "O" & " " & "0101"
                SendBuff = SendBuff & "000"     'Format$(gOrderTable.iOrdCnt, "000")
                SendBuff = SendBuff & "N" & "2"
                SendBuff = SendBuff & Left$(pSampleInfo.ID & Space(13), 13)
                SendBuff = SendBuff & Left$(Format(pSampleInfo.RACK, "00") & "-" & Format(pSampleInfo.POS, "00") & Space(7), 7) 'Tray-Pos
                SendBuff = SendBuff & Space$(16) & Space$(16) & "M" & Space$(3)
                SendBuff = SendBuff & Space$(8) & " 1.0" & "1" & "1"
                SendBuff = SendBuff & Space$(1) & Chr(3)
                                        
                'n개의 sSndPacket 구성
                ReDim Preserve sSndPacket(Val(iIdleFlag))
                sSndPacket(Val(iIdleFlag)) = Chr(2) & SendBuff & ChkSum_ASTM(SendBuff) & vbCr & vbLf
                                        
                pSampleInfo.ORDCNT = 0
                
                '조회된 내용이 없는 경우 환자정보 구조체 초기화
                Call Init_pResultInfo
        
                RaiseEvent SendOrderOK("", "", "", "")
                
                Exit Sub
            End If
            
            sTestCd = ""
            
            For i = 1 To pSampleInfo.ORDCNT
                sTestCd = sTestCd & Right(Space(3) & Trim$(pSampleInfo.IFCD(i)), 3) & "M"
            Next i
            
            iIdleFlag = CStr(Val(iIdleFlag) + 1)
        
            sFrameNo = CStr(Val(iIdleFlag) Mod 8)
                    
            SendBuff = sFrameNo & "O" & " " & "0101"
            SendBuff = SendBuff & Format$(pSampleInfo.ORDCNT, "000")
            SendBuff = SendBuff & "N"                                   'Sample classification
            SendBuff = SendBuff & "1"                                   'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
            SendBuff = SendBuff & Left$(pSampleInfo.ID & Space(13), 13)
            SendBuff = SendBuff & Left$(Format(pSampleInfo.RACK, "00") & "-" & Format(pSampleInfo.POS, "00") & Space(7), 7) 'Tray-Pos
            SendBuff = SendBuff & Space$(16) & Space$(16) & "M" & Space$(3)
            
'            SendBuff = SendBuff & Space$(8) & " 1.0" & "1" & "1"
            SendBuff = SendBuff & Space$(8)
            SendBuff = SendBuff & " 1.0"        'Dilution coefficient(4)
            
            If pSampleInfo.SPCCD = "2" Then
                SendBuff = SendBuff & "2"           'Sample classification(1:blood serum, 2:urine)
            Else
                SendBuff = SendBuff & "1"           'Sample classification(1:blood serum, 2:urine)
            End If
            
            SendBuff = SendBuff & "1"           'Container classification
            
            SendBuff = SendBuff & sTestCd & Space$(1) & Chr(3)
            
            'n개의 sSndPacket 구성
            ReDim Preserve sSndPacket(Val(iIdleFlag))
            sSndPacket(Val(iIdleFlag)) = Chr(2) & SendBuff & ChkSum_ASTM(SendBuff) & vbCr & vbLf
                         
            RaiseEvent SendOrderOK(pSampleInfo.ID, "", "", "")
                        
        Case "E"
            '처음 Packet 전송
            msComm.Output = sSndPacket(1)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(sSndPacket(1))
            End If
            
            iOrderFlag = 1
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "P"
            'Packet 전송
            iOrderFlag = iOrderFlag + 1
            
            msComm.Output = sSndPacket(iOrderFlag)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(sSndPacket(iOrderFlag))
            End If
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "L"
            'EOT 전송
            msComm.Output = Chr(4)
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If
            
            '초기화
            iOrderFlag = 0: iPendingFlag = 0: iIdleFlag = 0: iTotQueryFlag = 0
            
        Case Else
    End Select
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub Get_OrderString()
        
    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .SEQNO = m_p_sSeq
            .RACK = m_p_sRack
            .POS = m_p_sPos
            .ORDCNT = 0
            .SINDEX = False
            .SPCCD = ""
            .CMT1 = ""
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
        .SINDEX = m_p_bSIndex
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
        
        .SPCCD = m_p_sSpcCd     '검체구분(2: Urine)
        .CMT1 = m_p_sCmt1
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

Private Sub PhaseCfg_Protocol_ADVIA1650()

    Dim wkDat As String
    Dim ix1 As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5          'ENQ
                        RcvBuffer = ""
                        sRcvState = "": sSndState = ""
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                    Case 6
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 4          'EOT
                        Select Case sRcvState
                            Case "Q"
                                m_iPhase = 3
                                iTotQueryFlag = iPendingFlag
                                iPendingFlag = 0
                                
                                'Order전송 Start
                                msComm.Output = Chr(5)
                                sSndState = "E"
                                
                            Case "R"
                                m_iPhase = 1
                        End Select
                        
                        sRcvState = ""
                        
                    Case 5          'ENQ
                        msComm.Output = Chr(6)
                        RcvBuffer = ""
                        
                    Case 10         'LF
                        m_iPhase = 2
                        
                        Call DataEditResponse_ADVIA1650
                        
                        'TimeOut 방지 위해서 장비쪽 설정을 길게
                        msComm.Output = Chr(6)   'ACK
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                End Select
            
            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        
                        Select Case sSndState
                            Case "E"        '<ENQ> 전송 후의 상태
                                Call SendOrder_ADVIA1650
                        
                            Case "P"        '<Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1650
                                                
                            Case "L"        '마지막 <Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1650
                           
                                'Order관련 초기화
                                sSndState = ""
                                Erase sSndPacket        '2006/11/3 yk...
                                m_iPhase = 1
                        End Select
                        
                    Case 5      'ENQ
                        RcvBuffer = ""
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                        
                    Case 21     'NAK
                        Select Case sSndState
                            Case "E"
                                msComm.Output = Chr(5)
                                m_iPhase = 3
                            Case "P"
                                msComm.Output = sSndPacket(iOrderFlag)
                                m_iPhase = 3
                            Case "L"
                                msComm.Output = sSndPacket(iOrderFlag)
                                m_iPhase = 3
                        End Select
                        
                    Case 4      'EOT
                        RcvBuffer = ""
                        m_iPhase = 1
                        sRcvState = "": sSndState = ""
                        'Order관련 초기화
                        iPendingFlag = 0: iTotQueryFlag = 0
                        
                    Case Else
                End Select
        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_ADVIA1800()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5          'ENQ
                        RcvBuffer = ""
                        m_iPhase = 2
                     
                        msComm.Output = Chr(6)
                        
                    Case Else
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                
                    Case 4          'EOT
                        Call DataEditResponse_ADVIA1800("EOT")
                        RcvBuffer = ""
                        
                        Select Case sRcvState
                            Case "Q"
                                m_iPhase = 3
                                iTotQueryFlag = iPendingFlag
                                iPendingFlag = 0
                                
                                'Order전송 Start
                                msComm.Output = Chr(5)
                                sSndState = "E"
                                
                            Case "R"
                                m_iPhase = 1
                        End Select
                        
                        sRcvState = ""
                        
                    Case 5          'ENQ
                        msComm.Output = Chr(6)
                        RcvBuffer = ""
                        
                    Case 10         'LF
                        If RcvBuffer <> "" Then
                            Call DataEditResponse_ADVIA1800
                            RcvBuffer = ""
                        End If
                        
                        'TimeOut 방지 위해서 장비쪽 설정을 길게
                        msComm.Output = Chr(6)   'ACK
                        
                    Case 13     'CR
                        
                    Case 23    'ETB
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        
                End Select
            
            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        
                        Select Case sSndState
                            Case "E"        '<ENQ> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                        
                            Case "P"        '<Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                                                
                            Case "L"        '마지막 <Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1800
                           
                                'Order관련 초기화
                                sSndState = ""
                                Erase sSndPacket        '2006/11/3 yk...
                                m_iPhase = 1
                        End Select
                        
                    Case 5      'ENQ
                        RcvBuffer = ""
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                        
                    Case 21     'NAK
                        Select Case sSndState
                            Case "E"
                                msComm.Output = Chr(5)
                                m_iPhase = 3
                            Case "P"
                                msComm.Output = sSndPacket(iOrderFlag)
                                m_iPhase = 3
                            Case "L"
                                msComm.Output = sSndPacket(iOrderFlag)
                                m_iPhase = 3
                        End Select
                        
                    Case 4      'EOT
                        RcvBuffer = ""
                        m_iPhase = 1
                        sRcvState = "": sSndState = ""
                        'Order관련 초기화
                        iPendingFlag = 0: iTotQueryFlag = 0
                        
                    Case Else
                End Select
        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_ADVIA1800_Batch()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5          'ENQ
                        RcvBuffer = ""
                        m_iPhase = 2
                     
                        msComm.Output = Chr(6)
                        
                    Case Else
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                
                    Case 4          'EOT
                        Call DataEditResponse_ADVIA1800_Batch("EOT")
                        RcvBuffer = ""
                        
                        Select Case sRcvState
                            Case "Q"
                                m_iPhase = 3
                                iTotQueryFlag = iPendingFlag
                                iPendingFlag = 0
                                
                                'Order전송 Start
                                msComm.Output = Chr(5)
                                sSndState = "E"
                                
                            Case "R"
                                m_iPhase = 1
                        End Select
                        
                        sRcvState = ""
                        
                    Case 5          'ENQ
                        msComm.Output = Chr(6)
                        RcvBuffer = ""
                        
                    Case 10         'LF
                        If RcvBuffer <> "" Then
                            Call DataEditResponse_ADVIA1800_Batch
                            RcvBuffer = ""
                        End If
                        
                        'TimeOut 방지 위해서 장비쪽 설정을 길게
                        msComm.Output = Chr(6)   'ACK
                        
                    Case 13     'CR
                        
                    Case 23    'ETB
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        
                End Select
            
            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        
                        Select Case sSndState
                            Case "E"        '<ENQ> 전송 후의 상태
                                Call SendOrder_ADVIA1800_Batch
                        
                            Case "P"        '<Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1800_Batch
                                                
                            Case "L"        '마지막 <Packet> 전송 후의 상태
                                Call SendOrder_ADVIA1800_Batch
                           
                                'Order관련 초기화
                                sSndState = ""
                                Erase sSndPacket        '2006/11/3 yk...
                                m_iPhase = 1
                        End Select
                        
                    Case 5      'ENQ
                        RcvBuffer = ""
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                        
                    Case 21     'NAK
                        Select Case sSndState
                            Case "E"
                                msComm.Output = Chr(5)
                                m_iPhase = 3
                            Case "P"
                                msComm.Output = sSndPacket(iOrderFlag)
                                m_iPhase = 3
                            Case "L"
                                msComm.Output = sSndPacket(iOrderFlag)
                                m_iPhase = 3
                        End Select
                        
                    Case 4      'EOT
                        RcvBuffer = ""
                        m_iPhase = 1
                        sRcvState = "": sSndState = ""
                        'Order관련 초기화
                        iPendingFlag = 0: iTotQueryFlag = 0
                        
                    Case Else
                End Select
        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_ADVIA1800_Old()
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
'                    Case 5          'ENQ
'                        RcvBuffer = ""
'                        sRcvState = "": sSndState = ""
'                        m_iPhase = 2
'
'                        bEndChk = True: bSTXChk = False
'
'                        msComm.Output = Chr(6)
'
'                    Case Else
'                        m_iPhase = 1
'                End Select
'
'            Case 2
'                Select Case Asc(wkDat)
'                    Case 2      'STX
'                        If bEndChk = True Then
'                            RcvBuffer = ""
'                        Else
'                            bSTXChk = True
'                        End If
'
'                        bEndChk = True
'
'                    Case 4          'EOT
'                        Select Case sRcvState
'                            Case "Q"
'                                m_iPhase = 3
'                                iTotQueryFlag = iPendingFlag
'                                iPendingFlag = 0
'
'                                'Order전송 Start
'                                msComm.Output = Chr(5)
'                                sSndState = "E"
'
'                            Case "R"
'                                m_iPhase = 1
'                        End Select
'
'                        sRcvState = ""
'
'                    Case 5          'ENQ
''                        bEndChk = True: bSTXChk = True
'
'                        msComm.Output = Chr(6)
'                        RcvBuffer = ""
'
'                    Case 10         'LF
'                        m_iPhase = 2
'
''                        'TimeOut 방지 위해서 장비쪽 설정을 길게
''                        'Data 편집 전에 전송...2013/4/17 yk
''                        msComm.Output = Chr(6)   'ACK
'
'                        If bEndChk = True And RcvBuffer <> "" Then
'                            Call DataEditResponse_ADVIA1800
'                            RcvBuffer = ""
'                        End If
'
'                        'TimeOut 방지 위해서 장비쪽 설정을 길게
'                        msComm.Output = Chr(6)   'ACK
'
'                    Case 13     'CR
'                        If bEndChk = True Then
'                            Call DataEditResponse_ADVIA1800
'                            RcvBuffer = ""
'                        End If
'
'                    Case 23    'ETB
''                        bEndChk = False
'
'                        If sRcvState = "Q" Then     '2013/4/17 yk
'                        Else
'                            msComm.Output = Chr(6)
'                        End If
''
''                        If bEndChk = True Then
''                            Call DataEditResponse_ADVIA1800
''                            RcvBuffer = ""
''                        End If
'
'                    Case Else
'                        If bEndChk = True Then
'                            If bSTXChk = True Then
'                                bSTXChk = False
'                            Else
'                                RcvBuffer = RcvBuffer & wkDat
'                            End If
'                        End If
'
'                        m_iPhase = 2
'                End Select
'
'            Case 3
'                Select Case Asc(wkDat)
'                    Case 6      'ACK
'
'                        Select Case sSndState
'                            Case "E"        '<ENQ> 전송 후의 상태
'                                Call SendOrder_ADVIA1800
'
'                            Case "P"        '<Packet> 전송 후의 상태
'                                Call SendOrder_ADVIA1800
'
'                            Case "L"        '마지막 <Packet> 전송 후의 상태
'                                Call SendOrder_ADVIA1800
'
'                                'Order관련 초기화
'                                sSndState = ""
'                                Erase sSndPacket        '2006/11/3 yk...
'                                m_iPhase = 1
'                        End Select
'
'                    Case 5      'ENQ
'                        bEndChk = True: bSTXChk = False
'                        RcvBuffer = ""
'                        msComm.Output = Chr(6)
'                        m_iPhase = 2
'
'                    Case 21     'NAK
'                        Select Case sSndState
'                            Case "E"
'                                msComm.Output = Chr(5)
'                                m_iPhase = 3
'                            Case "P"
'                                msComm.Output = sSndPacket(iOrderFlag)
'                                m_iPhase = 3
'                            Case "L"
'                                msComm.Output = sSndPacket(iOrderFlag)
'                                m_iPhase = 3
'                        End Select
'
'                    Case 4      'EOT
'                        RcvBuffer = ""
'                        m_iPhase = 1
'                        sRcvState = "": sSndState = ""
'                        'Order관련 초기화
'                        iPendingFlag = 0: iTotQueryFlag = 0
'
'                    Case Else
'                End Select
'        End Select
'    Next ix1
    
End Sub
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
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintRcvLog(wkBuf)
            End If
                                
            If iSpaceCnt = 30 Then
                iSpaceCnt = 0
            End If
            iSpaceCnt = iSpaceCnt + 2
            
            RaiseEvent DispMsgComm(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
            
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
    m_p_bSIndex = PropBag.ReadProperty("p_bSIndex", m_def_p_bSIndex)
    m_p_sRerunGbn = PropBag.ReadProperty("p_sRerunGbn", m_def_p_sRerunGbn)
    m_p_sTSVol = PropBag.ReadProperty("p_sTSVol", m_def_p_sTSVol)
    m_p_sSpcCd = PropBag.ReadProperty("p_sSpcCd", m_def_p_sSpcCd)
    m_p_sCmt1 = PropBag.ReadProperty("p_sCmt1", m_def_p_sCmt1)
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
    Call PropBag.WriteProperty("p_bSIndex", m_p_bSIndex, m_def_p_bSIndex)
    Call PropBag.WriteProperty("p_sRerunGbn", m_p_sRerunGbn, m_def_p_sRerunGbn)
    Call PropBag.WriteProperty("p_sTSVol", m_p_sTSVol, m_def_p_sTSVol)
    Call PropBag.WriteProperty("p_sSpcCd", m_p_sSpcCd, m_def_p_sSpcCd)
    Call PropBag.WriteProperty("p_sCmt1", m_p_sCmt1, m_def_p_sCmt1)
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
    
    '변수 초기화(E-170/H-7600)
    RstEnd = "Y": bEndChk = True: bSTXChk = False
    
    '변수 초기화(Advia1650)
    iPendingFlag = 0: iTotQueryFlag = 0: iTmpPendingFlag = 0: iIdleFlag = 0
    iOrderFlag = 0: iResultFlag = 0
    sRcvState = "": sSndState = ""
    
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
'    m_iStartSampleNo = m_def_iStartSampleNo
    m_p_bSIndex = m_def_p_bSIndex
    m_p_sRerunGbn = m_def_p_sRerunGbn
    m_p_sTSVol = m_def_p_sTSVol
    m_p_sSpcCd = m_def_p_sSpcCd
    m_p_sCmt1 = m_def_p_sCmt1
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
'MemberInfo=0,0,0,0
Public Property Get p_bSIndex() As Boolean
    p_bSIndex = m_p_bSIndex
End Property

Public Property Let p_bSIndex(ByVal New_p_bSIndex As Boolean)
    m_p_bSIndex = New_p_bSIndex
    PropertyChanged "p_bSIndex"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sRerunGbn() As String
    p_sRerunGbn = m_p_sRerunGbn
End Property

Public Property Let p_sRerunGbn(ByVal New_p_sRerunGbn As String)
    m_p_sRerunGbn = New_p_sRerunGbn
    PropertyChanged "p_sRerunGbn"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sTSVol() As String
    p_sTSVol = m_p_sTSVol
End Property

Public Property Let p_sTSVol(ByVal New_p_sTSVol As String)
    m_p_sTSVol = New_p_sTSVol
    PropertyChanged "p_sTSVol"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get p_sSpcCd() As Variant
    p_sSpcCd = m_p_sSpcCd
End Property

Public Property Let p_sSpcCd(ByVal New_p_sSpcCd As Variant)
    m_p_sSpcCd = New_p_sSpcCd
    PropertyChanged "p_sSpcCd"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get p_sCmt1() As String
    p_sCmt1 = m_p_sCmt1
End Property

Public Property Let p_sCmt1(ByVal New_p_sCmt1 As String)
    m_p_sCmt1 = New_p_sCmt1
    PropertyChanged "p_sCmt1"
End Property

