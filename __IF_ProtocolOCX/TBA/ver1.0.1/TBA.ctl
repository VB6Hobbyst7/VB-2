VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl TBA 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1065
      Top             =   2415
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
   Begin MSCommLib.MSComm msComm 
      Left            =   255
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "TBA"
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
Event EnableExit()
Event SendOrderOK(sID$)
Event RaiseError(sError$)
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
Dim iSpaceCnt   As Integer

'FOR TBA
Dim sSndState   As String

Private Sub DataEditResponse_TBA200FR()
    On Error GoTo ErrRtn
    
    Dim sType   As String
    Dim sBufCnt As String
    Dim ii      As Integer
    Dim iAllCnt As Integer
    Dim tmpIFCd$, tmpRst$
    
    sType = Mid$(RcvBuffer, 1, 1)
    
    Select Case sType
        Case Chr(6)         'ACK
            Select Case sSndState
                Case "I"
                    Call TransferToken
                    sSndState = "S"
                
                Case "Z"
                    Call SendOrder_TBA200FR
            
            End Select
            
        Case Chr(21)        'NAK
            Select Case sSndState
                Case "I"
                    Call Send_Initial
                Case "S"
                    Call TransferToken
                Case "Y"
                    Call Send_Initial
            End Select
            
        Case "M"
            Call Sleep(20)
            msComm.Output = Chr(2) & Chr(6) & Chr(3)
            
            If sSndState = "Y" Then
                Call SendOrder_TBA200FR
            
                sSndState = "Z"
            Else
                '종료버튼 활성화
                RaiseEvent EnableExit
                
                Call TransferToken
                sSndState = "S"
            End If
            
        Case "R"
            '결과 구조체 초기화
            Call Init_pResultInfo
            
            'Packet 편집
            With pResultInfo
                .SEQNO = Trim(Mid(RcvBuffer, 3, 4))
                .ID = Trim(Mid(RcvBuffer, 7, 20))
                .RACK = Trim(Mid(RcvBuffer, 27, 4))
                .POS = Trim(Mid(RcvBuffer, 31, 2))
            End With
            
            sBufCnt = NoTrimGetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(23))
            sBufCnt = Mid(sBufCnt, 48)
            
            iAllCnt = Len(sBufCnt) / 13

            For ii = 1 To iAllCnt
                tmpIFCd = Trim(Mid(sBufCnt, (ii - 1) * 13 + 1, 4))
                tmpRst = Trim(Mid(sBufCnt, (ii - 1) * 13 + 5, 6))
                
                '결과값 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            Next ii
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

            Call Init_pResultInfo
            
            Call Sleep(20)
            
            msComm.Output = Chr(2) & Chr(6) & Chr(3)
        
        Case Else
        
    End Select
   
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Function NoTrimGetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim POS%

    POS = InStr(tStr, sUserSymbol)

    If POS = 0 Then
    Else
        NoTrimGetByOneUserSymbol = Mid$(tStr, 1, POS - 1)
        sOriginal = Mid$(sOriginal, POS + 1, Len(sOriginal) - POS)
    End If
End Function
Private Sub PhaseCfg_Protocol_TBA200FR()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
              
        Select Case Asc(wkDat)
            Case 2      '----- STX 수신
                m_iPhase = 2
                RcvBuffer = ""
                
            Case 3      '----- ETX 수신
                Call DataEditResponse_TBA200FR
                
            Case Else   '----- 문자 수신
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
End Sub


Private Sub SendOrder_TBA200FR()
    On Error GoTo ErrRtn
    
    Dim sSend   As String
    Dim sTestCd As String
    Dim ii      As Integer
    
    RaiseEvent RequestCurOrder("", "", "")
    
    Call Get_OrderString
    
    If pSampleInfo.ID = "" Or pSampleInfo.ORDCNT = 0 Then
        Call TransferToken
        sSndState = "S"
        Exit Sub
    End If
    
    'Order Send
    sSend = Chr(2) & "O " & pSampleInfo.ID & String(20 - Len(pSampleInfo.ID), " ")
    sSend = sSend & String(4 - Len(pSampleInfo.RACK), " ") & pSampleInfo.RACK
    sSend = sSend & String(2 - Len(pSampleInfo.POS), " ") & pSampleInfo.POS
    sSend = sSend & "  1"

    sTestCd = ""
    For ii = 1 To pSampleInfo.ORDCNT
        If Trim(pSampleInfo.IFCD(ii)) <> "" Then
            sTestCd = sTestCd & String(4 - Len(pSampleInfo.IFCD(ii)), " ") & pSampleInfo.IFCD(ii) & "1"
        End If
    Next ii

    sSend = sSend & sTestCd & Chr(23)
    sSend = sSend & Format(Now, "YYYYMMDDHHMM")
    sSend = sSend & Space$(30)      'NAME
    sSend = sSend & Space$(1)       'SEX
    sSend = sSend & Space(8)        'BIRTHDAY
    sSend = sSend & Space(20)       'LOCATION
    sSend = sSend & Space(20)       'DOCTOR
    sSend = sSend & Space(20)       'COMMENT
    sSend = sSend & Chr(23) & Chr(3)

    msComm.Output = sSend
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    
    'Order 전송 OK
    RaiseEvent SendOrderOK(pSampleInfo.ID)
    
ErrRtn:
   If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
   End If
End Sub

Private Sub Edit_Data_TBA200FR()
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
'    Dim vLabDate, vSlipCd, vLabSeq, vRstCnt, vBuf
'    Dim i%, j%, k%, iCRow%
'    Dim iAllCnt%, iRstCnt%, iCmtCnt%
'    Dim sIFCd$, sRst$, sRetVal$
'
'    Dim sTotTestCd  As String
'    Dim sTotTestNm  As String
'    Dim sTotRst     As String
'    Dim sTotCom     As String
'
'    Dim sMid$, sBuf$, sBufCnt$, sDiskID$, sDiskPos$
'
'    Dim iMatchFlag As Integer
'
'    sMid = Mid$(RcvBuffer, 1, 1)
'
'    Select Case sMid
'
'        Case Chr(6)
'
'            If sSndState = "I" Then
'                Call TransferToken
'                sSndState = "S"
'            ElseIf sSndState = "Z" Then
'
'                If GetNowOrderList = "NONE" Then
'                    Call TransferToken
'                    sSndState = "S"
'                Else
'
'                    'Order Send
'                    sTX = ""
'
'                    sTX = Chr(2) & "O " & gOrderTable.sSampID & String(20 - Len(gOrderTable.sSampID), " ")
'                    sTX = sTX & String(4 - Len(gOrderTable.sRack), " ") & gOrderTable.sRack
'                    sTX = sTX & String(2 - Len(gOrderTable.sPos), " ") & gOrderTable.sPos
'                    sTX = sTX & "  1"
'
'                    For i = 1 To gOrderTable.iOrdCnt
'                        If Trim(gOrderTable.sIFTestCd(i)) = "" Then
'                        Else
'                            sTX = sTX & String(4 - Len(gOrderTable.sIFTestCd(i)), " ") & gOrderTable.sIFTestCd(i) & "1"
'                        End If
'                    Next
'
'                    sTX = sTX & Chr(23) & gOrderTable.sLabDate
'                    sTX = sTX & Space$(30) & Space$(1) & Chr(23) & Chr(3)
'
'                    Comm1.Output = sTX
'
''                    'Log File에 쓰기
''                    If giTestMode = 77 Then
''                        Print #302, sTX;
''                    End If
'
'                    'Order 전송 OK 이므로 Order 성공을 화면에 표시
'                    Call spdIntList.SetText(1, gOrderTable.iCRow, "0")
'                    Call spdIntList.SetText(9, gOrderTable.iCRow, CStr(Val(gOrderTable.iOrdCnt)) & "")
'                    Call SpdForeBack(spdIntList, 5, 10, gOrderTable.iCRow, gOrderTable.iCRow, RGB(0, 0, 0), 연노랑)
'                End If
'
'            End If
'
'        Case Chr(21)
'
'            If sSndState = "I" Then
'                Call cmdInitial_Click
'            ElseIf sSndState = "S" Then
'                Call TransferToken
'            ElseIf sSndState = "Y" Then
''                Comm1.Output = sTX
'                Call cmdInitial_Click
'            End If
'
'        Case "M"
'
'            Call Sleep(20)
'            Comm1.Output = Chr(2) & Chr(6) & Chr(3)
'
'            If sSndState = "Y" Then
'                If GetNowOrderList = "NONE" Then
'                    Call TransferToken
'                    sSndState = "S"
'                Else
'
'                    'Order Send
'                    sTX = ""
'
'                    sTX = Chr(2) & "O " & gOrderTable.sSampID & String(20 - Len(gOrderTable.sSampID), " ")
'                    sTX = sTX & String(4 - Len(gOrderTable.sRack), " ") & gOrderTable.sRack
'                    sTX = sTX & String(2 - Len(gOrderTable.sPos), " ") & gOrderTable.sPos
'                    sTX = sTX & "  1"
'
'                    For i = 1 To gOrderTable.iOrdCnt
'                        If Trim(gOrderTable.sIFTestCd(i)) = "" Then
'                        Else
'                            sTX = sTX & String(4 - Len(gOrderTable.sIFTestCd(i)), " ") & gOrderTable.sIFTestCd(i) & "1"
'                        End If
'                    Next
'
'                    sTX = sTX & Chr(23) & gOrderTable.sLabDate
'                    sTX = sTX & Space$(30) & Space$(1) & Chr(23) & Chr(3)
'
'                    Comm1.Output = sTX
'
''                    'Log File에 쓰기
''                    If giTestMode = 77 Then
''                        Print #302, sTX;
''                    End If
'
'                    'Order 전송 OK 이므로 Order 성공을 화면에 표시
'                    Call spdIntList.SetText(1, gOrderTable.iCRow, "0")
'                    Call spdIntList.SetText(9, gOrderTable.iCRow, CStr(Val(gOrderTable.iOrdCnt)) & "")
'                    Call SpdForeBack(spdIntList, 5, 10, gOrderTable.iCRow, gOrderTable.iCRow, RGB(0, 0, 0), 연노랑)
'                End If
'
'                sSndState = "Z"
'            Else
'                cmdExit.Enabled = True
'                Call TransferToken
'                sSndState = "S"
'            End If
'
'        Case "R"
'
'            Dim sTP     As String
'            Dim sALB    As String
'            Dim sTBIL   As String
'            Dim sDBIL   As String
'            Dim sBUN    As String
'            Dim sCRE    As String
'            Dim sFE     As String
'            Dim sUIBC   As String
'
'            sTP = "":  sALB = "": sTBIL = "": sDBIL = ""
'            sBUN = "": sCRE = "": sFE = "":  sUIBC = ""
'            iRstCnt = 0: iAllCnt = 0
'
'            'Packet 편집
'            sSampNo = Trim(Mid(RcvBuffer, 7, 20))
'            sDiskID = Trim(Mid(RcvBuffer, 27, 4))
'            sDiskPos = Trim(Mid(RcvBuffer, 31, 2))
'
'            sLabDate = Mid(sSampNo, 1, 8)
'            sSlipCd = Mid(sSampNo, 9, 3)
'            sLabSeq = Mid(sSampNo, 12, 5)
'
'            '현재의 전송과 매칭되는 Row 찾기
'            iCRow = FindCurRow(sLabDate, sSlipCd, sLabSeq)
'
'            Call spdIntList.GetText(16, iCRow, vRstCnt)
'
''            Call NoTrimGetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(10))
'
'            sBufCnt = NoTrimGetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(23))
'            sBufCnt = Mid(sBufCnt, 48)
'
'            iAllCnt = Len(sBufCnt) / 13
'
'            For i = 1 To iAllCnt
'                sIFCd = Trim(Mid(sBufCnt, (i - 1) * 13 + 1, 4))
'                sRst = Trim(Mid(sBufCnt, (i - 1) * 13 + 5, 6))
'
'                For j = 1 To CInt(vRstCnt)
'                    If iMatchFlag = 1 Then
'                        iMatchFlag = 0
'                        Exit For
'                    End If
'                    Call spdIntList.GetText(16 + j, iCRow, vBuf)
'                    'vBuf 형태 TestCd/Result/
'                    sBuf = Trim(CStr(vBuf))
'                    sBuf = GetByOne(sBuf, sBuf)
'
'                    '전체 IFItem 중에서 현재의 TestCd에 해당하는 IFCd 구함
'                    For k = 1 To giIntItemCnt
'                        If sBuf = gIFItem(k).s02 & gIFItem(k).s03 & gIFItem(k).s04 & _
'                                    gIFItem(k).s05 & gIFItem(k).s06 Then
'                            If Trim(sIFCd) = gIFItem(k).s09 Then
'                                sTotTestCd = sTotTestCd & sBuf & Chr(124)
'                                sTotTestNm = sTotTestNm & gIFItem(k).s07 & Chr(124)
'                                sTotRst = sTotRst & sRst & Chr(124)
'                                iRstCnt = iRstCnt + 1
'                                iMatchFlag = 1
'                                Exit For
'                            End If
'                        End If
'                    Next k
'
'                Next j
'
''                '계산값을 위한 저장
''                If sIFCd = "1" Then
''                    sTP = sRst
''                ElseIf sIFCd = "2" Then
''                    sALB = sRst
''                ElseIf sIFCd = "10" Then
''                    sTBIL = sRst
''                ElseIf sIFCd = "9" Then
''                    sDBIL = sRst
''                ElseIf sIFCd = "13" Then
''                    sBUN = sRst
''                ElseIf sIFCd = "14" Then
''                    sCRE = sRst
''                ElseIf sIFCd = "22" Then
''                    sFE = sRst
''                ElseIf sIFCd = "23" Then
''                    sUIBC = sRst
''                End If
'
'            Next i
'
'            '계산값 적용
''            If IsNumeric(sTP) = True And IsNumeric(sALB) = True Then
''                iRstCnt = iRstCnt + 1
''                sTIFCd = sTIFCd & "1000" & Chr(124)
''                sTRst = sTRst & Trim(Str(Val(sTP) - Val(sALB))) & Chr(124)
''            End If
''
''            If IsNumeric(sALB) = True And IsNumeric(sTP) = True And IsNumeric(sALB) = True Then
''                iRstCnt = iRstCnt + 1
''                sTIFCd = sTIFCd & "1001" & Chr(124)
''                sTRst = sTRst & Trim(Format(Val(sALB) / (Val(sTP) - Val(sALB)), "0.000")) & Chr(124)
''            End If
''
''            If IsNumeric(sTBIL) = True And IsNumeric(sDBIL) = True Then
''                iRstCnt = iRstCnt + 1
''                sTIFCd = sTIFCd & "1002" & Chr(124)
''                sTRst = sTRst & Trim(Str(Val(sTBIL) - Val(sDBIL))) & Chr(124)
''            End If
''
''            If IsNumeric(sBUN) = True And IsNumeric(sCRE) = True Then
''                iRstCnt = iRstCnt + 1
''                sTIFCd = sTIFCd & "1003" & Chr(124)
''                sTRst = sTRst & Trim(Format(Val(sBUN) / Val(sCRE), "0.000")) & Chr(124)
''            End If
''
''            If IsNumeric(sFE) = True And IsNumeric(sUIBC) = True Then
''                iRstCnt = iRstCnt + 1
''                sTIFCd = sTIFCd & "1004" & Chr(124)
''                sTRst = sTRst & Trim(Str(Val(sFE) + Val(sUIBC))) & Chr(124)
''            End If
'
'
'            '결과등록
'            If iCRow > 0 Then
'                '현재 장비에서 전송된 작업번호 표시
'                    lblResult = sLabDate & "-" & sSlipCd & "-" & sLabSeq
'                    sRetVal = ViewResults(0, iRstCnt, sLabDate, sSlipCd, sLabSeq, sTotTestCd, sTotTestNm, sTotRst)
'
''                    If iCmtCnt = 0 Then
''                    Else
''                        If sRetVal = "NONE" Then
''                        Else
''                            Call ViewComments(gResultTable(1).iCRow, CStr(iCmtCnt) & Chr(124) & sTotCom)
''                        End If
''                    End If
'
'                    '--- DB 등록
'                    Call Append_Result(sLabDate, sSlipCd, sLabSeq, sRetVal)
'            End If
'
'            Call Sleep(20)
'
'            Comm1.Output = Chr(2) & Chr(6) & Chr(3)
'
'    End Select
'
'    Exit Sub
'ErrHandler:
'    ViewMsg "Edit_Data - " & Err.Description
End Sub


Private Sub TransferToken()

    If sSndState = "I" Then
        Timer1.Interval = 1000
        Timer1.Enabled = True
    Else
        Timer1.Interval = 10000
        Timer1.Enabled = True
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
        Case "TBA200FR"
            Call PhaseCfg_Protocol_TBA200FR
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub



Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    With pSampleInfo
        .ID = ""
        .ORDCNT = 0
    End With
    
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
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
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

Private Sub Timer1_Timer()
    On Error GoTo ErrRtn
    
    Dim sSendBuf$
   
    sSendBuf = Chr(2) & "M     " & Chr(3)
    
    msComm.Output = sSendBuf
    
    Timer1.Enabled = False
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Timer Error - " & Err.Description)
    End If
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
    
    Timer1.Enabled = False
    
    On Error GoTo ErrPortOpen
    If m_PortOpen = True Then
        msComm.PortOpen = True
    End If
    
    Call Send_Initial
    
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
    
    If m_iSendPhase = 1 Then
        sSndState = "Y"
    End If
    
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
'MemberInfo=14
Public Function Send_Initial() As Variant
    On Error GoTo ErrRtn
    
    Dim sSendBuf$
    
    sSndState = "I"

    sSendBuf = Chr(2) & "I " & Chr(3)
        
    Call Sleep(1000)
    
    msComm.Output = sSendBuf
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Initial 오류 - " & Err.Description)
    End If
End Function

