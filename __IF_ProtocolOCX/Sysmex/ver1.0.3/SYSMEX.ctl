VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl SYSMEX 
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
Attribute VB_Name = "SYSMEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_iTotalItemCnt = 0
Const m_def_iOrderFlag = 0
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
Dim m_iTotalItemCnt As Integer
Dim m_iOrderFlag As Integer
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
Event RequestCurOrder(sID$, sRack$, sPos$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event SendOrderOK(sID$)
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
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

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
        Case "SE9000"
            Call PhaseCfg_Protocol_SE9000
        
        Case "CA500"
            Call PhaseCfg_Protocol_CA500    '바코드 사용
        
        Case "CA1500"
            Call PhaseCfg_Protocol_CA1500   '바코드 사용
        
        Case "CA7000"
            Call PhaseCfg_Protocol_CA7000   '바코드 사용
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub DataEditResponse_CA7000()
    On Error GoTo ErrRtn
    
    Dim sBC     As String
    Dim sLC     As String
    
    Dim iTestStart  As Integer
    Dim tmpBuffer   As String
    Dim ii      As Integer
    Dim tmpIFCd As String
    Dim tmpRst  As String
    
    
    sBC = Mid$(RcvBuffer, 1, 2)
    sLC = Mid$(RcvBuffer, 3, 1)
    
    Select Case sBC
        Case "R2"
            pSampleInfo.RACK = Mid(RcvBuffer, 20, 6)
            pSampleInfo.POS = Mid(RcvBuffer, 26, 2)
            pSampleInfo.ID = Trim(Mid(RcvBuffer, 28, 15))
            
            'Order Request 요청 받은 후
            Call SendOrder_CA7000
            
            Exit Sub
            
        Case "D1"
            '결과정보 초기화
            Call Init_pResultInfo
            
            'SampleID
            With pResultInfo
                .ID = Trim(Mid(RcvBuffer, 28, 15))
                .RACK = Mid(RcvBuffer, 20, 6)
                .POS = Mid(RcvBuffer, 26, 2)
                
                If Trim(pResultInfo.ID) = "" Then
                    Exit Sub
                End If
            End With
            
            iTestStart = 59
             
            '--- 결과편집
            For ii = 1 To 100       '현재 장비 매뉴얼상엔 20항목임...
                tmpBuffer = Mid(RcvBuffer, iTestStart + 9 * (ii - 1), 1)
            
                If Asc(tmpBuffer) = 3 Then Exit For
                
                tmpIFCd = Mid(RcvBuffer, iTestStart + 9 * (ii - 1), 3)
                tmpRst = Mid(RcvBuffer, iTestStart + 9 * (ii - 1) + 3, 5)

                If tmpRst = Space(5) Then
                    tmpRst = "N"
                End If

                '결과값 누적
                If Trim(tmpIFCd) <> "" Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                End If
            Next ii
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With
                
        Case Else
        
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
    End If
End Sub
Private Sub DataEditResponse_CA1500()
    On Error GoTo ErrRtn
    
    Dim sBC     As String
    Dim sLC     As String
    
    Dim iTestStart  As Integer
    Dim tmpBuffer   As String
    Dim ii      As Integer
    Dim tmpIFCd As String
    Dim tmpRst  As String
    
    
    sBC = Mid$(RcvBuffer, 1, 2)
    sLC = Mid$(RcvBuffer, 3, 1)
    
    Select Case sBC
        Case "R2"
            pSampleInfo.RACK = Mid(RcvBuffer, 20, 6)
            pSampleInfo.POS = Mid(RcvBuffer, 26, 2)
            pSampleInfo.ID = Trim(Mid(RcvBuffer, 28, 15))
            
            'Order Request 요청 받은 후
            Call SendOrder_CA1500
            
            Exit Sub
            
        Case "D1"
            '결과정보 초기화
            Call Init_pResultInfo
            
            'SampleID
            With pResultInfo
                .ID = Trim(Mid(RcvBuffer, 28, 15))
                .RACK = Mid(RcvBuffer, 20, 6)
                .POS = Mid(RcvBuffer, 26, 2)
                
                If Trim(pResultInfo.ID) = "" Then
                    Exit Sub
                End If
            End With
            
            iTestStart = 59
             
            '--- 결과편집
            For ii = 1 To 100       '현재 장비 매뉴얼상엔 20항목임...
                tmpBuffer = Mid(RcvBuffer, iTestStart + 9 * (ii - 1), 1)
            
                If Asc(tmpBuffer) = 3 Then Exit For
                
                tmpIFCd = Mid(RcvBuffer, iTestStart + 9 * (ii - 1), 3)
                tmpRst = Mid(RcvBuffer, iTestStart + 9 * (ii - 1) + 3, 5)

                If tmpRst = Space(5) Then
                    tmpRst = "N"
                End If

                '결과값 누적
                If Trim(tmpIFCd) <> "" Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                End If
            Next ii
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With
                
        Case Else
        
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
    End If
End Sub

Private Sub DataEditResponse_CA500()
    On Error GoTo ErrRtn
    
    Dim sBC     As String
    Dim sLC     As String
    
    Dim iTestStart  As Integer
    Dim tmpBuffer   As String
    Dim ii      As Integer
    Dim tmpIFCd As String
    Dim tmpRst  As String
    
    
    sBC = Mid$(RcvBuffer, 1, 2)
    sLC = Mid$(RcvBuffer, 3, 1)
    
    Select Case sBC
        Case "R2"
            pSampleInfo.RACK = Mid(RcvBuffer, 20, 4)
            pSampleInfo.POS = Mid(RcvBuffer, 24, 2)
            pSampleInfo.ID = Trim(Mid(RcvBuffer, 26, 15))
            
            'Order Request 요청 받은 후
            Call SendOrder_CA500
            
            Exit Sub
            
        Case "D1"
            '결과정보 초기화
            Call Init_pResultInfo
            
            'SampleID
            With pResultInfo
                .ID = Trim(Mid(RcvBuffer, 26, 15))
                .RACK = Mid(RcvBuffer, 20, 4)
                .POS = Mid(RcvBuffer, 24, 2)
                
                If Trim(pResultInfo.ID) = "" Then
                    Exit Sub
                End If
            End With
            
            iTestStart = 53
             
            '--- 결과편집
            For ii = 1 To 17        '현재 장비 매뉴얼상엔 17항목임...
                tmpBuffer = Mid(RcvBuffer, iTestStart + 9 * (ii - 1), 1)
            
                If Asc(tmpBuffer) = 3 Then Exit For
                
                tmpIFCd = Mid(RcvBuffer, iTestStart + 9 * (ii - 1), 3)
                tmpRst = Mid(RcvBuffer, iTestStart + 9 * (ii - 1) + 3, 5)

                If tmpRst = Space(5) Then
                    tmpRst = "N"
                End If

                '결과값 누적
                If Trim(tmpIFCd) <> "" Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                End If
            Next ii
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With
                
        Case Else
        
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
    End If
End Sub

Private Sub PhaseCfg_Protocol_CA7000()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)
             
        Select Case Asc(wkDat)
            Case 2      '----- STX 수신
                RcvBuffer = ""
                
            Case 3      '----- ETX 수신 (ETX 도 문자열에 포함해야함)
                RcvBuffer = RcvBuffer & wkDat
                
                Call Sleep(200)     '0.2 sec or More Delay
                msComm.Output = Chr(6)
                
                Call DataEditResponse_CA7000
                
            Case 6      '----- ACK 수신
                'Order 전송 완료
                RaiseEvent SendOrderOK(pSampleInfo.ID)
                
            Case 21     '----- NCK 수신
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
End Sub

Private Sub SendOrder_CA7000()
    On Error GoTo ErrRtn

    Dim SendBuf As String
    Dim ii%
    Dim sTestCd$

    SendBuf = "S"
    SendBuf = SendBuf & "2"
    SendBuf = SendBuf & "21"
    SendBuf = SendBuf & "01"
    SendBuf = SendBuf & "01"
    SendBuf = SendBuf & "U"
    SendBuf = SendBuf & Format$(Date, "YYMMDD")
    SendBuf = SendBuf & Format$(Now, "HHMM")
    SendBuf = SendBuf & pSampleInfo.RACK
    SendBuf = SendBuf & pSampleInfo.POS

    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    Call Get_OrderString
    
    '검사항목 편집
    sTestCd = ""
    With pSampleInfo
        For ii = 1 To pSampleInfo.ORDCNT
            If Trim(.IFCD(ii)) <> "" Then
                If Right(.IFCD(ii), 1) = "0" Then
                    If InStr(sTestCd, .IFCD(ii)) = 0 Then
                        sTestCd = sTestCd & .IFCD(ii) & Space(6)
                    End If
                Else
                    If InStr(sTestCd, Mid(.IFCD(ii), 1, Len(.IFCD(ii)) - 1) & "0") = 0 Then
                        sTestCd = sTestCd & Mid(.IFCD(ii), 1, Len(.IFCD(ii)) - 1) & "0" & Space(6)
                    End If
                End If
            End If
        Next ii
    End With
    
    If pSampleInfo.ORDCNT = 0 Then
        SendBuf = SendBuf & Space(15)
        SendBuf = SendBuf & "C"
        SendBuf = SendBuf & Space(15)
        SendBuf = SendBuf & ""
    Else
        SendBuf = SendBuf & Right(Space(15) & pSampleInfo.ID, 15)
        SendBuf = SendBuf & "B"
        SendBuf = SendBuf & Space(15)
        SendBuf = SendBuf & sTestCd
    End If

    Call Sleep(500)     '0.2 sec or More Delay
    
    msComm.Output = Chr(2) & SendBuf & Chr(3)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & SendBuf & Chr(3))
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_CA1500()
    On Error GoTo ErrRtn

    Dim SendBuf As String
    Dim ii%
    Dim sTestCd$

    SendBuf = "S"
    SendBuf = SendBuf & "2"
    SendBuf = SendBuf & "21"
    SendBuf = SendBuf & "01"
    SendBuf = SendBuf & "01"
    SendBuf = SendBuf & "U"
    SendBuf = SendBuf & Format$(Date, "YYMMDD")
    SendBuf = SendBuf & Format$(Now, "HHMM")
    SendBuf = SendBuf & pSampleInfo.RACK
    SendBuf = SendBuf & pSampleInfo.POS

    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    Call Get_OrderString
    
    '검사항목 편집
    sTestCd = ""
    With pSampleInfo
        For ii = 1 To pSampleInfo.ORDCNT
            If Trim(.IFCD(ii)) <> "" Then
                If Right(.IFCD(ii), 1) = "0" Then
                    If InStr(sTestCd, .IFCD(ii)) = 0 Then
                        sTestCd = sTestCd & .IFCD(ii) & Space(6)
                    End If
                Else
                    If InStr(sTestCd, Mid(.IFCD(ii), 1, Len(.IFCD(ii)) - 1) & "0") = 0 Then
                        sTestCd = sTestCd & Mid(.IFCD(ii), 1, Len(.IFCD(ii)) - 1) & "0" & Space(6)
                    End If
                End If
            End If
        Next ii
    End With
    
    If pSampleInfo.ORDCNT = 0 Then
        SendBuf = SendBuf & Space(15)
        SendBuf = SendBuf & "C"
        SendBuf = SendBuf & Space(15)
        SendBuf = SendBuf & ""
    Else
        SendBuf = SendBuf & Right(Space(15) & pSampleInfo.ID, 15)
        SendBuf = SendBuf & "B"
        SendBuf = SendBuf & Space(15)
        SendBuf = SendBuf & sTestCd
    End If

    Call Sleep(500)
    
    msComm.Output = Chr(2) & SendBuf & Chr(3)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & SendBuf & Chr(3))
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
    End If
End Sub


Private Sub SendOrder_CA500()
    On Error GoTo ErrRtn

    Dim SendBuf As String
    Dim ii%
    Dim sTestCd$

    SendBuf = "S"
    SendBuf = SendBuf & "2"
    SendBuf = SendBuf & "21"
    SendBuf = SendBuf & "01"
    SendBuf = SendBuf & "01"
    SendBuf = SendBuf & "U"
    SendBuf = SendBuf & Format$(Date, "YYMMDD")
    SendBuf = SendBuf & Format$(Now, "HHMM")
    SendBuf = SendBuf & pSampleInfo.RACK
    SendBuf = SendBuf & pSampleInfo.POS

    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    Call Get_OrderString
    
    '검사항목 편집
    sTestCd = ""
    With pSampleInfo
        For ii = 1 To pSampleInfo.ORDCNT
            If Trim(.IFCD(ii)) <> "" Then
                If Right(.IFCD(ii), 1) = "0" Then
                    If InStr(sTestCd, .IFCD(ii)) = 0 Then
                        sTestCd = sTestCd & .IFCD(ii) & Space(6)
                    End If
                Else
                    If InStr(sTestCd, Mid(.IFCD(ii), 1, Len(.IFCD(ii)) - 1) & "0") = 0 Then
                        sTestCd = sTestCd & Mid(.IFCD(ii), 1, Len(.IFCD(ii)) - 1) & "0" & Space(6)
                    End If
                End If
            End If
        Next ii
    End With
    
    If pSampleInfo.ORDCNT = 0 Then
        SendBuf = SendBuf & Space(15)
        SendBuf = SendBuf & "C"
        SendBuf = SendBuf & Space(11)
        SendBuf = SendBuf & ""
    Else
        SendBuf = SendBuf & Right(Space(15) & pSampleInfo.ID, 15)
        SendBuf = SendBuf & "B"
        SendBuf = SendBuf & Space(11)
        SendBuf = SendBuf & sTestCd
    End If

    Call Sleep(500)
    
    msComm.Output = Chr(2) & SendBuf & Chr(3)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & SendBuf & Chr(3))
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
    End If
End Sub


Private Sub PhaseCfg_Protocol_CA1500()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
                 
        Select Case Asc(wkDat)
            Case 2      '----- STX 수신
                RcvBuffer = ""
                        
            Case 3      '----- ETX 수신 (ETX 도 문자열에 포함해야함)
                RcvBuffer = RcvBuffer & wkDat
                msComm.Output = Chr(6)
                
                Call DataEditResponse_CA1500
                
            Case 6      '----- ACK 수신
                'Order 전송 완료
                RaiseEvent SendOrderOK(pSampleInfo.ID)
                
            Case 21     '----- NCK 수신
                
            Case Else   '----- 문자 수신
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_CA500()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
                 
        Select Case Asc(wkDat)
            Case 2      '----- STX 수신
                RcvBuffer = ""
                        
            Case 3      '----- ETX 수신 (ETX 도 문자열에 포함해야함)
                RcvBuffer = RcvBuffer & wkDat
                
                Call Sleep(200)     '0.2 sec or More Delay
                msComm.Output = Chr(6)
                
                Call DataEditResponse_CA500
                
            Case 6      '----- ACK 수신
                'Order 전송 완료
                RaiseEvent SendOrderOK(pSampleInfo.ID)
                
            Case 21     '----- NCK 수신
                
            Case Else   '----- 문자 수신
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
End Sub

'
'   SE-9000 바코드 사용(용인세브란스병원用)
'
Private Sub DataEdit_SE9000()
    On Error GoTo ErrRtn

    Dim sBC     As String
    Dim sLC     As String

    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack As String
    Dim tmpPos  As String
    Dim ii      As Integer
    Dim tmpRst()    As String       '결과 임시 저장

    Dim sChk$, sChk2$


    sBC = Mid(RcvBuffer, 1, 2)
    sLC = Mid(RcvBuffer, 3, 1)

    Select Case sBC
        Case "R1"
            pSampleInfo.ID = ""

'''            sChk = Mid$(RcvBuffer, 3, 1)
'''            sChk2 = Mid$(RcvBuffer, 5, 1)
''            sChk = Mid$(RcvBuffer, 5, 1)
''            sChk2 = Mid$(RcvBuffer, 7, 1)
''
''            If sChk <> "+" And sChk <> "-" And sChk2 <> "Q" Then
''                Exit Sub
''            End If

            tmpBarCd = Mid(RcvBuffer, 3, 13)
'            tmpBarCd = Mid$(tmpBarCd, 3)
            pSampleInfo.ID = Trim(tmpBarCd)

            Call SendOrder_SE9000
            m_iPhase = 2

            Exit Sub

        Case "D1"
            Select Case sLC
                Case "U"
                    '결과정보 초기화
                    Call Init_pResultInfo

                    tmpRack = Mid(RcvBuffer, 10, 4)
                    tmpPos = Mid(RcvBuffer, 14, 2)
                    tmpBarCd = Trim(Mid(RcvBuffer, 22, 13))

''                    'For 세브란스 SE-9000
'''                    sChk = Mid$(RcvBuffer, 22, 1)
'''                    sChk2 = Mid$(RcvBuffer, 24, 1)
''                    sChk = Mid$(RcvBuffer, 24, 1)
''                    sChk2 = Mid$(RcvBuffer, 26, 1)
''
''                    If sChk <> "+" And sChk <> "-" And sChk2 <> "Q" Then
''                    Else
'''                        tmpBarCd = Mid$(tmpBarCd, 3, 7) & "0" & Mid(tmpBarCd, 10)
''                        tmpBarCd = Mid$(tmpBarCd, 3, 7) & Mid(tmpBarCd, 10)
''                    End If
''                    '--------------------

                    ReDim tmpRst(25) As String

                    'WBC
                    tmpRst(1) = Mid(RcvBuffer, 63, 5)
                    If tmpRst(1) = Space(5) Then
                        tmpRst(1) = "N"
                    Else
                        tmpRst(1) = Format(Val(Format(tmpRst(1), "@@@.@@")), "0.00")
                    End If

                    'RBC
                    tmpRst(2) = Mid(RcvBuffer, 69, 4)

                    If tmpRst(2) = Space(4) Then
                        tmpRst(2) = "N"
                    Else
                        tmpRst(2) = Format(Val(Format$(tmpRst(2), "@@.@@")), "0.00")
                    End If

                    'HGB, HCT, MCV, MCH, MCHC
                    For ii = 3 To 7
                        tmpRst(ii) = Mid(RcvBuffer, 74 + (ii - 3) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    'PLT
                    tmpRst(8) = Mid(RcvBuffer, 99, 4)

                    If tmpRst(8) = Space(4) Then
                        tmpRst(8) = "N"
                    Else
                        tmpRst(8) = Trim(Val(Format(tmpRst(8), "@@@@")))
                    End If

                    'LYMPH%, MONO%, NEUT%, EO%, BASO%
                    For ii = 9 To 13
                        tmpRst(ii) = Mid(RcvBuffer, 104 + (ii - 9) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    'LYMPH#, MONO#, NEUT#, EO#, BASO#
                    For ii = 14 To 18
                        tmpRst(ii) = Mid(RcvBuffer, 129 + (ii - 14) * 6, 5)

                        If tmpRst(ii) = Space(5) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@@")), "0.00")
                        End If
                    Next ii

                    'RDW-CV(%), RDW-SD(fL), PDW(fL), MPV(fL), P-LCR
                    For ii = 19 To 23
                        tmpRst(ii) = Mid(RcvBuffer, 159 + (ii - 19) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = ""
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    '--- 아래 항목들은 KX-21에선 검사하지 않음(SE-9000에서 검사)
                    'RET%
                    tmpRst(24) = Mid$(RcvBuffer, 189, 4)

                    If tmpRst(24) = "    " Then
                        tmpRst(24) = "N"
                    Else
                        tmpRst(24) = Format$(Val(Format$(tmpRst(24), "@@.@@")), "0.00")
                    End If
                    'RET#
                    tmpRst(25) = Mid$(RcvBuffer, 194, 4)

                    If tmpRst(25) = "    " Then
                        tmpRst(25) = "N"
                    Else
                        tmpRst(25) = Trim(Val("0." & tmpRst(25)))
                    End If
                    '이상 데이터 거르기
                    For ii = 24 To 25
                        If Val(tmpRst(ii)) = "0" Then
                            tmpRst(ii) = "-"
                        End If
                    Next ii
                    '--- 여기까지...SE-9000에서만 검사...

                    With pResultInfo
                        .ID = tmpBarCd
                        .RACK = tmpRack
                        .POS = tmpPos

                        For ii = 1 To 25
                            .RSTCNT = .RSTCNT + 1

                            .IFCD = .IFCD & Trim(ii) & Chr(124)
                            .RST1 = .RST1 & tmpRst(ii) & Chr(124)
                            .RST2 = .RST2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & Chr(124)
                        Next ii
                    End With

                    '결과 처리
                    With pResultInfo
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                    End With

                    'Query 도중 결과가 먼저 나온 경우를 위해
                    If m_iOrderFlag = 1 Then
                        Call SendOrder_SE9000
                        m_iPhase = 2
                    Else
                        m_iPhase = 1
                    End If

                Case "C"

                Case Else

            End Select

        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub
Private Sub DataEditResponse_SE9000()
    On Error GoTo ErrRtn

    Dim sBC     As String
    Dim sLC     As String

    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack As String
    Dim tmpPos  As String
    Dim ii      As Integer
    Dim tmpRst()    As String       '결과 임시 저장

    sBC = Mid(RcvBuffer, 1, 2)
    sLC = Mid(RcvBuffer, 3, 1)

    Select Case sBC
        Case "R1"
            pSampleInfo.ID = ""

            tmpBarCd = Trim(Mid(RcvBuffer, 3, 13))
            pSampleInfo.ID = Trim(tmpBarCd)

            Call SendOrder_SE9000
            m_iPhase = 2

            Exit Sub

        Case "D1"
            Select Case sLC
                Case "U"
                    '결과정보 초기화
                    Call Init_pResultInfo

                    tmpRack = Mid(RcvBuffer, 10, 4)
                    tmpPos = Mid(RcvBuffer, 14, 2)
                    tmpBarCd = Trim(Mid(RcvBuffer, 22, 13))

                    ReDim tmpRst(25) As String

                    'WBC
                    tmpRst(1) = Mid(RcvBuffer, 63, 5)
                    If tmpRst(1) = Space(5) Then
                        tmpRst(1) = "N"
                    Else
                        tmpRst(1) = Format(Val(Format(tmpRst(1), "@@@.@@")), "0.00")
                    End If

                    'RBC
                    tmpRst(2) = Mid(RcvBuffer, 69, 4)

                    If tmpRst(2) = Space(4) Then
                        tmpRst(2) = "N"
                    Else
                        tmpRst(2) = Format(Val(Format$(tmpRst(2), "@@.@@")), "0.00")
                    End If

                    'HGB, HCT, MCV, MCH, MCHC
                    For ii = 3 To 7
                        tmpRst(ii) = Mid(RcvBuffer, 74 + (ii - 3) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    'PLT
                    tmpRst(8) = Mid(RcvBuffer, 99, 4)

                    If tmpRst(8) = Space(4) Then
                        tmpRst(8) = "N"
                    Else
                        tmpRst(8) = Trim(Val(Format(tmpRst(8), "@@@@")))
                    End If

                    'LYMPH%, MONO%, NEUT%, EO%, BASO%
                    For ii = 9 To 13
                        tmpRst(ii) = Mid(RcvBuffer, 104 + (ii - 9) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    'LYMPH#, MONO#, NEUT#, EO#, BASO#
                    For ii = 14 To 18
                        tmpRst(ii) = Mid(RcvBuffer, 129 + (ii - 14) * 6, 5)

                        If tmpRst(ii) = Space(5) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@@")), "0.00")
                        End If
                    Next ii

                    'RDW-CV(%), RDW-SD(fL), PDW(fL), MPV(fL), P-LCR
                    For ii = 19 To 23
                        tmpRst(ii) = Mid(RcvBuffer, 159 + (ii - 19) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = ""
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    '--- 아래 항목들은 KX-21에선 검사하지 않음(SE-9000에서 검사)
                    'RET%
                    tmpRst(24) = Mid$(RcvBuffer, 189, 4)

                    If tmpRst(24) = "    " Then
                        tmpRst(24) = "N"
                    Else
                        tmpRst(24) = Format$(Val(Format$(tmpRst(24), "@@.@@")), "0.00")
                    End If
                    'RET#
                    tmpRst(25) = Mid$(RcvBuffer, 194, 4)

                    If tmpRst(25) = "    " Then
                        tmpRst(25) = "N"
                    Else
                        tmpRst(25) = Trim(Val("0." & tmpRst(25)))
                    End If
                    '이상 데이터 거르기
                    For ii = 24 To 25
                        If Val(tmpRst(ii)) = "0" Then
                            tmpRst(ii) = "-"
                        End If
                    Next ii
                    '--- 여기까지...SE-9000에서만 검사...

                    With pResultInfo
                        .ID = tmpBarCd
                        .RACK = tmpRack
                        .POS = tmpPos

                        For ii = 1 To 25
                            .RSTCNT = .RSTCNT + 1

                            .IFCD = .IFCD & Trim(ii) & Chr(124)
                            .RST1 = .RST1 & tmpRst(ii) & Chr(124)
                            .RST2 = .RST2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & Chr(124)
                        Next ii
                    End With

                    '결과 처리
                    With pResultInfo
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                    End With

                    'Query 도중 결과가 먼저 나온 경우를 위해
                    If m_iOrderFlag = 1 Then
                        Call SendOrder_SE9000
                        m_iPhase = 2
                    Else
                        m_iPhase = 1
                    End If

                Case "C"

                Case Else

            End Select

        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_SE9000()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                    
                    Case 3      'ETX
                        Call DataEdit_SE9000
                        
                        msComm.Output = Chr(6)       'ACK

                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                End Select
                
            Case 2
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        RaiseEvent SendOrderOK(pSampleInfo.ID)
                        
                        'Order를 보내고 다시 초기 상태
                        m_iPhase = 1
                        m_iOrderFlag = 0
                        
                    Case 21
                        Call SendOrder_SE9000
                    
                    Case Else
                        m_iPhase = 1
                        m_iOrderFlag = 0
                End Select
        End Select
    Next ix1
    
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

Private Sub SendOrder_SE9000()
    On Error GoTo ErrRtn
    
    Dim SendBuf$, sTestCd$, sBuf$
    Dim iPos%, i%
    Dim sOrder$
    
    sTestCd = String(31, "0")

    RaiseEvent RequestCurOrder(pSampleInfo.ID, "", "")
    
    Call Get_OrderString
    
    If pSampleInfo.ORDCNT = 0 Then
        RaiseEvent DispMsg("인터페이스 오더 항목이 존재하지 않습니다!!")
        Exit Sub
    End If
    
    For i = 1 To pSampleInfo.ORDCNT
        sOrder = sOrder & Trim(pSampleInfo.IFCD(i))
    Next i
    
    'ORDER 편집
    If InStr(sOrder, "C") > 0 Then      'CBC
        For i = 1 To 8
            Mid(sTestCd, i, 1) = "1"
        Next i
        For i = 19 To 23
            Mid(sTestCd, i, 1) = "1"
        Next i
    End If
    If InStr(sOrder, "D") > 0 Then      'DIFF
        For i = 9 To 18
            Mid(sTestCd, i, 1) = "1"
        Next i
        Mid(sTestCd, 24, 1) = "1"
    End If
    If InStr(sOrder, "R") > 0 Then      'RETI
        For i = 25 To 26
            Mid(sTestCd, i, 1) = "1"
        Next i
        For i = 28 To 30
            Mid(sTestCd, i, 1) = "1"
        Next i
    End If
    
    SendBuf = "S"
    SendBuf = SendBuf & "1"
    SendBuf = SendBuf & Format(Now, "YYYYMMDD")
    SendBuf = SendBuf & Right(String(13, "0") & pSampleInfo.ID, 13)
    SendBuf = SendBuf & Space$(4)
    SendBuf = SendBuf & Space$(2)
    SendBuf = SendBuf & "1"
    SendBuf = SendBuf & Right(String(13, "0") & pSampleInfo.ID, 13)
    SendBuf = SendBuf & Space$(25)
    SendBuf = SendBuf & "1"
    SendBuf = SendBuf & Space$(8)
    SendBuf = SendBuf & Space$(15)
    SendBuf = SendBuf & Space$(8)
    SendBuf = SendBuf & Space$(20)
    SendBuf = SendBuf & Space$(20)
    SendBuf = SendBuf & sTestCd
    
    msComm.Output = Chr(2) & SendBuf & Chr(3)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & SendBuf & Chr(3))
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
    End If
End Sub

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



Private Sub SE9000_TEMP()
''-----------
'    'Order 종합
'    If iOrderOpt = 0 Then
'        If iRealCnt > 0 Then    'CBC 항목이 낱개로 난 경우...
'            sBuf = sBuf & "1111111100000000001111100000000"   'CBC ONLY
'        Else
'            sBuf = sBuf & "0000000000000000000000000000000"   'No Order
'        End If
'
'    ElseIf iOrderOpt = 1 Then
'        'CBC ONLY
'        sBuf = sBuf & "1111111100000000001111100000000"   'CBC ONLY
'                      '1234567890123456789012345678901  <31>
'
'    ElseIf iOrderOpt = 2 Then
'        'DIFF ONLY
'        sBuf = sBuf & "0000000011111111110000010000000"   'DIFF ONLY
'                      '1234567890123456789012345678901  <31>
'    ElseIf iOrderOpt = 3 Then
'        'CBC + DIFF
'        sBuf = sBuf & "1111111111111111111111110000000"   'CBC + DIFF
'                      '1234567890123456789012345678901  <31>
'    ElseIf iOrderOpt = 4 Then
'        'RET|C ONLY
'        sBuf = sBuf & "0000000000000000000000001101110"   'RETIC ONLY
'                      '1234567890123456789012345678901  <31>
'    ElseIf iOrderOpt = 5 Then
'        'CBC + RETIC
'        sBuf = sBuf & "1111111100000000001111101101110"   'CBC + RETIC
'                      '1234567890123456789012345678901  <31>
'
'    ElseIf iOrderOpt = 6 Then
'        'DIFF + RETIC
'        sBuf = sBuf & "0000000011111111110000011101110"   'DIFF + RETIC
'                      '1234567890123456789012345678901  <31>
'
'    ElseIf iOrderOpt = 7 Then
'        'CBC + DIFF + RETIC
'        sBuf = sBuf & "1111111111111111111111111101110"   'CBC + DIFF + RETIC
'                      '1234567890123456789012345678901  <31>
'    End If
''-----------

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
    On Error GoTo ErrRtn
    
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
    m_iOrderFlag = PropBag.ReadProperty("iOrderFlag", m_def_iOrderFlag)
    
    m_iTotalItemCnt = PropBag.ReadProperty("iTotalItemCnt", m_def_iTotalItemCnt)
    
ErrRtn:
    If Err <> 0 Then
        MsgBox Err.Description
    End If
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
    Call PropBag.WriteProperty("iOrderFlag", m_iOrderFlag, m_def_iOrderFlag)
    Call PropBag.WriteProperty("iTotalItemCnt", m_iTotalItemCnt, m_def_iTotalItemCnt)
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
    
    m_iOrderFlag = 0
    
ErrPortOpen:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
        RaiseEvent RaiseError("PortOpen Error!!! " & Err.Description)
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
    m_iOrderFlag = m_def_iOrderFlag
    m_iTotalItemCnt = m_def_iTotalItemCnt
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iOrderFlag() As Integer
    iOrderFlag = m_iOrderFlag
End Property

Public Property Let iOrderFlag(ByVal New_iOrderFlag As Integer)
    m_iOrderFlag = New_iOrderFlag
    PropertyChanged "iOrderFlag"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iTotalItemCnt() As Integer
    iTotalItemCnt = m_iTotalItemCnt
End Property

Public Property Let iTotalItemCnt(ByVal New_iTotalItemCnt As Integer)
    m_iTotalItemCnt = New_iTotalItemCnt
    PropertyChanged "iTotalItemCnt"
End Property

