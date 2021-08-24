VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl BT3000 
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   3330
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   960
      Top             =   2625
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "TEST"
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Top             =   1620
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
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblPhase 
      Height          =   300
      Left            =   270
      TabIndex        =   2
      Top             =   2070
      Width           =   870
   End
End
Attribute VB_Name = "BT3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_SendFlag = 0
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
Dim m_SendFlag As Boolean
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
Event SendOrderOK(sID$, sRack$, sPos$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)
Event RequestNextOrder()

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
Dim miRealCnt   As Integer

Dim mbT_Flag    As Boolean      'for BT-3000


''Private Sub DataEditResponse_BT3000()
''    On Error GoTo ErrRtn
''
''    Dim sBC     As String
''    Dim sLC     As String
''    Dim iETBpos%, i%, iRstCnt%
''    Dim sTmpBuf1$, sTmpBuf2$, sTmp$, sTmp2$
''    Dim sSampNo$, sSampType$
''    Dim tmpIFCd$, tmpRst$, tmpFlag$, tmpRstDT$
''    Dim iPos%
''    Dim sUnitNo$
''    Dim sRecType$
''
''    sRecType = Mid(RcvBuffer, 1, 1)
''
''    Select Case sRecType
''        'T 110700001GOT01.73900.03868.506GPT01.75800.04174.535
''        '123456789012345678901234567890123456789012345678901234567890
''        '         1         2         3         4         5         6
''        Case "T"
''            pResultInfo.IFCD = ""
''            pResultInfo.RSTCNT = 0
''
''            pResultInfo.ID = Trim(Mid(RcvBuffer, 2, 10))
''            pResultInfo.RSTCNT = Asc(Mid(RcvBuffer, 12, 1))
''
''            sTmp = Mid(RcvBuffer, 1, 12)
''
''            sTmp2 = Mid(RcvBuffer, 13)
''            sTmp2 = Replace(sTmp2, Chr(1), "")
''            sTmp2 = Replace(sTmp2, Chr(2), "")
''            sTmp2 = Replace(sTmp2, Chr(5), "")
''            sTmp2 = Replace(sTmp2, Chr(14), "")
''
''            RcvBuffer = sTmp & sTmp2
''
''''            RcvBuffer = Replace(Mid(RcvBuffer, 14), Chr(2), "")
''            RcvBuffer = Replace(RcvBuffer, Space(1), "")
''
''            Do
''                tmpIFCd = Trim(Mid(RcvBuffer, 13 + (21 * (i - 1)), 3))
''
''                If tmpIFCd = "" Then
''                    Exit Do
''                End If
''
''                tmpRst = Trim(Mid(RcvBuffer, 28 + (21 * (i - 1)), 6))
''
''                If Left(tmpRst, 1) = "." Then
''                    tmpRst = "0" & tmpRst
''                End If
''
''                With pResultInfo
''                    .IFCD = .IFCD & tmpIFCd & Chr(124)
''                    .RST1 = .RST1 & tmpRst & Chr(124)
''                    .RST2 = .RST2 & Chr(124)
''                    .UNIT = .UNIT & Chr(124)
''                    .FLAG = .FLAG & tmpFlag & Chr(124)
''                End With
''            Loop
''
''            For i = 1 To pResultInfo.RSTCNT
''                tmpIFCd = Trim(Mid(RcvBuffer, 13 + (21 * (i - 1)), 3))
''
''                If tmpIFCd = "" Then
''                    Exit For
''                End If
''
''                tmpRst = Trim(Mid(RcvBuffer, 28 + (21 * (i - 1)), 6))
''
''                If Left(tmpRst, 1) = "." Then
''                    tmpRst = "0" & tmpRst
''                End If
''
''                With pResultInfo
''                    .IFCD = .IFCD & tmpIFCd & Chr(124)
''                    .RST1 = .RST1 & tmpRst & Chr(124)
''                    .RST2 = .RST2 & Chr(124)
''                    .UNIT = .UNIT & Chr(124)
''                    .FLAG = .FLAG & tmpFlag & Chr(124)
''                End With
''            Next i
''
''            '결과값 등록/화면 표시 처리...
''            With pResultInfo
''                ''If .RSTCNT > 0 Then
''                If pResultInfo.RSTCNT = UBound(Split(pResultInfo.IFCD, Chr(124))) Then
''                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .KIND, "", "")
''                End If
''            End With
''
''            Call Init_pResultInfo_BT3000
''
''    End Select
''
''    Exit Sub
''
''ErrRtn:
''    If Err <> 0 Then
''        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
''    End If
''End Sub

Private Sub DataEditResponse_BT3000()
    On Error GoTo ErrRtn
   
    Dim sBC     As String
    Dim sLC     As String
    Dim iETBpos%, i%, iRstCnt%
    Dim sTmpBuf1$, sTmpBuf2$, sTmp$, sTmp2$
    Dim sSampNo$, sSampType$
    Dim tmpIFCd$, tmpRst$, tmpFlag$, tmpRstDT$
    Dim iPos%
    Dim sUnitNo$
    Dim sRecType$
    Dim sTestInfo$
    
    sRecType = Mid(RcvBuffer, 1, 1)
    
    Select Case sRecType
        'T0810300127GOT00.00000.00011.131GPT00.00000.00009.760GGT00.00000.00017.271TBL00.00000.00000.548
        
        Case "T"
            miRealCnt = 0
            
            pResultInfo.ID = Trim(Mid(RcvBuffer, 2, 10))
            pResultInfo.RSTCNT = Asc(Mid(RcvBuffer, 12, 1))
            
            sTestInfo = Trim(Mid(RcvBuffer, 13))
            
'            '<S--- 2008/11/13 yk
'            sTestInfo = Replace(sTestInfo, Chr(1), "")
'            '>E--------------
            For i = 1 To 32    '44
                sTestInfo = Replace(sTestInfo, Chr(i), "")
            Next i
            
            iPos = 1
            
            Do
                tmpIFCd = Mid(sTestInfo, iPos, 3)
                
                If tmpIFCd = "" Or Len(tmpIFCd) < 3 Then Exit Do
                
'                iPos = iPos + 12
'                tmpRst = CStr(Val(Mid(sTestInfo, iPos, 9)))
                
                '<S--- 2008/11/12 yk
                iPos = iPos + 15
                tmpRst = CStr(Val(Mid(sTestInfo, iPos, 6)))
                '>E--------------
                
                If Left(tmpRst, 1) = "." Then
                    tmpRst = "0" & tmpRst
                End If
                                
                miRealCnt = miRealCnt + 1
                
                With pResultInfo
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & tmpFlag & Chr(124)
                End With
                
'                iPos = iPos + 9
                
                '<S--- 2008/11/12 yk
                iPos = iPos + 6
'                If m_SendFlag = True Then       'SendFlag On Mode
'                    iPos = iPos + 1
'                End If
                '>E--------------
                
''                tmpIFCd = Mid(sTestInfo, iPos, 3)
''
''                If tmpIFCd = "" Or Len(tmpIFCd) < 3 Then Exit Do
''
''                iPos = iPos + 12
''
''                tmpRst = CStr(Val(Mid(sTestInfo, iPos, 9)))
''
''                If Left(tmpRst, 1) = "." Then
''                    tmpRst = "0" & tmpRst
''                End If
''
''                miRealCnt = miRealCnt + 1
''
''                With pResultInfo
''                    .IFCD = .IFCD & tmpIFCd & Chr(124)
''                    .RST1 = .RST1 & tmpRst & Chr(124)
''                    .RST2 = .RST2 & Chr(124)
''                    .UNIT = .UNIT & Chr(124)
''                    .FLAG = .FLAG & tmpFlag & Chr(124)
''                End With
''
''                iPos = iPos + 11
            Loop
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If miRealCnt > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, miRealCnt, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .KIND, "", "")
                End If
            End With
'            With pResultInfo
'                ''If .RSTCNT > 0 Then
'                If pResultInfo.RSTCNT = miRealCnt Then
'                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .KIND, "", "")
'                End If
'            End With
            
            ''Call Init_pResultInfo
    
    End Select
    
    Exit Sub

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
    End If
End Sub

Private Sub PhaseCfg_Protocol_BT3000()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
                 
        Select Case m_iPhase
            Case 1      '
                Select Case Asc(wkDat)
                    Case 6      '----- ACK 수신
                        m_iPhase = 2
                        lblPhase = m_iPhase
                        Call SendOrder_BT3000

                    '<S--- 2008/11/12 yk
                    Case 2
                        msComm.Output = Chr(6)
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        m_iPhase = 4
                        lblPhase = m_iPhase
                    '>E--------------
                End Select
                
            Case 2      '
                Select Case wkDat
                    Case "Y", "N"
                        RaiseEvent RequestNextOrder
                        
                        '<S--- 2008/11/12 yk
                        If m_p_iOrdCnt = 0 Then
                            m_iPhase = 3
                            lblPhase = m_iPhase
                        Else
                            m_iPhase = 1
                            lblPhase = m_iPhase
                            msComm.Output = Chr(2)
                            If m_sTestMode = "77" Then
                                RaiseEvent PrintSendLog(Chr(2))
                            End If
                        End If
                        '>E--------------
                        
                        Exit For
                End Select
            
            Case 3
                Select Case Asc(wkDat)
                    Case 2
                        msComm.Output = Chr(6)
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        m_iPhase = 4
                        lblPhase = m_iPhase
                        
'                    '<S--- 2008/11/12 yk
'                    Case 6
'                        msComm.Output = Chr(6)
'                        If m_sTestMode = "77" Then
'                            RaiseEvent PrintSendLog(Chr(6))
'                        End If
'                        m_iPhase = 4
'                    '>E--------------
                    
'                    Case 6      '결과요구인정
'                        msComm.Output = "R" & Chr(4)
'                        m_iPhase = 4
                End Select
                
            Case 4
                Select Case Asc(wkDat)
                    Case 21
                        RcvBuffer = ""
                        m_iPhase = 3
                        lblPhase = m_iPhase
                        'Timer1.Enabled = True
                        mbT_Flag = False
                        
                    Case 4      'EOT
                        If Len(RcvBuffer) > 15 And mbT_Flag = True Then
                            RcvBuffer = RcvBuffer & wkDat
                                                        
                            Call DataEditResponse_BT3000
                            
                            If pResultInfo.RSTCNT = miRealCnt Then
                                RcvBuffer = ""
                                m_iPhase = 3
                                lblPhase = m_iPhase
                                mbT_Flag = False
                                
'                                '<S--- 2008/11/12 yk
'                                msComm.Output = Chr(6)
'                                If m_sTestMode = "77" Then
'                                    RaiseEvent PrintSendLog(Chr(6))
'                                End If
'                                '>E--------------
                            End If
                            
'                            '<S--- 2008/11/12 yk
'                            m_iPhase = 3
'                            lblPhase = m_iPhase
'                            RcvBuffer = ""
'                            mbT_Flag = False
'                            '>E--------------
                            
                            Call Init_pResultInfo
                        Else
                            RcvBuffer = RcvBuffer & wkDat
                        End If
                        
                    '<S--- 2008/11/12 yk
                    Case 2      'STX
                        If Trim(RcvBuffer) = "" Then
                            msComm.Output = Chr(6)
                            If m_sTestMode = "77" Then
                                RaiseEvent PrintSendLog(Chr(6))
                            End If
                            m_iPhase = 4
                            lblPhase = m_iPhase
                        Else
                            RcvBuffer = RcvBuffer & wkDat
                        End If
                    
                    Case 84     'T
                        RcvBuffer = RcvBuffer & wkDat
                        mbT_Flag = True
                    '>E---------------
                    
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                End Select
                
         End Select
    Next ix1
    
End Sub

Private Sub SendOrder_BT3000()
    On Error GoTo ErrRtn

    '환자의 Order 전송
    Dim sSendBuff   As String
    Dim sTestCd     As String
    Dim ii          As Integer
    Dim iCnt        As Integer
    Dim iRealCnt    As Integer
    Dim tmpData()   As String
    Dim sSendTemp   As String

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
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
    End With

    For iCnt = 1 To pSampleInfo.ORDCNT
        If Trim(pSampleInfo.IFCD(iCnt)) = "" Then
        Else
            iRealCnt = iRealCnt + 1
            sSendTemp = sSendTemp & Left(pSampleInfo.IFCD(iCnt) & Space(3), 3)
        End If
    Next
    
    sSendBuff = "PS" & Right(Space(10) & pSampleInfo.ID, 10) & "N" & Format(pSampleInfo.POS, "00") & Format(CStr(iRealCnt), "00") & sSendTemp & Chr(4)
    
    msComm.Output = sSendBuff

    'Order 전송 완료
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    RaiseEvent RequestNextOrder

    'Log 작성
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendBuff)
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
    End If
'''    On Error GoTo ErrRtn
'''
'''    '환자의 Order 전송
'''    Dim sSendBuff   As String
'''    Dim sTestCd     As String
'''    Dim ii          As Integer
'''    Dim iCnt        As Integer
'''    Dim iRealCnt    As Integer
'''    Dim tmpData()   As String
'''    Dim sSendTemp   As String
'''
'''    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
'''        With pSampleInfo
'''            .ID = m_p_sID
'''            .ORDCNT = 0
'''        End With
'''        Exit Sub
'''    End If
'''
'''    ReDim tmpData(m_p_iOrdCnt) As String
'''    tmpData() = Split(m_p_sTIFCd, Chr(124))
'''
'''    With pSampleInfo
'''        .ID = m_p_sID
'''        .SEQNO = m_p_sSeq
'''        .RACK = m_p_sRack
'''        .POS = m_p_sPos
'''        .ORDCNT = m_p_iOrdCnt
'''
'''        ReDim .IFCD(.ORDCNT)
'''        iCnt = 0
'''        For ii = 1 To .ORDCNT
'''            If Trim(tmpData(ii - 1)) <> "" Then
'''                iCnt = iCnt + 1
'''                .IFCD(iCnt) = tmpData(ii - 1)
'''            End If
'''        Next ii
'''        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
'''    End With
'''
'''    'Send Message 편집
'''    sSendBuff = sSendBuff & Left(pSampleInfo.ID & Space(15), 15) & "T" & "S" & "N"
'''    sSendBuff = sSendBuff & pSampleInfo.POS
'''
'''    For iCnt = 1 To pSampleInfo.ORDCNT
'''        If Trim(pSampleInfo.IFCD(iCnt)) = "" Then
'''        Else
'''            iRealCnt = iRealCnt + 1
'''            sSendTemp = sSendTemp & Left(pSampleInfo.IFCD(iCnt) & Space(4), 4)
'''        End If
'''    Next
'''
'''    sSendBuff = sSendBuff & Format(CStr(iRealCnt), "00") & sSendTemp
'''
'''    sSendBuff = sSendBuff & BT3000_CheckSum(sSendBuff) & Chr(4)
'''
'''    msComm.Output = sSendBuff
'''
'''    'Order 전송 완료
'''    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
'''
'''    'Log 작성
'''    If m_sTestMode = "77" Then
'''        RaiseEvent PrintSendLog(sSendBuff)
'''    End If
'''
'''ErrRtn:
'''    If Err <> 0 Then
'''        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
'''    End If
End Sub

Private Function BT3000_CheckSum(ByVal sBuf$) As String
    Dim iCnt As Long
    Dim iSum As Long
    
    For iCnt = 1 To Len(sBuf)
        iSum = iSum + Val(Asc(Mid(sBuf, iCnt, 1)))
    Next
    
    iSum = iSum Mod 256
    
    BT3000_CheckSum = Right("   " & CStr(iSum), 3)
End Function


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
        Case "BT3000"
            Call PhaseCfg_Protocol_BT3000
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
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
        .KIND = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
End Sub

Private Sub Init_pResultInfo_BT3000()
    
    With pResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        ''.RSTCNT = 0
        ''.IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .KIND = ""
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
    '대기상태가 아닐경우 skip
    If iPhase <> 3 Then Exit Sub

    msComm.Output = Chr(2)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2))
    End If
                        
    Timer1.Enabled = False
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
    m_iOrderFlag = PropBag.ReadProperty("iOrderFlag", m_def_iOrderFlag)
    m_iTotalItemCnt = PropBag.ReadProperty("iTotalItemCnt", m_def_iTotalItemCnt)
'    m_iLenID = PropBag.ReadProperty("iLenID", m_def_iLenID)
    m_SendFlag = PropBag.ReadProperty("SendFlag", m_def_SendFlag)
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
'    Call PropBag.WriteProperty("iLenID", m_iLenID, m_def_iLenID)
    Call PropBag.WriteProperty("SendFlag", m_SendFlag, m_def_SendFlag)
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
'    m_iLenID = m_def_iLenID
    m_SendFlag = m_def_SendFlag
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
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(iChr))
    End If
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
    End If
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get SendFlag() As Boolean
    SendFlag = m_SendFlag
End Property

Public Property Let SendFlag(ByVal New_SendFlag As Boolean)
    m_SendFlag = New_SendFlag
    PropertyChanged "SendFlag"
End Property

