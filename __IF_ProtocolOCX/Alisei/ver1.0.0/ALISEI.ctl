VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl ALISEI 
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
      InputMode       =   1
   End
End
Attribute VB_Name = "ALISEI"
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
Event RequestNextOrder()
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sOther1$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
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
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

'for ALISEI
Dim miCnt   As Integer
Dim miIDBlk     As Integer
Dim mvSndBuf1   As Variant
Dim mvSndBuf2   As Variant
Dim mvWkBuf As Variant
Dim msRcvBuf    As String



Private Sub SendOrder_ALISEI(ByRef iOrdCnt As Integer)
    On Error GoTo ErrSendOrd
    
    Dim sSndBuff    As String
    Dim btSBuf1()   As Byte
    Dim btSBuf2()   As Byte
    Dim sTestCd$, sDateInfo$
    Dim ii%, iCnt%
    Dim iChkSum%
    Dim sPrtLog As String
    
    '현재 전송할 오더 조회
    RaiseEvent RequestNextOrder

    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
        End With
        
        msComm.Output = Chr(4)
        RaiseEvent PrintSendLog(Chr(4))

        miIDBlk = 0     '2006/9/13 yk

        iOrdCnt = 0
        
        Exit Sub
    End If

    iOrdCnt = m_p_iOrdCnt
    
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
    
    'Order 편집
    sTestCd = ""
    For ii = 1 To pSampleInfo.ORDCNT
        sTestCd = sTestCd & "," & Trim(pSampleInfo.IFCD(ii))
    Next ii
    
    
    ReDim btSBuf1(131)
    ReDim btSBuf2(131)
    
    btSBuf1(0) = 1
    btSBuf2(0) = 1
    
    miIDBlk = miIDBlk + 1
    If miIDBlk > 255 Then
        miIDBlk = 0
    End If
    
'        btSBuf1(1) = miIDBlk - 254  '255
'        btSBuf1(2) = 255 - (miIDBlk - 254)  '255)
'    Else
        btSBuf1(1) = miIDBlk
        btSBuf1(2) = 255 - miIDBlk
'    End If
    
    miIDBlk = miIDBlk + 1
    If miIDBlk > 255 Then
        miIDBlk = 0
    End If
        
'        btSBuf2(1) = miIDBlk - 254  '255
'        btSBuf2(2) = 255 - (miIDBlk - 254)  '255)
'    Else
        btSBuf2(1) = miIDBlk
        btSBuf2(2) = 255 - miIDBlk
'    End If
        
    sDateInfo = Format(Now, "DDMMYYYY")
    
    sSndBuff = ""
    sSndBuff = sSndBuff & "[" & CStr(Val(pSampleInfo.POS)) & "]" & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "Id=" & pSampleInfo.ID & Chr(13) & Chr(10)
    
    '2004-04-02 KHS (Name 24자리밖에들어갈수없음)
    sSndBuff = sSndBuff & "Name=" & Chr(13) & Chr(10)
'    sSndBuff = sSndBuff & "Name=" & gOrderTable.sName & cCR & cLF

    sSndBuff = sSndBuff & "Surname=" & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "Ward=" & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "Sex=" & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "Pregnancy=" & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "B_Day=" & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "B_Month=" & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "B_Year=" & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "Ins_Day=" & Mid(sDateInfo, 1, 2) & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "Ins_Month=" & Mid(sDateInfo, 3, 2) & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "Ins_Year=" & Mid(sDateInfo, 5, 4) & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & "Test=" & Mid(sTestCd, 2) & Chr(13) & Chr(10)
    sSndBuff = sSndBuff & Chr(13) & Chr(10)
    
    sPrtLog = sSndBuff
    
    '1st SendData
    iChkSum = 0
    For ii = 1 To 128
        btSBuf1(ii + 2) = Asc(Mid(sSndBuff, ii, 1))
        iChkSum = iChkSum + Asc(Mid(sSndBuff, ii, 1))
    Next ii
    iChkSum = iChkSum Mod 256
    btSBuf1(131) = iChkSum
    
    mvSndBuf1 = btSBuf1
    
    '2nd SendData
    sSndBuff = Mid(sSndBuff, 129)
    sSndBuff = sSndBuff & Space(128)
    sSndBuff = Left(sSndBuff, 128)
    
    iChkSum = 0
    For ii = 1 To 128
        btSBuf2(ii + 2) = Asc(Mid(sSndBuff, ii, 1))
        iChkSum = iChkSum + Asc(Mid(sSndBuff, ii, 1))
    Next ii
    iChkSum = iChkSum Mod 256
    btSBuf2(131) = iChkSum
    
    mvSndBuf2 = btSBuf2
    
    
    'Log 작성
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Asc(btSBuf1(1)) & "/" & Asc(btSBuf1(2)) & "," & sSndBuff)
    End If
    
ErrSendOrd:
    If Err <> 0 Then
        m_iPhase = 1
        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
    End If
End Sub
Private Sub Order_Input()
'    On Error GoTo ErrHandler
'
''환자의 Order 전송
'    Dim sSendBuff As String
'    Dim i%, j%, k%, iOrdCnt%, iChkSum%
'    Dim vIFCnt, vTmp
'    Dim sTmp$, sTIFOrdCd$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$
'    Dim objOrd As Object
'    Dim sDateInfo$, sOrdTmp$
'
'    ReDim miBuff1(131)
'    ReDim miBuff2(131)
'
'    miBuff1(0) = 1
'    miBuff2(0) = 1
'
'    miNum = miNum + 1
'
'    If miNum > 255 Then
'        miBuff1(1) = miNum - 255
'        miBuff1(2) = 255 - (miNum - 255)
'    Else
'        miBuff1(1) = miNum
'        miBuff1(2) = 255 - miNum
'    End If
'
'    miNum = miNum + 1
'
'    If miNum > 255 Then
'        miBuff2(1) = miNum - 255
'        miBuff2(2) = 255 - (miNum - 255)
'    Else
'        miBuff2(1) = miNum
'        miBuff2(2) = 255 - miNum
'    End If
'
'    sSendBuff = ""
'    sOrdTmp = ""
'
'    sDateInfo = Format(Now, "DDMMYYYY")
'
'    sSendBuff = ""
'    sSendBuff = sSendBuff & "[" & CStr(Val(gOrderTable.sPos)) & "]" & cCR & cLF
'    sSendBuff = sSendBuff & "Id=" & gOrderTable.sSampID & cCR & cLF
'    '--- 2004-04-02 KHS Modified (Name 24자리밖에들어갈수없음)
''    sSendBuff = sSendBuff & "Name=" & cCR & cLF
'    sSendBuff = sSendBuff & "Name=" & gOrderTable.sName & cCR & cLF
'
'    sSendBuff = sSendBuff & "Surname=" & cCR & cLF
'    sSendBuff = sSendBuff & "Ward=" & cCR & cLF
'    sSendBuff = sSendBuff & "Sex=" & cCR & cLF
'    sSendBuff = sSendBuff & "Pregnancy=" & cCR & cLF
'    sSendBuff = sSendBuff & "B_Day=" & cCR & cLF
'    sSendBuff = sSendBuff & "B_Month=" & cCR & cLF
'    sSendBuff = sSendBuff & "B_Year=" & cCR & cLF
'    sSendBuff = sSendBuff & "Ins_Day=" & Mid(sDateInfo, 1, 2) & cCR & cLF
'    sSendBuff = sSendBuff & "Ins_Month=" & Mid(sDateInfo, 3, 2) & cCR & cLF
'    sSendBuff = sSendBuff & "Ins_Year=" & Mid(sDateInfo, 5, 4) & cCR & cLF
'
'    For i = 1 To gOrderTable.iOrdCnt
'        If gOrderTable.sIFTestCd(i) <> "" Then
'            sOrdTmp = sOrdTmp & "," & Trim(gOrderTable.sIFTestCd(i))
''            sOrdTmp = sOrdTmp & "," & ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))
'        End If
'    Next i
'
'    sSendBuff = sSendBuff & "Test=" & Mid(sOrdTmp, 2) & cCR & cLF
'    'sSendBuff = sSendBuff & "Test=1,2" & cCR & cLF  For Test
'
'    sSendBuff = sSendBuff & cCR & cLF
'
'    iChkSum = 0
'    For i = 1 To 128
'        miBuff1(i + 2) = Asc(Mid(sSendBuff, i, 1))
'        iChkSum = iChkSum + Asc(Mid(sSendBuff, i, 1))
'    Next i
'
'    iChkSum = iChkSum Mod 256
'    miBuff1(131) = iChkSum
'
'    mvSndBuf1 = miBuff1
'
'    sSendBuff = Mid(sSendBuff, 129)
'    sSendBuff = sSendBuff & Space(128)
'    sSendBuff = Left(sSendBuff, 128)
'
'    iChkSum = 0
'    For i = 1 To 128
'        miBuff2(i + 2) = Asc(Mid(sSendBuff, i, 1))
'        iChkSum = iChkSum + Asc(Mid(sSendBuff, i, 1))
'    Next i
'
'    iChkSum = iChkSum Mod 256
'    miBuff2(131) = iChkSum
'
'    mvSndBuf2 = miBuff2
'
'    pnlStatus.Caption = gOrderTable.sSampID & " Order OK!"
'
''    Call DisplayOrderOK("AFTER_ORDER")
'    With spdResult
'        .Col = -1: .Row = gOrderTable.iCRow
'        .BackColor = RGB(255, 245, 245)
'    End With
'
'    If piTestMode = 77 Then
'        Print #2, sSendBuff;
'    End If
'
'    Exit Sub
'ErrHandler:
'    OrderFlag = 0
'    miPhase = 1
'    pnlStatus.Caption = "Order_Input 오류 - (" & Err.Description & ")"
End Sub

'
'   comInputModeBinary
'
Private Sub PhaseCfg_Protocol_ALISEI_SIMPLEX()

    Dim wkDat   As String
    Dim ix1     As Integer
    Dim vWkBuf  As Variant
    Dim iCnt%
    
    vWkBuf = mvWkBuf
    
    If LenB(mvWkBuf) > 0 Then
        If vWkBuf(0) = Asc(Chr(21)) Then    'And m_iPhase <> 3 Then
            m_iPhase = 2
            miCnt = 0       '''2006/9/12 yk
        End If
    End If
    
    For ix1 = 0 To LenB(vWkBuf) - 1
'        'for Test
'        If sTestMode = "77" Then
'            RaiseEvent PrintRcvLog("<" & vWkBuf(ix1) & ">")
'        End If
    
        Select Case m_iPhase
            Case 1
                Select Case vWkBuf(ix1)
                    Case Asc(Chr(1))        'SOH
                        miCnt = miCnt + 1
'                        miCnt = 1
                        
                    Case Asc(Chr(4))        'EOT
                        miCnt = miCnt + 1
                        
                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
                        Else
                            If CStr(vWkBuf(ix1)) > 31 And CStr(vWkBuf(ix1)) < 123 Then
                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
                            ElseIf CStr(vWkBuf(ix1)) = 10 Or CStr(vWkBuf(ix1)) = 13 Then
                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
                            Else
                                RcvBuffer = RcvBuffer & " "
                            End If
                        End If

                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
                        Else
                            '2006/10/9 yk
                            If sTestMode = "77" Then
                                RaiseEvent PrintRcvLog(msRcvBuf)
                            End If
                            
                            Call DataEditResponse_ALISEI_SIMPLEX
                            
                            msComm.Output = Chr(6)
                                                        
                            miCnt = 0
                            RcvBuffer = ""
                            msRcvBuf = ""
                            
                            '결과받은후에는 ORDER나 RESULT ACTION 없으면 빠지게...
                            'm_iPhase = 4
                        End If
                        
                    Case Asc(Chr(6))
                        msComm.Output = Chr(6)
                    
                    Case Else
                        miCnt = miCnt + 1
                        
                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
                        Else
                            If CStr(vWkBuf(ix1)) > 31 And CStr(vWkBuf(ix1)) < 123 Then
                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
                            ElseIf CStr(vWkBuf(ix1)) = 10 Or CStr(vWkBuf(ix1)) = 13 Then
                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
                            Else
                                RcvBuffer = RcvBuffer & " "
                            End If
                        End If
                        
                End Select
            
''            Case 2
''                Select Case vWkBuf(ix1)
''                    Case Asc(Chr(21))   'NAK
''                        If LenB(mvSndBuf1) = 0 And LenB(mvSndBuf2) <> 0 Then
''                            iCnt = 1
''                        Else
''                            Call SendOrder_ALISEI(iCnt)
''                        End If
''
''                        If iCnt > 0 Then
''                            If Trim(mvSndBuf1) <> "" Then
''                                msComm.Output = mvSndBuf1
''                                mvSndBuf1 = ""
''                            ElseIf Trim(mvSndBuf2) <> "" Then
''                                msComm.Output = mvSndBuf2
''                                mvSndBuf2 = ""
''                            End If
''                        End If
''
''                        m_iPhase = 3
''                End Select
''
''            Case 3
''                Select Case vWkBuf(ix1)
''                    Case 21     'NAK
''                        If LenB(mvSndBuf2) = 0 Then
''                            msComm.Output = Chr(4)
''
''                        ElseIf Trim(mvSndBuf1) <> "" Then
''                            msComm.Output = mvSndBuf1
''                            mvSndBuf1 = ""
''
''                        ElseIf Trim(mvSndBuf2) <> "" Then
''                            msComm.Output = mvSndBuf2
''                            mvSndBuf2 = ""
''                        End If
''
''                    Case 6      'ACK
''                        If Trim(mvSndBuf1) <> "" Then
''                            msComm.Output = mvSndBuf1
''                            mvSndBuf1 = ""
''                        ElseIf Trim(mvSndBuf2) <> "" Then
''                            msComm.Output = mvSndBuf2
''                            mvSndBuf2 = ""
''                        Else
''                            If LenB(mvSndBuf1) = 0 And LenB(mvSndBuf2) <> 0 Then
''                                iCnt = 1
''                            Else
''                                Call SendOrder_ALISEI(iCnt)
''                            End If
''
''                            If iCnt > 0 Then
''                                If Trim(mvSndBuf1) <> "" Then
''                                    msComm.Output = mvSndBuf1
''                                    mvSndBuf1 = ""
''                                ElseIf Trim(mvSndBuf2) <> "" Then
''                                    msComm.Output = mvSndBuf2
''                                    mvSndBuf2 = ""
''                                End If
''                            End If
''                        End If
''
'''                    Case Else
'''                        miCnt = miCnt + 1
'''
'''                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
'''                        Else
'''                            If CStr(vWkBuf(ix1)) > 31 And CStr(vWkBuf(ix1)) < 123 Then
'''                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
'''                            ElseIf CStr(vWkBuf(ix1)) = 10 Or CStr(vWkBuf(ix1)) = 13 Then
'''                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
'''                            Else
'''                                RcvBuffer = RcvBuffer & " "
'''                            End If
'''                        End If
''                End Select
                
            Case Else
        End Select

        If miCnt = 132 Then
            msRcvBuf = RcvBuffer
            msComm.Output = Chr(6)
            RaiseEvent PrintSendLog(Chr(6))
            miCnt = 0
            
            Exit Sub
        End If
    Next ix1
    
End Sub

'
'   comInputModeBinary
'
Private Sub PhaseCfg_Protocol_ALISEI()

    Dim wkDat   As String
    Dim ix1     As Integer
    Dim vWkBuf  As Variant
    Dim iCnt%
    
    vWkBuf = mvWkBuf
    
    If LenB(mvWkBuf) > 0 Then
        If vWkBuf(0) = Asc(Chr(21)) Then    'And m_iPhase <> 3 Then
            m_iPhase = 2
            miCnt = 0       '''2006/9/12 yk
        End If
    End If
    
    For ix1 = 0 To LenB(vWkBuf) - 1
'        'for Test
'        If sTestMode = "77" Then
'            RaiseEvent PrintRcvLog("<" & vWkBuf(ix1) & ">")
'        End If
    
        Select Case m_iPhase
            Case 1
                Select Case vWkBuf(ix1)
                    Case Asc(Chr(1))        'SOH
                        miCnt = miCnt + 1
'                        miCnt = 1
                        
                    Case Asc(Chr(4))        'EOT
                        miCnt = miCnt + 1
                        
                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
                        Else
                            If CStr(vWkBuf(ix1)) > 31 And CStr(vWkBuf(ix1)) < 123 Then
                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
                            ElseIf CStr(vWkBuf(ix1)) = 10 Or CStr(vWkBuf(ix1)) = 13 Then
                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
                            Else
                                RcvBuffer = RcvBuffer & " "
                            End If
                        End If

                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
                        Else
                            '2006/10/9 yk
                            If sTestMode = "77" Then
                                RaiseEvent PrintRcvLog(msRcvBuf)
                            End If
                            
                            If UCase(m_EqName) = "ALISEI_OLD" Then
                                Call DataEditResponse_ALISEI_Old
                            Else
                                Call DataEditResponse_ALISEI
                            End If
                            
                            msComm.Output = Chr(6)
                                                        
                            miCnt = 0
                            RcvBuffer = ""
                            msRcvBuf = ""
                            
                            '결과받은후에는 ORDER나 RESULT ACTION 없으면 빠지게...
                            m_iPhase = 4
                        End If
                        
''                    Case Asc(Chr(6))
''                        msComm.Output = Chr(6)
                    
                    Case Else
                        miCnt = miCnt + 1
                        
                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
                        Else
                            If CStr(vWkBuf(ix1)) > 31 And CStr(vWkBuf(ix1)) < 123 Then
                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
                            ElseIf CStr(vWkBuf(ix1)) = 10 Or CStr(vWkBuf(ix1)) = 13 Then
                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
                            Else
                                RcvBuffer = RcvBuffer & " "
                            End If
                        End If
                        
                End Select
            
            Case 2
                Select Case vWkBuf(ix1)
                    Case Asc(Chr(21))   'NAK
                        If LenB(mvSndBuf1) = 0 And LenB(mvSndBuf2) <> 0 Then
                            iCnt = 1
                        Else
                            Call SendOrder_ALISEI(iCnt)
                        End If
                        
                        If iCnt > 0 Then
                            If Trim(mvSndBuf1) <> "" Then
                                msComm.Output = mvSndBuf1
                                mvSndBuf1 = ""
                            ElseIf Trim(mvSndBuf2) <> "" Then
                                msComm.Output = mvSndBuf2
                                mvSndBuf2 = ""
                            End If
                        End If
                        
                        m_iPhase = 3
                End Select
                
            Case 3
                Select Case vWkBuf(ix1)
                    Case 21     'NAK
                        If LenB(mvSndBuf2) = 0 Then
                            msComm.Output = Chr(4)
                            
                        ElseIf Trim(mvSndBuf1) <> "" Then
                            msComm.Output = mvSndBuf1
                            mvSndBuf1 = ""
                            
                        ElseIf Trim(mvSndBuf2) <> "" Then
                            msComm.Output = mvSndBuf2
                            mvSndBuf2 = ""
                        End If
                        
                    Case 6      'ACK
                        If Trim(mvSndBuf1) <> "" Then
                            msComm.Output = mvSndBuf1
                            mvSndBuf1 = ""
                        ElseIf Trim(mvSndBuf2) <> "" Then
                            msComm.Output = mvSndBuf2
                            mvSndBuf2 = ""
                        Else
                            If LenB(mvSndBuf1) = 0 And LenB(mvSndBuf2) <> 0 Then
                                iCnt = 1
                            Else
                                Call SendOrder_ALISEI(iCnt)
                            End If
                            
                            If iCnt > 0 Then
                                If Trim(mvSndBuf1) <> "" Then
                                    msComm.Output = mvSndBuf1
                                    mvSndBuf1 = ""
                                ElseIf Trim(mvSndBuf2) <> "" Then
                                    msComm.Output = mvSndBuf2
                                    mvSndBuf2 = ""
                                End If
                            End If
                        End If
                    
'                    Case Else
'                        miCnt = miCnt + 1
'
'                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
'                        Else
'                            If CStr(vWkBuf(ix1)) > 31 And CStr(vWkBuf(ix1)) < 123 Then
'                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
'                            ElseIf CStr(vWkBuf(ix1)) = 10 Or CStr(vWkBuf(ix1)) = 13 Then
'                                RcvBuffer = RcvBuffer & Chr(CStr(vWkBuf(ix1)))
'                            Else
'                                RcvBuffer = RcvBuffer & " "
'                            End If
'                        End If
                End Select
                
            Case Else
        End Select

        If miCnt = 132 Then
            msRcvBuf = RcvBuffer
            msComm.Output = Chr(6)
            RaiseEvent PrintSendLog(Chr(6))
            miCnt = 0
            
            Exit Sub
        End If
    Next ix1
    
End Sub



Private Sub PhaseCfg_Protocol_ALISEI_InputModeText()
'
'    Dim wkDat   As String
'    Dim ix1     As Integer
'    Dim iCnt%
'
'    For ix1 = 1 To LenB(wkBuf)
'        wkDat = MidB$(wkBuf, ix1, 1)
'
'        Select Case m_iPhase
'            Case 1
'                Select Case AscB(wkDat)
'                    Case 1      'SOH
'                        miCnt = 1
'
'                    Case 4      'EOT
'                        miCnt = miCnt + 1
'
'                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
'                        Else
'                            RcvBuffer = RcvBuffer & wkDat
'                        End If
'
'                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
'                        Else
'                            Call DataEditResponse_ALISEI
'
'                            msComm.Output = Chr(6)
'
'                            miCnt = 0
'                            RcvBuffer = ""
'
'                            '결과받은후에는 ORDER나 RESULT ACTION 없으면 빠지게...
'                            m_iPhase = 2
'                        End If
'
'                    Case 21     'NAK
'                        Call SendOrder_ALISEI(iCnt)
'
'                        If iCnt > 0 Then
'                            If Trim(mvSndBuf1) <> "" Then
'                                msComm.Output = mvSndBuf1
'
'                                mvSndBuf1 = ""
'                            ElseIf Trim(mvSndBuf2) <> "" Then
'                                msComm.Output = mvSndBuf2
'                                mvSndBuf2 = ""
'                            End If
'                        End If
'
'                    Case 6      'ACK
'                        If Trim(mvSndBuf1) <> "" Then
'                            msComm.Output = mvSndBuf1
'                            mvSndBuf1 = ""
'                        ElseIf Trim(mvSndBuf2) <> "" Then
'                            msComm.Output = mvSndBuf2
'                            mvSndBuf2 = ""
'                        Else
'                            Call SendOrder_ALISEI(iCnt)
'
'                            If iCnt > 0 Then
'                                If Trim(mvSndBuf1) <> "" Then
'                                    msComm.Output = mvSndBuf1
'                                    mvSndBuf1 = ""
'                                ElseIf Trim(mvSndBuf2) <> "" Then
'                                    msComm.Output = mvSndBuf2
'                                    mvSndBuf2 = ""
'                                End If
'                            End If
'                        End If
'
'                    Case Else
'                        miCnt = miCnt + 1
'
'                        If (miCnt = 2) Or (miCnt = 3) Or (miCnt = 132) Then
'                        Else
'                            RcvBuffer = RcvBuffer & wkDat
'                        End If
'                End Select
'
'            Case Else
'        End Select
'
'        If miCnt = 132 Then
'            msComm.Output = Chr(6)
'            miCnt = 0
'
'            Exit Sub
'        End If
'    Next ix1
    
End Sub
Private Sub DataEditResponse_ALISEI_Old()
    On Error GoTo ErrRtn

    Dim tmpBarCd$, tmpNumber$, tmpIFCd$, tmpRst1$, tmpRst2$
    Dim aRow()  As String
    Dim aIFCd() As String
    Dim aRst1() As String
    Dim ii%, kk%
    Dim sType   As String
    Dim bRstGbn As Boolean: bRstGbn = False
    Dim aData() As String
    Dim sRst1$, sRst2$
    
'    aRow() = Split(RcvBuffer, vbCrLf)
    aRow() = Split(msRcvBuf, vbCrLf)
    For ii = 0 To UBound(aRow()) - 1
'        If Trim(aRow(ii)) = "" Then
'            Exit For
'        End If
        
        sType = Left(aRow(ii), 3)
        Select Case sType
            Case "Id="
                tmpBarCd = Trim(Mid(aRow(ii), 4))
                
                With pResultInfo
                    .ID = tmpBarCd
                End With
                
            Case "Tes"
                tmpIFCd = Trim(Mid(aRow(ii), 6))

            Case "Res"
                bRstGbn = True
                
                tmpRst1 = Trim(Mid(aRow(ii), 8))
                
                If InStr(tmpIFCd, ",") > 0 Then
                    Erase aIFCd()
                    aIFCd() = Split(tmpIFCd, ",")
                    Erase aRst1()
                    aRst1() = Split(tmpRst1, ",")
                    
                    For kk = 0 To UBound(aIFCd())
                        If Trim(aIFCd(kk)) = "" Then Exit For
                        
                        If InStr(Trim(aRst1(kk)), Space(2)) > 0 Then
                            Erase aData()
                            aData() = Split(Trim(aRst1(kk)), Space(2))
                            
                            If UBound(aData()) > 1 Then
                                sRst1 = Trim(aData(0)) & Trim(aData(1))
                                sRst2 = Trim(aData(2))
                            Else
                                sRst1 = Trim(aData(0))
                                sRst2 = Trim(aData(1))
                            End If
                        Else
                            sRst1 = Trim(aRst1(kk)): sRst2 = ""
                        End If
                        
                        With pResultInfo
                            .RSTCNT = .RSTCNT + 1
                            
                            .IFCD = .IFCD & Trim(aIFCd(kk)) & Chr(124)
                            .RST1 = .RST1 & sRst1 & Chr(124)
                            .RST2 = .RST2 & sRst2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & Chr(124)
                        End With
                    Next kk
                Else
                    If InStr(Trim(tmpRst1), Space(2)) > 0 Then
                        Erase aData()
                        aData() = Split(Trim(tmpRst1), Space(2))
                        sRst1 = Trim(aData(0))
                        sRst2 = Trim(aData(1))
                    Else
                        sRst1 = Trim(tmpRst1): sRst2 = ""
                    End If
                        
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & sRst1 & Chr(124)
                        .RST2 = .RST2 & sRst2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                End If
                
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                    End If
                End With
                
                '결과정보 초기화
                Call Init_pResultInfo

            Case Else
                If Left(aRow(ii), 1) = "[" Then
                    '결과정보 초기화
                    Call Init_pResultInfo
                
                    tmpNumber = Trim(Mid(aRow(ii), 2))
                    tmpNumber = Replace(tmpNumber, "]", "")
                    
                    pResultInfo.SEQNO = tmpNumber
                    
                    tmpBarCd = "": tmpIFCd = "": tmpRst1 = "": tmpRst2 = ""
                    bRstGbn = False
                End If
        End Select
    Next ii

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
    End If
End Sub
Private Function ChkSum_RIAmat(ByVal Para As String) As String

'    Dim i   As Integer
'    Dim Tmp As Integer
'    Dim ChkS1   As Integer
'    Dim ChkS2   As String
'
'    Dim sC1$, sC2$
'
'    Dim aBuf()  As Byte
'
'    aBuf = StrConv(Para, vbFromUnicode)
'
'
'    For i = 0 To UBound(aBuf)
'        ChkS1 = ChkS1 + aBuf(i)
'    Next i
'
'    ChkS1 = ChkS1 Mod 256
'
''    ChkS2 = Right$("0" & Hex$(ChkS1), 2)
'    ChkS2 = Right$("0" & CStr(Hex$(ChkS1)), 2)
'
'    sC1 = Mid(ChkS2, 1, 1)
'    sC1 = "3" & sC1
'    sC1 = CDec("&H" & sC1)
'    sC1 = Chr(Val(sC1))
'    sC1 = Right(sC1, 1)
'
'    sC2 = Mid(ChkS2, 2, 1)
'    sC2 = "3" & sC2
'    sC2 = CDec("&H" & sC2)
'    sC2 = Chr(Val(sC2))
'    sC2 = Right(sC2, 1)
'
'    ChkSum_RIAmat = sC1 & sC2
    
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
        Case "ALISEI", "ALISEI_OLD", "BRIO"
            Call PhaseCfg_Protocol_ALISEI
'            Call PhaseCfg_Protocol_ALISEI_InputModeText

        Case "ALISEI_SIMPLEX"
            Call PhaseCfg_Protocol_ALISEI_SIMPLEX
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub DataEditResponse_ALISEI()
    On Error GoTo ErrRtn

    Dim tmpBarCd$, tmpNumber$, tmpIFCd$, tmpRst1$, tmpRst2$
    Dim aRow()  As String
    Dim aIFCd() As String
    Dim aRst1() As String
    Dim ii%, kk%
    Dim sType   As String
    Dim bRstGbn As Boolean: bRstGbn = False
    Dim aData() As String
    Dim sRst1$, sRst2$
    
'    aRow() = Split(RcvBuffer, vbCrLf)
    aRow() = Split(msRcvBuf, vbCrLf)
    For ii = 0 To UBound(aRow()) - 1
''        If Trim(aRow(ii)) = "" Then
''            Exit For
''        End If
        
        sType = Left(aRow(ii), 3)
        Select Case sType
            Case "Id="
                tmpBarCd = Trim(Mid(aRow(ii), 4))
                
                With pResultInfo
                    .ID = tmpBarCd
                End With
                
            Case "Tes"
                tmpIFCd = Trim(Mid(aRow(ii), 6))

            Case "Res"
                bRstGbn = True
                
                tmpRst1 = Trim(Mid(aRow(ii), 8))
                
                If InStr(tmpIFCd, ",") > 0 Then
                    Erase aIFCd()
                    aIFCd() = Split(tmpIFCd, ",")
                    Erase aRst1()
                    aRst1() = Split(tmpRst1, ",")
                    
                    For kk = 0 To UBound(aIFCd())
                        If Trim(aIFCd(kk)) = "" Then Exit For
                        
                        If InStr(Trim(aRst1(kk)), Space(2)) > 0 Then
                            Erase aData()
                            aData() = Split(Trim(aRst1(kk)), Space(2))
                            sRst1 = Trim(aData(0))
                            sRst2 = Trim(aData(1))
                        Else
                            sRst1 = Trim(aRst1(kk)): sRst2 = ""
                        End If
                        
                        With pResultInfo
                            .RSTCNT = .RSTCNT + 1
                            
                            .IFCD = .IFCD & Trim(aIFCd(kk)) & Chr(124)
                            .RST1 = .RST1 & sRst1 & Chr(124)
                            .RST2 = .RST2 & sRst2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & Chr(124)
                        End With
                    Next kk
                Else
                    If InStr(Trim(tmpRst1), Space(2)) > 0 Then
                        Erase aData()
                        aData() = Split(Trim(tmpRst1), Space(2))
                        sRst1 = Trim(aData(0))
                        sRst2 = Trim(aData(1))
                    Else
                        sRst1 = Trim(tmpRst1): sRst2 = ""
                    End If
                        
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & sRst1 & Chr(124)
                        .RST2 = .RST2 & sRst2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                End If
                
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                    End If
                End With
                
                '결과정보 초기화
                Call Init_pResultInfo

            Case Else
                If Left(aRow(ii), 1) = "[" Then
                    '결과정보 초기화
                    Call Init_pResultInfo
                
                    tmpNumber = Trim(Mid(aRow(ii), 2))
                    tmpNumber = Replace(tmpNumber, "]", "")
                    
                    pResultInfo.SEQNO = tmpNumber
                    
                    tmpBarCd = "": tmpIFCd = "": tmpRst1 = "": tmpRst2 = ""
                    bRstGbn = False
                End If
        End Select
    Next ii

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
    End If
End Sub

Private Sub DataEditResponse_ALISEI_SIMPLEX()
    On Error GoTo ErrRtn

    Dim tmpBarCd$, tmpNumber$, tmpIFCd$, tmpRst1$, tmpRst2$
    Dim aRow()  As String
    Dim aIFCd() As String
    Dim aRst1() As String
    Dim ii%, kk%
    Dim sType   As String
    Dim bRstGbn As Boolean: bRstGbn = False
    Dim aData() As String
    Dim sRst1$, sRst2$
    
'    aRow() = Split(RcvBuffer, vbCrLf)
    aRow() = Split(msRcvBuf, vbCrLf)
    For ii = 0 To UBound(aRow()) - 1
''        If Trim(aRow(ii)) = "" Then
''            Exit For
''        End If
        
        sType = Left(aRow(ii), 3)
        Select Case sType
            Case "Id="
                tmpBarCd = Trim(Mid(aRow(ii), 4))
                
                With pResultInfo
                    .ID = tmpBarCd
                End With
                
            Case "Tes"
                tmpIFCd = Trim(Mid(aRow(ii), 6))

            Case "Res"
                bRstGbn = True
                
                tmpRst1 = Trim(Mid(aRow(ii), 8))
                
                If InStr(tmpIFCd, ",") > 0 Then
                    Erase aIFCd()
                    aIFCd() = Split(tmpIFCd, ",")
                    Erase aRst1()
                    aRst1() = Split(tmpRst1, ",")
                    
                    For kk = 0 To UBound(aIFCd())
                        If Trim(aIFCd(kk)) = "" Then Exit For
                        
                        If InStr(Trim(aRst1(kk)), Space(2)) > 0 Then
                            Erase aData()
                            aData() = Split(Trim(aRst1(kk)), Space(2))
                            sRst1 = Trim(aData(0))
                            sRst2 = Trim(aData(1))
                        Else
                            sRst1 = Trim(aRst1(kk)): sRst2 = ""
                        End If
                        
                        With pResultInfo
                            .RSTCNT = .RSTCNT + 1
                            
                            .IFCD = .IFCD & Trim(aIFCd(kk)) & Chr(124)
                            .RST1 = .RST1 & sRst1 & Chr(124)
                            .RST2 = .RST2 & sRst2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & Chr(124)
                        End With
                    Next kk
                Else
                    If InStr(Trim(tmpRst1), Space(2)) > 0 Then
                        Erase aData()
                        aData() = Split(Trim(tmpRst1), Space(2))
                        sRst1 = Trim(aData(0))
                        sRst2 = Trim(aData(1))
                    Else
                        sRst1 = Trim(tmpRst1): sRst2 = ""
                    End If
                        
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & sRst1 & Chr(124)
                        .RST2 = .RST2 & sRst2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                End If
                
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                    End If
                End With
                
                '결과정보 초기화
                Call Init_pResultInfo

            Case Else
                If Left(aRow(ii), 1) = "[" Then
                    '결과정보 초기화
                    Call Init_pResultInfo
                
                    tmpNumber = Trim(Mid(aRow(ii), 2))
                    tmpNumber = Replace(tmpNumber, "]", "")
                    
                    pResultInfo.SEQNO = tmpNumber
                    
                    tmpBarCd = "": tmpIFCd = "": tmpRst1 = "": tmpRst2 = ""
                    bRstGbn = False
                End If
        End Select
    Next ii

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
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
            .IFCD(ii) = tmpData(ii - 1)
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

    'wkBuf = Text1
''    Dim iLen%, ix%
''    Dim aTmp()
''    If Text1 = "" Then Exit Sub
''
''    iLen = Len(Text1)
''
''    ReDim aTmp(iLen) As Variant
''
''    For ix = 0 To iLen - 1
''        aTmp(ix) = Mid(Text1, ix + 1, 1)
''    Next ix
''    mvWkBuf = aTmp
    
    Call PhaseCfg_Protocol

End Sub


Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
'            wkBuf = msComm.Input
            mvWkBuf = msComm.Input
            
'            If sTestMode = "77" Then
'                RaiseEvent PrintRcvLog(wkBuf)
'            End If
                                
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
    m_iOrderFlag = PropBag.ReadProperty("iOrderFlag", m_def_iOrderFlag)
    m_iTotalItemCnt = PropBag.ReadProperty("iTotalItemCnt", m_def_iTotalItemCnt)
'    m_iLenID = PropBag.ReadProperty("iLenID", m_def_iLenID)
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
'MemberInfo=14
Public Function mSetStandBy() As Variant
    
    wkBuf = ""
    RcvBuffer = ""
    msRcvBuf = ""
    
    m_iPhase = 1
    miCnt = 0
    
End Function

Private Function GetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim POS%

    POS = InStr(tStr, sUserSymbol)

    If POS = 0 Then
    Else
        GetByOneUserSymbol = Trim$(Mid$(tStr, 1, POS - 1))
        sOriginal = Trim$(Mid$(sOriginal, POS + 1, Len(sOriginal) - POS))
    End If
End Function
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function mSetStandByOrd() As Variant

    wkBuf = ""
    RcvBuffer = ""
    msRcvBuf = ""
    
    mvSndBuf1 = "": mvSndBuf2 = ""
    
    m_iPhase = 1
    miCnt = 0
    
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function mDebug() As Variant
    
'    MsgBox "m_iPhase:" & m_iPhase & ", miCnt:" & miCnt
    
End Function

