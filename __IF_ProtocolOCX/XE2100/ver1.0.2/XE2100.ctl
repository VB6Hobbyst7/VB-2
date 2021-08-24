VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl XE2100 
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
Attribute VB_Name = "XE2100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sPatInfo = "0"
Const m_def_p_sSampInfo = "0"
Const m_def_SiteNm = 0
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
Dim m_p_sPatInfo As String
Dim m_p_sSampInfo As String
Dim m_SiteNm As Variant
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTRstDT$, sOther1$)
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sRack$, sPos$)
Event DispMsg(sMsg$)
Event RequestNextOrder()

'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim wkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'Dim msBarCd As String
'Dim msRack As String
'Dim msPos As String
'Dim msSeqNo As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim iSpaceCnt   As Integer

'for XE-2100
Dim miFlagCnt   As Integer
Dim msFlagCd  As String
Dim msFlagTot   As String
Dim msFlagTot2  As String

Dim msOrdS2     As String

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
        Case "XE2100"
            Call PhaseCfg_Protocol_XE2100
        
        Case "XEALPHA"
            Call PhaseCfg_Protocol_XEAlpha
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
End Sub

Private Sub PhaseCfg_Protocol_XE2100()
    Dim sWkDat$
    Dim i%
    
    For i = 1 To Len(wkBuf)
        sWkDat = Mid(wkBuf, i, 1)
        
        Select Case m_iPhase
            Case 1
                Select Case Asc(sWkDat)
                    Case 2
                        RcvBuffer = ""
                        
                    Case 3
                        msComm.Output = Chr(6)
                        Call DataEditResponse_XE2100
                        RcvBuffer = ""
                        
                    Case Else
                        RcvBuffer = RcvBuffer & sWkDat
                End Select
            
            Case 2
                Select Case Asc(sWkDat)
                    Case 6
                        If msOrdS2 = "" Then
                            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

                            'Order를 보내고 다시 초기 상태
                            m_iPhase = 1
                        Else
                            msComm.Output = msOrdS2
                            msOrdS2 = ""
                        End If

                    Case 21
                        Call SendOrder_XE2100

                    Case Else
                        m_iPhase = 1
                End Select
                
            Case Else
        End Select
    Next i
    
End Sub
Private Sub DataEditResponse_XE2100()
    On Error GoTo ErrRtn

    Dim i%, iRealCnt%, ix1%
    Dim sRxData$, sBarCd$, sSeqNo$, sRack$, sPos$
    Dim sBC$, sLC$
    Dim s_aResult$(32)

    Dim sTotIFCd$, sTotRst$
    Dim FlagBuf$

    sRxData = RcvBuffer

    sBC = Mid(sRxData, 1, 2)
    sLC = Mid(sRxData, 3, 1)

    Select Case sBC
        Case "R1"
            With pSampleInfo
                .ID = Trim(Mid(RcvBuffer, 6, 15))
                .RACK = Trim(Mid(RcvBuffer, 23, 6))
                .POS = Trim(Mid(RcvBuffer, 29, 2))
            End With

            Call SendOrder_XE2100

            m_iPhase = 2

            Exit Sub

        Case "D1"
            If sLC = "B" Or sLC = "E" Then Exit Sub

            miFlagCnt = 0: msFlagCd = "": msFlagTot = "": msFlagTot2 = ""

            If Len(sRxData) > 200 Then
                FlagBuf = ""
                RcvBuffer = Mid$(RcvBuffer, 105, 96)

                '실제 IP FLAG만 취득
                FlagBuf = Mid$(RcvBuffer, 1, 10)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 13, 2)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 17, 3)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 21, 1)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 24, 3)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 33, 10)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 49, 4)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 54, 1)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 65, 4)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 83, 1)

                'Manual 잘못으로인한 두가지 경우의 수를 적용(일본이 실수를?!?!)
                If Mid$(RcvBuffer, 84, 1) = "1" Or Mid$(RcvBuffer, 85, 1) = "1" Then
                    FlagBuf = FlagBuf & "1"
                Else
                    FlagBuf = FlagBuf & "0"
                End If

                For ix1 = 1 To Len(FlagBuf)
                    If Mid$(FlagBuf, ix1, 1) = "1" Then
                        miFlagCnt = miFlagCnt + 1
                        msFlagCd = msFlagCd & Trim(Str(ix1 + 100 - 1)) & "|"
                        msFlagTot = msFlagTot & "Detected!" & "|"
                        msFlagTot2 = msFlagTot2 & "|"
                    End If
                Next ix1
            End If

            sSeqNo = Trim(Mid(sRxData, 20, 10))
            sBarCd = Trim(Mid(sRxData, 33, 15))
            sRack = Trim(Mid(sRxData, 62, 6))
            sPos = Trim(Mid(sRxData, 68, 2))

            '2004/10/14 Bug Fix
'            If Left(sBarCd, 3) = "ERR" Then Exit Sub

            With pSampleInfo
                .ID = sBarCd
                .SEQNO = sSeqNo
                .RACK = sRack
                .POS = sPos

'                RaiseEvent RequestCurOrder(.ID, .RACK, .POS)
            End With

            Exit Sub

        Case "D2"
            If sLC = "B" Or sLC = "E" Then Exit Sub

            Call Init_pResultInfo

            If Len(sRxData) > 253 Then
                RaiseEvent DispMsg("XE2100으로부터 전송된 문자열의 길이 (" & CStr(Len(sRxData)) & ")의 이상이 발생하였습니다!!")
                Exit Sub
            End If

            sSeqNo = Trim(Mid(sRxData, 20, 10))
'            If pSampleInfo.SEQNO <> "" And pSampleInfo.SEQNO <> sSeqNo Then      ' msSeqNo <> sSeqNo Then
'                RaiseEvent DispMsg("D1과 D2에서 다른 장비일련번호가 전송되었습니다!!")
'                miFlagCnt = 0: msFlagCd = "": msFlagTot = "": msFlagTot2 = ""
'                Exit Sub
'            End If
'
            sBarCd = Trim(Mid(sRxData, 33, 15))
'            If pSampleInfo.ID <> "" And pSampleInfo.ID <> sBarCd Then        ' msBarCd <> sBarCd Then
'                RaiseEvent DispMsg("D1과 D2에서 다른 바코드 정보가 전송되었습니다!!")
'                miFlagCnt = 0: msFlagCd = "": msFlagTot = "": msFlagTot2 = ""
'                Exit Sub
'            End If

            'WBC
            s_aResult(1) = Mid(sRxData, 48, 5)

            If s_aResult(1) = Space(5) Then
                s_aResult(1) = "N"
            Else
                If Left(s_aResult(1), 1) = "*" Then
                    s_aResult(1) = "*"
                Else
                    s_aResult(1) = Format(Val(Format(s_aResult(1), "@@@.@@")), "0.00")
                End If
            End If

            'RBC
            s_aResult(2) = Mid(sRxData, 54, 4)

            If s_aResult(2) = Space(4) Then
                s_aResult(2) = "N"
            Else
                If Left(s_aResult(2), 1) = "*" Then
                    s_aResult(2) = "*"
                Else
                    s_aResult(2) = Format(Val(Format(s_aResult(2), "@@.@@")), "0.00")
                End If
            End If

            'HGB, HCT, MCV, MCH, MCHC
            For i = 3 To 7
                s_aResult(i) = Mid(sRxData, 59 + (i - 3) * 5, 4)

                If s_aResult(i) = Space(4) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@")), "0.0")
                    End If
                End If
            Next

            'PLT
            s_aResult(8) = Mid(sRxData, 84, 4)

            If s_aResult(8) = Space(4) Then
                s_aResult(8) = "N"
            Else
                If Left(s_aResult(8), 1) = "*" Then
                    s_aResult(8) = "*"
                Else
                    s_aResult(8) = Format(Val(Format(s_aResult(8), "@@@@")), "0")
                End If
            End If

            'LYMPH%, MONO%, NEUT%, EO%, BASO%
            For i = 9 To 13
                s_aResult(i) = Mid(sRxData, 89 + (i - 9) * 5, 4)

                If s_aResult(i) = Space(4) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@")), "0.0")
                    End If
                End If
            Next

            'LYMPH#, MONO#, NEUT#, EO#, BASO#
            For i = 14 To 18
                s_aResult(i) = Mid(sRxData, 114 + (i - 14) * 6, 5)

                If s_aResult(i) = Space(5) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@@")), "0.00")
                    End If
                End If
            Next

            'RDW-CV(%), RDW-SD(fL), PDW(fL), MPV(fL), P-LCR
            For i = 19 To 23
                s_aResult(i) = Mid(sRxData, 144 + (i - 19) * 5, 4)

                If s_aResult(i) = Space(4) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@")), "0.0")
                    End If
                End If
            Next

            'RET% ***** Manual과 Format이 다름, 결과가 틀림 -> Manual @@@.@(ex 12.9) vs 실제 @@.@@(1.29)
            s_aResult(24) = Mid(sRxData, 169, 4)

            If s_aResult(24) = Space(4) Then
                s_aResult(24) = "N"
            Else
                If Left(s_aResult(24), 1) = "*" Then
                    s_aResult(24) = "*"
                Else
                    s_aResult(24) = Format(Val(Format(s_aResult(24), "@@.@@")), "0.00")
                End If
            End If

            'RET#
            s_aResult(25) = Mid(sRxData, 174, 4)

            If s_aResult(25) = Space(4) Then
                s_aResult(25) = "N"
            Else
                If Left(s_aResult(25), 1) = "*" Then
                    s_aResult(25) = "*"
                Else
                    s_aResult(25) = Format(Val("." & Format(s_aResult(25), "@@@@")), "0.0000")
                End If
            End If

            'IRF, LFR, MFR, HFR
            For i = 26 To 29
                s_aResult(i) = Mid(sRxData, 179 + (i - 26) * 5, 4)

                If s_aResult(i) = Space(4) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@")), "0.0")
                    End If
                End If
            Next

            'PCT
            s_aResult(30) = Mid(sRxData, 199, 4)

            If s_aResult(30) = Space(4) Then
                s_aResult(30) = "N"
            Else
                If Left(s_aResult(30), 1) = "*" Then
                    s_aResult(30) = "*"
                Else
                    If Left(s_aResult(30), 1) = "*" Then
                        s_aResult(30) = "*"
                    Else
                        s_aResult(30) = Format(Val(Format(s_aResult(30), "@@.@@")), "0.00")
                    End If
                End If
            End If

            'NRBC%
            s_aResult(31) = Mid(sRxData, 204, 5)

            If s_aResult(31) = Space(5) Then
                s_aResult(31) = "N"
            Else
                If Left(s_aResult(31), 1) = "*" Then
                    s_aResult(31) = "*"
                Else
                    s_aResult(31) = Format(Val(Format(s_aResult(31), "@@@@.@")), "0.0")
                End If
            End If

            'NRBC#
            s_aResult(32) = Mid(sRxData, 210, 5)

            If s_aResult(32) = Space(5) Then
                s_aResult(32) = "N"
            Else
                If Left(s_aResult(32), 1) = "*" Then
                    s_aResult(32) = "*"
                Else
                    s_aResult(32) = Format(Val(Format(s_aResult(32), "@@@.@@")), "0.00")
                End If
            End If

            '실제결과
            iRealCnt = 0
            sTotIFCd = ""
            sTotRst = ""

            For i = 1 To 32
                If Trim(s_aResult(i)) = "N" Then
                Else
                    iRealCnt = iRealCnt + 1

                    sTotIFCd = sTotIFCd & CStr(i) & Chr(124)
                    sTotRst = sTotRst & Trim(s_aResult(i)) & Chr(124)
                End If
            Next

            '--- Flag Result ADD ---
            iRealCnt = iRealCnt + miFlagCnt
            sTotIFCd = sTotIFCd & msFlagCd
            sTotRst = sTotRst & msFlagTot
'            sTRst2 = sTRst2 & msFlagTot2

            '결과정보 구조체에 저장
            With pResultInfo
                .ID = sBarCd        ' pSampleInfo.ID        'msBarCd
                .SEQNO = sSeqNo     'pSampleInfo.SEQNO  'msSeqNo
                .RACK = pSampleInfo.RACK    'msRack
                .POS = pSampleInfo.POS      'msPos
                .RSTCNT = iRealCnt
                .IFCD = sTotIFCd
                .RST1 = sTotRst
                .RST2 = String(iRealCnt, Chr(124))
                .UNIT = String(iRealCnt, Chr(124))
                .FLAG = String(iRealCnt, Chr(124))
            End With

            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "")
                End If
            End With
            Call Init_pResultInfo

            miFlagCnt = 0: msFlagCd = "": msFlagTot = "": msFlagTot2 = ""

        Case Else
    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub



Private Sub SendOrder_XE2100()
    On Error GoTo ErrRtn

    Dim sOrder$
    Dim iPos%, i%
    Dim sOrdBuf$, sBuf$
    
    msOrdS2 = ""
    sOrdBuf = String$(35, "0")
    
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

    Call Get_OrderString
    
    If pSampleInfo.ORDCNT = 0 Then
        RaiseEvent DispMsg("인터페이스 오더 항목이 존재하지 않습니다!!")
        Exit Sub
    End If
    
    For i = 1 To pSampleInfo.ORDCNT
        sOrder = sOrder & Trim(pSampleInfo.IFCD(i))
    Next i

    sBuf = Chr$(2)
    sBuf = sBuf & "S11"
    sBuf = sBuf & Format(Now, "YYYYMMDD")
    sBuf = sBuf & Space$(3)
    sBuf = sBuf & Right(Space(15) & pSampleInfo.ID, 15)
    sBuf = sBuf & Space$(2)
    sBuf = sBuf & Format(pSampleInfo.RACK, "000000")
    sBuf = sBuf & Format(pSampleInfo.POS, "00")
    sBuf = sBuf & "1"
    '--- Patient No (16)
    sBuf = sBuf & Right(Space(16) & pSampleInfo.SEQNO, 16)
    sBuf = sBuf & Space$(40)
    '--- SEX (1)
'    If gOrderTable.sEmer = "F" Then
'        sBuf = sBuf & "2"
'    Else
'        sBuf = sBuf & "1"
'    End If
    sBuf = sBuf & "1"
    sBuf = sBuf & Space$(8)
    '--- Doctor (20)
'    sBuf = sBuf & Right(Space(20) & gOrderTable.sEmer, 20)
    sBuf = sBuf & Space(20)
    '--- Ward (20)
'    sBuf = sBuf & Right(Space(20) & gOrderTable.sSex, 20)
    sBuf = sBuf & Space(20)
    '--- Sample Comment (40)
'    sBuf = sBuf & Right(Space(40) & gOrderTable.sReRun, 40)
    sBuf = sBuf & Space(40)
    sBuf = sBuf & Space$(18)
    
    'Order 만들기
    If InStr(sOrder, "C") > 0 Then      'CBC
        For i = 1 To 8
            Mid(sOrdBuf, i, 1) = "1"
        Next i
        For i = 19 To 23
            Mid(sOrdBuf, i, 1) = "1"
        Next i
        Mid(sOrdBuf, 33, 1) = "1"
    End If
    
    If InStr(sOrder, "D") > 0 Then      'DIFF
        For i = 9 To 18
            Mid(sOrdBuf, i, 1) = "1"
        Next i
    End If
    
    If InStr(sOrder, "R") > 0 Then      'RETI
        For i = 26 To 31
            Mid(sOrdBuf, i, 1) = "1"
        Next i
    End If
    
    If InStr(sOrder, "N") > 0 Then      'NRBC
        Mid(sOrdBuf, 34, 1) = "1"
        Mid(sOrdBuf, 35, 1) = "1"
    End If
    
    sBuf = sBuf & sOrdBuf
    sBuf = sBuf & Space$(15)
    sBuf = sBuf & Chr(3)
    
    msOrdS2 = Chr$(2)
    msOrdS2 = msOrdS2 & "S21"
    msOrdS2 = msOrdS2 & Format(Now, "YYYYMMDD")
    msOrdS2 = msOrdS2 & Space$(3)
    msOrdS2 = msOrdS2 & Right(Space(15) & pSampleInfo.ID, 15)
    msOrdS2 = msOrdS2 & Space$(2)
    msOrdS2 = msOrdS2 & Format(pSampleInfo.RACK, "000000")
    msOrdS2 = msOrdS2 & Format(pSampleInfo.POS, "00")
    msOrdS2 = msOrdS2 & "1"
    msOrdS2 = msOrdS2 & Right(Space(16) & pSampleInfo.SEQNO, 16)
    msOrdS2 = msOrdS2 & Space$(100)
    msOrdS2 = msOrdS2 & Space$(97)
    msOrdS2 = msOrdS2 & Chr(3)
    
    msComm.Output = sBuf

    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(sBuf)
    End If
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub





Private Sub Get_OrderString()
    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Or m_p_sTIFCd = "" Then
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
    End With
End Sub

'결과정보 구조체 초기화
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

Private Sub DataEditResponse_XEAlpha()
    On Error GoTo ErrRtn
    
    Dim i%, iRealCnt%, ix1%
    Dim sRxData$, sBarCd$, sSeqNo$, sRack$, sPos$, sRstDT$
    Dim sBC$, sLC$
    Dim s_aResult$(32)
    
    Dim sTotIFCd$, sTotRst$
    Dim FlagBuf$
    
    sRxData = RcvBuffer

    sBC = Mid(sRxData, 1, 2)
    sLC = Mid(sRxData, 3, 1)

    Select Case sBC
        Case "D1"
            If sLC = "B" Or sLC = "E" Then Exit Sub
            
            miFlagCnt = 0: msFlagCd = "": msFlagTot = "": msFlagTot2 = ""

            If Len(sRxData) > 200 Then
                FlagBuf = ""
                RcvBuffer = Mid$(RcvBuffer, 105, 96)
            
                '실제 IP FLAG만 취득
                FlagBuf = Mid$(RcvBuffer, 1, 10)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 13, 2)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 17, 3)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 21, 1)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 24, 3)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 33, 10)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 49, 4)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 54, 1)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 65, 4)
                FlagBuf = FlagBuf & Mid$(RcvBuffer, 83, 1)
                
                'Manual 잘못으로인한 두가지 경우의 수를 적용(일본이 실수를?!?!)
                If Mid$(RcvBuffer, 84, 1) = "1" Or Mid$(RcvBuffer, 85, 1) = "1" Then
                    FlagBuf = FlagBuf & "1"
                Else
                    FlagBuf = FlagBuf & "0"
                End If
                
                For ix1 = 1 To Len(FlagBuf)
                    If Mid$(FlagBuf, ix1, 1) = "1" Then
                        miFlagCnt = miFlagCnt + 1
                        msFlagCd = msFlagCd & Trim(Str(ix1 + 100 - 1)) & "|"
                        msFlagTot = msFlagTot & "Detected!" & "|"
                        msFlagTot2 = msFlagTot2 & "|"
                    End If
                Next ix1
            End If
            
            sSeqNo = Trim(Mid(sRxData, 20, 10))
            sBarCd = Trim(Mid(sRxData, 33, 15))
            sRstDT = Trim(Mid(sRxData, 48, 12))     '2005/6/23 Add yk
            sRack = Trim(Mid(sRxData, 62, 6))
            sPos = Trim(Mid(sRxData, 68, 2))

            '2004/10/14 Bug Fix
'            If Left(sBarCd, 3) = "ERR" Then Exit Sub

            With pSampleInfo
                .ID = sBarCd
                .SEQNO = sSeqNo
                .RACK = sRack
                .POS = sPos
                .RSTDT = sRstDT
                
                If Left(sBarCd, 3) = "ERR" Then     '2004/10/14 Update yk
                Else
                    RaiseEvent RequestCurOrder(.ID, .RACK, .POS)
                End If
            End With

            Exit Sub
            
        Case "D2"
            If sLC = "B" Or sLC = "E" Then Exit Sub
            
            Call Init_pResultInfo
            
            If Len(sRxData) > 253 Then
                RaiseEvent DispMsg("XE2100으로부터 전송된 문자열의 길이 (" & CStr(Len(sRxData)) & ")의 이상이 발생하였습니다!!")
                Exit Sub
            End If
            
            sSeqNo = Trim(Mid(sRxData, 20, 10))
'            If pSampleInfo.SEQNO <> "" And pSampleInfo.SEQNO <> sSeqNo Then      ' msSeqNo <> sSeqNo Then
'                RaiseEvent DispMsg("D1과 D2에서 다른 장비일련번호가 전송되었습니다!!")
'                miFlagCnt = 0: msFlagCd = "": msFlagTot = "": msFlagTot2 = ""
'                Exit Sub
'            End If
            
            sBarCd = Trim(Mid(sRxData, 33, 15))
'            If pSampleInfo.ID <> "" And pSampleInfo.ID <> sBarCd Then        ' msBarCd <> sBarCd Then
'                RaiseEvent DispMsg("D1과 D2에서 다른 바코드 정보가 전송되었습니다!!")
'                miFlagCnt = 0: msFlagCd = "": msFlagTot = "": msFlagTot2 = ""
'                Exit Sub
'            End If
            
            'WBC
            s_aResult(1) = Mid(sRxData, 48, 5)
            
            If s_aResult(1) = Space(5) Then
                s_aResult(1) = "N"
            Else
                If Left(s_aResult(1), 1) = "*" Then
                    s_aResult(1) = "*"
                Else
                    s_aResult(1) = Format(Val(Format(s_aResult(1), "@@@.@@")), "0.00")
                End If
            End If
            
            'RBC
            s_aResult(2) = Mid(sRxData, 54, 4)
            
            If s_aResult(2) = Space(4) Then
                s_aResult(2) = "N"
            Else
                If Left(s_aResult(2), 1) = "*" Then
                    s_aResult(2) = "*"
                Else
                    s_aResult(2) = Format(Val(Format(s_aResult(2), "@@.@@")), "0.00")
                End If
            End If
            
            'HGB, HCT, MCV, MCH, MCHC
            For i = 3 To 7
                s_aResult(i) = Mid(sRxData, 59 + (i - 3) * 5, 4)
                                  
                If s_aResult(i) = Space(4) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@")), "0.0")
                    End If
                End If
            Next
            
            'PLT
            s_aResult(8) = Mid(sRxData, 84, 4)
                        
            If s_aResult(8) = Space(4) Then
                s_aResult(8) = "N"
            Else
                If Left(s_aResult(8), 1) = "*" Then
                    s_aResult(8) = "*"
                Else
                    s_aResult(8) = Format(Val(Format(s_aResult(8), "@@@@")), "0")
                End If
            End If
            
            'LYMPH%, MONO%, NEUT%, EO%, BASO%
            For i = 9 To 13
                s_aResult(i) = Mid(sRxData, 89 + (i - 9) * 5, 4)
                                  
                If s_aResult(i) = Space(4) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@")), "0.0")
                    End If
                End If
            Next
            
            'LYMPH#, MONO#, NEUT#, EO#, BASO#
            For i = 14 To 18
                s_aResult(i) = Mid(sRxData, 114 + (i - 14) * 6, 5)
                                  
                If s_aResult(i) = Space(5) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@@")), "0.00")
                    End If
                End If
            Next
            
            'RDW-CV(%), RDW-SD(fL), PDW(fL), MPV(fL), P-LCR
            For i = 19 To 23
                s_aResult(i) = Mid(sRxData, 144 + (i - 19) * 5, 4)
                                  
                If s_aResult(i) = Space(4) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@")), "0.0")
                    End If
                End If
            Next
            
            'RET% ***** Manual과 Format이 다름, 결과가 틀림 -> Manual @@@.@(ex 12.9) vs 실제 @@.@@(1.29)
            s_aResult(24) = Mid(sRxData, 169, 4)
                        
            If s_aResult(24) = Space(4) Then
                s_aResult(24) = "N"
            Else
                If Left(s_aResult(24), 1) = "*" Then
                    s_aResult(24) = "*"
                Else
                    s_aResult(24) = Format(Val(Format(s_aResult(24), "@@.@@")), "0.00")
                End If
            End If
            
            'RET#
            s_aResult(25) = Mid(sRxData, 174, 4)
                        
            If s_aResult(25) = Space(4) Then
                s_aResult(25) = "N"
            Else
                If Left(s_aResult(25), 1) = "*" Then
                    s_aResult(25) = "*"
                Else
                    s_aResult(25) = Format(Val("." & Format(s_aResult(25), "@@@@")), "0.0000")
                End If
            End If
            
            'IRF, LFR, MFR, HFR
            For i = 26 To 29
                s_aResult(i) = Mid(sRxData, 179 + (i - 26) * 5, 4)
                                  
                If s_aResult(i) = Space(4) Then
                    s_aResult(i) = "N"
                Else
                    If Left(s_aResult(i), 1) = "*" Then
                        s_aResult(i) = "*"
                    Else
                        s_aResult(i) = Format(Val(Format(s_aResult(i), "@@@.@")), "0.0")
                    End If
                End If
            Next
            
            'PCT
            s_aResult(30) = Mid(sRxData, 199, 4)
            
            If s_aResult(30) = Space(4) Then
                s_aResult(30) = "N"
            Else
                If Left(s_aResult(30), 1) = "*" Then
                    s_aResult(30) = "*"
                Else
                    If Left(s_aResult(30), 1) = "*" Then
                        s_aResult(30) = "*"
                    Else
                        s_aResult(30) = Format(Val(Format(s_aResult(30), "@@.@@")), "0.00")
                    End If
                End If
            End If
            
            'NRBC%
            s_aResult(31) = Mid(sRxData, 204, 5)
                                  
            If s_aResult(31) = Space(5) Then
                s_aResult(31) = "N"
            Else
                If Left(s_aResult(31), 1) = "*" Then
                    s_aResult(31) = "*"
                Else
                    s_aResult(31) = Format(Val(Format(s_aResult(31), "@@@@.@")), "0.0")
                End If
            End If
            
            'NRBC#
            s_aResult(32) = Mid(sRxData, 210, 5)
                                  
            If s_aResult(32) = Space(5) Then
                s_aResult(32) = "N"
            Else
                If Left(s_aResult(32), 1) = "*" Then
                    s_aResult(32) = "*"
                Else
                    s_aResult(32) = Format(Val(Format(s_aResult(32), "@@@.@@")), "0.00")
                End If
            End If
            
            '실제결과
            iRealCnt = 0
            sTotIFCd = ""
            sTotRst = ""
            
            For i = 1 To 32
                If Trim(s_aResult(i)) = "N" Then
                Else
                    iRealCnt = iRealCnt + 1
                    
                    sTotIFCd = sTotIFCd & CStr(i) & Chr(124)
                    sTotRst = sTotRst & Trim(s_aResult(i)) & Chr(124)
                End If
            Next
            
            '--- Flag Result ADD ---
            iRealCnt = iRealCnt + miFlagCnt
            sTotIFCd = sTotIFCd & msFlagCd
            sTotRst = sTotRst & msFlagTot
'            sTRst2 = sTRst2 & msFlagTot2
        
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = sBarCd        ' pSampleInfo.ID        'msBarCd
                .SEQNO = sSeqNo     'pSampleInfo.SEQNO  'msSeqNo
                .RACK = pSampleInfo.RACK     'msRack
                .POS = pSampleInfo.POS      'msPos
                .RSTCNT = iRealCnt
                .IFCD = sTotIFCd
                .RST1 = sTotRst
                .RST2 = String(iRealCnt, Chr(124))
                .UNIT = String(iRealCnt, Chr(124))
                .FLAG = String(iRealCnt, Chr(124))
                '결과시간 추가...2005/6/23 yk
                For i = 1 To iRealCnt
                    .RSTDT = .RSTDT & pSampleInfo.RSTDT & Chr(124)
                Next i
            End With
            
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .RSTDT, "")
                End If
            End With
            Call Init_pResultInfo
            
            miFlagCnt = 0: msFlagCd = "": msFlagTot = "": msFlagTot2 = ""
            
        Case Else
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_XEAlpha()
    Dim sWkDat$
    Dim i%
    
    For i = 1 To Len(wkBuf)
        sWkDat = Mid(wkBuf, i, 1)
        
        Select Case m_iPhase
            Case 1
                Select Case Asc(sWkDat)
                    Case 2
                        RcvBuffer = ""
                        
                    Case 3
                        Call DataEditResponse_XEAlpha
                        msComm.Output = Chr(6)
                        
                    Case Else
                        RcvBuffer = RcvBuffer & sWkDat
                End Select
            Case Else
        End Select
    Next
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
    m_SiteNm = PropBag.ReadProperty("SiteNm", m_def_SiteNm)
    m_p_sPatInfo = PropBag.ReadProperty("p_sPatInfo", m_def_p_sPatInfo)
    m_p_sSampInfo = PropBag.ReadProperty("p_sSampInfo", m_def_p_sSampInfo)
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
    Call PropBag.WriteProperty("SiteNm", m_SiteNm, m_def_SiteNm)
    Call PropBag.WriteProperty("p_sPatInfo", m_p_sPatInfo, m_def_p_sPatInfo)
    Call PropBag.WriteProperty("p_sSampInfo", m_p_sSampInfo, m_def_p_sSampInfo)
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
    m_SiteNm = m_def_SiteNm
    m_p_sPatInfo = m_def_p_sPatInfo
    m_p_sSampInfo = m_def_p_sSampInfo
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
Public Property Get SiteNm() As Variant
    SiteNm = m_SiteNm
End Property

Public Property Let SiteNm(ByVal New_SiteNm As Variant)
    m_SiteNm = New_SiteNm
    PropertyChanged "SiteNm"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sPatInfo() As String
    p_sPatInfo = m_p_sPatInfo
End Property

Public Property Let p_sPatInfo(ByVal New_p_sPatInfo As String)
    m_p_sPatInfo = New_p_sPatInfo
    PropertyChanged "p_sPatInfo"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sSampInfo() As String
    p_sSampInfo = m_p_sSampInfo
End Property

Public Property Let p_sSampInfo(ByVal New_p_sSampInfo As String)
    m_p_sSampInfo = New_p_sSampInfo
    PropertyChanged "p_sSampInfo"
End Property
