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
Event RequestCurOrder(sID$, sRack$, sPos$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
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

Private Function ConvertDataAlarmCode(ByVal sEqNm As String, ByVal sCode As String) As String
'
'    Dim sTmp    As String
'
'    ConvertDataAlarmCode = "": sTmp = ""
'
'    Select Case UCase(sEqNm)
'        Case "HITACHI7600"
'            Select Case Trim(sCode)
'                Case "0": sTmp = ""
'                Case "1": sTmp = "ADC?"
'                Case "2": sTmp = "Cell?"
'                Case "3": sTmp = "Sampl"
'                Case "4": sTmp = "Reagn"
'                Case "5": sTmp = "ABS?"
'                Case "6": sTmp = "Prozon"
'                Case "7": sTmp = "Limt0"
'                Case "8": sTmp = "Limt1"
'                Case "9": sTmp = "Limt2"
'                Case "10": sTmp = "Lin."
'                Case "11": sTmp = "Lin8."
'                Case "12": sTmp = "S1Abs?"
'                Case "13": sTmp = "Dup"
'                Case "14": sTmp = "Std?"
'                Case "15": sTmp = "Sens"
'                Case "16": sTmp = "Calib"
'                Case "17": sTmp = "SDI"
'                Case "18": sTmp = "Noise"
'                Case "19": sTmp = "Level"
'                Case "20": sTmp = "Slope?"
'                Case "21": sTmp = "Margin"
'                Case "22": sTmp = "I.Std"
'                Case "23": sTmp = "R.Over"
'                Case "24": sTmp = "Cmp.T"
'                Case "25": sTmp = "Cmp.TI"
'                Case "26": sTmp = "LIMTH"
'                Case "27": sTmp = "LIMTL"
'                Case "28": sTmp = "Random"
'                Case "29": sTmp = "Systm1"
'                Case "30": sTmp = "Systm2"
'                Case "31": sTmp = "Systm3"
'                Case "32": sTmp = "Systm4"
'                Case "33": sTmp = "Systm5"
'                Case "34": sTmp = "Systm6"
'                Case "35": sTmp = "QCErr1"
'                Case "36": sTmp = "QCErr2"
'                Case "37": sTmp = "Calc?"
'                Case "38": sTmp = "Over"
'                Case "39": sTmp = "???"
'                Case "42": sTmp = "Edited"
'                Case "44": sTmp = "ReptH"
'                Case "45": sTmp = "ReptL"
'                Case "51": sTmp = "Resp1"
'                Case "52": sTmp = "Resp2"
'                Case "53": sTmp = "Condi"
'            End Select
'
'        Case Else
'
'    End Select
'
'    ConvertDataAlarmCode = Trim(sTmp)
'
End Function
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
        Case "TBA80FR"
            Call PhaseCfg_Protocol_TBA80FR
            
        Case "TBA120FR"
            Call PhaseCfg_Protocol_TBA120FR
            
        Case "TBA200FR"
            Call PhaseCfg_Protocol_TBA200FR
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_TBA80FR()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX 수신
                m_iPhase = 2
                RcvBuffer = ""

            Case 3      'ETX 수신
                Call DataEditResponse_TBA80FR

            Case Else   '문자 수신
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1

End Sub

Private Sub PhaseCfg_Protocol_TBA120FR()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX 수신
                m_iPhase = 2
                RcvBuffer = ""

            Case 3      'ETX 수신
                Call DataEditResponse_TBA120FR

            Case Else   '문자 수신
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1

End Sub
Private Sub PhaseCfg_Protocol_TBA200FR()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX 수신
                m_iPhase = 2
                RcvBuffer = ""

            Case 3      'ETX 수신
                Call DataEditResponse_TBA200FR

            Case Else   '문자 수신
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1

End Sub

Private Sub SendOrder_TBA80FR()
    On Error GoTo Err_Rtn

    Dim sBuf As String
    Dim i%, j%, k%, iOrdCnt%
    Dim sTIFOrdCd$
    Dim sTmp$, sTestBuf$, sOrdList$, sIFSeq$, sTIFSeq$
    Dim sSendBuff$

    Dim sSend   As String
    Dim sID     As String
    Dim sDate   As String

    Dim aRerun()    As String
    
    '----- 검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.ID, "", "")

    Call Get_OrderString

    'Order Packet 구성
    sSendBuff = ""

    'Order Send
    sSend = Chr(2) & "Y C  " & pSampleInfo.ID & Space(14 - Len(pSampleInfo.ID)) & Space(6) & Space(1)
    sSend = sSend & String(2 - Len(pSampleInfo.RACK), " ") & pSampleInfo.RACK & "/"
    sSend = sSend & String(4 - Len(pSampleInfo.POS), " ") & pSampleInfo.POS & "/0" & Space(1) & "  1" & Space(1) & "1"
    
    '검사항목 관련 Packet
    sTestBuf = ""
    For i = 1 To pSampleInfo.ORDCNT
''        '재검방법에 따라 재검 오더 편집...2005/6/29 yk
''        If InStr(pSampleInfo.IFCD(i), "^") > 0 Then
''            aRerun() = Split(pSampleInfo.IFCD(i), "^")
''            pSampleInfo.IFCD(i) = Trim(aRerun(0))
''
''            Select Case Trim(aRerun(1))
''                Case "1"        '표준 샘플량
''                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "1"
''                Case "2"        '재검 샘플량 1
''                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "2"
''                Case "3"        '재검 샘플량 2
''                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "3"
''                Case Else
''                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "1"
''            End Select
''        Else
            '표준 샘플량
            sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & Space(1)
''        End If
    Next i
    
    sSend = sSend & sTestBuf & vbCr & vbLf
    sDate = Format(Now, "YYYYMMDD")
    sSend = sSend & Mid(sDate, 3, 2) & "/" & Mid(sDate, 5, 2) & "/" & Mid(sDate, 7, 2) & Space(1) & Space(5) & Space(1)
    sSend = sSend & Space(30) & Space(1) & Space(30) & Chr(3)
    
    msComm.Output = sSend

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    
    If pSampleInfo.ORDCNT > 0 Then
        RaiseEvent SendOrderOK(pSampleInfo.ID, "", "", "")
    Else
        RaiseEvent SendOrderOK("", "", "", "")
    End If
    
Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub DataEditResponse_TBA120FR()
    On Error GoTo ErrRtn

    Dim sType   As String
    Dim sBufCnt As String
    Dim ii      As Integer
    Dim iAllCnt As Integer
    Dim tmpIFCd$, tmpRst$

    Dim sBarCd$, sDiskID$, sDiskPos$, sQCLevel$, sQCDt$
    Dim iPosQC%

    sType = Mid$(RcvBuffer, 1, 1)

    Select Case sType
        Case Chr(6)         'ACK

        Case Chr(21)        'NAK

        Case "Q"
            'Q 12345678901234       40/ 121/0 ABCD
            'Q 0001000000000201740001408003 3
            'Q 000120165000030            048
            sBarCd = Trim(Mid(RcvBuffer, 7, 20))
            sDiskID = Trim(Mid(RcvBuffer, 27, 4))
            sDiskPos = Trim(Mid(RcvBuffer, 31, 3))

'            If Mid(sBarCd, 1, 1) = "0" Then
'                sBarCd = Mid(sBarCd, 2)
'            End If

            pSampleInfo.ID = sBarCd
            pSampleInfo.RACK = sDiskID
            pSampleInfo.POS = sDiskPos

            Call SendOrder_TBA120FR

        Case "M"
            msComm.Output = Chr(2) & Chr(6) & Chr(3)

        Case "R"
'            iPosQC = InStr(RcvBuffer, Chr(23))
'            sQCLevel = ""
'
'            If iPosQC > 0 Then
'                sQCLevel = Trim(Mid(RcvBuffer, iPosQC + 13, 30))
'            End If

            msComm.Output = Chr(2) & Chr(6) & Chr(3)

            '결과 구조체 초기화
            Call Init_pResultInfo

            'Packet 편집
            With pResultInfo
                .SEQNO = Trim(Mid(RcvBuffer, 3, 4))
                .ID = Trim(Mid(RcvBuffer, 7, 20))
                .RACK = Trim(Mid(RcvBuffer, 27, 4))
                .POS = Trim(Mid(RcvBuffer, 31, 3))
                sQCDt = Trim(Mid(RcvBuffer, 33, 12))
                
'                If Trim(sQCLevel) <> "" Then
'                    .ID = sQCLevel
'                    .KIND = "QC"
'                Else
                    If Mid(.ID, 1, 1) = "0" Then
                        .ID = Mid(.ID, 2)
                    End If
                    .KIND = ""
'                End If
            End With

            sBufCnt = NoTrimGetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(23))
            sBufCnt = Mid(sBufCnt, 49)

            iAllCnt = Len(sBufCnt) / 15

            If Len(sBufCnt) Mod 15 = 0 Then
                For ii = 1 To iAllCnt
                    tmpIFCd = Trim(Mid(sBufCnt, (ii - 1) * 15 + 1, 4))
                    tmpRst = Trim(Mid(sBufCnt, (ii - 1) * 15 + 5, 6))

                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                        .RSTDT = .RSTDT & sQCDt & Chr(124)
                    End With
                Next ii

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, .RSTDT, "")
                    End If
                End With
            Else
                RaiseEvent DispMsg(pResultInfo.ID & "결과 길이 이상... 재전송해주십시요.")
            End If

            Call Init_pResultInfo
            
        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Sub SendOrder_TBA200FR()
    On Error GoTo Err_Rtn

    Dim sBuf As String
    Dim i%, j%, k%, iOrdCnt%
    Dim sTIFOrdCd$
    Dim sTmp$, sTestBuf$, sOrdList$, sIFSeq$, sTIFSeq$
    Dim sSendBuff$

    Dim sSend   As String
    Dim sID     As String

    Dim aRerun()    As String
    
    '----- 검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.ID, "", "")

    Call Get_OrderString

    'Order Packet 구성
    sSendBuff = ""

    'Order Send
    sSend = Chr(2) & "O " & pSampleInfo.ID & String(20 - Len(pSampleInfo.ID), " ")
    sSend = sSend & String(4 - Len(pSampleInfo.RACK), " ") & pSampleInfo.RACK
    sSend = sSend & String(2 - Len(pSampleInfo.POS), " ") & pSampleInfo.POS
    sSend = sSend & "  1"       '수작업 희석 배율치
    
    '검사항목 관련 Packet
    sTestBuf = ""
    For i = 1 To pSampleInfo.ORDCNT
        '재검방법에 따라 재검 오더 편집...2005/6/29 yk
        If InStr(pSampleInfo.IFCD(i), "^") > 0 Then
            aRerun() = Split(pSampleInfo.IFCD(i), "^")
            pSampleInfo.IFCD(i) = Trim(aRerun(0))
            
            Select Case Trim(aRerun(1))
                Case "1"        '표준 샘플량
                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "1"
                Case "2"        '재검 샘플량 1
                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "2"
                Case "3"        '재검 샘플량 2
                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "3"
                Case Else
                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "1"
            End Select
        Else
            '표준 샘플량
            sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "1"
        End If
    Next i
    sSend = sSend & sTestBuf & Chr(23)
    
    sSend = sSend & Format(Now, "yyyyMMddhhmm")
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
    
    If pSampleInfo.ORDCNT > 0 Then
        RaiseEvent SendOrderOK(pSampleInfo.ID, "", "", "")
    Else
        RaiseEvent SendOrderOK("", "", "", "")
    End If
    
Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub DataEditResponse_TBA80FR()
    On Error GoTo ErrRtn

    Dim sType   As String
    Dim sBufCnt As String
    Dim ii      As Integer
    Dim iAllCnt As Integer
    Dim tmpIFCd$, tmpRst$

    Dim sBarCd$, sDiskID$, sDiskPos$, sQCLevel$, sQCDt$
    Dim iPosQC%
    
    Dim sBuf$

    sType = Mid$(RcvBuffer, 1, 1)

    Select Case sType
        Case Chr(6)         'ACK

        Case Chr(21)        'NAK

        Case "Q"
            'Q 12345678901234       40/ 121/0 ABCD
            'Q 0001000000000201740001408003 3
            'Q 000120165000030            048
            sBarCd = Trim(Mid(RcvBuffer, 3, 14))
            sDiskID = Trim(Mid(RcvBuffer, 24, 2))
            sDiskPos = Trim(Mid(RcvBuffer, 27, 4))

'            If Mid(sBarCd, 1, 1) = "0" Then
'                sBarCd = Mid(sBarCd, 2)
'            End If

            pSampleInfo.ID = sBarCd
            pSampleInfo.RACK = sDiskID
            pSampleInfo.POS = sDiskPos

            Call SendOrder_TBA80FR

        Case "R"
            iPosQC = InStr(RcvBuffer, Chr(23))
            sQCLevel = ""

            If iPosQC > 0 Then
                sQCLevel = Trim(Mid(RcvBuffer, iPosQC + 13, 30))
            End If

            msComm.Output = Chr(2) & Chr(6) & Chr(3)

            '결과 구조체 초기화
            Call Init_pResultInfo

            'Packet 편집
            With pResultInfo
                '.SEQNO = Trim(Mid(RcvBuffer, 3, 4))
                .ID = Trim(Mid(RcvBuffer, 3, 14))
                .RACK = Trim(Mid(RcvBuffer, 24, 2))
                .POS = Trim(Mid(RcvBuffer, 27, 4))
                sQCDt = Trim(Mid(RcvBuffer, 39, 8))
                
                If Trim(sQCLevel) <> "" Then
                    .ID = sQCLevel
                    .KIND = "QC"
                Else
                    If Mid(.ID, 1, 1) = "0" Then
                        .ID = Mid(.ID, 2)
                    End If
                    .KIND = ""
                End If
            End With

''            sBufCnt = NoTrimGetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(23))
''            sBufCnt = Mid(sBufCnt, 48)
            
            Call GetByOneUserSymbol(RcvBuffer, RcvBuffer, vbLf)
            
            '{TestCode(4)+TestResult(6)+ResultFlag(3)}*n
            sBuf = GetByOneUserSymbol(RcvBuffer, RcvBuffer, vbCr)
            
            sBuf = Replace(sBuf, Chr(1), " ")
            
            iAllCnt = Len(sBuf) \ 13

            If Len(sBufCnt) Mod 13 = 0 Then
                For ii = 1 To iAllCnt
                    tmpIFCd = Trim(Mid(sBuf, (ii - 1) * 13 + 1, 4))
                    tmpRst = Trim(Mid(sBuf, (ii - 1) * 13 + 5, 6))

                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                        .RSTDT = .RSTDT & sQCDt & Chr(124)
                    End With
                Next ii

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, .RSTDT, "")
                    End If
                End With
            Else
                RaiseEvent DispMsg(pResultInfo.ID & "결과 길이 이상... 재전송해주십시요.")
            End If

            Call Init_pResultInfo
            
        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEditResponse_TBA200FR()
    On Error GoTo ErrRtn

    Dim sType   As String
    Dim sBufCnt As String
    Dim ii      As Integer
    Dim iAllCnt As Integer
    Dim tmpIFCd$, tmpRst$

    Dim sBarCd$, sDiskID$, sDiskPos$, sQCLevel$, sQCDt$
    Dim iPosQC%

    sType = Mid$(RcvBuffer, 1, 1)

    Select Case sType
        Case Chr(6)         'ACK

        Case Chr(21)        'NAK

        Case "Q"
            'Q 12345678901234       40/ 121/0 ABCD
            'Q 0001000000000201740001408003 3
            'Q 000120165000030            048
            sBarCd = Trim(Mid(RcvBuffer, 7, 20))
            sDiskID = Trim(Mid(RcvBuffer, 27, 4))
            sDiskPos = Trim(Mid(RcvBuffer, 31, 2))

'            If Mid(sBarCd, 1, 1) = "0" Then
'                sBarCd = Mid(sBarCd, 2)
'            End If

            pSampleInfo.ID = sBarCd
            pSampleInfo.RACK = sDiskID
            pSampleInfo.POS = sDiskPos

            Call SendOrder_TBA200FR

        Case "M"
            msComm.Output = Chr(2) & Chr(6) & Chr(3)

        Case "R"
            iPosQC = InStr(RcvBuffer, Chr(23))
            sQCLevel = ""

            If iPosQC > 0 Then
                sQCLevel = Trim(Mid(RcvBuffer, iPosQC + 13, 30))
            End If

            msComm.Output = Chr(2) & Chr(6) & Chr(3)

            '결과 구조체 초기화
            Call Init_pResultInfo

            'Packet 편집
            With pResultInfo
                .SEQNO = Trim(Mid(RcvBuffer, 3, 4))
                .ID = Trim(Mid(RcvBuffer, 7, 20))
                .RACK = Trim(Mid(RcvBuffer, 27, 4))
                .POS = Trim(Mid(RcvBuffer, 31, 2))
                sQCDt = Trim(Mid(RcvBuffer, 33, 12))
                
                If Trim(sQCLevel) <> "" Then
                    .ID = sQCLevel
                    .KIND = "QC"
                Else
                    If Mid(.ID, 1, 1) = "0" Then
                        .ID = Mid(.ID, 2)
                    End If
                    .KIND = ""
                End If
            End With

            sBufCnt = NoTrimGetByOneUserSymbol(RcvBuffer, RcvBuffer, Chr(23))
            sBufCnt = Mid(sBufCnt, 48)

            iAllCnt = Len(sBufCnt) / 13

            If Len(sBufCnt) Mod 13 = 0 Then
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
                        .RSTDT = .RSTDT & sQCDt & Chr(124)
                    End With
                Next ii

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, .RSTDT, "")
                    End If
                End With
            Else
                RaiseEvent DispMsg(pResultInfo.ID & "결과 길이 이상... 재전송해주십시요.")
            End If

            Call Init_pResultInfo
            
        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_TBA120FR()
    On Error GoTo Err_Rtn

    Dim sBuf As String
    Dim i%, j%, k%, iOrdCnt%
    Dim sTIFOrdCd$
    Dim sTmp$, sTestBuf$, sOrdList$, sIFSeq$, sTIFSeq$
    Dim sSendBuff$

    Dim sSend   As String
    Dim sID     As String

    Dim aRerun()    As String
    
    '----- 검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

    Call Get_OrderString

    'Order Packet 구성
    sSendBuff = ""

    'Order Send
    sSend = Chr(2) & "O " & pSampleInfo.ID & String(20 - Len(pSampleInfo.ID), " ")
    sSend = sSend & String(4 - Len(pSampleInfo.RACK), " ") & pSampleInfo.RACK
    sSend = sSend & String(3 - Len(pSampleInfo.POS), " ") & pSampleInfo.POS
    sSend = sSend & "  1"       '수작업 희석 배율치
    
    '검사항목 관련 Packet
    sTestBuf = ""
    For i = 1 To pSampleInfo.ORDCNT
        '재검방법에 따라 재검 오더 편집...2005/6/29 yk
        If InStr(pSampleInfo.IFCD(i), "^") > 0 Then
            aRerun() = Split(pSampleInfo.IFCD(i), "^")
            pSampleInfo.IFCD(i) = Trim(aRerun(0))
            
            Select Case Trim(aRerun(1))
                Case "1"        '표준 샘플량
                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "1"
                Case "2"        '재검 샘플량 1
                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "2"
                Case "3"        '재검 샘플량 2
                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "3"
                Case Else
                    sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "1"
            End Select
        Else
            '표준 샘플량
            sTestBuf = sTestBuf & String(4 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & "1"
        End If
    Next i
    sSend = sSend & sTestBuf & Chr(23)
    
'    sSend = sSend & Format(Now, "yyyyMMddhhmm")
'    sSend = sSend & Space$(42)      'NAME
'    sSend = sSend & Space$(20)      'PID
'    sSend = sSend & Space$(1)       'SEX
'    sSend = sSend & Space(8)        'BIRTHDAY
'    sSend = sSend & Space(20)       'LOCATION
'    sSend = sSend & Space(20)       'DOCTOR
'    sSend = sSend & Space(20)       'COMMENT
    sSend = sSend & Chr(23) & Chr(3)

    msComm.Output = sSend

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    
    If pSampleInfo.ORDCNT > 0 Then
        RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    Else
        RaiseEvent SendOrderOK("", "", "", "")
    End If
    
Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

Public Function GetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim POS%

    POS = InStr(tStr, sUserSymbol)

    If POS = 0 Then
    Else
        GetByOneUserSymbol = Trim$(Mid$(tStr, 1, POS - 1))
        sOriginal = Trim$(Mid$(sOriginal, POS + 1, Len(sOriginal) - POS))
    End If
End Function

Private Function NoTrimGetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim POS%

    POS = InStr(tStr, sUserSymbol)

    If POS = 0 Then
    Else
        NoTrimGetByOneUserSymbol = Mid$(tStr, 1, POS - 1)
        sOriginal = Mid$(sOriginal, POS + 1, Len(sOriginal) - POS)
    End If
End Function
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
        .KIND = ""
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

Private Sub Get_OrderString()
        
    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .SEQNO = m_p_sSeq
'            .RACK = m_p_sRack
'            .POS = m_p_sPos
            .ORDCNT = 0
            .KIND = m_p_sRerunGbn
        End With
    
        Exit Sub
    End If
    
    ReDim tmpData(m_p_iOrdCnt) As String
    tmpData() = Split(m_p_sTIFCd, Chr(124))
    
    With pSampleInfo
        .ID = m_p_sID
        .SEQNO = m_p_sSeq
'        .RACK = m_p_sRack
'        .POS = m_p_sPos
        .ORDCNT = m_p_iOrdCnt
        .KIND = m_p_sRerunGbn
        
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
            
            If m_sTestMode = "77" Then
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
    m_p_bSIndex = PropBag.ReadProperty("p_bSIndex", m_def_p_bSIndex)
    m_p_sRerunGbn = PropBag.ReadProperty("p_sRerunGbn", m_def_p_sRerunGbn)
    m_p_sTSVol = PropBag.ReadProperty("p_sTSVol", m_def_p_sTSVol)
    m_p_sSpcCd = PropBag.ReadProperty("p_sSpcCd", m_def_p_sSpcCd)
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

