VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl RAPIDLAB 
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
      Width           =   1485
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
Attribute VB_Name = "RAPIDLAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_iTotalItemCnt = 0
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sInstID$, sKind$, sTRstDT$, sOther1$)
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sKind$, sTRstDT$, sOther1$)
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

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

'RapidLab 800 Series Interface Message Character
Dim ENQ$
Dim STX$
Dim ETX$
Dim EOT$
Dim LF$
Dim Cr$
Dim RS$
Dim FS$
Dim GS$
Dim ACK$
Dim NAK$
Dim FS2$
Dim AckOn   As Boolean
Dim Sample_Seq  As String
Dim aMod    As String
Dim iIID    As String
Private Sub DispData_Link(ByVal MsgBuf As String)
    On Error GoTo ErrRtn

    Dim C   As Integer
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim Result  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String

    Dim iACC    As String: iACC = ""    '2005/11/10 yk
    Dim iOID    As String: iOID = ""    '2006/01/09 yk

    Dim sRstDate$, sRstTime$
    Dim sInstID$

    Dim siFIO2  As String: siFIO2 = ""  '2006/5/29 yk
    

    '결과구조체 초기화
    Call Init_pResultInfo

    'aMod
    x1 = 1
    x1 = InStr(x1, MsgBuf, "aMod") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        aMod = Mid(MsgBuf, x1, x2 - x1)
    End If

    'iIID
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iIID") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        iIID = Mid(MsgBuf, x1, x2 - x1)
    End If

    'rSEQ
    x1 = 1
    x1 = InStr(x1, MsgBuf, "rSEQ") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        rSeq = Mid(MsgBuf, x1, x2 - x1)
    End If

    'PID
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iPID") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        iPID = Mid(MsgBuf, x1, x2 - x1)
    End If

    x1 = 1
    x1 = InStr(x1, MsgBuf, "rDATE") + 6
    If x1 <> 6 Then
        x2 = InStr(x1, MsgBuf, GS)
        sRstDate = Mid(MsgBuf, x1, x2 - x1)
        sRstDate = ConvertDateType(sRstDate)
    End If

    x1 = 1
    x1 = InStr(x1, MsgBuf, "rTIME") + 6
    If x1 <> 6 Then
        x2 = InStr(x1, MsgBuf, GS)
        sRstTime = Mid(MsgBuf, x1, x2 - x1)
        sRstTime = Format(sRstTime, "HHNNSS")
    End If

    'ACC
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iACC") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        iACC = Mid(MsgBuf, x1, x2 - x1)
    End If

    'OID
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iOID") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        iOID = Mid(MsgBuf, x1, x2 - x1)
    End If

    'System
    x1 = 1
    x1 = InStr(x1, MsgBuf, "rSYSTEM") + 8
    If x1 <> 8 Then
        x2 = InStr(x1, MsgBuf, GS)
        sInstID = Mid(MsgBuf, x1, x2 - x1)
    End If
    
    'iFIO2...2006/5/29 yk
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iFIO2") + 6
    If x1 <> 6 Then
        x2 = InStr(x1, MsgBuf, GS)
        siFIO2 = Mid(MsgBuf, x1, x2 - x1)
    End If
    
    
    x2 = 0

    '접수번호, SeqNo
    pResultInfo.ID = Trim(iPID)
    pResultInfo.SEQNO = Trim(rSeq)
    pResultInfo.OTHER = Trim(iACC)
    '2006/2/9 yk
    pResultInfo.INSTID = Trim(sInstID)
    If Trim(iOID) <> "" Then
        pResultInfo.INSTID = pResultInfo.INSTID & Chr(124) & Trim(iOID)
    End If

    '-----------
    '   Measured Data
    '-----------
    x1 = 1
    Do While InStr(x1, MsgBuf, FS & "m") <> 0
        x1 = InStr(x1, MsgBuf, FS & "m")
        x2 = InStr(x1, MsgBuf, GS)

'        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
        'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
        AssayNm = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))

        x2 = x2 + 1
        x1 = InStr(x2, MsgBuf, GS)
        Result = Mid(MsgBuf, x2, x1 - x2)

        'Data 누적 편집
        With pResultInfo
            .RSTCNT = .RSTCNT + 1
            .IFCD = .IFCD & AssayNm & Chr(124)
            .RST1 = .RST1 & Result & Chr(124)
            .RST2 = .RST2 & Chr(124)
            .UNIT = .UNIT & Chr(124)
            .FLAG = .FLAG & Chr(124)
            .RSTDT = .RSTDT & sRstDate & sRstTime & Chr(124)
        End With
    Loop

    '-----------
    '   Calibrated Data
    '-----------
    x1 = 1
    Do While InStr(x1, MsgBuf, FS & "c") <> 0
        x1 = InStr(x1, MsgBuf, FS & "c")
        x2 = InStr(x1, MsgBuf, GS)

'        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
        'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
        AssayNm = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))

        x2 = x2 + 1
        x1 = InStr(x2, MsgBuf, GS)
        Result = Mid(MsgBuf, x2, x1 - x2)

        'Data 누적 편집
        With pResultInfo
            .RSTCNT = .RSTCNT + 1
            .IFCD = .IFCD & AssayNm & Chr(124)
            .RST1 = .RST1 & Result & Chr(124)
            .RST2 = .RST2 & Chr(124)
            .UNIT = .UNIT & Chr(124)
            .FLAG = .FLAG & Chr(124)
            .RSTDT = .RSTDT & sRstDate & sRstTime & Chr(124)
        End With
    Loop

    'FIO2 존재할 경우 항목에 추가...2006/5/29 yk
    If siFIO2 <> "" Then
        With pResultInfo
            .RSTCNT = .RSTCNT + 1
            .IFCD = .IFCD & "iFIO2" & Chr(124)
            .RST1 = .RST1 & siFIO2 & Chr(124)
            .RST2 = .RST2 & Chr(124)
            .UNIT = .UNIT & Chr(124)
            .FLAG = .FLAG & Chr(124)
            .RSTDT = .RSTDT & sRstDate & sRstTime & Chr(124)
        End With
    End If

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, "", .RSTDT, .OTHER)
        End If
    End With

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DispData 에러발생(" & Err.Description & ")")
    End If
End Sub
Private Sub TEMP()

    'Message #5:
    '<STX>PATIENT_DATA_AV<FS><RS>aMOD<GS>DMS<GS>
    '<GS><GS><FS>iIID<GS>COMP2<GS><GS><GS><FS>rSEQ
    '<GS>104<GS><GS><GS><FS><RS><ETX>46<EOT>
    
    'Message #6:
    '<STX>PATIENT_DATA_REQ<FS><RS>aMOD<GS>DMS<GS>
    '<GS><GS><FS>iIID<GS>COMP2<GS><GS><GS><FS>rSEQ
    '<GS>104<GS><GS><GS><FS><RS><ETX>97<EOT>



    'Message #7:
    '<STX>PATIENT_DATA<FS><RS>aMOD<GS>DMS<GS><GS>
    '<GS><FS>iIID<GS>COMP2<GS><GS><GS><FS>rSEQ<GS>
    '104<GS><GS><GS><FS>rDATE<GS>06Feb1997<GS><GS><GS>
    '<FS>rTIME<GS>20:23:00<GS><GS><GS><FS>rSYSTEM<GS>
    '840-ICU<GS><GS><GS><FS>iPID<GS>879562<GS><GS><GS>
    '<FS>iSEX<GS>MALE<GS><GS><GS><FS>iDOB<GS>12May1934
    '<GS><GS><GS><FS>iACC<GS>110345<GS><GS><GS><FS>
    'iROOM<GS>SURGICAL_ICU<GS><GS><GS><FS>iDID<GS>
    '0473<GS><GS><GS><FS>iSOURCE<GS>ARTERIAL<GS><GS>
    '<GS><FS>iOID<GS>4456<GS><GS><GS><FS>mpH<GS>7.433
    '<GS><GS><GS><FS>mPCO2<GS>38.3<GS>mmHg<GS><GS>
    '<FS>mPO2<GS>89.9<GS>mmHg<GS><GS><FS>mBP<GS>761
    '<GS>mmHg<GS><GS><FS>cHCO3act<GS>25.0<GS>mmol/L<GS>
    '<GS><FS>cHCO3std<GS>22.9<GS>mmol/L<GS><GS><FS>
    'cBE (vt) < GS > 1# < GS > mmol / L < GS <> GS <> FS > cBE(vv) < GS > 0.8
    '<GS>mmol/L<GS><GS><FS>ctCO2<GS>26.2<GS>mmol/L<GS>
    '<GS><FS>cO2SAT<GS>97.1<GS>%<GS><GS><FS>cO2(CT)<GS>
    '20.5<GS>mL/dL<GS><GS><FS>cpH(T)<GS>7.408<GS><GS>
    '<GS><FS>cPCO2(T)<GS>41.3<GS>mmHg<GS><GS><FS>cPO2(T)
    '<GS>100.1<GS>mmHg<GS><GS><FS>iTEMP<GS>38.7<GS>C
    '<GS><GS><FS><RS><ETX>9F<EOT>

End Sub
Private Sub DispData_340(ByVal MsgBuf As String)
    On Error GoTo ErrRtn

    Dim C   As Integer
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim Result  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String

    Dim sRstDate$, sRstTime$
    

    '결과구조체 초기화
    Call Init_pResultInfo

    'aMod
    x1 = 1
    x1 = InStr(x1, MsgBuf, "aMOD") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        aMod = Mid(MsgBuf, x1, x2 - x1)
    End If

    'iIID
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iIID") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        iIID = Mid(MsgBuf, x1, x2 - x1)
    End If

    'rSEQ
    x1 = 1
    x1 = InStr(x1, MsgBuf, "rSEQ") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        rSeq = Mid(MsgBuf, x1, x2 - x1)
    End If

    'PID
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iPID") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        iPID = Mid(MsgBuf, x1, x2 - x1)
    End If

    x1 = 1
    x1 = InStr(x1, MsgBuf, "rDATE") + 6
    If x1 <> 6 Then
        x2 = InStr(x1, MsgBuf, GS)
        sRstDate = Mid(MsgBuf, x1, x2 - x1)
        sRstDate = ConvertDateType(sRstDate)
    End If
    
    x1 = 1
    x1 = InStr(x1, MsgBuf, "rTIME") + 6
    If x1 <> 6 Then
        x2 = InStr(x1, MsgBuf, GS)
        sRstTime = Mid(MsgBuf, x1, x2 - x1)
        sRstTime = Format(sRstTime, "HHNNSS")
    End If
    

    x2 = 0

    '접수번호, SeqNo
    pResultInfo.ID = Trim(iPID)
    pResultInfo.SEQNO = Trim(rSeq)


    '-----------
    '   Measured Data
    '-----------
    x1 = 1
    Do While InStr(x1, MsgBuf, FS & "m") <> 0
        x1 = InStr(x1, MsgBuf, FS & "m")
        x2 = InStr(x1, MsgBuf, GS)

'        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
        'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
        AssayNm = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))

        x2 = x2 + 1
        x1 = InStr(x2, MsgBuf, GS)
        Result = Mid(MsgBuf, x2, x1 - x2)

        'Data 누적 편집
        With pResultInfo
            .RSTCNT = .RSTCNT + 1
            .IFCD = .IFCD & AssayNm & Chr(124)
            .RST1 = .RST1 & Result & Chr(124)
            .RST2 = .RST2 & Chr(124)
            .UNIT = .UNIT & Chr(124)
            .FLAG = .FLAG & Chr(124)
            .RSTDT = .RSTDT & sRstDate & sRstTime & Chr(124)
        End With
    Loop

    '-----------
    '   Calibrated Data
    '-----------
    x1 = 1
    Do While InStr(x1, MsgBuf, FS & "c") <> 0
        x1 = InStr(x1, MsgBuf, FS & "c")
        x2 = InStr(x1, MsgBuf, GS)

'        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
        'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
        AssayNm = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))

        x2 = x2 + 1
        x1 = InStr(x2, MsgBuf, GS)
        Result = Mid(MsgBuf, x2, x1 - x2)

        'Data 누적 편집
        With pResultInfo
            .RSTCNT = .RSTCNT + 1
            .IFCD = .IFCD & AssayNm & Chr(124)
            .RST1 = .RST1 & Result & Chr(124)
            .RST2 = .RST2 & Chr(124)
            .UNIT = .UNIT & Chr(124)
            .FLAG = .FLAG & Chr(124)
            .RSTDT = .RSTDT & sRstDate & sRstTime & Chr(124)
        End With
    Loop

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", .RSTDT, "")
        End If
    End With

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DispData 에러발생(" & Err.Description & ")")
    End If
End Sub


Private Function ConvertDateType(ByVal sDate As String) As String
    On Error GoTo ErrRtn
    
    Dim kk%
    Dim sTmp$
    Dim tmpYYYY$, tmpMM$, tmpDD$
    
    ConvertDateType = sDate
    
    tmpYYYY = Right(sDate, 4)
    sDate = Mid(sDate, 1, Len(sDate) - 4)
    
    For kk = 1 To Len(sDate)
        sTmp = Mid(sDate, kk, 1)
        If IsNumeric(sTmp) Then
            tmpDD = tmpDD & sTmp
        Else
            tmpMM = tmpMM & sTmp
        End If
    Next kk
    
    sTmp = tmpDD & Space(1) & tmpMM & Space(1) & tmpYYYY
    
    ConvertDateType = Format(sTmp, "YYYYMMDD")
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("ConvertDateType - " & Err.Description)
    End If
End Function

Private Sub GetaModiIID(ByVal sMsg As String)

    Dim tmpData()   As String
    
    '<STX>SYS_READY<FS><RS>aMOD<GS>1265<GS><GS><GS><FS>iIID
    '<GS>12345<GS><GS><GS><FS>aDATE<GS>20Jan2004<GS><GS><GS>
    '<FS>aTIME<GS>13:35:32<GS><GS><GS><FS>iOID<GS>3<GS><GS><GS><FS>
    '<ETX>{chksum}<EOT>

    tmpData() = Split(sMsg, GS)
    
    'aMod
    aMod = Trim(tmpData(1))
    
    'iIID
    iIID = Trim(tmpData(5))

End Sub

Private Sub SendMessage_Link(ByVal MsgHead As String)
    On Error GoTo SendMessage_Error
    
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
    Dim R As Integer
    Dim Tmp     As String
    Dim OrdVal  As String
    Dim OrdNm   As Variant

    Dim sSendData$
    
    Select Case MsgHead
        Case "ID_DATA"
            'ID_DATAaMODDMSiIIDCOMP248
            Buffer = STX & "ID_DATA" & FS & RS _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & RS _
                                    & ETX
                                    
        Case "PATIENT_DATA_REQ"
            Buffer = STX & "PATIENT_DATA_REQ" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & RS & ETX
                                        
'        Case "SMP_REQ"
'            Buffer = STX & "SMP_REQ" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
'                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
'                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
'                                        & FS & RS & ETX
            
        Case "SMP_ORD"
'            With spdList
'                For c = 1 To 3 Step 2
'                    For r = 1 To .MaxRows
'
'                        .Col = c: .Row = r
'                        OrdVal = CStr(.Value Xor 1)
'                        OrdNm = .TypeCheckText
'                        If InStr(1, OrdNm, "Hb") Then OrdNm = "tHb"
'                        If OrdNm <> "" Then
'                            Tmp = Tmp & FS & "sDIS" & CStr(OrdNm) & GS & OrdVal & GS & GS & GS
'                        End If
'                    Next
'                Next
'            End With
'
'            Buffer = STX & "SMP_ORD" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
'                                            & FS & "iIID" & GS & iIID & GS & GS & GS _
'                                            & FS & "iPID" & GS & mskLabNo.ClipText & GS & GS & GS _
'                                            & Tmp _
'                                            & FS & ETX
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    sSendData = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
    msComm.Output = sSendData
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendData)
    End If
    
SendMessage_Error:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendMessage Error : " & Err.Description)
    End If
End Sub
Private Sub SendMessage_340(ByVal MsgHead As String)
    On Error GoTo SendMessage_Error
    
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
    Dim R As Integer
    Dim Tmp     As String
    Dim OrdVal  As String
    Dim OrdNm   As Variant

    Select Case MsgHead
        Case "ID_DATA"
            Buffer = STX & "ID_DATA" & FS & RS _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & RS _
                                    & ETX
        Case "SMP_REQ"
            Buffer = STX & "SMP_REQ" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & RS & ETX
            
        Case "SMP_ORD"
'            With spdList
'                For c = 1 To 3 Step 2
'                    For r = 1 To .MaxRows
'
'                        .Col = c: .Row = r
'                        OrdVal = CStr(.Value Xor 1)
'                        OrdNm = .TypeCheckText
'                        If InStr(1, OrdNm, "Hb") Then OrdNm = "tHb"
'                        If OrdNm <> "" Then
'                            Tmp = Tmp & FS & "sDIS" & CStr(OrdNm) & GS & OrdVal & GS & GS & GS
'                        End If
'                    Next
'                Next
'            End With
'
'            Buffer = STX & "SMP_ORD" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
'                                            & FS & "iIID" & GS & iIID & GS & GS & GS _
'                                            & FS & "iPID" & GS & mskLabNo.ClipText & GS & GS & GS _
'                                            & Tmp _
'                                            & FS & ETX
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    msComm.Output = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
SendMessage_Error:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendMessage Error : " & Err.Description)
    End If
End Sub

Private Sub SendMessage(ByVal MsgHead As String)
    On Error GoTo SendMessage_Error
    
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
    Dim R As Integer
    Dim Tmp     As String
    Dim OrdVal  As String
    Dim OrdNm   As Variant

    Select Case MsgHead
        Case "ID_DATA"
            Buffer = STX & "ID_DATA" & FS & RS _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & RS _
                                    & ETX
        Case "SMP_REQ"
            Buffer = STX & "SMP_REQ" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & RS & ETX
            
        Case "SMP_ORD"
'            With spdList
'                For c = 1 To 3 Step 2
'                    For r = 1 To .MaxRows
'
'                        .Col = c: .Row = r
'                        OrdVal = CStr(.Value Xor 1)
'                        OrdNm = .TypeCheckText
'                        If InStr(1, OrdNm, "Hb") Then OrdNm = "tHb"
'                        If OrdNm <> "" Then
'                            Tmp = Tmp & FS & "sDIS" & CStr(OrdNm) & GS & OrdVal & GS & GS & GS
'                        End If
'                    Next
'                Next
'            End With
'
'            Buffer = STX & "SMP_ORD" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
'                                            & FS & "iIID" & GS & iIID & GS & GS & GS _
'                                            & FS & "iPID" & GS & mskLabNo.ClipText & GS & GS & GS _
'                                            & Tmp _
'                                            & FS & ETX
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    msComm.Output = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
SendMessage_Error:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendMessage Error : " & Err.Description)
    End If
End Sub

Private Sub DispData(ByVal MsgBuf As String)
    On Error GoTo ErrRtn

    Dim C   As Integer
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim Result  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String

    Dim sRstDate$, sRstTime$
    

    '결과구조체 초기화
    Call Init_pResultInfo

    'aMod
    x1 = 1
    x1 = InStr(x1, MsgBuf, "aMod") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        aMod = Mid(MsgBuf, x1, x2 - x1)
    End If

    'iIID
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iIID") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        iIID = Mid(MsgBuf, x1, x2 - x1)
    End If

    'rSEQ
    x1 = 1
    x1 = InStr(x1, MsgBuf, "rSEQ") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        rSeq = Mid(MsgBuf, x1, x2 - x1)
    End If

    'PID
    x1 = 1
    x1 = InStr(x1, MsgBuf, "iPID") + 5
    If x1 <> 5 Then
        x2 = InStr(x1, MsgBuf, GS)
        iPID = Mid(MsgBuf, x1, x2 - x1)
    End If

    x1 = 1
    x1 = InStr(x1, MsgBuf, "rDATE") + 6
    If x1 <> 6 Then
        x2 = InStr(x1, MsgBuf, GS)
        sRstDate = Mid(MsgBuf, x1, x2 - x1)
        sRstDate = ConvertDateType(sRstDate)
    End If
    
    x1 = 1
    x1 = InStr(x1, MsgBuf, "rTIME") + 6
    If x1 <> 6 Then
        x2 = InStr(x1, MsgBuf, GS)
        sRstTime = Mid(MsgBuf, x1, x2 - x1)
        sRstTime = Format(sRstTime, "HHNNSS")
    End If
    

    x2 = 0

    '접수번호, SeqNo
    pResultInfo.ID = Trim(iPID)
    pResultInfo.SEQNO = Trim(rSeq)


    '-----------
    '   Measured Data
    '-----------
    x1 = 1
    Do While InStr(x1, MsgBuf, FS & "m") <> 0
        x1 = InStr(x1, MsgBuf, FS & "m")
        x2 = InStr(x1, MsgBuf, GS)

'        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
        'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
        AssayNm = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))

        x2 = x2 + 1
        x1 = InStr(x2, MsgBuf, GS)
        Result = Mid(MsgBuf, x2, x1 - x2)

        'Data 누적 편집
        With pResultInfo
            .RSTCNT = .RSTCNT + 1
            .IFCD = .IFCD & AssayNm & Chr(124)
            .RST1 = .RST1 & Result & Chr(124)
            .RST2 = .RST2 & Chr(124)
            .UNIT = .UNIT & Chr(124)
            .FLAG = .FLAG & Chr(124)
            .RSTDT = .RSTDT & sRstDate & sRstTime & Chr(124)
        End With
    Loop

    '-----------
    '   Calibrated Data
    '-----------
    x1 = 1
    Do While InStr(x1, MsgBuf, FS & "c") <> 0
        x1 = InStr(x1, MsgBuf, FS & "c")
        x2 = InStr(x1, MsgBuf, GS)

'        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
        'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
        AssayNm = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))

        x2 = x2 + 1
        x1 = InStr(x2, MsgBuf, GS)
        Result = Mid(MsgBuf, x2, x1 - x2)

        'Data 누적 편집
        With pResultInfo
            .RSTCNT = .RSTCNT + 1
            .IFCD = .IFCD & AssayNm & Chr(124)
            .RST1 = .RST1 & Result & Chr(124)
            .RST2 = .RST2 & Chr(124)
            .UNIT = .UNIT & Chr(124)
            .FLAG = .FLAG & Chr(124)
            .RSTDT = .RSTDT & sRstDate & sRstTime & Chr(124)
        End With
    Loop

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", .RSTDT, "")
        End If
    End With

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DispData 에러발생(" & Err.Description & ")")
    End If
End Sub



Private Sub INIT_CHARACTER()

    ENQ = Chr$(5)
    STX = Chr$(2)
    ETX = Chr$(3)
    EOT = Chr$(4)
    LF = Chr$(10)
    Cr = Chr$(13)
    RS = Chr$(30)
    FS = Chr$(28)
    GS = Chr$(29)
    ACK = Chr$(6)
    NAK = Chr$(21)
    FS2 = Chr$(124)

End Sub
Private Sub PhaseCfg_Protocol_Rapidlab1200()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2  'STX
                RcvBuffer = wkDat
                AckOn = False
                
            Case 4  'EOT
                If AckOn = False Then
                    msComm.Output = STX & ACK & ETX & "0B" & EOT        'Ack Message
                    
                    Select Case UCase(m_EqName)
                        Case "RAPIDLAB1200"
                            Call Analysis_Message_1200
                        Case "RAPIDLINK"
                            Call Analysis_Message_Link
                    End Select
                End If
                
            Case 6  'ACK
                AckOn = True
                RcvBuffer = RcvBuffer & wkDat
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_Rapidlab340()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2  'STX
                RcvBuffer = wkDat
                AckOn = False
                
            Case 4  'EOT
                If AckOn = False Then
                    msComm.Output = STX & ACK & ETX & "0B" & EOT        'Ack Message
                    Call Analysis_Message_340
                End If
                
            Case 6  'ACK
                AckOn = True
                RcvBuffer = RcvBuffer & wkDat
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_Rapidlab800()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2  'STX
                RcvBuffer = wkDat
                AckOn = False
                
            Case 4  'EOT
                If AckOn = False Then
                    msComm.Output = STX & ACK & ETX & "0B" & EOT        'Ack Message
                    Call Analysis_Message
                End If
                
            Case 6  'ACK
                AckOn = True
                RcvBuffer = RcvBuffer & wkDat
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub
Private Sub Analysis_Message_Link()
    On Error GoTo Analysis_Error
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    
    X = InStr(1, RcvBuffer, FS)
    If RcvBuffer <> "" Then MsgID = Mid(RcvBuffer, 2, X - 2)
    
    Select Case MsgID
        Case "ID_REQ"
            Call SendMessage_Link("ID_DATA")
            RaiseEvent DispMsg("검사를 시작할 준비가 되었습니다.")
            
        Case "SMP_START"
        
        Case "PATIENT_DATA_AV"
            Do Until X = 0
                X = InStr(X, RcvBuffer, "r")
                If X = 0 Then Exit Do
                If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                    X = X + 5
                    C = InStr(X, RcvBuffer, GS)
                    Sample_Seq = Mid(RcvBuffer, X, C - X)
                End If
                
                '2005-07-23 KHS Modified
                Call GetaModiIID(RcvBuffer)
                Call SendMessage_Link("PATIENT_DATA_REQ")
            Loop
            
'        Case "SMP_NEW_AV"
'            Do Until X = 0
'                X = InStr(X, RcvBuffer, "r")
'                If X = 0 Then Exit Do
'                If Mid(RcvBuffer, X, 4) = "rSEQ" Then
'                    X = X + 5
'                    C = InStr(X, RcvBuffer, GS)
'                    Sample_Seq = Mid(RcvBuffer, X, C - X)
'                End If
'
'                '2005-07-23 KHS Modified
'                Call GetaModiIID(RcvBuffer)
'                Call SendMessage_Link("SMP_REQ")
'            Loop
        
        Case "SYS_READY"
                '2005-07-23 KHS Modified
'''            Call GetaModiIID(RcvBuffer)
'''            Call SendMessage_1200("SMP_REQ")
        
        Case "SYS_NOT_READY"
        
        '2005-07-23 KHS Modified
        Case "PATIENT_DATA", "SMP_NEW_DATA", "SMP_EDIT_DATA"
            Call DispData_Link(RcvBuffer)
            
        Case "CAL_ABORT"
        
    End Select
    
Analysis_Error:
    If Err <> 0 Then
        RaiseEvent DispMsg("Analysis_Message ERROR : " & Err.Description)
    End If
End Sub
Private Sub SendMessage_1200(ByVal MsgHead As String)
    On Error GoTo SendMessage_Error
    
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
    Dim R As Integer
    Dim Tmp     As String
    Dim OrdVal  As String
    Dim OrdNm   As Variant

    Dim sSendData$
    
    Select Case MsgHead
        Case "ID_DATA"
            Buffer = STX & "ID_DATA" & FS & RS _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & RS _
                                    & ETX
        Case "SMP_REQ"
            Buffer = STX & "SMP_REQ" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & RS & ETX
            
        Case "SMP_ORD"
'            With spdList
'                For c = 1 To 3 Step 2
'                    For r = 1 To .MaxRows
'
'                        .Col = c: .Row = r
'                        OrdVal = CStr(.Value Xor 1)
'                        OrdNm = .TypeCheckText
'                        If InStr(1, OrdNm, "Hb") Then OrdNm = "tHb"
'                        If OrdNm <> "" Then
'                            Tmp = Tmp & FS & "sDIS" & CStr(OrdNm) & GS & OrdVal & GS & GS & GS
'                        End If
'                    Next
'                Next
'            End With
'
'            Buffer = STX & "SMP_ORD" & FS & RS & "aMOD" & GS & aMod & GS & GS & GS _
'                                            & FS & "iIID" & GS & iIID & GS & GS & GS _
'                                            & FS & "iPID" & GS & mskLabNo.ClipText & GS & GS & GS _
'                                            & Tmp _
'                                            & FS & ETX
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    sSendData = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
    msComm.Output = sSendData
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendData)
    End If
    
SendMessage_Error:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendMessage Error : " & Err.Description)
    End If
End Sub

Private Sub Analysis_Message_340()
    On Error GoTo Analysis_Error
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    
    X = InStr(1, RcvBuffer, FS)
    If RcvBuffer <> "" Then MsgID = Mid(RcvBuffer, 2, X - 2)
    
    Select Case MsgID
        Case "ID_REQ"
            Call SendMessage_340("ID_DATA")
            RaiseEvent DispMsg("검사를 시작할 준비가 되었습니다.")
            
        Case "SMP_START"
        
        Case "SMP_NEW_AV"
            Do Until X = 0
                X = InStr(X, RcvBuffer, "r")
                If X = 0 Then Exit Do
                If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                    X = X + 5
                    C = InStr(X, RcvBuffer, GS)
                    Sample_Seq = Mid(RcvBuffer, X, C - X)
                End If
            Loop
        
        Case "SYS_READY"
            Call SendMessage_340("SMP_REQ")
        
        Case "SYS_NOT_READY"
        
        Case "SMP_NEW_DATA", "SMP_EDIT_DATA"
            Call DispData_340(RcvBuffer)
            
        Case "CAL_ABORT"
        
    End Select
    
Analysis_Error:
    If Err <> 0 Then
        RaiseEvent DispMsg("Analysis_Message ERROR : " & Err.Description)
    End If
End Sub
Private Sub Analysis_Message()
    On Error GoTo Analysis_Error
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    
    X = InStr(1, RcvBuffer, FS)
    If RcvBuffer <> "" Then MsgID = Mid(RcvBuffer, 2, X - 2)
    
    Select Case MsgID
        Case "ID_REQ"
            Call SendMessage("ID_DATA")
            RaiseEvent DispMsg("검사를 시작할 준비가 되었습니다.")
            
        Case "SMP_START"
        
        Case "SMP_NEW_AV"
            Do Until X = 0
                X = InStr(X, RcvBuffer, "r")
                If X = 0 Then Exit Do
                If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                    X = X + 5
                    C = InStr(X, RcvBuffer, GS)
                    Sample_Seq = Mid(RcvBuffer, X, C - X)
                End If
            Loop
        
        Case "SYS_READY"
            Call SendMessage("SMP_REQ")
        
        Case "SYS_NOT_READY"
        
        Case "SMP_NEW_DATA", "SMP_EDIT_DATA"
            Call DispData(RcvBuffer)
            
        Case "CAL_ABORT"
        
    End Select
    
Analysis_Error:
    If Err <> 0 Then
        RaiseEvent DispMsg("Analysis_Message ERROR : " & Err.Description)
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
        Case "RAPIDLAB840", "RAPIDLAB850"
            Call PhaseCfg_Protocol_Rapidlab800
        
        Case "RAPIDLAB1200", "RAPIDLINK"
            Call PhaseCfg_Protocol_Rapidlab1200
        
        Case "RAPIDLAB348"
            Call PhaseCfg_Protocol_Rapidlab340
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub Analysis_Message_1200()
    On Error GoTo Analysis_Error
    
    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    
    X = InStr(1, RcvBuffer, FS)
    If RcvBuffer <> "" Then MsgID = Mid(RcvBuffer, 2, X - 2)
    
    Select Case MsgID
        Case "ID_REQ"
            Call SendMessage_1200("ID_DATA")
            RaiseEvent DispMsg("검사를 시작할 준비가 되었습니다.")
            
        Case "SMP_START"
        
        Case "SMP_NEW_AV"
            Do Until X = 0
                X = InStr(X, RcvBuffer, "r")
                If X = 0 Then Exit Do
                If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                    X = X + 5
                    C = InStr(X, RcvBuffer, GS)
                    Sample_Seq = Mid(RcvBuffer, X, C - X)
                End If
                
                '2005-07-23 KHS Modified
                Call GetaModiIID(RcvBuffer)
                Call SendMessage_1200("SMP_REQ")
            Loop
        
        Case "SYS_READY"
                '2005-07-23 KHS Modified
'''            Call GetaModiIID(RcvBuffer)
'''            Call SendMessage_1200("SMP_REQ")
        
        Case "SYS_NOT_READY"
        
        '2005-07-23 KHS Modified
        Case "SMP_NEW_DATA", "SMP_EDIT_DATA"
            Call DispData(RcvBuffer)
            
        Case "CAL_ABORT"
        
    End Select
    
Analysis_Error:
    If Err <> 0 Then
        RaiseEvent DispMsg("Analysis_Message ERROR : " & Err.Description)
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
        .INSTID = ""
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
    m_iTotalItemCnt = PropBag.ReadProperty("iTotalItemCnt", m_def_iTotalItemCnt)
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
    
    'For Rapidlab
    Call INIT_CHARACTER
    
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
Public Property Get iTotalItemCnt() As Integer
    iTotalItemCnt = m_iTotalItemCnt
End Property

Public Property Let iTotalItemCnt(ByVal New_iTotalItemCnt As Integer)
    m_iTotalItemCnt = New_iTotalItemCnt
    PropertyChanged "iTotalItemCnt"
End Property

