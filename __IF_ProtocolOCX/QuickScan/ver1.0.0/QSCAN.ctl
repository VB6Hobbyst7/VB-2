VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl QSCAN 
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
Attribute VB_Name = "QSCAN"
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

'토탈, 환자번호, 환자이름, 병원명, 의뢰과/병동, 나이, 성별, 검체명
Const m_def_p_Total = "0"
Const m_def_p_RegNo = "0"
Const m_def_p_PatName = "0"
Const m_def_p_HosName = "0"
Const m_def_p_DepNm = "0"
Const m_def_p_Age = "0"
Const m_def_p_Sex = "0"
Const m_def_p_SpcNo = "0"

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

Dim m_p_Total As String
Dim m_p_RegNo As String
Dim m_p_PatName As String
Dim m_p_HosName As String
Dim m_p_DepNm As String
Dim m_p_Age As String
Dim m_p_Sex As String
Dim m_p_SpcNo As String

'이벤트 선언:
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

'For QuickScan
Dim m_sQuery  As String
Dim m_iField  As Integer
Dim m_sMPeak As String
Dim m_iMPeak1     As Integer
Dim m_iMPeak2     As Integer

Private Type TYPE_Result
    strBarno    As String
    strSeqno    As String
    strExcde    As String
    strChtno    As String
    strPicrt    As String
    strSldrt    As String
    strFracd    As String
    strFrart    As String
    strTotal    As String
End Type
Private f_typRst    As TYPE_Result

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
    
    Dim sTmp    As String
    
    ConvertDataAlarmCode = "": sTmp = ""
    
    Select Case UCase(sEqNm)
        Case "HITACHI7600"
            Select Case Trim(sCode)
                Case "0": sTmp = ""
                Case "1": sTmp = "ADC?"
                Case "2": sTmp = "Cell?"
                Case "3": sTmp = "Sampl"
                Case "4": sTmp = "Reagn"
                Case "5": sTmp = "ABS?"
                Case "6": sTmp = "Prozon"
                Case "7": sTmp = "Limt0"
                Case "8": sTmp = "Limt1"
                Case "9": sTmp = "Limt2"
                Case "10": sTmp = "Lin."
                Case "11": sTmp = "Lin8."
                Case "12": sTmp = "S1Abs?"
                Case "13": sTmp = "Dup"
                Case "14": sTmp = "Std?"
                Case "15": sTmp = "Sens"
                Case "16": sTmp = "Calib"
                Case "17": sTmp = "SDI"
                Case "18": sTmp = "Noise"
                Case "19": sTmp = "Level"
                Case "20": sTmp = "Slope?"
                Case "21": sTmp = "Margin"
                Case "22": sTmp = "I.Std"
                Case "23": sTmp = "R.Over"
                Case "24": sTmp = "Cmp.T"
                Case "25": sTmp = "Cmp.TI"
                Case "26": sTmp = "LIMTH"
                Case "27": sTmp = "LIMTL"
                Case "28": sTmp = "Random"
                Case "29": sTmp = "Systm1"
                Case "30": sTmp = "Systm2"
                Case "31": sTmp = "Systm3"
                Case "32": sTmp = "Systm4"
                Case "33": sTmp = "Systm5"
                Case "34": sTmp = "Systm6"
                Case "35": sTmp = "QCErr1"
                Case "36": sTmp = "QCErr2"
                Case "37": sTmp = "Calc?"
                Case "38": sTmp = "Over"
                Case "39": sTmp = "???"
                Case "42": sTmp = "Edited"
                Case "44": sTmp = "ReptH"
                Case "45": sTmp = "ReptL"
                Case "51": sTmp = "Resp1"
                Case "52": sTmp = "Resp2"
                Case "53": sTmp = "Condi"
            End Select
            
        Case Else
        
    End Select
    
    ConvertDataAlarmCode = Trim(sTmp)
    
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
        Case "QUICKSCAN"
            If m_bUseBarcode = True Then
                'Call PhaseCfg_Protocol_QuickScan    '바코드사용
            Else
                Call PhaseCfg_Protocol_QuickScan_Batch    'Batch Mode
            End If
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_QuickScan_Batch()
    On Error GoTo ErrRtn
    
    Dim wkdat As String
    Dim ix1 As Long

    For ix1 = 1 To Len(wkBuf)
        
        wkdat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase

            Case 1
                Select Case Asc(wkdat)
                    Case 5          'ENQ
                        bEndChk = True: bSTXChk = False
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                    Case 6
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkdat)
                    Case 2          'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True
                        
                    'Case 3         'ETX
                    Case 10         'LF
                        If bEndChk = True Then
                            Call DataEditResponse_QuickScan
                            RcvBuffer = ""
                        End If
                        
                        msComm.Output = Chr(6)
                    Case 4          'EOT
                        If m_sQuery = "Q" Then
                            m_iPhase = 3
                            m_iFrameN = 0
                            m_iField = 0
                            
                            Call SendOrder_QuickScan_Batch
                        Else
                            m_iPhase = 1
                        End If
                    Case 5          'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                    Case 21         'NAK
                        msComm.Output = Chr(5)
                        m_iPhase = 1
                    Case 23         'ETB
                        bEndChk = False
                        msComm.Output = Chr(6)
                    
                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkdat
                            End If
                        End If
                End Select

            Case 3
                Select Case Asc(wkdat)
                    Case 3  '-- EOT
                        'Call SendOrder_QuickScan_Batch
                            
                    Case 6          'ACK
'                        Call Order_Input
                        
                        If m_iField = 2 Then
                            'txtSeq = Trim(CStr(Val(txtSeq)))
                            Call SendOrder_QuickScan_Batch
                            
                            ''m_iField = m_iField + 1
                            
                        ElseIf m_iField = -1 Then
                            m_iField = 0
                            m_iPhase = 1
                            msComm.Output = Chr(4)

''                            If miLogMode = 1 Then
''                                Print #3, Chr(4);
''                            End If
                            
                        Else
                            m_iField = m_iField + 1

                            Call SendOrder_QuickScan_Batch
                        End If
                        
                    Case 5          'ENQ
                        bEndChk = True: bSTXChk = False
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                    Case 21         'NAK
                        msComm.Output = Chr(5)
                        m_iPhase = 1
'                    Case 4          'EOT
'                        m_iPhase = 1
                End Select
        End Select
        
    Next ix1

    Exit Sub
    
ErrRtn:
    MsgBox "PhaseCfg_Protocol : " & Err.Description

End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_QuickScan()
        
    On Error GoTo ErrRtn
    
    Dim strDta()    As String
    
    If InStr(RcvBuffer, "|") > 0 Then
        RcvBuffer = Mid$(RcvBuffer, InStr(RcvBuffer, "|") - 1)
    End If
    
    Select Case Mid$(RcvBuffer, 1, 1)
        Case "H"
        Case "P"
            m_sMPeak = ""
            m_iMPeak1 = 0:    m_iMPeak2 = 0
            strDta() = Split(RcvBuffer, "|")
            
            pResultInfo.ID = strDta(3)
            
''            With f_typRst
''                .strBarno = strDta(3)
''                .strChtno = ""
''                .strSeqno = ""
''                .strExcde = ""
''                .strFracd = ""
''                .strFrart = ""
''                .strPicrt = ""
''                .strSldrt = ""
''
''                If UBound(strDta) > 14 Then
''                    .strTotal = strDta(14)
''                Else
''                    .strTotal = ""
''                End If
''            End With
                    
        Case "M"
            strDta() = Split(RcvBuffer, "|")
            'f_typRst.strPicrt = strDta(4) & "|" & strDta(5) & "|" & strDta(6)
                    
        Case "O"
            strDta() = Split(RcvBuffer, "|")
            strDta() = Split(strDta(4), "^")
            
            pResultInfo.OTHER = strDta(3)
            'f_typRst.strExcde = strDta(3)
                    
        Case "R"
            strDta() = Split(RcvBuffer, "|")
            
            With pResultInfo
                If InStr(strDta(2), "A/G") > 0 Then
                    
                    .RSTCNT = .RSTCNT + 1
                    .RST1 = .RST1 & strDta(3) & Chr(124)
                    
                    strDta() = Split(strDta(2), "^")
                    
                    .IFCD = .IFCD & strDta(4) & Chr(124)
                    .RST2 = .RST2 & "" & Chr(124)
                    .UNIT = .UNIT & "" & Chr(124)
                    .FLAG = .FLAG & "" & Chr(124)
                    .INSTID = .INSTID & "" & Chr(124)        'Inst ID...(2005/1/2) yk
                    .RSTDT = .RSTDT & pSampleInfo.RSTDT & Chr(124)  '결과일시(2005/6/10) yk
                    
                Else
                    If strDta(4) = "%" Then
                        .RSTCNT = .RSTCNT + 1
                        .RST1 = .RST1 & strDta(3) & Chr(124)
                                                    
                        strDta() = Split(strDta(2), "^")
                        
                        If InStr(strDta(4), "M-Spike") > 0 Then
                            If InStr(strDta(4), "Gamma") > 0 Then
                                m_iMPeak1 = m_iMPeak1 + 1
                                strDta(4) = strDta(4) & CStr(m_iMPeak1)
                            Else
                                m_iMPeak2 = m_iMPeak2 + 1
                                strDta(4) = strDta(4) & CStr(m_iMPeak2)
                            End If
                        End If
                        
                        .IFCD = .IFCD & Trim(strDta(4)) & Chr(124)
                        .RST2 = .RST2 & "" & Chr(124)
                        .UNIT = .UNIT & "" & Chr(124)
                        .FLAG = .FLAG & "" & Chr(124)
                        .INSTID = .INSTID & "" & Chr(124)        'Inst ID...(2005/1/2) yk
                        .RSTDT = .RSTDT & pSampleInfo.RSTDT & Chr(124)  '결과일시(2005/6/10) yk
                    End If
                End If
            End With
            
            
''            strDta() = Split(RcvBuffer, "|")
''            If InStr(strDta(2), "A/G") > 0 Then
''                f_typRst.strFrart = f_typRst.strFrart & strDta(3) & "|"
''
''                strDta() = Split(strDta(2), "^")
''                f_typRst.strFracd = f_typRst.strFracd & strDta(4) & "|"
''            Else
''                If strDta(4) = "%" Then
''                    f_typRst.strFrart = f_typRst.strFrart & strDta(3) & "|"
''
''                    strDta() = Split(strDta(2), "^")
''
''                    If InStr(strDta(4), "M-Spike") > 0 Then
''                        If InStr(strDta(4), "Gamma") > 0 Then
''                            m_iMPeak1 = m_iMPeak1 + 1
''                            strDta(4) = strDta(4) & CStr(m_iMPeak1)
''                        Else
''                            m_iMPeak2 = m_iMPeak2 + 1
''                            strDta(4) = strDta(4) & CStr(m_iMPeak2)
''                        End If
''                    End If
''                    f_typRst.strFracd = f_typRst.strFracd & strDta(4) & "|"
''                End If
''            End If
                    
        Case "Q"
            'm_iPhase = 3
            m_iField = 0
            m_sQuery = "Q"
                    
        Case "C"
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .KIND, .RSTDT, .OTHER)
                End If
            End With

            Call Init_pResultInfo
            
''            With f_typRst
''                .strBarno = ""
''                .strChtno = ""
''                .strSeqno = ""
''                .strExcde = ""
''                .strFracd = ""
''                .strFrart = ""
''                .strPicrt = ""
''                .strSldrt = ""
''            End With

    End Select
    
    Exit Sub

ErrRtn:
    MsgBox "Edit_Data - " & Err.Description & "(" & CStr(Val(Err.Number)) & ")"
    
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
'
'   환자 Order 전송
'
Private Sub SendOrder_QuickScan_Batch()
    Dim i%
    
    Dim varTmp  As Variant
    
    Dim strBrno$, strLbno$, strExcd$, strTPro$, strItem$, strSeq$
    Dim strPtno$, strPsex$, strPage$, strDept$, strWard$, strLab$, strPtnm$, strSpNm$
    
    Dim strSend$
    Dim sTotTest As String
   
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.KIND)
    
    Call Get_OrderString

    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If
    
    If pSampleInfo.ID = "" Then
        strSend = Trim(CStr(m_iFrameN)) & "L|1" & vbCr & Chr(3)
        msComm.Output = Chr(2) & strSend & CheckSumASTM(strSend) & vbCr & vbLf
        
        If m_sTestMode = "77" Then
            RaiseEvent PrintSendLog(Chr(2) & strSend & CheckSumASTM(strSend) & vbCr & vbLf)
        End If

        m_iField = -1
         
        Exit Sub
    End If
    
    Select Case m_iField
        Case 0:
            msComm.Output = Chr(5)
            m_iFrameN = 0
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(5))
            End If
            
        Case 1:
            strSend = Trim(CStr(m_iFrameN)) & "H|" & vbCr & Chr(3)
            msComm.Output = Chr(2) & strSend & CheckSumASTM(strSend) & vbCr & vbLf
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(2) & strSend & CheckSumASTM(strSend) & vbCr & vbLf)
            End If
                    
        Case 2:
            '2P|1||2005-L1234-1|123456780|가|||||10/M||||2005071712345|12.3
            'CA
            For i = 1 To UBound(pSampleInfo.IFCD)
                sTotTest = sTotTest & pSampleInfo.IFCD(i) & ", "
            Next
            sTotTest = Mid(sTotTest, 1, Len(sTotTest) - 2)
            
            strSend = Trim(CStr(m_iFrameN)) & "P|||" & "" & pSampleInfo.ID & "|" & pSampleInfo.HosName & "|" & pSampleInfo.PatName & _
                        "|||||||||" & pSampleInfo.Total & "||||" & sTotTest & "|" & vbCr & Chr(3)
                        
            msComm.Output = Chr(2) & strSend & CheckSumASTM(strSend) & vbCr & vbLf
            
            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
            
            pSampleInfo.ID = ""
            pSampleInfo.SEQNO = ""
            pSampleInfo.RACK = ""
            pSampleInfo.POS = ""
            pSampleInfo.PatName = ""
            pSampleInfo.HosName = ""
            pSampleInfo.Total = ""
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(2) & strSend & CheckSumASTM(strSend) & vbCr & vbLf)
            End If
    End Select
    
    m_iFrameN = m_iFrameN + 1
    m_iPhase = 3
        
End Sub

Public Function CheckSumASTM(ByVal sBuf$) As String

'    Dim i   As Integer
'    Dim Tmp As Integer
'    Dim ChkS1   As Integer
'    Dim ChkS2   As String
'
''    For i = 1 To Len(Para)
''        Tmp = Asc(Mid$(Para, i, 1))
''        ChkS1 = ChkS1 + Tmp
''    Next i
'
'    For i = 0 To UBound(Para) - 1
'        Tmp = Para(i)
'        ChkS1 = ChkS1 + Tmp
'    Next
'
'    ChkS1 = ChkS1 Mod 256
'    ChkS2 = Right$("0" & Hex$(ChkS1), 2)
'
'    ChkSum = ChkS2
    
    Dim iCnt As Integer
    Dim iSum As Integer
    
    Dim a_Buf() As Byte
    
    a_Buf = StrConv(sBuf, vbFromUnicode)
    
    For iCnt = 0 To UBound(a_Buf)
        iSum = iSum + a_Buf(iCnt)
    Next
        
    iSum = iSum Mod 256
    
    CheckSumASTM = Right("0" & CStr(Hex(iSum)), 2)
    
    
End Function

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
            .PatName = m_p_PatName
            .Age = m_p_Age
            .Sex = m_p_Sex
            .HosName = m_p_HosName
            .Total = m_p_Total
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
        .PatName = m_p_PatName
        .Age = m_p_Age
        .Sex = m_p_Sex
        .HosName = m_p_HosName
        .Total = m_p_Total
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
    m_p_sCmt1 = PropBag.ReadProperty("p_sCmt1", m_def_p_sCmt1)
    
    m_p_Total = PropBag.ReadProperty("p_Total", m_def_p_Total)
    m_p_RegNo = PropBag.ReadProperty("p_RegNo", m_def_p_RegNo)
    m_p_PatName = PropBag.ReadProperty("p_PatName", m_def_p_PatName)
    m_p_HosName = PropBag.ReadProperty("p_HosName", m_def_p_HosName)
    m_p_DepNm = PropBag.ReadProperty("p_DepNm", m_def_p_DepNm)
    m_p_Age = PropBag.ReadProperty("p_Age", m_def_p_Age)
    m_p_Sex = PropBag.ReadProperty("p_Sex", m_def_p_Sex)
    m_p_SpcNo = PropBag.ReadProperty("p_SpcNo", m_def_p_SpcNo)

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
    
    Call PropBag.WriteProperty("p_Total", m_p_Total, m_def_p_Total)
    Call PropBag.WriteProperty("p_RegNo", m_p_RegNo, m_def_p_RegNo)
    Call PropBag.WriteProperty("p_PatName", m_p_PatName, m_def_p_PatName)
    Call PropBag.WriteProperty("p_HosName", m_p_HosName, m_def_p_HosName)
    Call PropBag.WriteProperty("p_DepNm", m_p_DepNm, m_def_p_DepNm)
    Call PropBag.WriteProperty("p_Age", m_p_Age, m_def_p_Age)
    Call PropBag.WriteProperty("p_Sex", m_p_Sex, m_def_p_Sex)
    Call PropBag.WriteProperty("p_SpcNo", m_p_SpcNo, m_def_p_SpcNo)
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
    m_p_sCmt1 = m_def_p_sCmt1
    
    m_p_Total = m_def_p_Total
    m_p_RegNo = m_def_p_RegNo
    m_p_PatName = m_def_p_PatName
    m_p_HosName = m_def_p_HosName
    m_p_DepNm = m_def_p_DepNm
    m_p_Age = m_def_p_Age
    m_p_Sex = m_def_p_Sex
    m_p_SpcNo = m_def_p_SpcNo
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

Public Property Get p_Total() As String
    p_Total = m_p_Total
End Property
Public Property Let p_Total(ByVal New_p_Total As String)
    m_p_Total = New_p_Total
    PropertyChanged "p_Total"
End Property

Public Property Get p_RegNo() As String
    p_RegNo = m_p_RegNo
End Property
Public Property Let p_RegNo(ByVal New_p_RegNo As String)
    m_p_RegNo = New_p_RegNo
    PropertyChanged "p_RegNo"
End Property

Public Property Get p_PatName() As String
    p_PatName = m_p_PatName
End Property
Public Property Let p_PatName(ByVal New_p_PatName As String)
    m_p_PatName = New_p_PatName
    PropertyChanged "p_PatName"
End Property

Public Property Get p_HosName() As String
    p_HosName = m_p_HosName
End Property
Public Property Let p_HosName(ByVal New_p_HosName As String)
    m_p_HosName = New_p_HosName
    PropertyChanged "p_HosName"
End Property

Public Property Get p_DepNm() As String
    p_DepNm = m_p_DepNm
End Property
Public Property Let p_DepNm(ByVal New_p_DepNm As String)
    m_p_DepNm = New_p_DepNm
    PropertyChanged "p_DepNm"
End Property

Public Property Get p_Age() As String
    p_Age = m_p_Age
End Property
Public Property Let p_Age(ByVal New_p_Age As String)
    m_p_Age = New_p_Age
    PropertyChanged "p_Age"
End Property

Public Property Get p_Sex() As String
    p_Sex = m_p_Sex
End Property
Public Property Let p_Sex(ByVal New_p_Sex As String)
    m_p_Sex = New_p_Sex
    PropertyChanged "p_Sex"
End Property

Public Property Get p_SpcNo() As String
    p_SpcNo = m_p_SpcNo
End Property
Public Property Let p_SpcNo(ByVal New_p_SpcNo As String)
    m_p_SpcNo = New_p_SpcNo
    PropertyChanged "p_SpcNo"
End Property

