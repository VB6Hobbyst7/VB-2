VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl UF100 
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
      ScrollBars      =   2  '����
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
      Handshaking     =   1
   End
End
Attribute VB_Name = "UF100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�⺻ �Ӽ� ��:
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
'�Ӽ� ����:
Dim m_iBCLen As Integer
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
'�̺�Ʈ ����:
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event SendOrderOK(sID$)
Event DispMsg(sMsg$)


'===== User Define
'�������̽����� ���
Dim RcvBuffer   As String
Dim wkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'����ü ����
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'��Ÿ
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

Dim miRstCnt%, msTotIFRstCd$, msTotRst$, msTotRst2$
Dim msUnit$, msFlag$, msReview$, msCMatch$, msComment$, msRBCInfo$
Dim msBarCd$, msRack$, msPos$

'For Urisys2400
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim RstEnd  As String

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,CommPort
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "��� ��Ʈ ��ȣ�� ��ȯ�ϰų� �����մϴ�."
    CommPort = msComm.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    msComm.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Private Sub PhaseCfg_Protocol()
    On Error GoTo ErrHandler
    
    '--- ����� Ȯ��
    If m_EditPW <> pEditPW Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�. (��)���̾����̷� ������ �ֽʽÿ�!!!", vbCritical, "����� Ȯ��"
        Exit Sub
    End If
    '---------------
    
    If m_EqName = "0" Or m_EqName = "" Then
        RaiseEvent DispMsg("�˻������� ������ �ֽʽÿ�.!!!")
        Exit Sub
    End If
    
    Select Case UCase(m_EqName)
        Case "UF100"
            Call PhaseCfg_Protocol_UF100
        
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
    Exit Sub
    
ErrHandler:
    RaiseEvent DispMsg("PhaseCfg_Protocol ���� - " & Err.Description)
End Sub
Private Sub DataEditResponse_UF100()
    On Error GoTo ErrRtn

    Dim sBC         As String
    Dim sLC         As String

    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim ii          As Integer
    Dim sData()     As String
    Dim tmpIFCd$, tmpRst$, tmpRstDT$
    Dim sTIFCd$, sTRst$, sTRst2$, sTUnit$, sTFlag$
    Dim tmpCnt%
    Dim sUFNeed     As String
    Dim sMach       As String

    Dim iChk%, iCnt%, sIFRstCd$, sRst$

    sBC = Mid(RcvBuffer, 1, 1)
    sLC = Mid(RcvBuffer, 2, 5)

    Select Case sBC
        Case Chr(13)    'CLINITEK200+ Format

'            
'            #0-001      05-02-22
'            ID = 4095700
'            COLOR=Dk yellow     '
'            Clarity=Turbid      '
'            GLU*              1+
'            BIL*              3+
'            KET*              1+
'            SG 1.01
'            BLO*              3+
'            pH             >=9.0
'            PRO*              2+
'            URO*   >=8.0 E.U./dL
'            NIT*        POSITIVE
'            LEU*              1+
'            

            '������� �ʱ�ȭ
            Call Init_pResultInfo

            iChk = 1: tmpCnt = 0: sTIFCd = "": sTRst = ""
            sUFNeed = ""
            sMach = "ATLAS"

            'chr(10)�� ����
            RcvBuffer = Replace(RcvBuffer, Chr(10), "")

            'Chr(13)���� �и�
            sData = Split(RcvBuffer, Chr(13))

            For ii = 1 To UBound(sData) - 1
                If Left(sData(ii), 1) = "#" Then
                    'SEQ����
                    tmpSeqNo = Trim(Mid(sData(ii), 2, 5))
                ElseIf UCase(Left(sData(ii), 2)) = "ID" Then
                    'ID����
                    tmpBarCd = Trim(Mid(sData(ii), 4))
                ElseIf UCase(Left(sData(ii), 5)) = "COLOR" Then
                    'Color����
                    tmpCnt = tmpCnt + 1
                    sTIFCd = sTIFCd & "COLOR" & Chr(124)
                    sTRst = sTRst & Trim(Mid(sData(ii), 7, 14)) & Chr(124)
                    sTRst2 = sTRst2 & Chr(124)
                    sTUnit = sTUnit & Chr(124)
                    sTFlag = sTFlag & Chr(124)
                ElseIf UCase(Left(sData(ii), 7)) = "CLARITY" Then
                    'Clarity����
                    'CL500���� ����ڵ庯�� - ATLAS�� Clarity�ȳ����Լ�����
                    sMach = "CL500"
                    tmpRack = Left(tmpSeqNo, 1)
                    tmpPos = Right(tmpSeqNo, 3)
                    tmpCnt = tmpCnt + 1
                    sTIFCd = sTIFCd & "CLARITY" & Chr(124)
                    sTRst = sTRst & Trim(Mid(sData(ii), 9, 12)) & Chr(124)
                    sTRst2 = sTRst2 & Chr(124)
                    sTUnit = sTUnit & Chr(124)
                    sTFlag = sTFlag & Chr(124)
                Else
                    'URINE10�����
                    tmpIFCd = Trim(Left(sData(ii), 3))
                    'tmpRst = Trim(Mid(sData(ii), 7))
                    tmpRst = Mid(sData(ii), 7)

                    If tmpIFCd = "URO" Then
                        tmpRst = Trim(Left(tmpRst, 6))
                        If tmpRst = "" Then tmpRst = "ERROR"
                    End If

                    tmpCnt = tmpCnt + 1
                    sTIFCd = sTIFCd & tmpIFCd & Chr$(124)
                    'sTRst = sTRst & tmpRst & Chr$(124)
                    sTRst = sTRst & Trim(tmpRst) & Chr$(124)
                    sTRst2 = sTRst2 & Chr(124)
                    sTUnit = sTUnit & Chr(124)
                    sTFlag = sTFlag & Chr(124)
                End If
            Next ii

            tmpRstDT = Format(Now, "YYYYMMDDHHNNSS")

            With pResultInfo
                .ID = tmpBarCd
                .SEQNO = tmpSeqNo
                .RACK = tmpRack
                .POS = tmpPos
                .RSTCNT = tmpCnt

                .IFCD = sTIFCd
                .RST1 = sTRst
                .RST2 = sTRst2
                .UNIT = sTUnit
                .FLAG = sTFlag
                For ii = 1 To tmpCnt
                    .RSTDT = .RSTDT & tmpRstDT & Chr(124)
                Next ii
            End With

            '��� ó��
            With pResultInfo
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", .RSTDT, sMach)
            End With

            msComm.Output = Chr(6)

            If sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(6))
            End If

        Case "O"

            Call Sleep(500)

            pSampleInfo.ID = ""

            msComm.Output = Chr(6)

            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(6))
            End If

            Call Sleep(500)

            'O00101........0-001
            pSampleInfo.RACK = Mid(RcvBuffer, 2, 3)
            pSampleInfo.POS = Mid(RcvBuffer, 5, 2)
            pSampleInfo.ID = Trim(Mid(RcvBuffer, 7, 13))

            'Bacode Length ������ �������ڵ��ȣ����...
            pSampleInfo.ID = Right(pSampleInfo.ID, m_iBCLen)

            Call SendOrder_UF100

            If pSampleInfo.ORDCNT = 0 Then
                msComm.Output = Chr(2) & "SNG" & Mid(RcvBuffer, 7, 13) & Chr(3)
                If m_sTestMode = "77" Then
                    RaiseEvent PrintSendLog(Chr(2) & "SNG" & Mid(RcvBuffer, 7, 13) & Chr(3))
                End If
            Else
                msComm.Output = Chr(2) & "SGO" & Mid(RcvBuffer, 7, 13) & Chr(3)
                If m_sTestMode = "77" Then
                    RaiseEvent PrintSendLog(Chr(2) & "SGO" & Mid(RcvBuffer, 7, 13) & Chr(3))
                End If
            End If

            Exit Sub

        Case "D"
            Select Case sLC
                Case "S4101"    'Sample Information Block
                    '������� �ʱ�ȭ
                    Call Init_pResultInfo

                    miRstCnt = 0: msTotIFRstCd = "": msTotRst = "": msTotRst2 = ""
                    msUnit = "": msFlag = "": msReview = "": msCMatch = "": msRBCInfo = ""
                    msBarCd = "": msRack = "": msPos = ""

                    msRack = Mid(RcvBuffer, 20, 4)
                    msPos = Mid(RcvBuffer, 24, 2)
                    msBarCd = Trim(Mid(RcvBuffer, 26, 13))

                    If IsNumeric(msBarCd) = True Then
                        msBarCd = Trim(CStr(Val(msBarCd)))
                    End If

                    'CROSS-MATCH
                    If Mid(RcvBuffer, 98, 1) = "?" Then
                        msCMatch = msCMatch & "/RBC"
                    End If
                    If Mid(RcvBuffer, 99, 1) = "?" Then
                        msCMatch = msCMatch & "/WBC"
                    End If
                    If Mid(RcvBuffer, 100, 1) = "?" Then
                        msCMatch = msCMatch & "/EC"
                    End If
                    If Mid(RcvBuffer, 101, 1) = "?" Then
                        msCMatch = msCMatch & "/CAST"
                    End If
                    If Mid(RcvBuffer, 102, 1) = "?" Then
                        msCMatch = msCMatch & "/BACT"
                    End If

                    'REVIEW
                    If Mid(RcvBuffer, 41, 1) = "1" Then
                        msReview = "Y" & msCMatch
                    Else
                        msReview = "" & msCMatch
                    End If
                    
                    'Comment���������� �߰���
'                    miRstCnt = miRstCnt + 1
'                    msTotIFRstCd = msTotIFRstCd & "REVIEW" & Chr(124)
'                    msTotRst = msTotRst & msReview & Chr(124)
'                    msTotRst2 = msTotRst2 & Chr(124)
'                    msUnit = msUnit & Chr(124)
'                    msFlag = msFlag & Chr(124)

                    'RBC INFO
                    Select Case Mid(RcvBuffer, 46, 1)
                        Case "0"
                            msRBCInfo = "Inadequate RBC Count"
                        Case "1"
                            msRBCInfo = "Isomorphic RBC"
                        Case "2"
                            msRBCInfo = "Dysmorphic RBC"
                        Case "3"
                            msRBCInfo = "Mixed RBC"
                        Case Else
                            msRBCInfo = "ERROR"
                    End Select
                    miRstCnt = miRstCnt + 1
                    msTotIFRstCd = msTotIFRstCd & "INFO" & Chr(124)
                    msTotRst = msTotRst & msRBCInfo & Chr(124)
                    msTotRst2 = msTotRst2 & Chr(124)
                    msUnit = msUnit & Chr(124)
                    msFlag = msFlag & Chr(124)

                    'CONDUCTIVITY
                    miRstCnt = miRstCnt + 1
                    msTotIFRstCd = msTotIFRstCd & "COND" & Chr(124)
                    sRst = Mid(RcvBuffer, 47, 4)
                    If IsNumeric(sRst) = True Then
                        sRst = Trim(CStr(Val(sRst)))
                    Else
                        sRst = "ERROR"
                    End If

                    msTotRst = msTotRst & sRst & Chr(124)
                    msTotRst2 = msTotRst2 & Chr(124)
                    msUnit = msUnit & Chr(124)
                    msFlag = msFlag & Chr(124)

                    'Total Cell Count
                    miRstCnt = miRstCnt + 1
                    msTotIFRstCd = msTotIFRstCd & "CELL" & Chr(124)
                    sRst = Mid(RcvBuffer, 87, 6)
                    If IsNumeric(sRst) = True Then
                        sRst = Trim(CStr(Val(sRst)))
                    Else
                        sRst = "ERROR"
                    End If

                    msTotRst = msTotRst & sRst & Chr(124)
                    msTotRst2 = msTotRst2 & Chr(124)
                    msUnit = msUnit & Chr(124)
                    msFlag = msFlag & Chr(124)

                Case "P4102"    'Cell Count Block

                    iCnt = Mid(RcvBuffer, 9, 2)
                    RcvBuffer = Mid(RcvBuffer, 11)

                    If IsNumeric(iCnt) Then
                        For ii = 1 To Val(iCnt)
                            sIFRstCd = Mid(RcvBuffer, 1 + 12 * (ii - 1), 4)
                            sRst = Mid(RcvBuffer, 1 + 12 * (ii - 1) + 4, 8)

                            If IsNumeric(sRst) = True Then
                                sRst = Trim(CStr(Val(sRst)))
                            Else
                                sRst = "ERROR"
                            End If

                            miRstCnt = miRstCnt + 1
                            msTotIFRstCd = msTotIFRstCd & sIFRstCd & Chr(124)
                            msTotRst = msTotRst & sRst & Chr(124)
                            msTotRst2 = msTotRst2 & Chr(124)
                            msUnit = msUnit & Chr(124)
                            msFlag = msFlag & Chr(124)
                        Next ii
                    End If

                Case "C4103"    'Comment Block

                    'DC4103030200D90107
                    iCnt = Mid(RcvBuffer, 9, 2)
                    RcvBuffer = Mid(RcvBuffer, 11)
                    
                    msComment = ""
                    
                    If IsNumeric(iCnt) Then
                        For ii = 1 To Val(iCnt)
                            sIFRstCd = Mid(RcvBuffer, 1 + 4 * (ii - 1), 4)
                            Select Case sIFRstCd
                                Case "00D9" 'P.Cast
                                    msComment = msComment & ">P.Cast"
                                Case "0107" 'SRC
                                    msComment = msComment & ">SRC"
                                Case "0402" 'YLC
                                    msComment = msComment & ">YLC"
                                Case "0300" 'X'TAL
                                    msComment = msComment & ">XTAL"
                                Case "0501" 'SPERM
                                    msComment = msComment & ">SPERM"
                            End Select
                        Next ii
                    End If
                    
                    'Review + Cross �� Comment(FLAG)�߰���
                    msReview = msReview & msComment

                    miRstCnt = miRstCnt + 1
                    msTotIFRstCd = msTotIFRstCd & "REVIEW" & Chr(124)
                    msTotRst = msTotRst & msReview & Chr(124)
                    msTotRst2 = msTotRst2 & Chr(124)
                    msUnit = msUnit & Chr(124)
                    msFlag = msFlag & Chr(124)

                    tmpRstDT = Format(Now, "YYYYMMDDHHNNSS")

                    With pResultInfo
                        .ID = msBarCd
                        .RACK = msRack
                        .POS = msPos
                        .RSTCNT = miRstCnt

                        .IFCD = msTotIFRstCd
                        .RST1 = msTotRst
                        .RST2 = msTotRst2
                        .UNIT = msUnit
                        .FLAG = msFlag
                        For ii = 1 To miRstCnt
                            .RSTDT = .RSTDT & tmpRstDT & Chr(124)
                        Next ii
                    End With

                    '��� ó��
                    With pResultInfo
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", .RSTDT, "UF100")
                    End With

                    miRstCnt = 0: msTotIFRstCd = "": msTotRst = "": msTotRst2 = ""
                    msUnit = "": msFlag = "": msReview = "": msCMatch = "": msRBCInfo = ""
                    msBarCd = "": msRack = "": msPos = ""
            End Select

            msComm.Output = Chr(6)

            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(6))
            End If

        Case Else

            msComm.Output = Chr(6)

            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(6))
            End If
    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit ���� �߻� - " & Err.Description)
    End If
End Sub


'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iBCLen() As Integer
    iBCLen = m_iBCLen
End Property

Public Property Let iBCLen(ByVal New_iBCLen As Integer)
    m_iBCLen = New_iBCLen
    PropertyChanged "iBCLen"
End Property

Private Sub PhaseCfg_Protocol_UF100_Old()
    On Error GoTo ErrHandler
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX
                RcvBuffer = ""
            
            Case 3      'ETX
                Call DataEditResponse_UF100
                
                msComm.Output = Chr(6)       'ACK

                If sTestMode = "77" Then
                    RaiseEvent PrintSendLog(Chr(6))
                End If
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
    Exit Sub
    
ErrHandler:
    RaiseEvent DispMsg("PhaseCfg_Protocol_UF100 ���� - " & Err.Description)
End Sub

Private Sub PhaseCfg_Protocol_UF100()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
        Select Case m_iPhase
'            Case 1
'                Select Case Asc(wkDat)
'                    Case 5  'ENQ
'                        m_iPhase = 2
'                        RstEnd = "Y"
'                        bEndChk = True: bSTXChk = False
'
'                        msComm.Output = Chr(6)
'
'                    Case Else
'
'                End Select

            Case 1, 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If Mid(wkBuf, ix1 + 1, 1) = "D" Then
                            'uf-100
                            RcvBuffer = ""
                        Else
                            If bEndChk = True Then
                                RcvBuffer = ""
                            Else
                                bSTXChk = True
                            End If
                            bEndChk = True
                        End If
                        
                    Case 10     'LF
                        If bEndChk = True Then
                            Call DataEditResponse_URISYS2400
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                        
                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_URISYS2400
                            RcvBuffer = ""
                        End If
                    
                    Case 3      'ETX
                        If Left(RcvBuffer, 1) = "D" Or Left(RcvBuffer, 1) = "O" Then
                            Call DataEditResponse_UF100
                        End If
                        
                        msComm.Output = Chr(6)
                    
                        If sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                    
                    Case 4      'EOT
                        RcvBuffer = ""
                        m_iPhase = 1
                    
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                    
                    Case 21     'NAK
                        Call DataEditResponse_URISYS2400
                    
                    Case 23     'ETB
                        bEndChk = False
                        
                    Case Else
                        If Left(RcvBuffer, 1) = "D" Then
                            RcvBuffer = RcvBuffer & wkDat
                        Else
                            If bEndChk = True Then
                                If bSTXChk = True Then
                                    bSTXChk = False
                                Else
                                    RcvBuffer = RcvBuffer & wkDat
                                End If
                            End If
                        End If
                End Select

            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
'                        Call Order_Input
                    
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    
                    Case 21     'NAK
                        msComm.Output = Chr(5)   'ENQ
'                        Call Order_Input
                        m_iPhase = 3
                        
                    Case 4      'EOT
                        m_iPhase = 1
                        
                End Select
        End Select
    Next ix1

End Sub

' *=====================================================*
' *               Data���� & ����ó��                   *
' *=====================================================*
Private Sub DataEditResponse_URISYS2400()
    On Error GoTo ErrRtn
    
    Dim RecType As String   'Record Type
    Dim ii      As Integer
    
    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpType$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$
    
    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If
    
    Select Case RecType
        Case "H"        'Header Record
        Case "M"
        Case "P"        'Patient Record
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
            ' URISYS2400 ��񿡼� ������� ����...

        Case "O"
            tmpBarCd = ""
            tmpSeqNo = "": tmpRack = "": tmpPos = "": tmpType = ""
            
            tmpField() = Split(RcvBuffer, Chr(124))
            
            tmpBarCd = Trim(tmpField(2))
            
            If Trim(tmpField(3)) = "" Then Exit Sub
            ii = InStr(tmpField(3), "^")
            If ii <> 0 Then
                tmpData() = Split(Trim(tmpField(3)), "^")
                
                tmpSeqNo = Trim(tmpData(0))
                tmpRack = Trim(tmpData(1))
                tmpPos = Trim(tmpData(2))
                tmpType = Trim(tmpData(4))      'SAMPLE/CONTROL
            End If
            
            With pSampleInfo
                .ID = UCase(tmpBarCd)
                .SEQNO = tmpSeqNo
                .RACK = tmpRack
                .POS = tmpPos
                .KIND = tmpType
            End With
            
        Case "R"        'Result Record
            '--- �������Ÿ ����
            '2: TEST ID
            '3: Result
            '4: UNIT
            '6: Result Abnormal Flag
            '8: Result Status
            
            'ex) R|5|^^^5|0.25|g/L|||||(CR)
            '# Test No (fixation)
            '1: SG, 2: pH, 3: LEU, 4: NIT, 5: PRO, 6: GLU,7: KET, 8: UBG, 9: BIL, 10:ERY, 11: COL, 12: CLA
            
            '# Sample flags
            '!: ID Edit
            'N: Sample Short
            'E: Sample Empty
            'R: Test Strip
            
            '# Result status: No value with .X.

            Erase tmpField()
            tmpField() = Split(RcvBuffer, "|")
            
            tmpData() = Split(tmpField(2), "^")
            tmpIFCd = Trim(tmpData(3))
            
            tmpRst = Trim(tmpField(3))
            tmpUnit = Trim(tmpField(4))
            tmpFlag = Trim(tmpField(6))
            
            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If

            '������� ����ü�� ����
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS
                .KIND = pSampleInfo.KIND

                '����� ����
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
            End With

        Case "C"        'Comment Record

        Case "L"
            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", .RSTDT, "URISYS")
                End If
            End With

            Call Init_pResultInfo
    
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit �����߻� - " & Err.Description)
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
'   ������� ����ü �ʱ�ȭ
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
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
End Sub
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,RTSEnable
Public Property Get RTSEnable() As Boolean
Attribute RTSEnable.VB_Description = "���� ��û ���� ���������� ���θ� �����մϴ�."
    RTSEnable = msComm.RTSEnable
End Property

Public Property Let RTSEnable(ByVal New_RTSEnable As Boolean)
    msComm.RTSEnable() = New_RTSEnable
    PropertyChanged "RTSEnable"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,RThreshold
Public Property Get RThreshold() As Integer
Attribute RThreshold.VB_Description = "������ ������ ���� ��ȯ�ϰų� �����մϴ�."
    RThreshold = msComm.RThreshold
End Property

Public Property Let RThreshold(ByVal New_RThreshold As Integer)
    msComm.RThreshold() = New_RThreshold
    PropertyChanged "RThreshold"
End Property

Private Sub SendOrder_UF100()
    On Error GoTo ErrRtn
    
    Dim SendBuf$, sBuf$
    Dim iPos%, i%
    Dim sOrder$
    
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    Call Get_OrderString
    
    If pSampleInfo.ORDCNT = 0 Then
        RaiseEvent DispMsg("�������̽� ���� �׸��� �������� �ʽ��ϴ�!!")
    Else
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder �����߻� - " & Err.Description)
    End If
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,Settings
Public Property Get Settings() As String
Attribute Settings.VB_Description = "���� �ӵ�, �и�Ƽ, ������ ��Ʈ, �ߴ� ��Ʈ �Ű� ������ ��ȯ�ϰų� �����մϴ�."
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
    On Error GoTo ErrHandler
        
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
            
            RaiseEvent DispMsg(Space(iSpaceCnt) & "���� Interface �۾� ��...")
            
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
    
    Exit Sub
    
ErrHandler:
    RaiseEvent DispMsg("msComm_OnComm ���� - " & Err.Description)
End Sub

'����ҿ��� �Ӽ����� �ε��մϴ�.
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

'�Ӽ����� ����ҿ� ����մϴ�.
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0,0,0,0
Public Property Get PortOpen() As Boolean
    PortOpen = m_PortOpen
End Property

Public Property Let PortOpen(ByVal New_PortOpen As Boolean)
    m_PortOpen = New_PortOpen
    PropertyChanged "PortOpen"
    
    '--- PortOpen�� ��ȣ Ȯ��
    If m_OpenPW <> pOpenPW Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�. (��)���̾����̷� ������ �ֽʽÿ�!!!", vbCritical, "����� Ȯ��"
        Exit Property
    End If
    '-----------------------
    
    On Error GoTo ErrPortOpen
    If m_PortOpen = True Then
        msComm.PortOpen = True
    End If
    On Error GoTo 0
    
    m_iOrderFlag = 0
    
    '���� �ʱ�ȭ
    RstEnd = "Y": bEndChk = True: bSTXChk = False
    
ErrPortOpen:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
        RaiseEvent RaiseError("PortOpen Error!!! " & Err.Description)
    End If
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get OpenPW() As String
    OpenPW = m_OpenPW
End Property

Public Property Let OpenPW(ByVal New_OpenPW As String)
    m_OpenPW = New_OpenPW
    PropertyChanged "OpenPW"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get EditPW() As String
    EditPW = m_EditPW
End Property

Public Property Let EditPW(ByVal New_EditPW As String)
    m_EditPW = New_EditPW
    PropertyChanged "EditPW"
End Property

'����� ���� ��Ʈ�ѿ� ���� �Ӽ��� �ʱ�ȭ�մϴ�.
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get EqName() As String
    EqName = m_EqName
End Property

Public Property Let EqName(ByVal New_EqName As String)
    m_EqName = New_EqName
    PropertyChanged "EqName"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0,0,0,0
Public Property Get bUseBarcode() As Boolean
    bUseBarcode = m_bUseBarcode
End Property

Public Property Let bUseBarcode(ByVal New_bUseBarcode As Boolean)
    m_bUseBarcode = New_bUseBarcode
    PropertyChanged "bUseBarcode"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iPhase() As Integer
    iPhase = m_iPhase
End Property

Public Property Let iPhase(ByVal New_iPhase As Integer)
    m_iPhase = New_iPhase
    PropertyChanged "iPhase"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iSendPhase() As Integer
    iSendPhase = m_iSendPhase
End Property

Public Property Let iSendPhase(ByVal New_iSendPhase As Integer)
    m_iSendPhase = New_iSendPhase
    PropertyChanged "iSendPhase"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get sTestMode() As String
    sTestMode = m_sTestMode
End Property

Public Property Let sTestMode(ByVal New_sTestMode As String)
    m_sTestMode = New_sTestMode
    PropertyChanged "sTestMode"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iFrameN() As Integer
    iFrameN = m_iFrameN
End Property

Public Property Let iFrameN(ByVal New_iFrameN As Integer)
    m_iFrameN = New_iFrameN
    PropertyChanged "iFrameN"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sID() As String
    p_sID = m_p_sID
End Property

Public Property Let p_sID(ByVal New_p_sID As String)
    m_p_sID = New_p_sID
    PropertyChanged "p_sID"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sSeq() As String
    p_sSeq = m_p_sSeq
End Property

Public Property Let p_sSeq(ByVal New_p_sSeq As String)
    m_p_sSeq = New_p_sSeq
    PropertyChanged "p_sSeq"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sRack() As String
    p_sRack = m_p_sRack
End Property

Public Property Let p_sRack(ByVal New_p_sRack As String)
    m_p_sRack = New_p_sRack
    PropertyChanged "p_sRack"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sPos() As String
    p_sPos = m_p_sPos
End Property

Public Property Let p_sPos(ByVal New_p_sPos As String)
    m_p_sPos = New_p_sPos
    PropertyChanged "p_sPos"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get p_iOrdCnt() As Integer
    p_iOrdCnt = m_p_iOrdCnt
End Property

Public Property Let p_iOrdCnt(ByVal New_p_iOrdCnt As Integer)
    m_p_iOrdCnt = New_p_iOrdCnt
    PropertyChanged "p_iOrdCnt"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sTIFCd() As String
    p_sTIFCd = m_p_sTIFCd
End Property

Public Property Let p_sTIFCd(ByVal New_p_sTIFCd As String)
    m_p_sTIFCd = New_p_sTIFCd
    PropertyChanged "p_sTIFCd"
End Property
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=14
Public Function Send_Chr(iChr%) As Variant
    On Error GoTo ErrComm
    msComm.Output = Chr(iChr)
    On Error GoTo 0
ErrComm:
    RaiseEvent DispMsg("Send_Chr ���� - " & Err.Description)
End Function

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iOrderFlag() As Integer
    iOrderFlag = m_iOrderFlag
End Property

Public Property Let iOrderFlag(ByVal New_iOrderFlag As Integer)
    m_iOrderFlag = New_iOrderFlag
    PropertyChanged "iOrderFlag"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iTotalItemCnt() As Integer
    iTotalItemCnt = m_iTotalItemCnt
End Property

Public Property Let iTotalItemCnt(ByVal New_iTotalItemCnt As Integer)
    m_iTotalItemCnt = New_iTotalItemCnt
    PropertyChanged "iTotalItemCnt"
End Property

