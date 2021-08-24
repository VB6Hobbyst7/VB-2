VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl COBASCORE 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   1980
      Top             =   2205
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   1500
      Top             =   2205
   End
   Begin VB.Timer Timer1 
      Left            =   990
      Top             =   2205
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
Attribute VB_Name = "COBASCORE"
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
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim WkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim iSpaceCnt   As Integer

'For Integra Series
Dim iTimerFlag      As Integer
Dim iIdleFlag       As Integer
Dim iPendingFlag    As Integer
Dim iConnectFlag    As Integer
Dim iOrderFlag      As Integer
Dim iResultFlag     As Integer
Dim iTimerCnt       As Integer
Dim iResultCnt      As Integer
Dim iOrdRstCnt      As Integer

Dim iNoTestFlag     As Integer
Dim iPendOrderCnt   As Integer

Dim iOrdState   As Integer
Dim iRstState   As Integer


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
        Case "COBASCORE2"
            Call PhaseCfg_Protocol_CobasCore2
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_CobasCore2()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(WkBuf)
        wkDat = Mid$(WkBuf, ix1, 1)
             
        Select Case Asc(wkDat)
            Case 1         ' SOH
                RcvBuffer = ""
                
            Case 4         ' EOT
                Call EditData_CobasCore2
                RcvBuffer = ""
            
            Case 17, 19    ' DC1, DC3 (XON, XOFF) 삭제
           
            Case Else      ' Data
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
End Sub
Private Sub EditData_CobasCore2()
    On Error GoTo ErrRtn
    
'<---- COBAS 장비에서 주로 사용 S --->
    Dim sBC          As String
    Dim sLC          As String
    Dim iBCpos       As Integer
    Dim iLCpos       As Integer
    
    Dim iErrCode     As Integer
'<---- COBAS 장비에서 주로 사용 E --->

    Dim tmpBarCd$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$
    
    Dim sRst     As String
    Dim sRst2    As String
    Dim sExpFlag    As String
    Dim sSignFlag   As String
    Dim sIFRstCd    As String
    
    Dim sControlNm  As String
    Dim sJNo$, sRack$, sPos$
    
    iErrCode = 0
    iBCpos = 22
    sBC = Mid(RcvBuffer, iBCpos, 2)
    
    Select Case sBC
        '### Idle Block, No more result Block ###
        Case "00"
            iIdleFlag = 1
        
        '### CAL Result Block ###
        Case "02"
        
        '### Control Result Block ###
        Case "03"
        
        '### Patient Result Block ###
        Case "04"
        
        '### Order Manipulation response Block ###
        Case "19"
            iErrCode = 99
        
        '### pending Sample Tubes Response Block ###
        Case "62"
            iPendingFlag = 1
            
        '### No More pending Sample Tubes Response Block ###
        Case "69"

        Case Else
        
    End Select
    
    iLCpos = iBCpos + 5
    
    Do
        If Asc(Mid(RcvBuffer, iLCpos, 1)) = 3 Then  'ETX(END OF DATA BLOCK)
            Exit Do
        End If

        sLC = Mid(RcvBuffer, iLCpos, 2)
        
        Select Case sLC
            Case "00"       'RESULT DATA
                sSignFlag = Trim(Mid(RcvBuffer, iLCpos + 4, 1))
                sRst = Trim(Mid(RcvBuffer, iLCpos + 5, 7))
                sExpFlag = Mid(RcvBuffer, iLCpos + 12, 2)

                If sRst = "9.999999" Then
                    If sSignFlag = "-" Then
                        sRst = "LOWER LIMIT"
                    Else
                        sRst = "UPPER LIMIT"
                    End If
                End If
                Select Case sExpFlag
                    Case "19", "20"
                        sRst2 = "Negative"
                    Case "21", "25"
                        sRst2 = "Positive"
                    Case "24", "28"
                        sRst2 = "GrayZone"
                    Case Else
                        sRst2 = ""
                End Select
                
                If Left(sRst, 1) = "." Then
                    sRst = "0" & sRst
                End If
                
                iLCpos = iLCpos + 23
                
                iResultFlag = 1
                
                Exit Do
            Case "01"       'Result Time --> CAL, QC 일때만 전송됨
                iLCpos = iLCpos + 12
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "02"       'Control ID --> CAL, QC 일때만 전송됨
                sControlNm = Trim(Mid(RcvBuffer, iLCpos + 3, 4))
                iLCpos = iLCpos + 9
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "03"       'Standard Rates --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "04"       'Calibration Curve --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
            
            Case "07"       'ABS Sample Check --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "09"       'QC 전송 Mode
                iLCpos = iLCpos + 9
                
            Case "41"       'Slot State
                'Example "41 023 128 000 000 050<LF>"
                Exit Do
            Case "42"       'Tube Information
                Select Case m_EqName
                    Case "COBASCORE2"
                        tmpBarCd = Trim(Mid(RcvBuffer, iLCpos + 3, 9))
                    Case "INTEGRA400", "INTEGRA700"
                        tmpBarCd = Trim(Mid(RcvBuffer, iLCpos + 12, 15))
                    Case "INTEGRA800"
                        tmpBarCd = Trim(Mid(RcvBuffer, iLCpos + 14, 15))
                End Select
                
                If Len(tmpBarCd) = 0 Then
                Else
                    pSampleInfo.ID = tmpBarCd
                    
                    Select Case m_EqName
                        Case "COBASCORE2"
                            pSampleInfo.RACK = Trim(Mid(RcvBuffer, iLCpos + 13, 3))
                            pSampleInfo.POS = Trim(Mid(RcvBuffer, iLCpos + 17, 2))
                        Case "INTEGRA400", "INTEGRA700"
                            pSampleInfo.RACK = Trim(Mid(RcvBuffer, iLCpos + 3, 3))
                            pSampleInfo.POS = Trim(Mid(RcvBuffer, iLCpos + 7, 2))
                        Case "INTEGRA800"
                            pSampleInfo.RACK = Trim(Mid(RcvBuffer, iLCpos + 3, 5))
                            pSampleInfo.POS = Trim(Mid(RcvBuffer, iLCpos + 9, 2))
                    End Select
                    
                    'Order 가져오는 부분
                    Call SendOrder_CobasCore2
                End If
                
                Select Case m_EqName
                    Case "COBASCORE2"
                        iLCpos = iLCpos + 42
                    Case "INTEGRA400", "INTEGRA700"
                        iLCpos = iLCpos + 28
                    Case "INTEGRA800"
                        iLCpos = iLCpos + 30
                End Select
                
            Case "43"       'Test State
                'Example "43 032 1<LF>"
                
            Case "44"       'Cal/CS State
            
            Case "50"       'Patient ID
            
            Case "51"       'Patient Information
            
            Case "52"       'Special Order Selection
            
            Case "11"   '"53"       'Order ID
                'Version 1.0
                'slipno = Trim(Mid(msRcvBuffer, iLCpos + 3, 9))
                
                'Version 2.0
                sJNo = Trim(Mid(RcvBuffer, iLCpos + 3, 9))      '15))
                
                sRack = Trim$(Mid$(RcvBuffer, iLCpos + 13, 3))
                sPos = Trim$(Mid$(RcvBuffer, iLCpos + 17, 2))
                                
'                iLCpos = iLCpos + 41        'Sample InfoString 사용
                iLCpos = iLCpos + 20        'Sample InfoString 사용 안함
                
                'Version 1.0
                'iLCpos = iLCpos + 24  'Sample type 옵션을 No
                'iLCpos = iLCpos + 28  'Sample type 옵션을 Ok
                
                'Version 2.0
                'iLCpos = iLCpos + 30  'Sample type 옵션을 No
                'iLCpos = iLCpos + 34  'Sample type 옵션을 Ok
                
            Case "12"   '"55"       'Test ID
                sIFRstCd = Trim(Mid(RcvBuffer, iLCpos + 3, 2))  '3))
                
                iLCpos = iLCpos + 6     '7
                
            Case "59"       '"96"       'Error Code
                If iOrderFlag = 0 Then
                    'Pending Sample Request후 Response에 대한 것
                    If Mid(RcvBuffer, iLCpos + 3, 2) = "61" Then
                        iTimerFlag = 0
                    End If
                    
                    Exit Do
                Else
                'Order를 내린 후 Response에 대한 것
                    If Mid(RcvBuffer, iLCpos + 3, 2) = "00" Then
                        iErrCode = 0     'Order Input Accepted
                        Exit Do
                    Else
                        If Mid(RcvBuffer, iLCpos + 3, 2) = "22" Then
                            iErrCode = 1     'Order already available
                            Exit Do
                        ElseIf Mid(RcvBuffer, iLCpos + 3, 2) = "24" Then
                            'Test not defined - all other tests will be performed
                            iErrCode = 0
                            RaiseEvent DispMsg("일부 항목의 IF 오더코드가 잘못 설정되었습니다!!")
                            Exit Do
                        Else
                            iErrCode = 2     '기타 에러로 검사중, ID 오류, ORDER NO 오류, SAMPLE TYPE 오류 등의 에러
                            RaiseEvent DispMsg("Tx Warning : " & Mid(RcvBuffer, iLCpos + 3, 2))
                            Exit Do
                        End If
                    End If
                End If
                
            Case "98"       'Protocol Version
                RaiseEvent DispMsg("Protocol Version - " & Mid(RcvBuffer, iLCpos + 3, 4))
                Exit Do
            
            Case "99"       'General Error Code
                RaiseEvent DispMsg("Ge Warning : " & Mid(RcvBuffer, iLCpos + 3, 2))
                Exit Do
                
            Case Else
                Exit Do
        End Select
    Loop
            
'### Pending Sample Request ####################################################
    If iPendingFlag = 1 And sBC = "62" Then
        iPendingFlag = 0
    End If
            
'### CONNECTION CHECK ##########################################################
    If iIdleFlag = 1 And sBC = "00" Then
        iIdleFlag = 0
        
        'Ver 1.0
        Timer1.Interval = 10000

        'Ver 2.0
        'Timer1.Interval = 30000
    End If

'### NO MORE PENDING SAMPLE #####################################################
    If iPendingFlag = 1 And sBC = "00" Then     '"69" Then
        iPendingFlag = 0
    End If
    
    
'### ORDER INPUT RESPONSE ################################################################
    'OrdState = 1 --> From Host To Integra : Sample Order 내린 상태
    'OrdState = 2 --> From Host To Integra : Order Delete를 요청한 상태
    'OrdState = 0 --> Order 전송이 제대로 끝난 상태
    
    If sBC = "19" And iErrCode = 0 Then
        If iOrderFlag = 1 Then
            RaiseEvent DispMsg(pSampleInfo.ID & "  Order OK!")
            iOrderFlag = 0      'Order 전송이 제대로 끝난 상태
            iTimerFlag = 1
        ElseIf iOrderFlag = 2 Then
            RaiseEvent DispMsg(pSampleInfo.ID & "  Delete OK!")
            'Order 재전송
            Call SendOrder_INTEGRA
        End If
        
    ElseIf sBC = "19" And iErrCode = 1 Then
        'LineCode 22의 에러발생
        RaiseEvent DispMsg("지금 Order가 이미 존재하거나 Full(50개)인 상태입니다.!!")
            
        iTimerFlag = 0
        RcvBuffer = ""
        Call ConnectionMsg
        Exit Sub
        
    ElseIf sBC = "19" And iErrCode = 2 Then
        'LineCode 22를 제외한 에러발생
        RaiseEvent DispMsg("Order 거부!! " & _
                        "TestNo Err, Already Running, ID Err, OrderNo Err, SampleType Err 등의 에러발생...")
        
        iTimerFlag = 0
        RcvBuffer = ""
        Call ConnectionMsg
        Exit Sub
    End If
    
'### SAMPLE RESULT 보기 & 등록 #####################################################
    If Len(sJNo) > 0 And sIFRstCd <> "" Then
        If iResultFlag = 1 And sBC = "04" Then
            iResultFlag = 0
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = sJNo
                .SEQNO = ""
                .RACK = ""
                .POS = ""

                '결과값 누적
                .RSTCNT = 1
                .IFCD = sIFRstCd & Chr(124)
                .RST1 = sRst & Chr(124)
                .RST2 = sRst2 & Chr(124)
                .UNIT = "" & Chr(124)
                .FLAG = "" & Chr(124)
            End With
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
            End With

            Call Init_pResultInfo
            
            iTimerFlag = 0
            
            'Ver 1.0
            'Timer1.Interval = 6000
            
            'Ver 2.0
            Timer1.Interval = 1000
        End If
    Else
        If iResultFlag = 1 And sBC = "04" Then
            iResultFlag = 0
        End If
    End If
    
    iTimerFlag = 0
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
        iTimerFlag = 0
        RcvBuffer = ""
        Call ConnectionMsg
    End If
End Sub
Private Sub Edit_Data()
'    On Error GoTo ErrHandler
'
''<---- COBAS 장비에서 주로 사용 S --->
'    Dim sBC          As String
'    Dim sLC          As String
'    Dim iBCpos       As Integer
'    Dim iLCpos       As Integer
'
'    Dim iErrCode     As Integer
'    Dim sGeneralErrCode    As String
''<---- COBAS 장비에서 주로 사용 E --->
'
'    Dim sJDate     As String
'    Dim sJGbn      As String
'    Dim sJNo      As String
'    Dim sIFSpcCd     As String   '인터페이스시 검체코드
'    Dim sIFRstCd    As String   '인터페이스시 검사항목코드
'
'    Dim sRack      As String
'    Dim sPos       As String
'    Dim sSendBuf     As String
'
'    Dim sRst     As String
'    Dim sRst2    As String
'    Dim sExpFalg     As String
'    Dim sSignFlag    As String
'
'    Dim sTestCd       As String
'    Dim sTestNm      As String
'
'    Dim sBarCd      As String
'    Dim i           As Integer
'    Dim sTmpBuffer   As String
'    Dim sRetVal     As String
'
'    Dim bRetVal As Boolean
'
'    Dim lngRetVal As Long
'    Dim sBuf      As String
'    Dim sSvrCd    As String
'    Dim iRetVal   As Integer
'
'    iErrCode = 0
'    iBCpos = 22
'    sBC = Mid$(RcvBuffer, iBCpos, 2)
'
'    If sBC = "19" Then       '### Order Manipulation response Block ###
'        iErrCode = 99
'    ElseIf sBC = "00" Then   '### Idle Block ###
'        IdleFlag = 1
''!--- Result Output Mode 에서 Transmit 되는 것이 Samples Only로 설정시 해당없음---!
'    ElseIf sBC = "02 " Then  '### CAL Result Block ###
'    ElseIf sBC = "03 " Then  '### Control Result Block ###
''!--------------------------------------------------------------------------------!
'
'    ElseIf sBC = "04" Then   '### Patient Result Block ###
'    ElseIf sBC = "49" Then   '### Archive Manipulation response Block ###
'    ElseIf sBC = "62" Then   '### pending Sample Tubes Response Block ###
'        PendingFlag = 1
'    ElseIf sBC = "69" Then   '### No More pending Sample Tubes Response Block ###
'    End If
'
'    iLCpos = iBCpos + 5
'
'    Do
'        If Asc(Mid$(RcvBuffer, iLCpos, 1)) = 3 Then  'ETX(END OF DATA BLOCK)
'            Exit Do
'        End If
'
'        sLC = Mid$(RcvBuffer, iLCpos, 2)
'
'        Select Case sLC
'            Case "00"       'RESULT DATA
''                sSignFlag = Trim(Mid$(RcvBuffer, iLCpos + 4, 1))
'                sRst = Trim(Mid$(RcvBuffer, iLCpos + 5, 7))
'                sExpFalg = Mid$(RcvBuffer, iLCpos + 12, 2)
'
''                If sSignFlag = "-" Then
''                    If sRst = "9.999999" And Trim(sExpFalg) = "99" Then
''                        sRst = "LOWER LIMIT"
''                    Else
''                        sRst = "-" & ConvertResult_Core2(sSignFlag, sExpFalg, sRst, sIFRstCd)
''                    End If
''                Else
''                    If sRst = "9.999999" And Trim(sExpFalg) = "99" Then
''                        sRst = "UPPER LIMIT"
''                    Else
'                        sRst = ConvertResult_Core2(sRst, sIFRstCd)      '2002/7/15 yk
''                    End If
''                End If
'
'                If Left$(sRst, 1) = "." Then
'                    sRst = "0" & sRst
'                End If
'
'                Call SpecificProcessResult(sIFRstCd, sRst, sRst2)
'
''                sRst = JudgeResult1(sIFRstCd, sRst, sRst2)
'
'                ResultFlag = 1
'
'                Exit Do
'            Case "01"       'Result Time --> CAL, QC 일때만 전송됨
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
'
'            Case "02"       'Control ID --> CAL, QC 일때만 전송됨
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
'
'            Case "03"       'Standard Rates --> CAL, QC 일때만 전송됨
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
'
'            Case "04"       'Calibration Curve --> CAL, QC 일때만 전송됨
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
'
'            Case "07"       'ABS Sample Check --> CAL, QC 일때만 전송됨
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
'
'            Case "09"       'QC 전송 Mode
'                iLCpos = iLCpos + 9
'
'            Case "41"       'Slot State
'                'Example "41 023 128 000 000 050<LF>"
'                Exit Do
'
'            Case "42"       'Tube Information
'                'Integra400
'                'Example "42 001 25 1 .....BARCD.....<LF>"
'                'Integra700
'                'Example "42 001 25 1 .....BARCD.....<LF>"
'                'Integra800
'                'Example "42 K0001 25 1 .....BARCD.....<LF>"
'
'                'Cobas Core2
'                sBarCd = Trim$(Mid$(RcvBuffer, iLCpos + 3, 9))
'                If Trim(sBarCd) <> "" Then
'                    sBarCd = Edit_BarCode_DSH(1, sBarCd)
'                End If
'
'                'Integra400
'                'sBarCd = Trim$(Mid$(RcvBuffer, iLCpos + 12, 15))
'
'                'Integra700
''                sBarCd = Trim$(Mid$(RcvBuffer, iLCpos + 12, 15))
'
'                'Integra800
'                'sBarCd = Trim$(Mid$(RcvBuffer, iLCpos + 14, 15))
'
'                If Len(sBarCd) = 0 Then
'                Else
'                    gOrderTable.sSampID = sBarCd
'
'                    'CobasCore2
'                    gOrderTable.sRack = Trim$(Mid$(RcvBuffer, iLCpos + 13, 3))
'                    gOrderTable.sPos = Trim$(Mid$(RcvBuffer, iLCpos + 17, 2))
'
'                    'Integra400
'                    'gOrderTable.sRack = Trim$(Mid$(RcvBuffer, iLCpos + 3, 3))
'                    'gOrderTable.sPos = Trim$(Mid$(RcvBuffer, iLCpos + 7, 2))
'
''                    'Integra700
''                    gOrderTable.sRack = Trim$(Mid$(RcvBuffer, iLCpos + 3, 3))
''                    gOrderTable.sPos = Trim$(Mid$(RcvBuffer, iLCpos + 7, 2))
'
'                    'Integra800
'                    'gOrderTable.sRack = Trim$(Mid$(RcvBuffer, iLCpos + 3, 5))
'                    'gOrderTable.sPos = Trim$(Mid$(RcvBuffer, iLCpos + 9, 2))
'
'                    'Order 가져오는 부분
'                    Call Order_Input("B")
'                End If
'
'                'CobasCore2
'                iLCpos = iLCpos + 42
'
'                'Integra400
'                'iLCpos = iLCpos + 28
'
'                'Integra700
''                iLCpos = iLCpos + 28
'
'                'Integra800
'                'iLCpos = iLCpos + 30
'            Case "43"       'Test State
'                'Example "43 032 1<LF>"
'
'            Case "44"       'Cal/CS State
'
'            Case "50"       'Patient ID
'
'            Case "51"       'Patient Information
'
'            Case "52"       'Special Order Selection
'
'            Case "11"   '53"       'Order ID
'                'Version 1.0
'                'slipno = Trim$(Mid$(RcvBuffer, iLCpos + 3, 9))
'
'                'Version 2.0
'                sJNo = Trim$(Mid$(RcvBuffer, iLCpos + 3, 9))    '15))
'                sJNo = Edit_BarCode_DSH(1, sJNo)        '대전성모병원용 바코드 편집
'
'                gOrderTable.sRack = Trim$(Mid$(RcvBuffer, iLCpos + 13, 3))
'                gOrderTable.sPos = Trim$(Mid$(RcvBuffer, iLCpos + 17, 2))
'
'                sIFSpcCd = ""
'
'                iLCpos = iLCpos + 41
'
'                'Version 1.0
'                'iLCpos = iLCpos + 24  'Sample type 옵션을 No
'                'iLCpos = iLCpos + 28  'Sample type 옵션을 Ok
'
'                'Version 2.0
''                iLCpos = iLCpos + 30  'Sample type 옵션을 No
'                'iLCpos = iLCpos + 34  'Sample type 옵션을 Ok
'
'            Case "12"   '55"       'Test ID
'                sIFRstCd = Trim$(Mid$(RcvBuffer, iLCpos + 3, 2))    '3))
'
'                iLCpos = iLCpos + 6     '7
'
'            Case "59"   '96"       'Error Code
'                If Mid$(RcvBuffer, iLCpos + 3, 2) = "00" Then
'                    iErrCode = 0     'Order Input Accepted
'                    Exit Do
'                ElseIf Mid$(RcvBuffer, iLCpos + 3, 2) = "02" Then
'                    iLCpos = iLCpos + 7
'
'                    If Left(RcvBuffer, 2) = "58" Then
'                        If Mid$(RcvBuffer, iLCpos + 3, 2) = "04" Then
'                            iErrCode = 1     'Order already available
'                            Exit Do
'                        Else
'                            iErrCode = 2     '기타 에러로 검사중, ID 오류, ORDER NO 오류, SAMPLE TYPE 오류 등의 에러
'                            'MsgBox "Order 전송 중 에러발생!! 에러코드 : " & Mid$(RcvBuffer, LCpos + 3, 2)
'                            Exit Do
'                        End If
'                    End If
'                Else
'                    iErrCode = 2     '기타 에러로 검사중, ID 오류, ORDER NO 오류, SAMPLE TYPE 오류 등의 에러
'                    'MsgBox "Order 전송 중 에러발생!! 에러코드 : " & Mid$(RcvBuffer, LCpos + 3, 2)
'                    Exit Do
'                End If
'
''                If OrderFlag = 0 Then
''                'Pending Sample Request후 Response에 대한 것
''                    If Mid$(RcvBuffer, iLCpos + 3, 2) = "61" Then
''                        TimerFlag = 0
''                        Exit Do
''                    End If
''
''                    Exit Do
''                Else
''                'Order를 내린 후 Response에 대한 것
''                    If Mid$(RcvBuffer, iLCpos + 3, 2) = "00" Then
''                        iErrCode = 0     'Order Input Accepted
''                        Exit Do
''                    Else
''                        If Mid$(RcvBuffer, iLCpos + 3, 2) = "22" Then
''                            iErrCode = 1     'Order already available
''                            Exit Do
''                        ElseIf Mid$(RcvBuffer, iLCpos + 3, 2) = "24" Then
''                            'Test not defined - all other tests will be performed
''                            iErrCode = 0
''                            ViewMsgLog "일부 항목의 IF 오더코드가 잘못 설정되었습니다!!"
''                            Exit Do
''                        Else
''                            iErrCode = 2     '기타 에러로 검사중, ID 오류, ORDER NO 오류, SAMPLE TYPE 오류 등의 에러
''                            ViewMsgLog "Tx Warning : " & Mid$(RcvBuffer, iLCpos + 3, 2)
''                            Exit Do
''                        End If
''                    End If
''                End If
'
'            Case "98"       'Protocol Version
'                MsgBox Mid$(RcvBuffer, iLCpos + 3, 4)
'                Exit Do
'
'            Case "99"       'General Error Code
'                sGeneralErrCode = Mid$(RcvBuffer, iLCpos + 3, 2)
'                ViewMsgLog "Ge Warning : " & sGeneralErrCode
'                Exit Do
'
'            Case Else
'                Exit Do
'        End Select
'    Loop
'
''### Pending Sample Request ####################################################
'    If PendingFlag = 1 And sBC = "62" Then
''        PendingFlag = 0
'    End If
'
''### CONNECTION CHECK ##########################################################
'    If IdleFlag = 1 And sBC = "00" Then
'        IdleFlag = 0
'
'        'Ver 1.0
'        Timer1.Interval = 10000
'        Timer1.Enabled = True
'
'        'Ver 2.0
'        'Timer1.Interval = 30000
'    End If
'
''### NO MORE PENDING SAMPLE #####################################################
'    If PendingFlag = 1 And sBC = "00" Then  '69" Then
'        PendingFlag = 0
'    End If
'
''### ORDER INPUT RESPONSE ################################################################
'    'OrderFlag = 1 --> From Host To Integra : Sample Order 내린 상태
'    'OrderFlag = 2 --> From Host To Integra : Order Delete를 요청한 상태
'    'OrderFlag = 0 --> Order 전송이 제대로 끝난 상태
'
'    If sBC = "19" And iErrCode = 0 Then
'        If OrderFlag = 1 Then
'            ViewMsg gOrderTable.sSampID & "   Order OK!"
'            OrderFlag = 0   'Order 전송이 제대로 끝난 상태
'
'            '''Call DisplayOrderOK
'
'            TimerFlag = 0
'        ElseIf OrderFlag = 2 Then
'            ViewMsg gOrderTable.sSampID & "   Delete OK!"
'
'            Call Order_Input
'        End If
'    ElseIf sBC = "19" And iErrCode = 1 Then
''        'LineCode 22의 에러발생
''        ViewMsgLog "지금 Order가 이미 존재하거나 Full(50개)인 상태입니다.!!"
'
'        TimerFlag = 0
'        RcvBuffer = ""
'        Call cmdInitial.DoClick
'
'        Exit Sub
'    ElseIf sBC = "19" And iErrCode = 2 Then
'        'LineCode 22를 제외한 에러발생
'        ViewMsgLog "Order 거부!! " & _
'           "TestNo Err, Already Running, ID Err, OrderNo Err, SampleType Err 등의 에러발생..."
'
'        TimerFlag = 0
'        RcvBuffer = ""
'        Call cmdInitial.DoClick
'
'        Exit Sub
'    End If
'
''### SAMPLE RESULT 보기 & 등록 #####################################################
'    If Len(sJNo) > 0 And sIFRstCd <> "" Then
'        If ResultFlag = 1 And sBC = "04" Then
'            ResultFlag = 0
'
'            Call DisplayResultOK(3, Format(dtpLabDate.Value, "YYYYMMDD"), "", _
'                                            "", "", sJNo, "", "", "", "", "", "", "", "", _
'                                            1, sIFRstCd & Chr$(124), sRst & Chr$(124), sRst2 & Chr$(124), _
'                                            "", "")
'
'            TimerFlag = 0
'
'            'Ver 1.0
'            'Timer1.Interval = 6000
'
'            'Ver 2.0
'            Timer1.Interval = 10000
'            Timer1.Enabled = True
'        End If
'    Else
'        If ResultFlag = 1 And sBC = "04" Then
'            ResultFlag = 0
'            Timer1.Enabled = True
'        End If
'    End If
'
'    TimerFlag = 0
'
'    Exit Sub
'ErrHandler:
'    ViewMsg "Edit_Data 에러 발생" & "(" & CStr(Err.Number) & " : " & Err.Description & ")"
'
'    TimerFlag = 0
'
'    RcvBuffer = ""
'    cmdInitial.DoClick
End Sub

Private Sub EditMessage()
'
'    Dim RtnCd            As Integer
'    Dim BlCode           As String               ' Block Code
'    Dim LiCode           As String               ' Line  Code
'    Dim BcAddr           As Integer              ' Block Code Address in Receive Buffer
'    Dim LcAddr           As Integer              ' Line  Code Address in Receive Buffer
'    Dim RackNo           As String               ' Rack No
'    Dim PosiNo           As String               ' Postion No
'    Dim TestNm           As String               ' Test Number
'    Dim Result           As String               ' Result
'    Dim CutF             As String
'    Dim ErCode           As Integer              ' Error Code(only BC 19)
'    Dim wk1              As String
'    Dim wk2              As Long
'    Dim BarCode          As String
'    Dim PoNeGr           As String               ' Postive, Negative, Gray Zone Result
'    Dim EditPoNeGr       As String
'    Dim i, ix            As Integer
'    Dim SendBuf          As String
'    Dim SxYmd            As String
'    Dim SlipCd           As String
'    Dim LabNo            As String
'
'    Dim TestName    As String       '검사명
'
'    BcAddr = 23
'    BlCode = Mid$(RcvBuffer, BcAddr, 2)             ' Get Block Code
'    '--- Order manipulation response
'    If BlCode = "19" Then
'       ErCode = 99                                  ' initial Value
'    End If
'    LcAddr = BcAddr + 5
'
'    Do
'        If Asc(Mid$(RcvBuffer, LcAddr, 1)) = 3 Then  ' ETX(Data Block End)?
'            Exit Do
'        End If
'
'        LiCode = Mid$(RcvBuffer, LcAddr, 2)          ' Get Line  Code
'
'        Select Case LiCode
'            Case "00"                                      ' Final Result
'                Result = Mid$(RcvBuffer, LcAddr + 4, 7)    ' Result(Od)
'                PoNeGr = Mid$(RcvBuffer, LcAddr + 12, 2)   ' Positive, Negative , gray Zone
'                LcAddr = LcAddr + 23
'                ResltFlg = 1
'                Exit Do
'            Case "03"
'                LcAddr = LcAddr + 31
'            Case "06"
'                ResltFlg = 1
'                Exit Do
'            Case "09"
'                LcAddr = LcAddr + 9
'            Case "11"                                       ' Sample IDentification
'                RackNo = Mid$(RcvBuffer, LcAddr + 13, 3)   ' Rack No
'                PosiNo = Mid$(RcvBuffer, LcAddr + 17, 2)   ' Position No
'                SampNo = Mid$(RcvBuffer, LcAddr + 3, 9)
'
'                If Len(Trim(SampNo)) < 8 Then   '9 Then     'Q.C는 길이가 8
'                    SampNo = ""
'                End If
'                LcAddr = LcAddr + 41
'            Case "12"                                       ' Test Number
'                TestNm = Mid$(RcvBuffer, LcAddr + 3, 2)     ' Test Number
'                LcAddr = LcAddr + 6
'            Case "42"                                       ' Sample ID
'                RackNo = Mid$(RcvBuffer, LcAddr + 13, 3)   ' Rack No
'                PosiNo = Mid$(RcvBuffer, LcAddr + 17, 2)   ' Position No
'                SampNo = Mid$(RcvBuffer, LcAddr + 20, 13)  ' Bar Code
'
'                LcAddr = LcAddr + 42
'
'                Call SAMPLE_INFO(SampNo)
'
'                With frmSCobas.spdSCobas
'                    Call .SetText(7, .MaxRows, RackNo)
'                    Call .SetText(8, .MaxRows, PosiNo)
'                End With
'                Exit Do
'            Case "59"                                          ' Order manipulation response
'                If Mid$(RcvBuffer, LcAddr + 3, 2) = "00" Then ' Order Check
'                    ErCode = 0
'                End If
'                Exit Do
'        End Select
'    Loop
'
'    '--- Pending Sample ID ?
'    If PendFlag = 1 And BlCode = "62" Then
'        ' *-------------------------------------------------------------------*
'        ' *  Panding sample request                                 (BC 60)   *
'        ' *-------------------------------------------------------------------*
'        SendBuf = Chr(1) & Chr(10)                              ' SOH & LF
'        SendBuf = SendBuf & "06 HOSTNAMESIXTEENX 60" & Chr(10)  ' Header Block
'        SendBuf = SendBuf & Chr(2) & Chr(10)                    ' STX & LF
'        SendBuf = SendBuf & "40 1" & Chr(10)                    ' Data Block(LC 40)
'        SendBuf = SendBuf & Chr(3) & Chr(10)                    ' ETX & LF
'        SendBuf = SendBuf & Chr(4) & Chr(10)                    ' EOT & LF
'        Comm1.Output = SendBuf
'
'        Do
'           'DoEvents
'        Loop Until Comm1.OutBufferCount = 0
'    End If
'
'    '--- Connection Check
'    If IdleFlg = 1 And BlCode = "00" Then
'        IdleFlg = 0
'        Timer1.Enabled = True                                    ' Timer Enable
'    End If
'    '=== Connection Check(yk 8/28)
'    If IdleFlg = 0 And BlCode = "00" And ResltFlg = 0 Then
'        Timer1.Enabled = True
'    End If
'
'    '--- No more Sample ID ?
'    If PendFlag = 1 And BlCode = "00" Then
'        PendFlag = 0
'    End If
'
'    '--- Order Input
'    If OrderFlg = 1 And BlCode = "19" Then
'        OrderCnt = OrderCnt + 1
'        If OrderCnt <= frmSCobas.spdSCobas.MaxRows Then
'            Call OrderSend
'        Else
'            OrderFlg = 0
'        End If
'    End If
'
'    '--- Sample Result ?
'    If Trim(SampNo) <> "" Then
'        If ResltFlg = 1 And BlCode = "04" Then
'
'            SampNo = Edit_ID(Trim(SampNo))
'
'            CutF = TestNameTable(Val(TestNm)).CutF
'            TestName = TestNameTable(Val(TestNm)).Name
'
'            Select Case PoNeGr
'                Case "19", "20"
'                    EditPoNeGr = "Negative"
'                Case "21", "25"
'                    EditPoNeGr = "Positive"
'                Case "24", "28"
'                    EditPoNeGr = "GrayZone"
'                Case Else
'                    EditPoNeGr = ""
'            End Select
'
'            With spdResult
'                .MaxRows = .MaxRows + 1
'                Call .SetText(1, .MaxRows, SampNo)
'                Call .SetText(2, .MaxRows, RackNo)
'                Call .SetText(3, .MaxRows, PosiNo)
'                Call .SetText(4, .MaxRows, TestName)
'                Call .SetText(5, .MaxRows, EditPoNeGr & Space(1) & Result)
'                Call .SetText(6, .MaxRows, CutF)
'            End With
'
'            '--- 일반/Q.C 구분하여 Local DB에 결과 입력
'            If Mid(SampNo, 9, 1) = "Q" Then     'Q.C
'                TestHeaderTable.QC_Level = Get_QCLevel(SampNo, TestNameTable(Val(TestNm)).code)
'                Call Add_Db_QC(SampNo, TestNameTable(Val(TestNm)).code, _
'                                TestHeaderTable.QC_Level, Result)
'            Else    '일반
'                If RecordExists(TbResult, "PrimaryKey", SampNo, TestNm) Then
'                    TbResult.Edit
'                Else
'                    TbResult.AddNew
'                    TbResult!Lab_ID = SampNo
'                    TbResult!EqCode = Trim(TestNm)
'                End If
'                TbResult!Result = EditPoNeGr
'                TbResult!Od = Result
'                TbResult!CutF = CutF
'                TbResult.Update
'                TbResult.MoveLast
'                '--- Local DB에 Sample ID입력
'                If RecordExist(TbSample, "PrimaryKey", SampNo) <> True Then
'                    TbSample.AddNew
'                    TbSample!Lab_ID = SampNo
'                    TbSample!ReckNo = Format(Val(Trim(RackNo)), "000")
'                    TbSample!PosiNo = Format(Val(Trim(PosiNo)), "00")
'                    TbSample.Update
'                    TbSample.MoveLast
'                End If
'            End If
'            '----------------------------
'
'            ResltFlg = 0
'            Timer1.Enabled = True                                 ' Timer Enable
'        End If
'    Else
'        If ResltFlg = 1 And BlCode = "04" Then
'            ResltFlg = 0                                          ' Timer Enable
'            Timer1.Enabled = True
'        End If
'    End If
'
'    '--- No more Sample Result ?
'    If ResltFlg = 1 And BlCode = "00" Then
'        ResltFlg = 0                                          ' Timer Enable
'        Timer1.Enabled = True
'    End If
'    If ResltFlg = 1 And BlCode = "01" Then
'        ResltFlg = 0                                          ' Timer Enable
'        Timer1.Enabled = True
'    End If
End Sub
Private Sub EditData_INTEGRA()
    On Error GoTo ErrRtn
    
'<---- COBAS 장비에서 주로 사용 S --->
    Dim sBC          As String
    Dim sLC          As String
    Dim iBCpos       As Integer
    Dim iLCpos       As Integer
    
    Dim iErrCode     As Integer
'<---- COBAS 장비에서 주로 사용 E --->

    Dim tmpBarCd$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$
    
    Dim sRst     As String
    Dim sRst2    As String
    Dim sExpFlag    As String
    Dim sSignFlag   As String
    Dim sIFRstCd    As String
    
    Dim sControlNm  As String
    Dim sJNo$
    
    iErrCode = 0
    iBCpos = 22
    sBC = Mid(RcvBuffer, iBCpos, 2)
    
    Select Case sBC
        '### Idle Block, No more result Block ###
        Case "00"
            iIdleFlag = 1
        
        '### CAL Result Block ###
        Case "02"
        
        '### Control Result Block ###
        Case "03"
        
        '### Patient Result Block ###
        Case "04"
        
        '### Order Manipulation response Block ###
        Case "19"
            iErrCode = 99
        
        '### pending Sample Tubes Response Block ###
        Case "62"
            iPendingFlag = 1
            
        '### No More pending Sample Tubes Response Block ###
        Case "69"

        Case Else
        
    End Select
    
    iLCpos = iBCpos + 5
    
    Do
        If Asc(Mid(RcvBuffer, iLCpos, 1)) = 3 Then  'ETX(END OF DATA BLOCK)
            Exit Do
        End If

        sLC = Mid(RcvBuffer, iLCpos, 2)
        
        Select Case sLC
            Case "00"       'RESULT DATA
                sSignFlag = Trim(Mid(RcvBuffer, iLCpos + 3, 1))
                sRst = Trim(Mid(RcvBuffer, iLCpos + 4, 8))
                sExpFlag = Mid(RcvBuffer, iLCpos + 12, 4)
                
                If sSignFlag = "-" Then
                    If sRst = "9.999999" And Mid(sExpFlag, 3, 2) = "99" Then
                        sRst = "LOWER LIMIT"
                    Else
                        sRst = "-" & ConvertResult1(Mid(sExpFlag, 2, 1), Mid(sExpFlag, 3, 2), sRst, sIFRstCd)
                    End If
                Else
                    If sRst = "9.999999" And Mid(sExpFlag, 3, 2) = "99" Then
                        sRst = "UPPER LIMIT"
                    Else
                        sRst = ConvertResult1(Mid(sExpFlag, 2, 1), Mid(sExpFlag, 3, 2), sRst, sIFRstCd)
                    End If
                End If
                
                If Left(sRst, 1) = "." Then
                    sRst = "0" & sRst
                End If
                
                iResultFlag = 1
                
                Exit Do
            Case "01"       'Result Time --> CAL, QC 일때만 전송됨
                iLCpos = iLCpos + 12
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "02"       'Control ID --> CAL, QC 일때만 전송됨
                sControlNm = Trim(Mid(RcvBuffer, iLCpos + 3, 4))
                iLCpos = iLCpos + 9
'                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "03"       'Standard Rates --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "04"       'Calibration Curve --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
            
            Case "07"       'ABS Sample Check --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "41"       'Slot State
                'Example "41 023 128 000 000 050<LF>"
                Exit Do
            Case "42"       'Tube Information
                'Integra400
                'Example "42 001 25 1 .....BARCD.....<LF>"
                'Integra700
                'Example "42 001 25 1 .....BARCD.....<LF>"
                'Integra800
                'Example "42 K0001 25 1 .....BARCD.....<LF>"
                
                Select Case m_EqName
                    Case "INTEGRA400", "INTEGRA700"
                        tmpBarCd = Trim(Mid(RcvBuffer, iLCpos + 12, 15))
                    Case "INTEGRA800"
                        tmpBarCd = Trim(Mid(RcvBuffer, iLCpos + 14, 15))
                End Select
                
                If Len(tmpBarCd) = 0 Then
                Else
                    pSampleInfo.ID = tmpBarCd
                    
                    Select Case m_EqName
                        Case "INTEGRA400", "INTEGRA700"
                            pSampleInfo.RACK = Trim(Mid(RcvBuffer, iLCpos + 3, 3))
                            pSampleInfo.POS = Trim(Mid(RcvBuffer, iLCpos + 7, 2))
                        Case "INTEGRA800"
                            pSampleInfo.RACK = Trim(Mid(RcvBuffer, iLCpos + 3, 5))
                            pSampleInfo.POS = Trim(Mid(RcvBuffer, iLCpos + 9, 2))
                    End Select
                    
                    'Order 가져오는 부분
                    Call SendOrder_INTEGRA
                End If
                
                Select Case m_EqName
                    Case "INTEGRA400", "INTEGRA700"
                        iLCpos = iLCpos + 28
                    Case "INTEGRA800"
                        iLCpos = iLCpos + 30
                End Select
                
            Case "43"       'Test State
                'Example "43 032 1<LF>"
                
            Case "44"       'Cal/CS State
            
            Case "50"       'Patient ID
            
            Case "51"       'Patient Information
            
            Case "52"       'Special Order Selection
            
            Case "53"       'Order ID
                'Version 1.0
                'slipno = Trim(Mid(msRcvBuffer, iLCpos + 3, 9))
                
                'Version 2.0
                sJNo = Trim(Mid(RcvBuffer, iLCpos + 3, 15))
                
                'Version 1.0
                'iLCpos = iLCpos + 24  'Sample type 옵션을 No
                'iLCpos = iLCpos + 28  'Sample type 옵션을 Ok
                
                'Version 2.0
                iLCpos = iLCpos + 30  'Sample type 옵션을 No
                'iLCpos = iLCpos + 34  'Sample type 옵션을 Ok
                
            Case "55"       'Test ID
                sIFRstCd = Trim(Mid(RcvBuffer, iLCpos + 3, 3))
                
                iLCpos = iLCpos + 7
                
            Case "96"       'Error Code
                If iOrderFlag = 0 Then
                    'Pending Sample Request후 Response에 대한 것
                    If Mid(RcvBuffer, iLCpos + 3, 2) = "61" Then
                        iTimerFlag = 0
                    End If
                    
                    Exit Do
                Else
                'Order를 내린 후 Response에 대한 것
                    If Mid(RcvBuffer, iLCpos + 3, 2) = "00" Then
                        iErrCode = 0     'Order Input Accepted
                        Exit Do
                    Else
                        If Mid(RcvBuffer, iLCpos + 3, 2) = "22" Then
                            iErrCode = 1     'Order already available
                            Exit Do
                        ElseIf Mid(RcvBuffer, iLCpos + 3, 2) = "24" Then
                            'Test not defined - all other tests will be performed
                            iErrCode = 0
                            RaiseEvent DispMsg("일부 항목의 IF 오더코드가 잘못 설정되었습니다!!")
                            Exit Do
                        Else
                            iErrCode = 2     '기타 에러로 검사중, ID 오류, ORDER NO 오류, SAMPLE TYPE 오류 등의 에러
                            RaiseEvent DispMsg("Tx Warning : " & Mid(RcvBuffer, iLCpos + 3, 2))
                            Exit Do
                        End If
                    End If
                End If
                
            Case "98"       'Protocol Version
                RaiseEvent DispMsg("Protocol Version - " & Mid(RcvBuffer, iLCpos + 3, 4))
                Exit Do
            
            Case "99"       'General Error Code
                RaiseEvent DispMsg("Ge Warning : " & Mid(RcvBuffer, iLCpos + 3, 2))
                Exit Do
                
            Case Else
                Exit Do
        End Select
    Loop
            
'### Pending Sample Request ####################################################
    If iPendingFlag = 1 And sBC = "62" Then
        iPendingFlag = 0
    End If
            
'### CONNECTION CHECK ##########################################################
    If iIdleFlag = 1 And sBC = "00" Then
        iIdleFlag = 0
        
        'Ver 1.0
        Timer1.Interval = 10000

        'Ver 2.0
        'Timer1.Interval = 30000
    End If

'### NO MORE PENDING SAMPLE #####################################################
    If iPendingFlag = 1 And sBC = "69" Then
        iPendingFlag = 0
    End If
    
    
'### ORDER INPUT RESPONSE ################################################################
    'OrdState = 1 --> From Host To Integra : Sample Order 내린 상태
    'OrdState = 2 --> From Host To Integra : Order Delete를 요청한 상태
    'OrdState = 0 --> Order 전송이 제대로 끝난 상태
    
    If sBC = "19" And iErrCode = 0 Then
        If iOrderFlag = 1 Then
            RaiseEvent DispMsg(pSampleInfo.ID & "  Order OK!")
            iOrderFlag = 0      'Order 전송이 제대로 끝난 상태
            iTimerFlag = 1
        ElseIf iOrderFlag = 2 Then
            RaiseEvent DispMsg(pSampleInfo.ID & "  Delete OK!")
            'Order 재전송
            Call SendOrder_INTEGRA
        End If
        
    ElseIf sBC = "19" And iErrCode = 1 Then
        'LineCode 22의 에러발생
        RaiseEvent DispMsg("지금 Order가 이미 존재하거나 Full(50개)인 상태입니다.!!")
            
        iTimerFlag = 0
        RcvBuffer = ""
        Call ConnectionMsg
        Exit Sub
        
    ElseIf sBC = "19" And iErrCode = 2 Then
        'LineCode 22를 제외한 에러발생
        RaiseEvent DispMsg("Order 거부!! " & _
                        "TestNo Err, Already Running, ID Err, OrderNo Err, SampleType Err 등의 에러발생...")
        
        iTimerFlag = 0
        RcvBuffer = ""
        Call ConnectionMsg
        Exit Sub
    End If
    
'### SAMPLE RESULT 보기 & 등록 #####################################################
    If Len(sJNo) > 0 And sIFRstCd <> "" Then
        If iResultFlag = 1 And sBC = "04" Then
            iResultFlag = 0
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = sJNo
                .SEQNO = ""
                .RACK = ""
                .POS = ""

                '결과값 누적
                .RSTCNT = 1
                .IFCD = sIFRstCd & Chr(124)
                .RST1 = sRst & Chr(124)
                .RST2 = "" & Chr(124)
                .UNIT = "" & Chr(124)
                .FLAG = "" & Chr(124)
            End With
            
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
            End With

            Call Init_pResultInfo
            
            iTimerFlag = 0
            
            'Ver 1.0
            'Timer1.Interval = 6000
            
            'Ver 2.0
            Timer1.Interval = 1000
        End If
    Else
        If iResultFlag = 1 And sBC = "04" Then
            iResultFlag = 0
        End If
    End If
    
    iTimerFlag = 0
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
        iTimerFlag = 0
        RcvBuffer = ""
        Call ConnectionMsg
    End If
End Sub

Private Function ConvertResult1(ByVal sSign As String, ByVal sExp As String, ByVal sRst As String, ByVal sIFRstCd As String, Optional ByVal sIFSeq As String) As String
    Dim i%
    Dim sValue$
    
    If IsNumeric(sRst) = False Then
        ConvertResult1 = sRst
        Exit Function
    End If
    
    If sSign = "" Then
        sSign = "+"
    End If
    
    If sSign = "+" Then
        sValue = CStr(Val(sRst) * (10 ^ Val(sExp)))
    ElseIf sSign = "-" Then
        sValue = CStr(Val(sRst) / (10 ^ Val(sExp)))
    End If
    
    If Left(sValue, 1) = "." Then
        sValue = "0" & sValue
    End If
    
    ConvertResult1 = sValue
    
End Function

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

Private Sub SendOrder_CobasCore2()
    On Error GoTo ErrRtn
    
    Dim sTmp    As String
    Dim i       As Integer
    Dim SendBuff    As String
    Dim tmpDate As String
    
    tmpDate = Format(Now, "YYYYMMDD")
    
    If Len(pSampleInfo.RACK) = 5 Then       'For Integra800
        pSampleInfo.RACK = Mid(pSampleInfo.RACK, 2)
    End If
    
    '----- 검사항목 조회/편집
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    
    Call Get_OrderString
    
    If pSampleInfo.ORDCNT = 0 Then
        RaiseEvent DispMsg("인터페이스 오더 항목이 존재하지 않습니다!!")
        
        Exit Sub
    End If
    
    SendBuff = ""

    'Order Packet 구성
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
    
    Select Case m_EqName
        Case "COBASCORE2"
            SendBuff = SendBuff & "06" & " " & "HOSTNAMESIXTEENX" & " " & "10" & Chr(10)     '<LF>
        Case "INTEGRA400"
            SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "10" & Chr(10)     '<LF>
        Case "INTEGRA700"
            SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "10" & Chr(10)     '<LF>
        Case "INTEGRA800"
            SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "10" & Chr(10)     '<LF>
    End Select
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
    SendBuff = SendBuff & "50" & " R " & Left(Trim(pSampleInfo.ID) & Space(10), 10) & Chr(10)    '<LF>
    
'    'Sample Type No
'    SendBuff = SendBuff & "53" & " " & pSampleInfo.ID & String(15 - Len(Trim(pSampleInfo.ID)), " ") & _
'                          " " & Right(tmpDate, 2) & "/" & Mid(tmpDate, 5, 2) & "/" & Left(tmpDate, 4) & _
'                          Chr(10)      '<LF>
   
    'Order 편집
    For i = 1 To pSampleInfo.ORDCNT
'        SendBuff = SendBuff & "52" & " " & String(3 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & Chr(10)
        SendBuff = SendBuff & "52" & " " & Format(pSampleInfo.IFCD(i), "00") & Chr(10)
    Next i
        
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
        
    msComm.Output = SendBuff
    
    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuff)
    End If
    
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    iOrderFlag = 1
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub
Private Sub SendOrder_INTEGRA()
    On Error GoTo ErrRtn
    
    Dim sTmp    As String
    Dim i       As Integer
    Dim SendBuff    As String
    Dim tmpDate As String
    
    tmpDate = Format(Now, "YYYYMMDD")
    
    If Len(pSampleInfo.RACK) = 5 Then       'For Integra800
        pSampleInfo.RACK = Mid(pSampleInfo.RACK, 2)
    End If
    
    '----- 검사항목 조회/편집
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    
    Call Get_OrderString
    
    If pSampleInfo.ORDCNT = 0 Then
        RaiseEvent DispMsg("인터페이스 오더 항목이 존재하지 않습니다!!")
        
        Exit Sub
    End If
    
    SendBuff = ""

    'Order Packet 구성
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
    
    Select Case m_EqName
        Case "INTEGRA400"
            SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "10" & Chr(10)     '<LF>
        Case "INTEGRA700"
            SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "10" & Chr(10)     '<LF>
        Case "INTEGRA800"
            SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "10" & Chr(10)     '<LF>
    End Select
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
    SendBuff = SendBuff & "50" & " " & String(15, " ") & Chr(10)     '<LF>
    
    'Sample Type No
    SendBuff = SendBuff & "53" & " " & pSampleInfo.ID & String(15 - Len(Trim(pSampleInfo.ID)), " ") & _
                          " " & Right(tmpDate, 2) & "/" & Mid(tmpDate, 5, 2) & "/" & Left(tmpDate, 4) & _
                          Chr(10)      '<LF>
        
    'Barcode type
    Select Case m_EqName
        Case "INTEGRA400", "INTEGRA700"
            SendBuff = SendBuff & "54" & " " & "000 00" & _
                                " " & "A" & " " & Space(21) & _
                                " " & Space(21) & Chr(10)    '<LF>
        Case "INTEGRA800"
            SendBuff = SendBuff & "54" & " " & "00000 00" & _
                                " " & "A" & " " & Space(21) & _
                                " " & Space(21) & Chr(10)    '<LF>
    End Select
    
    'Order 편집
    For i = 1 To pSampleInfo.ORDCNT
        SendBuff = SendBuff & "55" & " " & String(3 - Len(pSampleInfo.IFCD(i)), " ") & pSampleInfo.IFCD(i) & Chr(10)
    Next i
        
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
        
    msComm.Output = SendBuff
    
    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuff)
    End If
    
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    iOrderFlag = 1
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
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

Private Sub cmdTest_Click()

    WkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            WkBuf = msComm.Input
            
            If sTestMode = "77" Then
                RaiseEvent PrintRcvLog(WkBuf)
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
    
    Call RequestResultMsg
    
End Sub
Private Sub RequestPendingMsg()
    On Error GoTo ErrHandler
    
    Dim SendBuff    As String

    If iTimerFlag = 1 Then
        Exit Sub
    End If
    
 '########### PENDING BARCODE SAMPLES REQUEST ######################
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
    
    Select Case m_EqName
        Case "COBASCORE2"
            SendBuff = SendBuff & "06" & " " & "HOSTNAMESIXTEENX" & " " & "60" & Chr(10)
        Case "INTEGRA400"
            SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "60" & Chr(10)
        Case "INTEGRA700        "
            SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "60" & Chr(10)
        Case "INTEGRA800"
            SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "60" & Chr(10)
    End Select
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
    SendBuff = SendBuff & "40" & " " & "1" & Chr(10)
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
    
    iTimerFlag = 1
    
    msComm.Output = SendBuff
    
    If sTestMode = 77 Then
        RaiseEvent PrintSendLog(SendBuff)
    End If
    
    Exit Sub
ErrHandler:
    iTimerFlag = 0
End Sub
Private Sub RequestResultMsg()
    On Error GoTo ErrHandler
    
    Dim SendBuff    As String

    If iTimerFlag = 1 Then
        Exit Sub
    End If
    
 '########### ALL TYPES OF FINAL RESULTS ARE TRANSFFERD TO THE HOST ######################
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>

    Select Case m_EqName
        Case "COBASCORE2"
'            SendBuff = SendBuff & "06" & " " & "HOSTNAMESIXTEENX" & " " & "09" & Chr(10)
            SendBuff = SendBuff & "06" & " " & "COBAS CORE II   " & " " & "09" & Chr(10)
        Case "INTEGRA400"
            SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "09" & Chr(10)
        Case "INTEGRA700"
            SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "09" & Chr(10)
        Case "INTEGRA800"
            SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "09" & Chr(10)
    End Select
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
'    SendBuff = SendBuff & "10" & " " & "01" & Chr(10)
    SendBuff = SendBuff & "10" & " " & "09" & Chr(10)
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
    
    iTimerFlag = 1
    
    msComm.Output = SendBuff
    
    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuff)
    End If

    Exit Sub
ErrHandler:
    iTimerFlag = 0
End Sub

Private Sub Timer2_Timer()

    Call RequestPendingMsg
    
End Sub

Private Sub Timer3_Timer()

    If iTimerFlag = 1 Then
        Call ConnectionMsg
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
    
    On Error GoTo ErrPortOpen
    If m_PortOpen = True Then
        msComm.PortOpen = True
    End If
    On Error GoTo 0
    
    'For Integra
    iTimerFlag = 0
    iIdleFlag = 0
    iPendingFlag = 0
    iOrderFlag = 0
    iResultFlag = 0
    
    Call ConnectionMsg
    
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
'
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function ConnectionMsg() As Variant
    On Error GoTo ErrHandler
    
    Dim SendBuff    As String

 '########### CONNECTION ESTABLISH ######################
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
    
    Select Case m_EqName
        Case "COBASCORE2"
            SendBuff = SendBuff & "06" & " " & "HOSTNAMESIXTEENX" & " " & "00" & Chr(10)
        Case "INTEGRA400"
            SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "00" & Chr(10)
        Case "INTEGRA700"
            SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "00" & Chr(10)
        Case "INTEGRA800"
            SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "00" & Chr(10)
    End Select
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
    
    iTimerFlag = 1
    
    msComm.Output = SendBuff
    
    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuff)
    End If
        
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("ConnectionMsg Err - " & Err.Description)
    End If
End Function
