VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl HITACHI 
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
Attribute VB_Name = "HITACHI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
'Const m_def_iStartSampleNo = 0
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
'Dim m_iStartSampleNo As Integer
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$)
Event RaiseError(sError$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
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
        Case "HITACHI7600"
            If m_bUseBarcode = True Then
                '바코드 사용
                Call PhaseCfg_Protocol_HITACHI7600
            Else
                '바코드 사용 안할 경우...
                Call PhaseCfg_Protocol_HITACHI7600_Batch
            End If
        
        Case "HITACHI7020"
            Call PhaseCfg_Protocol_HITACHI7020      '바코드 사용안함
            
        Case "HITACHI747"
            Call PhaseCfg_Protocol_HITACHI747       '바코드 사용
        
        Case "HITACHI7180"
            If m_bUseBarcode = True Then
                Call PhaseCfg_Protocol_HITACHI7180
            Else
                Call PhaseCfg_Protocol_HITACHI7180_Batch
            End If
            
        Case "HITACHI7170"
            Call PhaseCfg_Protocol_HITACHI7170
            
        Case "HITACHI7150"          '단방향
            Call PhaseCfg_Protocol_HITACHI7150
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_HITACHI7180_Batch()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
           Case 2
              RcvBuffer = ""
           
           Case 3

           Case 10, 13
              Call DataEditResponse_Hitachi7180_Batch
              
           Case 21

           Case Else
              RcvBuffer = RcvBuffer + wkDat
        End Select
    Next ix1
    
End Sub
Private Sub DataEditResponse_Hitachi7180_Batch()
    On Error GoTo ErrRtn
    
    Dim sFrame  As String
    Dim sFunc   As String
    Dim i       As Integer
    Dim StartPtr    As Integer
    Dim tmpSeqNo$, tmpID$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarm$
    Dim iTestCnt    As Integer
    Dim sRstData    As String
    
    sFrame = Left(RcvBuffer, 1)
   
    Select Case sFrame
        '<REP>
        Case "?"
            '<MOR> 전송
            Call Send_MOR
        
        '<ANY>
        Case ">"
            '<MOR> 전송
            Call Send_MOR
        
        '<SPE>
        Case ";"
            '<SPE> 전송
'            sFunc = Mid(RcvBuffer, 2, 40)
            sFunc = Mid(RcvBuffer, 2, 12) & String(13, "#") & Mid(RcvBuffer, 27, 15)
            
            tmpSeqNo = Mid(RcvBuffer, 4, 5)
            tmpRack = Mid(RcvBuffer, 9, 1)
            tmpPos = Mid(RcvBuffer, 10, 3)
            tmpID = Trim(Mid(RcvBuffer, 14, 13))
            
            'ID 저장
            With pSampleInfo
                .ID = Trim(tmpID)
                .SEQNO = Trim(tmpSeqNo)
                .RACK = Trim(tmpRack)
                .POS = Trim(tmpPos)
            End With
            
            '--- ORDER 조회/전송
            Call SendOrder_Hitachi7180_Batch(sFunc, pSampleInfo.SEQNO)
        
        '<END>
        Case ":"
            '변수 초기화
            Call Init_pResultInfo
            
            sFunc = Mid$(RcvBuffer, 2, 1)      ' Function
            
            If sFunc = "K" Or sFunc = "L" Then
               '<MOR> 전송
                Call Send_MOR
                Exit Sub
            End If
    
'            If sFunc <> "@" And sFunc <> "N" And sFunc <> "M" Then
            If sFunc <> "@" And sFunc <> "M" Then
                tmpSeqNo = Trim(Mid(RcvBuffer, 4, 5))
                tmpRack = Trim(Mid(RcvBuffer, 9, 1))
                tmpPos = Trim(Mid(RcvBuffer, 10, 3))
                tmpID = Trim(Mid(RcvBuffer, 14, 13))
                
                If tmpID = "" Then
                    tmpID = tmpSeqNo
                End If
                
                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = tmpID
                    .SEQNO = tmpSeqNo
                    .RACK = tmpRack
                    .POS = tmpPos
                End With
                
                iTestCnt = Val(Trim(Mid(RcvBuffer, 48, 3)))
                
                StartPtr = 51

                For i = 1 To iTestCnt
                    sRstData = Mid(RcvBuffer, StartPtr, 10)
                    
                    tmpIFCd = CStr(Val(Trim(Mid(sRstData, 1, 3))))
                    tmpRst = Trim(Mid(sRstData, 4, 6))
                    tmpFlag = Trim(Mid(sRstData, 10, 1))
                    
                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                        .ALARMCD = .ALARMCD & Trim(tmpFlag) & Chr(124)
                    End With
                    
                    StartPtr = StartPtr + 10
                Next i
                 
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, "", .ALARMCD)
                    End If
                End With
                
                Call Init_pResultInfo
            End If
            
            '<MOR> 전송
            Call Send_MOR
            
        '<FR1> ~ <FR9>
        Case "1" To "9"
            '변수 초기화
            Call Init_pResultInfo
            
            sFunc = Mid$(RcvBuffer, 2, 1)      ' Function
            
            If sFunc = "K" Or sFunc = "L" Then
               '<MOR> 전송
                Call Send_MOR
                Exit Sub
            End If
    
'            If sFunc <> "@" And sFunc <> "N" And sFunc <> "M" Then
            If sFunc <> "@" And sFunc <> "M" Then
                tmpSeqNo = Trim(Mid(RcvBuffer, 4, 5))
                tmpRack = Trim(Mid(RcvBuffer, 9, 1))
                tmpPos = Trim(Mid(RcvBuffer, 10, 3))
                tmpID = Trim(Mid(RcvBuffer, 14, 13))
                
                If tmpID = "" Then
                    tmpID = tmpSeqNo
                End If
                
                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = tmpID
                    .SEQNO = tmpSeqNo
                    .RACK = tmpRack
                    .POS = tmpPos
                End With
                
                iTestCnt = Val(Trim(Mid(RcvBuffer, 48, 3)))     'Test Count
                StartPtr = 51

                For i = 1 To iTestCnt
                    sRstData = Mid(RcvBuffer, StartPtr, 10)
                    
                    tmpIFCd = CStr(Val(Trim(Mid(sRstData, 1, 3))))  'TEST NUMBER
                    tmpRst = Trim(Mid(sRstData, 4, 6))              'RESULT
                    tmpFlag = Trim(Mid(sRstData, 10, 1))            'FLAG
                    
                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                        .ALARMCD = .ALARMCD & Trim(tmpFlag) & Chr(124)
                    End With
                    
                    StartPtr = StartPtr + 10
                Next i
                 
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, "", .ALARMCD)
                    End If
                End With
                
                Call Init_pResultInfo
            End If
            
            '<MOR> 전송
            Call Send_MOR
            
        Case Else
            '<MOR> 전송
            Call Send_MOR
    End Select
    
    Exit Sub
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub
Private Sub DataEditResponse_Hitachi7180()
    On Error GoTo ErrRtn
    
    Dim sFrame  As String
    Dim sFunc   As String
    Dim i       As Integer
    Dim StartPtr    As Integer
    Dim tmpSeqNo$, tmpID$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarm$
    Dim iTestCnt    As Integer
    Dim sRstData    As String
    
    sFrame = Left(RcvBuffer, 1)
   
    Select Case sFrame
        '<REP>
        Case "?"
            '<MOR> 전송
            Call Send_MOR
        
        '<ANY>
        Case ">"
            '<MOR> 전송
            Call Send_MOR
        
        '<SPE>
        Case ";"
            '<SPE> 전송
            sFunc = Mid(RcvBuffer, 2, 40)
            tmpSeqNo = Mid(RcvBuffer, 4, 5)
            tmpRack = Mid(RcvBuffer, 9, 1)
            tmpPos = Mid(RcvBuffer, 10, 3)
            tmpID = Trim(Mid(RcvBuffer, 14, 13))
            
            'ID 저장
            With pSampleInfo
                .ID = Trim(tmpID)
                .SEQNO = Trim(tmpSeqNo)
                .RACK = Trim(tmpRack)
                .POS = Trim(tmpPos)
            End With
            
            If Len(tmpID) > 0 Then
                '--- ORDER 조회/전송
                Call SendOrder_Hitachi7180(sFunc, pSampleInfo.SEQNO)
            Else
                Call Send_MOR
            End If
        
        '<END>
        Case ":"
            '변수 초기화
            Call Init_pResultInfo
            
            sFunc = Mid$(RcvBuffer, 2, 1)      ' Function
            
            If sFunc = "K" Or sFunc = "L" Then
               '<MOR> 전송
                Call Send_MOR
                Exit Sub
            End If
    
'            If sFunc <> "@" And sFunc <> "N" And sFunc <> "M" Then
            If sFunc <> "@" And sFunc <> "M" Then
                tmpSeqNo = Trim(Mid(RcvBuffer, 4, 5))
                tmpRack = Trim(Mid(RcvBuffer, 9, 1))
                tmpPos = Trim(Mid(RcvBuffer, 10, 3))
                tmpID = Trim(Mid(RcvBuffer, 14, 13))
                
                If tmpID = "" Then
                    tmpID = tmpSeqNo
                End If
                
                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = tmpID
                    .SEQNO = tmpSeqNo
                    .RACK = tmpRack
                    .POS = tmpPos
                End With
                
                iTestCnt = Val(Trim(Mid(RcvBuffer, 48, 3)))
                
                StartPtr = 51

                For i = 1 To iTestCnt
                    sRstData = Mid(RcvBuffer, StartPtr, 10)
                    
                    tmpIFCd = CStr(Val(Trim(Mid(sRstData, 1, 3))))
                    tmpRst = Trim(Mid(sRstData, 4, 6))
                    tmpFlag = Trim(Mid(sRstData, 10, 1))
                    
                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                        .ALARMCD = .ALARMCD & Trim(tmpFlag) & Chr(124)
                    End With
                    
                    StartPtr = StartPtr + 10
                Next i
                 
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, "", .ALARMCD)
                    End If
                End With
                
                Call Init_pResultInfo
            End If
            
            '<MOR> 전송
            Call Send_MOR
            
        '<FR1> ~ <FR9>
        Case "1" To "9"
            '변수 초기화
            Call Init_pResultInfo
            
            sFunc = Mid$(RcvBuffer, 2, 1)      ' Function
            
            If sFunc = "K" Or sFunc = "L" Then
               '<MOR> 전송
                Call Send_MOR
                Exit Sub
            End If
    
'            If sFunc <> "@" And sFunc <> "N" And sFunc <> "M" Then
            If sFunc <> "@" And sFunc <> "M" Then
                tmpSeqNo = Trim(Mid(RcvBuffer, 4, 5))
                tmpRack = Trim(Mid(RcvBuffer, 9, 1))
                tmpPos = Trim(Mid(RcvBuffer, 10, 3))
                tmpID = Trim(Mid(RcvBuffer, 14, 13))
                
                If tmpID = "" Then
                    tmpID = tmpSeqNo
                End If
                
                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = tmpID
                    .SEQNO = tmpSeqNo
                    .RACK = tmpRack
                    .POS = tmpPos
                End With
                
                iTestCnt = Val(Trim(Mid(RcvBuffer, 48, 3)))     'Test Count
                StartPtr = 51

                For i = 1 To iTestCnt
                    sRstData = Mid(RcvBuffer, StartPtr, 10)
                    
                    tmpIFCd = CStr(Val(Trim(Mid(sRstData, 1, 3))))  'TEST NUMBER
                    tmpRst = Trim(Mid(sRstData, 4, 6))              'RESULT
                    tmpFlag = Trim(Mid(sRstData, 10, 1))            'FLAG
                    
                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                        .ALARMCD = .ALARMCD & Trim(tmpFlag) & Chr(124)
                    End With
                    
                    StartPtr = StartPtr + 10
                Next i
                 
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, "", .ALARMCD)
                    End If
                End With
                
                Call Init_pResultInfo
            End If
            
            '<MOR> 전송
            Call Send_MOR
            
        Case Else
            '<MOR> 전송
            Call Send_MOR
    End Select
    
    Exit Sub
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub


Private Sub PhaseCfg_Protocol_HITACHI7170()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
           Case 2
              RcvBuffer = ""
           
           Case 3

           Case 13
              Call DataEditResponse_Hitachi7170
              
           Case 21

           Case Else
              RcvBuffer = RcvBuffer + wkDat
        End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_HITACHI7150()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
       
        Select Case Asc(wkDat)
            Case 2          ' STX
                If bEndChk = True Then
                    RcvBuffer = ""
                Else
                    bSTXChk = True
                End If
                bEndChk = True
                
            Case 3          ' ETX 수신
                If bEndChk = True Then
                    Call DataEditResponse_Hitachi7150
                    RcvBuffer = ""
                End If
                msComm.Output = Chr(6)
               
            Case 23         ' ETB
                bEndChk = False
                msComm.Output = Chr(6)
               
            Case 4          ' EOT
'                msComm.Output = Chr(6)
            
            Case 21         'NAK
'                Call Order_Input
                
            Case Else
                If bEndChk = True Then
                    If bSTXChk = True Then
                        bSTXChk = False
                    Else
                        RcvBuffer = RcvBuffer & wkDat
                    End If
                End If
        
        End Select
    Next ix1

End Sub

Private Sub DataEditResponse_Hitachi7150()
    On Error GoTo ErrRtn

    Dim sFuncNo     As String   'Function Number
    Dim tmpGbn$, tmpSeqNo$, tmpID$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$
    Dim sRstData    As String
    Dim i%, iStartPos%, iBufLen%
        
    '' Function Number 얻기
    sFuncNo = Left(RcvBuffer, 2)
    
    Select Case sFuncNo
        Case "01"
            '' Function Number 01 - Test selection inquiry for routine sample(real time)
            
        Case "11"
            '' Function Number 11 - Test selection inquiry for rerun sample(real time)
            
        Case "02", "03", "52"
            '' 02 - Result transmission of routine, rerun, Stat and control sample (1) (real time)
            '' 03 - Result transmission of routine, rerun, Stat and control sample (2) (real time)
            '' 52 - Results transmission or routine, rerun, Stat and control sample (batch communication)
            
            tmpGbn = Trim(Mid$(RcvBuffer, 3, 1))      '응급, 일반 구분.
            tmpSeqNo = Trim(Mid$(RcvBuffer, 4, 4))
            tmpRack = Trim(Mid$(RcvBuffer, 9, 1))
            tmpPos = Trim(Mid$(RcvBuffer, 10, 2))
            tmpID = Trim(Mid$(RcvBuffer, 13, 11))
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = tmpID
                .SEQNO = tmpGbn & tmpSeqNo
                .RACK = tmpRack
                .POS = tmpPos
            End With
            
            iBufLen = Len(RcvBuffer)
            
            iStartPos = 25
            
            Do Until iStartPos >= iBufLen
                sRstData = Mid(RcvBuffer, iStartPos, 9)
            
                tmpIFCd = Trim(Mid$(sRstData, 1, 2))
                tmpRst = Trim(Mid$(sRstData, 3, 6))
                tmpFlag = Trim(Mid(sRstData, 9, 1))
               
                '결과값 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                    .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .FLAG = .FLAG & Trim(tmpFlag) & Chr(124)
                    .ALARMCD = .ALARMCD & Chr(124)
                End With
                
                iStartPos = iStartPos + 9
            Loop
        
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD)
                End If
            End With
            
            Call Init_pResultInfo
            
        Case "04"
            '' Calibration results transmission (1) (real time)
        Case "05"
            '' Calibration results transmission (2) (real time)
        Case "06"
            '' Original absorbance data transmission (real time)
        Case "51"
            '' Test selection inquiry for routine sample (batch communication)
'            Call Order_Input
        
        Case "61"
            '' Test selection inquiry for rerun sample (batch communication)
        Case "53"
            '' Edited data transmission (batch communication)
        Case "55"
            '' Test selection inquiry for routine sample when sample ID accessory is provided and
            '' ALL us designated (batch communication)
        
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_HITACHI7180()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2
                RcvBuffer = ""
           
            Case 3

'            Case 13
            Case 10, 13
                Call DataEditResponse_Hitachi7180
              
            Case 21

            Case Else
                RcvBuffer = RcvBuffer + wkDat
        End Select
    Next ix1
    
End Sub
Private Sub DataEditResponse_Hitachi7170()
    On Error GoTo ErrRtn

    Dim sFrame  As String
    Dim sFunc   As String
    Dim i       As Integer
    Dim StartPtr    As Integer
    Dim tmpSeqNo$, tmpID$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarm$
    Dim iTestCnt    As Integer
    Dim sRstData    As String

    sFrame = Left(RcvBuffer, 1)
    
    Select Case sFrame
        Case ">"                        ' ANY
            '<MOR> 전송
            Call Send_MOR2

        Case "?"                        ' REP
            '<MOR> 전송
            Call Send_MOR2

        Case ";"                        ' SPE
            sFunc = Mid(RcvBuffer, 2, 39)       ' Function
            tmpSeqNo = Mid(RcvBuffer, 4, 5)     ' Sample No.
            tmpRack = Mid(RcvBuffer, 9, 1)      ' Disk No
            tmpPos = Mid(RcvBuffer, 10, 3)      ' Position No.
            tmpID = Trim(Mid(RcvBuffer, 13, 13))    ' Id No.
         
            'ID 저장
            With pSampleInfo
                .ID = Trim(tmpID)
                .SEQNO = Trim(tmpSeqNo)
                .RACK = Trim(tmpRack)
                .POS = Trim(tmpPos)
            End With
            
            If Len(tmpID) > 0 Then
                '--- ORDER 조회/전송
                Call SendOrder_Hitachi7170(sFunc, pSampleInfo.SEQNO)
            Else
                Call Send_MOR2
            End If

        Case "1" To "9"               ' FR1 to FR9
            '변수 초기화
            Call Init_pResultInfo
            
            sFunc = Mid$(RcvBuffer, 2, 1)      ' Function
            
            If sFunc = "K" Or sFunc = "L" Or sFunc = "G" Or sFunc = "H" Then
               '<MOR> 전송
                Call Send_MOR2
                Exit Sub
            End If
            
            If sFunc <> "@" And sFunc <> "N" And sFunc <> "M" Then
                tmpSeqNo = Mid(RcvBuffer, 4, 5)
                tmpRack = Mid(RcvBuffer, 9, 1)
                tmpPos = Mid(RcvBuffer, 10, 3)
                tmpID = Trim(Mid(RcvBuffer, 13, 13))
            
                If tmpID = "" Then
                    tmpID = tmpSeqNo
                End If
                
                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = tmpID
                    .SEQNO = tmpSeqNo
                    .RACK = tmpRack
                    .POS = tmpPos
                End With
                
                iTestCnt = Val(Trim(Mid$(RcvBuffer, 41, 3)))    'Test Count
                StartPtr = 44
            
                For i = 1 To iTestCnt
                    sRstData = Mid(RcvBuffer, StartPtr, 10)
                    
                    tmpIFCd = CStr(Val(Trim(Mid(sRstData, 1, 3))))
                    tmpRst = Trim(Mid(sRstData, 4, 6))
                    tmpFlag = Trim(Mid(sRstData, 10, 1))
                    
                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .ALARMCD = .ALARMCD & Trim(tmpFlag) & Chr(124)
                    End With
                    
                    StartPtr = StartPtr + 10     ' Start Pointer Increament
                Next i
                
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, "", .ALARMCD)
                    End If
                End With
                
                Call Init_pResultInfo
            End If
            
            '<MOR> 전송
            Call Send_MOR2

        Case ":"                        ' END
            '변수 초기화
            Call Init_pResultInfo
            
            sFunc = Mid$(RcvBuffer, 2, 1)      ' Function
            
            If sFunc = "K" Or sFunc = "L" Or sFunc = "G" Or sFunc = "H" Then
               '<MOR> 전송
                Call Send_MOR2
                Exit Sub
            End If
            
            If sFunc <> "@" And sFunc <> "N" And sFunc <> "M" Then
                tmpSeqNo = Trim(Mid(RcvBuffer, 4, 5))
                tmpRack = Trim(Mid(RcvBuffer, 9, 1))
                tmpPos = Trim(Mid(RcvBuffer, 10, 3))
                tmpID = Trim(Mid(RcvBuffer, 13, 13))
                
                If tmpID = "" Then
                    tmpID = tmpSeqNo
                End If
                
                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = tmpID
                    .SEQNO = tmpSeqNo
                    .RACK = tmpRack
                    .POS = tmpPos
                End With
                
                iTestCnt = Val(Trim$(Mid$(RcvBuffer, 41, 3)))
                
                StartPtr = 44
                
                For i = 1 To iTestCnt
                    sRstData = Mid(RcvBuffer, StartPtr, 10)
                    
                    tmpIFCd = CStr(Val(Trim(Mid(sRstData, 1, 3))))
                    tmpRst = Trim(Mid(sRstData, 4, 6))
                    tmpFlag = Trim(Mid(sRstData, 10, 1))
                    
                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .ALARMCD = .ALARMCD & Trim(tmpFlag) & Chr(124)
                    End With
                    
                    StartPtr = StartPtr + 10    ' Start Pointer Increament
                Next i
            
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, "", .ALARMCD)
                    End If
                End With
                
                Call Init_pResultInfo
            End If
            
            '<MOR> 전송
            Call Send_MOR2
            
        Case Else
            '<MOR> 전송
            Call Send_MOR2
        
    End Select
       
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_Hitachi7180_Batch(ByVal sFunc As String, ByVal sSeqNo As String)
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim iCnt        As Integer
    Dim sTestDat    As String
    Dim ChkSum      As Integer
    Dim ChkSumS     As String
    Dim i   As Integer
    
    '검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    '검사항목 편집
    Call Get_OrderString
    
    pSampleInfo.SEQNO = sSeqNo
    
    sFunc = Replace(sFunc, String(13, "#"), Left(pSampleInfo.ID & Space(13), 13))
            
    sTestDat = String$(88, "0")

    '검사항목 Order코드 추가
    For iCnt = 1 To pSampleInfo.ORDCNT
        If Trim$(pSampleInfo.IFCD(iCnt)) <> "" Then
            Mid$(sTestDat, Val(pSampleInfo.IFCD(iCnt)), 1) = "1"
        End If
    Next iCnt
    
    'ORDER 전송
    sSendBuff = ";" & sFunc
    sSendBuff = sSendBuff & " 88"
    sSendBuff = sSendBuff & Mid(sTestDat, 1, 88)
'    sSendBuff = sSendBuff & "100000" & Space(30)
    'COMMENT란에 BARCODE 표시
    sSendBuff = sSendBuff & "100000" & Left(pSampleInfo.ID & Space(30), 30)
    
    Call Sleep(100)
    
    ' SPE Send
    msComm.Output = Chr$(2) & sSendBuff & Chr$(3) & HITACHI_CheckSum(sSendBuff) & Chr$(13)
    Do
    '   DoEvents
    Loop Until msComm.OutBufferCount = 0
    
    If m_sTestMode = 77 Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If
    
    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 Then
        RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_Hitachi7170(ByVal sFunc As String, ByVal sSeqNo As String)
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim iCnt        As Integer
    Dim sTestDat    As String
    Dim ChkSum      As Integer
    Dim ChkSumS     As String
    Dim i   As Integer
    
    '검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    '검사항목 편집
    Call Get_OrderString
    
    pSampleInfo.SEQNO = sSeqNo
    
            
    sTestDat = String$(88, "0")

    '검사항목 Order코드 추가
    For iCnt = 1 To pSampleInfo.ORDCNT
        If Trim$(pSampleInfo.IFCD(iCnt)) <> "" Then
            Mid$(sTestDat, Val(pSampleInfo.IFCD(iCnt)), 1) = "1"
        End If
    Next iCnt
    
    'ORDER 전송
    sSendBuff = ";" & sFunc
    sSendBuff = sSendBuff & " 88"
    sSendBuff = sSendBuff & Mid(sTestDat, 1, 88)
    sSendBuff = sSendBuff & "00000"
    
    Call Sleep(100)
    
    ' SPE Send
    msComm.Output = Chr$(2) & sSendBuff & Chr$(3) & HITACHI_CheckSum(sSendBuff) & Chr$(13)
    Do
    '   DoEvents
    Loop Until msComm.OutBufferCount = 0
    
    If m_sTestMode = 77 Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If
    
    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 Then
        RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub
Private Sub SendOrder_Hitachi7180(ByVal sFunc As String, ByVal sSeqNo As String)
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim iCnt        As Integer
    Dim sTestDat    As String
    Dim ChkSum      As Integer
    Dim ChkSumS     As String
    Dim i   As Integer
    
    '검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    '검사항목 편집
    Call Get_OrderString
    
    pSampleInfo.SEQNO = sSeqNo
    
            
    sTestDat = String$(88, "0")

    '검사항목 Order코드 추가
    For iCnt = 1 To pSampleInfo.ORDCNT
        If Trim$(pSampleInfo.IFCD(iCnt)) <> "" Then
            Mid$(sTestDat, Val(pSampleInfo.IFCD(iCnt)), 1) = "1"
        End If
    Next iCnt
    
    'ORDER 전송
    sSendBuff = ";" & sFunc
    sSendBuff = sSendBuff & " 88"
    sSendBuff = sSendBuff & Mid(sTestDat, 1, 88)
    sSendBuff = sSendBuff & "100000" & Space(30)
    
    Call Sleep(100)
    
    ' SPE Send
    msComm.Output = Chr$(2) & sSendBuff & Chr$(3) & HITACHI_CheckSum(sSendBuff) & Chr$(13)
    Do
    '   DoEvents
    Loop Until msComm.OutBufferCount = 0
    
    If m_sTestMode = 77 Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If
    
    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 Then
        RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub
Private Function HITACHI_CheckSum(ByVal sBuf$) As String
    Dim iCnt%
    Dim iSum%
    
    For iCnt = 1 To Len(sBuf)
        iSum = iSum + Val(Asc(Mid(sBuf, iCnt, 1)))
    Next
    
    HITACHI_CheckSum = Right("0" & CStr(Hex(iSum)), 2)
End Function
'
'   기존 Send_MOR에 DO/LOOP 문 추가
'
Private Sub Send_MOR2()
    Call Sleep(100)
    msComm.Output = Chr(2) & ">" & Chr(3) & "3E" & vbCr
    
    Do
    '   DoEvents
    Loop Until msComm.OutBufferCount = 0
End Sub
Private Sub Send_MOR()
    Call Sleep(100)
    msComm.Output = Chr(2) & ">" & Chr(3) & "3E" & vbCr
End Sub
'
'   HITACHI7020 NoBarcode
'
Private Sub PhaseCfg_Protocol_HITACHI7020()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
             
        Select Case Asc(wkDat)
            Case 2          ' STX
                RcvBuffer = ""
              
            Case 10, 13     '
                Call DataEditResponse_Hitachi7020       'Data 편집
                RcvBuffer = ""
              
            Case 21         ' Nak
           
            Case Else       ' Data
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
End Sub
Private Sub DataEditResponse_Hitachi7020()
    On Error GoTo ErrRtn

    Dim sFrame  As String
    Dim sFunc   As String
    Dim i       As Integer
    Dim sSampInfo   As String
    Dim StartPtr    As Integer
    Dim tmpSeqNo$, tmpID$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$
    Dim iTestCnt    As Integer
    Dim sRstData    As String


    sFrame = Left$(RcvBuffer, 1)        ' get Frame of RcvBuffer.

    Select Case sFrame
        Case ">"                        ' ANY
            For i = 1 To 10000          ' loop
            Next i
            ' MOR Send
            msComm.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Do
            '   DoEvents
            Loop Until msComm.OutBufferCount = 0

        Case "?"                        ' REP
            For i = 1 To 10000          ' loop
            Next i
            ' MOR Send
            msComm.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Do
            '   DoEvents
            Loop Until msComm.OutBufferCount = 0

        Case ";"                        ' SPE
            sFunc = Mid$(RcvBuffer, 2, 1)       ' Function
            tmpSeqNo = Mid$(RcvBuffer, 4, 5)    ' Sample No.
            tmpRack = Mid$(RcvBuffer, 9, 1)     ' Rack No.
            tmpPos = Mid$(RcvBuffer, 10, 3)     ' Position No.
            tmpID = Mid$(RcvBuffer, 13, 13)     ' Id No.
            'ID 저장
            With pSampleInfo
                .ID = Trim(tmpID)
                .SEQNO = tmpSeqNo
                .RACK = tmpRack
                .POS = tmpPos
            End With
            sSampInfo = Mid$(RcvBuffer, 4, 37)  ' Sample Information

            '--- Order 전송
            Call SendOrder_Hitachi7020(sFunc)


        Case "1" To "9"               ' FR1 to FR9(검사항목 25개 이상일 경우 처리)
            '변수 초기화
            Call Init_pResultInfo

            For i = 1 To 10000        ' loop
            Next i

            sFunc = Mid$(RcvBuffer, 2, 1)     ' Function

            If sFunc = "K" Or sFunc = "L" Or sFunc = "G" Or sFunc = "H" Then
                GoTo Nxt2
            End If

            If sFunc <> "@" Then              ' TNNA ?
                tmpID = Trim$(Mid$(RcvBuffer, 13, 13))          ' Id No.
                tmpPos = Trim$(Mid$(RcvBuffer, 10, 3))          ' Position No

                '결과정보 구조체에 저장
                pResultInfo.ID = tmpID
                pResultInfo.POS = tmpPos

                iTestCnt = Val(Trim$(Mid$(RcvBuffer, 41, 3)))   ' Test Count
                StartPtr = 44
                For i = 1 To iTestCnt
                    sRstData = Mid$(RcvBuffer, StartPtr, 10)

                    tmpIFCd = Trim(Mid$(sRstData, 1, 3))    ' Test Number
                    tmpRst = Trim(Mid$(sRstData, 4, 6))

                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                    End With
                    StartPtr = StartPtr + 10   ' Start Pointer Increament
                Next i

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                    End If
                End With
            End If

            ' MOR Send
            msComm.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Do
            '   DoEvents
            Loop Until msComm.OutBufferCount = 0

        Case ":"                        ' END
            '변수 초기화
            Call Init_pResultInfo

            sFunc = Mid$(RcvBuffer, 2, 1)     ' Function
            'Calibration 값이 넘어올 경우 예외처리...
            If sFunc = "K" Or sFunc = "L" Or sFunc = "G" Or sFunc = "H" Then
               GoTo Nxt2
            End If

            If sFunc <> "@" Then               ' TNNA
                tmpID = Trim$(Mid$(RcvBuffer, 13, 13))          ' Id No.
                tmpPos = Trim$(Mid$(RcvBuffer, 10, 3))          ' Position No

                '결과정보 구조체에 저장
                pResultInfo.ID = tmpID
                pResultInfo.POS = tmpPos

                iTestCnt = Val(Trim$(Mid$(RcvBuffer, 41, 3)))   ' Test Count
                StartPtr = 44
                For i = 1 To iTestCnt
                    sRstData = Mid$(RcvBuffer, StartPtr, 10)

                    tmpIFCd = Trim(Mid$(sRstData, 1, 3))        ' Test Number
                    tmpRst = Trim(Mid$(sRstData, 4, 6))

                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                    End With
                    StartPtr = StartPtr + 10   ' Start Pointer Increament
                Next i

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                    End If
                End With

                Call Init_pResultInfo
           End If

Nxt2:
           ' MOR Send
           msComm.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
           Do
           '   DoEvents
           Loop Until msComm.OutBufferCount = 0

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub
'
'   HITACHI747 Barcode
'
Private Sub PhaseCfg_Protocol_HITACHI747()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2          ' STX
                RcvBuffer = ""

            Case 3          ' ETX

            Case 10, 13
                Call DataEditResponse_Hitachi747

            Case 21         ' NAK

            Case Else       ' Data
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1

End Sub
' *=====================================================*
' *         Data편집 & 응답처리                         *
' *=====================================================*
Private Sub DataEditResponse_Hitachi747()
    On Error GoTo ErrRtn

    Dim sFrame  As String
    Dim sFunc   As String
    Dim i       As Integer
    Dim StartPtr    As Integer
    Dim tmpSeqNo$, tmpID$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$
    Dim iTestCnt    As Integer
    Dim sRstData    As String


    sFrame = Left$(RcvBuffer, 1)        ' get Frame of RcvBuffer.

    Select Case sFrame
        Case ">"                        ' ANY
            For i = 1 To 10000          ' loop
            Next i
            ' MOR Send
            msComm.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Do
            '   DoEvents
            Loop Until msComm.OutBufferCount = 0

        Case "?"                        ' REP
            For i = 1 To 10000          ' loop
            Next i
            ' MOR Send
            msComm.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Do
            '   DoEvents
            Loop Until msComm.OutBufferCount = 0

        Case ";"                        ' SPE
            sFunc = Mid$(RcvBuffer, 2, 1)       ' Function
            tmpRack = Mid$(RcvBuffer, 4, 4)     ' Rack No.
            tmpSeqNo = Mid$(RcvBuffer, 8, 5)    ' Sample No.
            tmpPos = Mid$(RcvBuffer, 13, 1)     ' Position No.
            tmpID = Mid$(RcvBuffer, 14, 13)     ' Id No.

            'ID 저장
            With pSampleInfo
                .ID = Trim(tmpID)
                .SEQNO = Trim(tmpSeqNo)
                .RACK = Trim(tmpRack)
                .POS = Trim(tmpPos)
            End With

            '--- ORDER 조회/전송
            Call SendOrder_Hitachi747(sFunc, pSampleInfo.SEQNO)


        Case "1" To "9"               ' FR1 to FR9
            '변수 초기화
            Call Init_pResultInfo

            For i = 1 To 10000        ' loop
            Next i

            sFunc = Mid$(RcvBuffer, 2, 1)     ' Function
            If sFunc = "K" Or sFunc = "L" Then
               GoTo Nxt2
            End If

            If sFunc <> "@" Then              ' TNNA
                tmpRack = Trim$(Mid$(RcvBuffer, 4, 4))      ' Rack No.
                tmpSeqNo = Trim$(Mid$(RcvBuffer, 8, 5))     ' Sample No.
                tmpPos = Trim$(Mid$(RcvBuffer, 13, 1))      ' Position No.
                tmpID = Trim$(Mid$(RcvBuffer, 14, 13))      ' BARCODE ID

                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = tmpID
                    .SEQNO = tmpSeqNo
                    .RACK = tmpRack
                    .POS = tmpPos
                End With

                iTestCnt = Val(Trim$(Mid$(RcvBuffer, 27, 2)))   ' Test Count
                StartPtr = 29
                For i = 1 To iTestCnt
                    sRstData = Mid$(RcvBuffer, StartPtr, 9)

                    tmpIFCd = Mid$(sRstData, 1, 2)  ' Test Number
                    tmpRst = Mid$(sRstData, 3, 6)   ' Result
                    tmpFlag = Mid(sRstData, 9, 1)   ' FLAG

                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                        .FLAG = .FLAG & Trim(tmpFlag) & Chr(124)
                    End With
                    StartPtr = StartPtr + 9         ' Start Pointer Increament
                Next i

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                    End If
                End With
            End If

            ' MOR Send
            msComm.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
            Do
            '   DoEvents
            Loop Until msComm.OutBufferCount = 0

        Case ":"                    ' END
            '변수 초기화
            Call Init_pResultInfo

            sFunc = Mid$(RcvBuffer, 2, 1)       ' Function
            If sFunc = "K" Or sFunc = "L" Then
               GoTo Nxt2
            End If

           If sFunc <> "@" Then              ' TNNA
                tmpRack = Trim$(Mid$(RcvBuffer, 4, 4))      ' Rack No.
                tmpSeqNo = Trim$(Mid$(RcvBuffer, 8, 5))     ' Sample No.
                tmpPos = Trim$(Mid$(RcvBuffer, 13, 1))      ' Position No.
                tmpID = Trim$(Mid$(RcvBuffer, 14, 13))

                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = tmpID
                    .SEQNO = tmpSeqNo
                    .RACK = tmpRack
                    .POS = tmpPos
                End With

                iTestCnt = Val(Trim$(Mid$(RcvBuffer, 27, 2)))   ' Test Count
                StartPtr = 29
                For i = 1 To iTestCnt
                    sRstData = Mid$(RcvBuffer, StartPtr, 9)

                    tmpIFCd = Mid$(sRstData, 1, 2)  ' Test Number
                    tmpRst = Mid$(sRstData, 3, 6)   ' Result
                    tmpFlag = Mid(sRstData, 9, 1)   ' FLAG

                    '결과값 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                        .RST1 = .RST1 & Trim(tmpRst) & Chr(124)
                        .FLAG = .FLAG & Trim(tmpFlag) & Chr(124)
                    End With
                    StartPtr = StartPtr + 9         ' Start Pointer Increament
                Next i

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                    End If
                End With
           End If

Nxt2:
           ' MOR Send
           msComm.Output = Chr$(2) & ">" & Chr$(3) & "3E" & Chr$(13)
           Do
           '   DoEvents
           Loop Until msComm.OutBufferCount = 0
    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub
'
'   HITACHI7600 바코드 사용 안하는 경우...(Rack/Pos 로 오더 전송)
'
Private Sub PhaseCfg_Protocol_HITACHI7600_Batch()
        
    Dim wkDat   As String
    Dim ix1 As Integer
    Dim i   As Integer
            
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False
                        
                        msComm.Output = Chr(6)
                        
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     '<LF>
                        If bEndChk = True Then
                            Call DataEditResponse_Hitachi7600_Batch
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                    
                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_Hitachi7600_Batch
                            RcvBuffer = ""
                        End If
                        
                    Case 4      'EOT
                        RcvBuffer = ""
                        m_iPhase = 1
                        
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)   'Send ACK
                        
                    Case 21     'NAK
                        Call DataEditResponse_Hitachi7600_Batch
                        
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        
                    Case 23     ' ETB
                        bEndChk = False
                        
                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkDat
                            End If
                        End If
                        
                End Select
            
            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        Call SendOrder_Hitachi7600_Batch
                                                
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                        
                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        msComm.Output = Chr(5)
                        m_iPhase = 3
                    
                    Case 4      'EOT
                        m_iPhase = 1
                
                End Select
        End Select
    Next ix1
    
End Sub
'
'   HITACHI7600 바코드 사용
'
Private Sub PhaseCfg_Protocol_HITACHI7600()
        
    Dim wkDat   As String
    Dim ix1 As Integer
    Dim i   As Integer
            
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False
                        
                        msComm.Output = Chr(6)
                        
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     '<LF>
                        If bEndChk = True Then
                            Call DataEditResponse_Hitachi7600
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                    
                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_Hitachi7600
                            RcvBuffer = ""
                        End If
                        
                    Case 4      'EOT
                        If sState = "Q" Then
                            msComm.Output = Chr(5)
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3
                        
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        msComm.Output = Chr(6)   'Send ACK
                        
                    Case 21     'NAK
                        Call DataEditResponse_Hitachi7600
                        
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        
                        msComm.Output = Chr(5)   'Send ENQ
                        
                    Case 23     ' ETB
                        bEndChk = False
                        
                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkDat
                            End If
                        End If
                        
                End Select
            
            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        If sState = "Q" Then
                            Call SendOrder_Hitachi7600
                        End If
                        
                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                        
                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        msComm.Output = Chr(5)
                        m_iPhase = 3
                    
                    Case 4      'EOT
                        m_iPhase = 1
                
                End Select
        End Select
    Next ix1
    
End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_Hitachi7600_Batch()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpComment1 As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpData()   As String
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
            
        Case "O"
            tmpSeqNo = "": tmpComment1 = "": tmpRack = "": tmpPos = ""
            tmpData() = Split(RcvBuffer, "|")
            tmpComment1 = tmpData(19)
            ii = InStr(1, tmpData(2), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpData(2), "^")
                tmpSeqNo = Trim(tmpData(0))
                tmpRack = Trim(tmpData(3))
                tmpPos = Trim(tmpData(4))
            End If

            ii = InStr(1, tmpComment1, "^")
            If ii <> 0 Then
                tmpData() = Split(tmpComment1, "^")
                tmpComment1 = Trim(tmpData(0))
            End If
            
'            pSampleInfo.ID = UCase(tmpComment1)
'            pSampleInfo.RACK = tmpRack
'            pSampleInfo.POS = tmpPos
            pResultInfo.ID = UCase(tmpComment1)
            pResultInfo.RACK = tmpRack
            pResultInfo.POS = tmpPos
                                    
        Case "R"        'Result Record
            '--- 결과데이타 편집
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '5:Reference Ranges
            '6:Result Abnormal Flags
            '8:Result Status(F:First,C:Rerun)
            tmpData() = Split(RcvBuffer, "|")
            
            tmpIFCd = Trim(tmpData(2))
            tmpIFCd = Mid(tmpIFCd, 4)
            tmpIFCd = Mid(tmpIFCd, 1, InStr(1, tmpIFCd, "/") - 1)
            tmpRst = Trim(tmpData(3))
            tmpUnit = Trim(tmpData(4))
            tmpFlag = Trim(tmpData(6))

            '--- 결과값에 "^" 들어갈 경우 편집
            ii = InStr(1, tmpRst, "^")
            If ii <> 0 Then tmpRst = Mid(tmpRst, ii + 1)

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
            
            '결과정보 구조체에 저장
            With pResultInfo
'                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
'                .RACK = pSampleInfo.RACK
'                .POS = pSampleInfo.POS

                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
            End With

        Case "C"        'Comment Record
        
        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub
' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_Hitachi7600()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpData()   As String
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarmCd$
   
    
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
            tmpData() = Split(RcvBuffer, "|")
            sReqStatusCd = Trim(tmpData(12))    'Order Request Status Code
            tmpData() = Split(tmpData(2), "/")
            
            tmpBarCd = Trim(tmpData(1))
            tmpSeqNo = tmpData(0)
            tmpRack = tmpData(3)
            tmpPos = tmpData(4)
            tmpData() = Split(tmpSeqNo, "^")
            tmpSeqNo = Trim(tmpData(2))
            
            If tmpBarCd <> "" Then    'BarCode ID가 잘 넘어왔는지 검사
                sState = "Q"
                pSampleInfo.ID = UCase(tmpBarCd)
            Else
                sState = ""
                pSampleInfo.ID = ""
            End If
                
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            
        Case "O"
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpData() = Split(RcvBuffer, "|")
            ii = InStr(1, tmpData(2), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpData(2), "^")
                tmpSeqNo = Trim(tmpData(0))
                tmpBarCd = Trim(tmpData(1))
                tmpRack = Trim(tmpData(3))
                tmpPos = Trim(tmpData(4))
            End If

            pSampleInfo.ID = UCase(tmpBarCd)
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
                                    
        Case "R"        'Result Record
            '--- 결과데이타 편집
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '5:Reference Ranges
            '6:Result Abnormal Flags
            '8:Result Status(F:First,C:Rerun)
            tmpData() = Split(RcvBuffer, "|")
            
            tmpIFCd = Trim(tmpData(2))
            tmpIFCd = Mid(tmpIFCd, 4)
            tmpIFCd = Mid(tmpIFCd, 1, InStr(1, tmpIFCd, "/") - 1)
            tmpRst = Trim(tmpData(3))
            tmpUnit = Trim(tmpData(4))
            tmpFlag = Trim(tmpData(6))

            '--- 결과값에 "^" 들어갈 경우 편집
            ii = InStr(1, tmpRst, "^")
            If ii <> 0 Then tmpRst = Mid(tmpRst, ii + 1)

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS

                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
            End With

        Case "C"        'Comment Record
            'Data Alarm 편집
            tmpData() = Split(RcvBuffer, Chr(124))
            
            tmpAlarmCd = ConvertDataAlarmCode(UCase(m_EqName), Trim(tmpData(3)))
            pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)
            
        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD)
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_Hitachi7020(ByVal sFunc As String)
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim iCnt        As Integer
    Dim sTestDat    As String
    Dim ChkSum      As Integer
    Dim ChkSumS     As String
    Dim i   As Integer
    
        
    '--- 검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.SEQNO, "", pSampleInfo.POS)
    
    '--- 검사항목 편집
    Call Get_OrderString
            
            
    sTestDat = String$(60, "0")
    
    '검사항목 Order코드 추가
    For iCnt = 1 To pSampleInfo.ORDCNT
        If Trim$(pSampleInfo.IFCD(iCnt)) <> "" Then
            Mid$(sTestDat, Val(pSampleInfo.IFCD(iCnt)), 1) = "1"
        End If
    Next iCnt
        
    '==== Order전송 ====
    sSendBuff = ";" & sFunc & " "
    sSendBuff = sSendBuff & Right(Space(5) & pSampleInfo.SEQNO, 5) & Space(1) _
            & Right(Space(3) & pSampleInfo.POS, 3) _
            & Right(Space(13) & pSampleInfo.ID, 13) _
            & Space(15)
    sSendBuff = sSendBuff & " 37"
    sSendBuff = sSendBuff & Mid$(sTestDat, 1, 37)
    sSendBuff = sSendBuff & "00000"

    ChkSum = 0
    For i = 1 To Len(sSendBuff)
        ChkSum = ChkSum + Asc(Mid$(sSendBuff, i, 1))
    Next i
    ChkSumS = Hex$(ChkSum)
    
    ' SPE Send
    msComm.Output = Chr$(2) & sSendBuff & Chr$(3) & Right$(ChkSumS, 2) & Chr$(13)
    
    If m_sTestMode = 77 Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If
    
    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 Then
        RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If
            
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub
Private Sub SendOrder_Hitachi747(ByVal sFunc As String, ByVal sSeqNo As String)
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim iCnt        As Integer
    Dim sTestDat    As String
    Dim ChkSum      As Integer
    Dim ChkSumS     As String
    Dim i   As Integer
    
    '검사항목 조회
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    '검사항목 편집
    Call Get_OrderString
    
    pSampleInfo.SEQNO = sSeqNo
    
            
    sTestDat = String$(60, "0")

    '검사항목 Order코드 추가
    For iCnt = 1 To pSampleInfo.ORDCNT
        If Trim$(pSampleInfo.IFCD(iCnt)) <> "" Then
            Mid$(sTestDat, Val(pSampleInfo.IFCD(iCnt)), 1) = "1"
        End If
    Next iCnt
    
    'ORDER 전송
    sSendBuff = ";" & sFunc & " "
    sSendBuff = sSendBuff & Right(Space(4) & pSampleInfo.RACK, 4) _
            & Right(Space(5) & pSampleInfo.SEQNO, 5) _
            & Right(Space(1) & pSampleInfo.POS, 1) _
            & Right$(Space(13) + Trim$(pSampleInfo.ID), 13)         ''''''2003/1/9 확인필요...
    sSendBuff = sSendBuff & "60"
    sSendBuff = sSendBuff & Mid(sTestDat, 1, 60)
    sSendBuff = sSendBuff & "0"
    
    ChkSum = 0
    For i = 1 To Len(sSendBuff)
        ChkSum = ChkSum + Asc(Mid$(sSendBuff, i, 1))
    Next i
    ChkSumS = Hex$(ChkSum)
    
    ' SPE Send
    msComm.Output = Chr$(2) & sSendBuff & Chr$(3) & Right$(ChkSumS, 2) & Chr$(13)
    Do
    '   DoEvents
    Loop Until msComm.OutBufferCount = 0
    
    If m_sTestMode = 77 Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If
    
    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 Then
        RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

'
'   환자 Order 전송(바코드 사용 안하는 경우)
'   -> 주의: 샘플 ID를 COMMENT1에 넣어 전송함.따라서 필히 장비에서 COMMENT 전송 옵션 체크해야 함.
'
Private Sub SendOrder_Hitachi7600_Batch()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String

    If m_p_sID = "" And m_p_iOrdCnt = 0 Then
        Exit Sub
    End If
    
    Select Case m_iSendPhase
        Case 1
            'Header Record
            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^BATCH|P|1" & vbCr

            'Patient Record
            sSendBuff = sSendBuff & "P|1" & vbCr
            
            '--- 시작번호/검사항목 편집
            Call Get_OrderString
            
            'Order Record
            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Space(13) & "^1^" _
                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|R1|"

            '검사항목 Order코드 추가
            For iCnt = 1 To pSampleInfo.ORDCNT
                '정상 오더
                sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "/\"
            Next iCnt
            If pSampleInfo.ORDCNT > 0 Then
                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
            End If

            sSendBuff = sSendBuff & "|R||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
                    & Left(Trim(pSampleInfo.ID) & Space(30), 30) _
                    & "^^^^||||||O" & vbCr

            'Terminator Record
            sSendBuff = sSendBuff & "L|1|N"


            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 242 Then
                sNextSend = Mid(sSendBuff, 242)
                sSendBuff = Left(sSendBuff, 241)
                sSendBuff = sSendBuff & Chr(23)

                m_iFrameN = m_iFrameN + 1
                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If

        Case 2
            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
            sNextSend = ""

Send_Terminate:
            m_iSendPhase = 3

        Case 3      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1
            
            '전송된 오더가 있는 경우 화면표시
            If pSampleInfo.ORDCNT > 0 Then
                If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
                    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
                End If
            Else
                '조회된 내용이 없는 경우 환자정보 구조체 초기화
                Call Init_pResultInfo
        
                RaiseEvent SendOrderOK("", "", "", "")
            End If
            
            Exit Sub
    End Select
    
    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If
    
Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub
'
'   환자 Order 전송
'
Private Sub SendOrder_Hitachi7600()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sStat   As String
    
    Select Case m_iSendPhase
        Case 1
            'Header Record
            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^REPLY|P|1" & vbCr

            'Patient Record
            sSendBuff = sSendBuff & "P|1" & vbCr
                    
            'Order Record
            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) & "^1^" _
                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|R1|"

            '----- 검사항목 조회
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

            Call Get_OrderString

            '검사항목 Order코드 추가
            For iCnt = 1 To pSampleInfo.ORDCNT
                'Request Information Code에 따라 검사항목을 추가하거나 취소한다.
                If Trim(sReqStatusCd) = "O" Then
                    '정상 오더
                    Select Case Trim$(pSampleInfo.IFCD(iCnt))
                        Case "989", "990", "991"
                            'ISE 항목이 있는 경우 전체 검사(ISE 검사는 2가지 조합만 오더 가능)
                            If Val(InStr(1, sSendBuff, "989" & "/")) = 0 Then
                                sSendBuff = sSendBuff & "^^^989/\"
                            End If
                            If Val(InStr(1, sSendBuff, "990" & "/")) = 0 Then
                                sSendBuff = sSendBuff & "^^^990/\"
                            End If
                            If Val(InStr(1, sSendBuff, "991" & "/")) = 0 Then
                                sSendBuff = sSendBuff & "^^^991/\"
                            End If
                            
                        Case Else
                            '일반항목
                            sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "/\"
                    End Select
                    
                ElseIf Trim(sReqStatusCd) = "A" Then
                    '오더 취소
                End If
            Next iCnt
            If pSampleInfo.ORDCNT > 0 And Trim(sReqStatusCd) <> "A" Then
                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
            End If
            
            'STAT RACK에 대한 처리추가
            If Left(pSampleInfo.RACK, 1) = "4" Then
                sStat = "S"
            Else
                sStat = "R"
            End If

            sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
                    & "^^^^||||||O" & vbCr

            'Terminator Record
            sSendBuff = sSendBuff & "L|1|N"


            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 241 Then       '242 Then
'                sNextSend = Mid(sSendBuff, 242)
'                sSendBuff = Left(sSendBuff, 241)
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iFrameN = m_iFrameN + 1
                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If

        Case 2
            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
            sNextSend = ""

Send_Terminate:
            m_iSendPhase = 3

        Case 3      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1

            sState = "": sReqStatusCd = ""

            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If

    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
        End If
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
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
        .ALARMCD = ""
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
'    m_iStartSampleNo = PropBag.ReadProperty("iStartSampleNo", m_def_iStartSampleNo)
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
'    Call PropBag.WriteProperty("iStartSampleNo", m_iStartSampleNo, m_def_iStartSampleNo)
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
