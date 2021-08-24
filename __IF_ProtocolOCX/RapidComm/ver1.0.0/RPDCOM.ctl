VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl RPDCOM 
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
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   300
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "RPDCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_Port = 0
Const m_def_IPAddress = 0
Const m_def_p_sCmt1 = ""
Const m_def_p_sSpcCd = 0
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
Const m_def_OpenPW = "0"
Const m_def_EditPW = "0"
'속성 변수:
Dim m_Port As Variant
Dim m_IPAddress As Variant
Dim m_p_sCmt1 As String
Dim m_p_sSpcCd As Variant
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
Dim m_OpenPW As String
Dim m_EditPW As String
'이벤트 선언:
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTRstDT$, sTAlarmCd$, sKind$, sSpcDesc$, sOperID$, sTInstID$, sTInstNm$, sOther1$)
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTRstDT$, sTAlarmCd$, sKind$, sSpcDesc$, sTInstID$, sTInstNm$, sOther1$)
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

Dim miSckMax As Integer
Dim msMsgType   As String

Dim msRackType  As String
Dim msCupType   As String
Dim msStatGbn   As String

Dim mlCtlID As Long
Dim msCtlID As String

Private Function ConvertDataAlarmCode(ByVal sEqNm As String, ByVal Scode As String) As String
    
    Dim sTmp    As String
    
    ConvertDataAlarmCode = "": sTmp = ""
    
    Select Case UCase(sEqNm)
        Case "HITACHI7600"
            Select Case Trim(Scode)
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
        Case "RAPIDCOMM"
            Call PhaseCfg_Protocol_RAPIDCOMM
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_RAPIDCOMM()
    On Error GoTo ErrPhase
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)
             
        Select Case Asc(wkDat)
            Case 10     'LF
            Case 11     'VT, 
                RcvBuffer = ""
                                                
            Case 28     'FS
                Call DataEditResponse_RAPIDCOMM
                RcvBuffer = ""
               
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .RSTDT, .ALARMCD, .Kind, .SPCCD, .OPERID, .INSTID, .INSTNM, .OTHER)
        End If
    End With

    Call Init_pResultInfo
    
ErrPhase:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub

Private Sub Send_ACK()
    
    Dim sSndBuf As String
    
    mlCtlID = mlCtlID + 1
    
    'MSH
    sSndBuf = "MSH|^~\&|host||cobas 8000||" & Format(Now(), "YYYYMMDDhhmmss") & "||ACK|" & Trim(mlCtlID) & "||2.5||||ER||UNICODE UTF-8" & Chr(13)
'    sSndBuf = "MSH|^~\&|host||cobas 8000||" & Format(Now(), "YYYYMMDDhhmmss") & "||ACK|" & Trim(mlCtlID) & "||2.5||||SU||UNICODE UTF-8" & Chr(13)       '2012/4/26 yk
    'MSA
    sSndBuf = sSndBuf & "MSA|AA|" & msCtlID & "||" & Chr(13)

    
    sSndBuf = Chr(11) & sSndBuf & Chr(28) & Chr(13)
    
    EqOutput_Socket (sSndBuf)

End Sub

Private Sub EqOutput_Socket(ByVal sMsg As String)
    sckServer(miSckMax).SendData (sMsg)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sMsg)
    End If
End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_RAPIDCOMM()
    On Error GoTo ErrRtn

    Dim sHeader     As String   'Record Type
    Dim sMsgType    As String
    Dim ii          As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpKind     As String
    Dim tmpSampType As String
    Dim tmpContType As String
    
    Dim aFrame()    As String
    Dim aField()    As String
    Dim aData()     As String
    
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpSpcCd$, tmpAlarmCd$, tmpOperID$, tmpInstID$, tmpDil$, tmpRststatus$
    Dim tmpRstDT$, tmpCmt$, tmpSrcCd$
    Dim sRcvBuffer  As String
    
    aFrame = Split(RcvBuffer, Chr(13))
    
    For ii = 0 To UBound(aFrame) - 1
        If Trim(aFrame(ii)) <> "" Then
            sRcvBuffer = Trim(aFrame(ii))
            
            aField = Split(sRcvBuffer, Chr(124))
    
            sHeader = Trim(aField(0))
            
            Select Case sHeader
                Case "MSH"          'Message heading
                    Call Init_pResultInfo
                    
                    'MSH|^~\&|Rapidcomm|Hospital|HIS|Hospital|20101202172736||ORU^R32|2T16:11:49D53I2S2533|P|2.4|||AL|AL|
                    msMsgType = Trim(aField(8))
                    msCtlID = Trim(aField(9))
                    
                    If InStr(msCtlID, ":") > 0 Then
                        msCtlID = Trim(Split(msCtlID, ":")(1))
                    End If
                    
                    Select Case Trim(Split(msMsgType, "^")(1))
                        Case "R01"  'QC
                            pResultInfo.Kind = "Q"
                            pResultInfo.ID = msCtlID
                        Case Else
                            pResultInfo.Kind = ""
                    End Select
              
                Case "PID"          'Patient Identification Segment - PID
                    'PID|||3250513|||||U|
                    tmpBarCd = Trim(aField(3))
                    pResultInfo.ID = tmpBarCd
              
                Case "OBR"          'Observation Request Segment - OBR
                    'OBR|1||PT2T16:11:49D53I2S2533||R||||||O||||BLDA^^^^^^P||||||||||F|
                    tmpSpcCd = Trim(aField(15)) '검체코드
                    pResultInfo.SPCCD = tmpSpcCd
                    
                Case "OBX"          'Observation Result Segment - OBX
                    'OBX|1|ST|pH||7.489||7.350-7.450|H|||F|||20130723141714||205018||^^17324^Rapidlab 1265|20130723141714|
                    'OBX|2|ST|pCO2||39.8|mmHg|35.0^mmHg-45.0^mmHg||||F|
                    tmpIFCd = Trim(aField(3))
                    tmpRst = Trim(aField(5))
                    tmpUnit = Trim(aField(6))
                    tmpFlag = Trim(aField(8))
                    tmpRststatus = Trim(aField(11)) 'F First run result, C Corrected result 자동Dilution 시 두번째 결과도 AMR값을 벗어나면 Flag를 'H' 표시하기 위함
                    
                    If Left$(tmpRst, 1) = "." Then
                        tmpRst = "0" & tmpRst
                    End If
                    
                    If Trim(aField(1)) = "1" Then
                        tmpRstDT = Trim(aField(14))
                        tmpOperID = Trim(aField(16))
                        tmpInstID = Trim(Split((aField(18)), "^")(2))
                        
                        pResultInfo.RSTDT = tmpRstDT
                        pResultInfo.OPERID = tmpOperID
                        pResultInfo.INSTID = tmpInstID
                    End If
                    
                    '결과정보 구조체에 저장
                    With pResultInfo
                        '결과값 누적
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & tmpUnit & Chr(124)
                        .FLAG = .FLAG & tmpFlag & Chr(124)
                        .STATUS = .STATUS & tmpRststatus & Chr(124)
                    End With
               
                Case "NTE"          'Comment Segment - NTE
                    'NTE|1|L|QUES|
                    tmpSrcCd = Trim(aField(2))
                    aData() = Split(aField(3), "^")
                    If tmpSrcCd = "I" Then
                        If Trim(aData(0)) = "0" Then
                        Else
                            ''tmpAlarmCd = Trim(aData(1))
                            tmpAlarmCd = Trim(aData(0))
                        End If
                    ElseIf tmpSrcCd = "L" Then
                        tmpAlarmCd = Trim(aData(0))
                    End If
                    pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)
                    tmpAlarmCd = ""
        
            End Select
        End If
    Next
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit_Cobas8000 - " & Err.Description)
    End If
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
        .Kind = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .INSTID = ""
        .INSTNM = ""
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
        .SPCCD = ""
        .OPERID = ""
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

Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub sckServer_Close(Index As Integer)
    sckServer(Index).Close
    Unload sckServer(Index)
End Sub

Private Sub sckServer_Connect(Index As Integer)
    RaiseEvent DispMsg("Connect Server...")
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Index = 0 Then
        miSckMax = miSckMax + 1
        
        Load sckServer(miSckMax)
        sckServer(miSckMax).LocalPort = 0
        sckServer(miSckMax).Accept requestID
        
        RaiseEvent DispMsg("sckServer(" & miSckMax & ").Accept " & requestID)
    End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo ErrSck
        
    sckServer(Index).GetData wkBuf

    If m_sTestMode = "77" Then
        RaiseEvent PrintRcvLog(wkBuf)
    End If
                        
    If iSpaceCnt = 30 Then
        iSpaceCnt = 0
    End If
    iSpaceCnt = iSpaceCnt + 2
    
    RaiseEvent DispMsg(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
    
    Call PhaseCfg_Protocol
    
ErrSck:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    With sckServer(Index)
        If .State <> sckClosed Then
            RaiseEvent DispMsg(Description & "(" & Number & ")")
        End If
    End With
End Sub

'저장소에서 속성값을 로드합니다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

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
    m_p_sSpcCd = PropBag.ReadProperty("p_sSpcCd", m_def_p_sSpcCd)
    m_p_sCmt1 = PropBag.ReadProperty("p_sCmt1", m_def_p_sCmt1)
    m_Port = PropBag.ReadProperty("Port", m_def_Port)
    m_IPAddress = PropBag.ReadProperty("IPAddress", m_def_IPAddress)
End Sub

'속성값을 저장소에 기록합니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

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
    Call PropBag.WriteProperty("p_sSpcCd", m_p_sSpcCd, m_def_p_sSpcCd)
    Call PropBag.WriteProperty("p_sCmt1", m_p_sCmt1, m_def_p_sCmt1)
    Call PropBag.WriteProperty("Port", m_Port, m_def_Port)
    Call PropBag.WriteProperty("IPAddress", m_IPAddress, m_def_IPAddress)
End Sub

'Public Property Let PortOpen(ByVal New_PortOpen As Boolean)
'    m_PortOpen = New_PortOpen
'    PropertyChanged "PortOpen"
'
'    '--- PortOpen시 암호 확인
'    If m_OpenPW <> pOpenPW Then
'        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
'        Exit Property
'    End If
'    '-----------------------
'
'    '변수 초기화(E-170/H-7600)
'    RstEnd = "Y": bEndChk = True: bSTXChk = False
'
'
'    On Error GoTo ErrPortOpen
'    If m_PortOpen = True Then
'        msComm.PortOpen = True
'    End If
'    On Error GoTo 0
'ErrPortOpen:
'    If Err <> 0 Then
'        MsgBox "PortOpen Error!!! " & Err.Description, vbCritical
'        RaiseEvent DispMsg(Err.Description)
'    End If
'End Property

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
    m_p_sSpcCd = m_def_p_sSpcCd
    m_p_sCmt1 = m_def_p_sCmt1
    m_Port = m_def_Port
    m_IPAddress = m_def_IPAddress
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
    
    Call EqOutput_Socket(Chr(iChr))
    On Error GoTo 0
    
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
    End If
End Function

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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get Port() As Variant
    Port = m_Port
End Property

Public Property Let Port(ByVal New_Port As Variant)
    m_Port = New_Port
    PropertyChanged "Port"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get IPAddress() As Variant
    IPAddress = m_IPAddress
End Property

Public Property Let IPAddress(ByVal New_IPAddress As Variant)
    m_IPAddress = New_IPAddress
    PropertyChanged "IPAddress"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function Connect() As Variant

End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function ConnectWinSock(Optional ByVal iGbn As Integer) As Variant
    On Error GoTo ErrRtn

    If iGbn = 0 Then
        '가장처음 Connect시 암호 확인
        If m_OpenPW <> pOpenPW Then
            MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
            Exit Function
        End If
        
        '변수 초기화(E-170/H-7600)
        RstEnd = "Y": bEndChk = True: bSTXChk = False
    End If
        
    miSckMax = 0
    With sckServer(miSckMax)
        .LocalPort = Val(m_Port)
        .Listen
    End With
            
    RaiseEvent DispMsg("WinSock State: " & sckServer(miSckMax).State)
        
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("ConnectWinSock Err - " & Err.Description)
    End If
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function CloseWinSock() As Variant
    On Error GoTo ErrClose
    
    sckServer(miSckMax).Close
    
ErrClose:
    If Err <> 0 Then
        RaiseEvent DispMsg("CloseWinSock Err - " & Err.Description)
    End If
End Function

