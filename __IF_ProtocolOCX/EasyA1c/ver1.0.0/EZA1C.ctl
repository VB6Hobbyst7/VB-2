VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl EZA1C 
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   ScaleHeight     =   2220
   ScaleWidth      =   1740
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
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
   Begin MSCommLib.MSComm msComm 
      Left            =   -30
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "EZA1C"
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
        Case "EASYA1C"
                
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Public Sub DataEditResponse_EasyA1C_Excel(ByVal sFileNm As String)
    On Error GoTo ErrRtn
    
    Dim intIdx    As Integer
    Dim RecType     As String       'Record Type
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpKind$, tmpName$, tmpAge$, TmpBirthDT$
    Dim tmpIFCd$, tmpRst$, tmpRst2$, tmpUnit$, tmpRef$, tmpFlag$, tmpRstDT$, tmpRstTM$, tmpCmt$
    Dim bRstFlag As Boolean
    Dim iSpaceCnt As Integer
    
    Dim objExcell   As Object
    
    Set objExcell = CreateObject("Excel.Application")
    
    objExcell.Workbooks.Open sFileNm
    
    objExcell.Visible = False
            
    bRstFlag = False
    intIdx = 0
        
    Do
        intIdx = intIdx + 1
        
        RecType = objExcell.Worksheets(1).Range("A" & CStr(intIdx)).Value
        
        If InStr(RecType, "환자") > 0 Then
            intIdx = intIdx + 2
            
            Do
                If iSpaceCnt = 30 Then
                    iSpaceCnt = 0
                End If
                iSpaceCnt = iSpaceCnt + 2
                
                RaiseEvent DispMsg(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
            
                tmpSeqNo = objExcell.Worksheets(1).Range("A" & CStr(intIdx)).Value      'Seq
                
                If IsNumeric(tmpSeqNo) Then
                    tmpRstDT = objExcell.Worksheets(1).Range("B" & CStr(intIdx)).Value      'Result Date
                    tmpRstTM = objExcell.Worksheets(1).Range("C" & CStr(intIdx)).Value      'Result Time
                    tmpRst = objExcell.Worksheets(1).Range("D" & CStr(intIdx)).Value        'Result
                    tmpBarCd = objExcell.Worksheets(1).Range("E" & CStr(intIdx)).Value      'BarCode
                    If tmpBarCd = 0 Then tmpBarCd = ""
                    tmpName = objExcell.Worksheets(1).Range("F" & CStr(intIdx)).Value       'Patient Name
                    tmpAge = objExcell.Worksheets(1).Range("G" & CStr(intIdx)).Value        'Patient Age
                    TmpBirthDT = objExcell.Worksheets(1).Range("H" & CStr(intIdx)).Value    'Patient BirthDay
                    tmpRst2 = objExcell.Worksheets(1).Range("I" & CStr(intIdx)).Value       'HbA1C Kind
                    tmpCmt = objExcell.Worksheets(1).Range("J" & CStr(intIdx)).Value        'Commet
                    
                    With pResultInfo
                        .ID = Trim(tmpBarCd)
                        .RST1 = Trim(tmpRst) & Chr(124)
                        .RST2 = Trim(tmpRst2) & Chr(124)
                        .FLAG = "" & Chr(124)
                        .IFCD = "HBA1C" & Chr(124)
                        .RSTDT = tmpRstDT
                        .RSTCNT = "1"
                        .OTHER = "EXCEL"
                        
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .KIND, .RSTDT, .OTHER)
        
                    End With
        
                    Call Init_pResultInfo
                    
                    bRstFlag = True
                Else
                    Exit Do
                End If
                
                intIdx = intIdx + 1
            Loop
        End If
        
        If bRstFlag = True Then
            Exit Do
        End If
    Loop
                
    objExcell.Workbooks.Close
    Set objExcell = Nothing
    
    Exit Sub
       
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
    
    objExcell.Workbooks.Close
    Set objExcell = Nothing

End Sub

Public Sub DataEditResponse_EasyA1C_File(ByVal sRcvBuf As String)
    On Error GoTo ErrRtn
    
    Dim i As Integer
    Dim intIdx    As Integer
    Dim RecType     As String       'Record Type
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpKind$, tmpName$, tmpAge$, TmpBirthDT$
    Dim tmpIFCd$, tmpRst$, tmpRst2$, tmpUnit$, tmpRef$, tmpFlag$, tmpRstDT$, tmpRstTM$, tmpCmt$
    Dim bRstFlag As Boolean
    Dim iSpaceCnt As Integer
    
    Dim aData() As String
    Dim aField() As String
    
''    200903070227,057,00000000000000,,,-1,,
''    200903082022,059,88011160004560,sd,32,2,d,
''    200903082036,073,88014036000000,sd,3,0,,
''    200903090232,062,00000000000000,,,-1,,
        
    aData = Split(sRcvBuf, Chr(3))
    
    For i = 0 To UBound(aData) - 1
        If aData(i) <> "" Then
            aField = Split(aData(i), ",")
            
            tmpRstDT = Trim(aField(0))
            tmpRst = Trim(CStr((Val(aField(1)))))
            tmpBarCd = Trim(CStr((Val(aField(2))))): If tmpBarCd = 0 Then tmpBarCd = ""
            tmpCmt = Trim(aField(6))
            
            With pResultInfo
                .ID = Trim(tmpBarCd)
                .RST1 = Trim(tmpRst) & Chr(124)
                .RST2 = Trim(tmpRst2) & Chr(124)
                .FLAG = "" & Chr(124)
                .IFCD = "HBA1C" & Chr(124)
                .RSTDT = tmpRstDT
                .RSTCNT = "1"
                .OTHER = "FILE"
                
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .KIND, .RSTDT, .OTHER)
    
            End With
    
            Call Init_pResultInfo
        End If
    Next
       
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If

End Sub

Public Sub Delete_ResultFile(ByVal sFileRoot As String)
    On Error GoTo ErrHandler
    
    Dim fs, f, f1, fc, s, FileRt, file
    Dim iFileCnt    As Integer
    Dim sRcvBuf As String
    Dim sRecType As String
            
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(sFileRoot)
    Set fc = f.Files
    
    For Each f1 In fc
        s = f1.Name
        
        If IsNumeric(Split(s, ".")(0)) = True And UCase(Trim(Split(s, ".")(1))) = "HDT" Then
            Set file = fs.GetFile(sFileRoot & "\" & s)
    
            file.Delete True
        End If
    Next
    
    Exit Sub
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("Delete_ResultFile Error - " & Err.Description)
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
Private Sub SendOrder_DPE_Batch()
'    On Error GoTo Err_Rtn
'
'    Dim sSendBuff   As String
'    Dim iCnt    As Integer
'    Dim ChkSum  As String
'    Dim sStat   As String
'
'    Select Case m_iSendPhase
'        Case 0
'            m_iSendPhase = 1
'            msComm.Output = Chr(5)
'            Exit Sub
'
'        Case 1
'            'Header Record
''            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^REPLY|P|1" & vbCr
'            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^BATCH|P|1" & vbCr
'
'            'Patient Record
'            sSendBuff = sSendBuff & "P|1" & vbCr
'
''            '----- 검사항목 조회
''            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.KIND)
'
'            Call Get_OrderString
'
'            'Order Record
''            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) _
''                    & "^" & pSampleInfo.SPCCD & "^" _
''                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|" & pSampleInfo.KIND & "|"
'            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) _
'                    & "^1^" _
'                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|R1|"
'
'            '검사항목 Order코드 추가
'            For iCnt = 1 To pSampleInfo.ORDCNT
'                '정상 오더
'                Select Case Trim$(pSampleInfo.IFCD(iCnt))
'                    Case "989", "990", "991"
'                        'ISE 항목이 있는 경우 전체 검사(ISE 검사는 2가지 조합만 오더 가능)
'                        If Val(InStr(1, sSendBuff, "989" & "/")) = 0 Then
'                            sSendBuff = sSendBuff & "^^^989/\"
'                        End If
'                        If Val(InStr(1, sSendBuff, "990" & "/")) = 0 Then
'                            sSendBuff = sSendBuff & "^^^990/\"
'                        End If
'                        If Val(InStr(1, sSendBuff, "991" & "/")) = 0 Then
'                            sSendBuff = sSendBuff & "^^^991/\"
'                        End If
'
'                    Case Else
'                        '일반항목
'                        sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "/\"
'                End Select
'            Next iCnt
'
'            If pSampleInfo.SINDEX = True Then           'Serum Index
'                sSendBuff = sSendBuff & "^^^992/\^^^993/\^^^994/\"
'            End If
'
'            If pSampleInfo.ORDCNT > 0 And Trim(sReqStatusCd) <> "A" Then
'                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
'            End If
'
'            'STAT RACK에 대한 처리추가
''            If Left(pSampleInfo.RACK, 1) = "4" Then
''                sStat = "S"
''            Else
'                sStat = "R"
''            End If
'
'            If Trim(pSampleInfo.CMT1) = "" Then
'                sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
'                        & "^^^^||||||O" & vbCr
'            Else
'                'Comment 정보 존재할 경우 Comment1 란에 특정정보 전송(2005/8/1 yk)
'                sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
'                        & Trim(pSampleInfo.CMT1) & "^^^^||||||O" & vbCr
'            End If
'
'            'Terminator Record
'            sSendBuff = sSendBuff & "L|1|N"
'
'
'            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
'            If Len(sSendBuff) >= 241 Then
'                sNextSend = Mid(sSendBuff, 241)
'                sSendBuff = Left(sSendBuff, 240)
'                sSendBuff = sSendBuff & Chr(23)
'
'                m_iFrameN = m_iFrameN + 1
'                m_iSendPhase = 2
'            Else
'                sSendBuff = sSendBuff & Chr(13) & Chr(3)
'                GoTo Send_Terminate
'            End If
'
'        Case 2
'            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
'            sNextSend = ""
'
'Send_Terminate:
'            m_iSendPhase = 3
'
'        Case 3      'EOT
'            msComm.Output = Chr(4)   'EOT
'            m_iFrameN = 1
'            m_iPhase = 3
'            m_iSendPhase = 1
'
'            sState = "": sReqStatusCd = ""
'
'            'BarCode Mode가 아닌 경우 다음 오더 조회
'            RaiseEvent RequestNextOrder
'
'            Exit Sub
'    End Select
'
'    ChkSum = ChkSum_ASTM(sSendBuff)
'    sSendBuff = sSendBuff & ChkSum
'    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)
'
'    If m_sTestMode = "77" Then
'        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
'    End If
'
''    '전송된 오더가 있는 경우 화면표시
''    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
''        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
''            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
''        End If
''    Else
''        '조회된 내용이 없는 경우 환자정보 구조체 초기화
''        Call Init_pResultInfo
''
''        RaiseEvent SendOrderOK("", "", "", "")
''    End If
'
'Err_Rtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
'    End If
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
            .SINDEX = False
            .CMT1 = ""
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
        .SINDEX = m_p_bSIndex
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
        
        .CMT1 = m_p_sCmt1
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
    
    '변수 초기화
    bEndChk = True: bSTXChk = False
    
    
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

