VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl COULTER 
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
Attribute VB_Name = "COULTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sPID = "0"
Const m_def_p_sData = "0"
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
Dim m_p_sPID As String
Dim m_p_sData As String
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
Event SendOrderOK(sID$, sRetCd$)
Event RequestCurOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
'Event SendOrderOK(sID$)
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

'For ACT5diff_ASTM
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim sNextSend   As String
Dim RstEnd      As String

Private Sub PhaseCfg_Protocol_LH750()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1      ''SYN, Blockcount 대기(datablock 이전의 대기상태)
                Select Case Asc(wkDat)
                    Case 22     'SYN에 해당
                        msComm.Output = Chr(22)     'SYN
                        RcvBuffer = RcvBuffer & wkDat   'wkBuf
                        m_iPhase = 1

                    Case Else   'blockcount-> 2 chars에 해당
                        msComm.Output = Chr(6)      'ACK
                        m_iPhase = 2
                End Select

            Case 2      ''datablock 수신 상태(one datablock의 끝인 ETX 이전까지)
                Select Case Asc(wkDat)
                    Case 3      'ETX
                        msComm.Output = Chr(6)      'ACK
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 3

                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                End Select

            Case 3      ''전송이 끝인지 or 다른 datablock 전송의 시작인지 판단하여 상태 변환
                Select Case Asc(wkDat)
                    Case 22     'SYN, 즉 전송의 끝
                        msComm.Output = Chr(6)      'ACK
                        RcvBuffer = RcvBuffer & wkDat

                        Call DataEdit_LH750

                        RcvBuffer = ""
                        m_iPhase = 1

                    Case 2  'STX, 즉 다른 datablock 전송 시작
                        'ix1 = ix1 + 3   'manual dataformat 참조 p.11
                        ''일단은 다 전송받고 edit_data에서 걸러내는 것으로 바꿈. 1998-05-21 김태윤
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2

                End Select

            '--- ORDER 전송 관련
            Case 4
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        Call SendOrder_LH750
                        m_iPhase = 5

                    Case 22     'SYN
                        msComm.Output = Chr(22)
                        m_iPhase = 1

                End Select

            Case 5
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        Call SendOrder_LH750
                        m_iPhase = 6

                    Case Else   'NAK -> RECEIVER ABORT
                        m_iPhase = 1

                End Select

            Case 6
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        msComm.Output = Chr(5)      'ENQ
                        m_iPhase = 7

                    Case Else
                        m_iPhase = 1

                End Select

            Case 7
                Select Case Asc(wkDat)
                    Case 6      'ACK

                    Case 16     'DLE
                        m_iPhase = 8

                End Select

            Case 8      'RETURN CODE 대기
                Select Case Asc(wkDat)
                    Case 65, 66, 67, 68, 69, 70     'A, B, C, D, E, F
                        m_iPhase = 1
                        m_iSendPhase = 1
                        RaiseEvent SendOrderOK(pSampleInfo.ID, wkDat)

                    Case Else
                        m_iPhase = 1
                        m_iSendPhase = 1
                        RaiseEvent SendOrderOK("", wkDat)

                End Select

        End Select
    Next ix1

End Sub
Private Sub SendOrder_LH750()
    On Error GoTo ErrRtn
    
    Dim sSend   As String * 256
    Dim sSendStr    As String
    Dim sChkSum As String
    
    Select Case m_iSendPhase
        Case 1
            RaiseEvent RequestCurOrder
            
            Call Get_OrderString
            
            If pSampleInfo.ORDCNT = 0 Then
                Exit Sub
            End If
                
            msComm.Output = "01"
            m_iSendPhase = m_iSendPhase + 1
            
            Exit Sub

        Case 2
            sSend = pSampleInfo.KIND
                
            sChkSum = ChkSum_LH750(sSend)
            
            sSendStr = Chr(2) & Format(1, "00") & sSend & sChkSum & Chr(3)
            
            msComm.Output = sSendStr
            
            If sTestMode = "77" Then
                RaiseEvent PrintSendLog(sSendStr)
            End If

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub

Private Function ChkSum_LH750(ByVal sPara As String) As String

   Dim crcmsb%, xt  As Single, tmpx%, tmpx2 As Integer
   Dim crclsb%, cr$, X%, test$, i%
   
   crcmsb = 255
   crclsb = 255
   
   For i = 1 To Len(sPara)
      X = Asc(Mid$(sPara, i, 1)) Xor crcmsb
      tmpx = Int(X / 16)
      X = X Xor tmpx
      xt = X
      tmpx = (xt * 16) Mod 256
      tmpx2 = Int(X / 8)
      crcmsb = (crclsb Xor tmpx2 Xor tmpx) Mod 256
      tmpx = (xt * 32) Mod 256
      crclsb = (X Xor tmpx) Mod 256
   Next i

   crclsb = crclsb Xor 255
   crcmsb = crcmsb Xor 255
   ChkSum_LH750 = Right("0" & Hex$(crcmsb), 2) & Right$("0" & Hex$(crclsb), 2)
    
End Function
'
'   STKS 1G.1 Format
'
Private Sub DataEdit_LH750_STKS()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim sTmp$, sTmp1$, sTmp2$, sTotIFCd$
    Dim sIFCd() As String
    Dim iPos%, iPos2%, ii%
    
        
    ''Data를 Edit하기 편리하도록
    ''<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>에서
    ''[DATA Block]부분만 제외하고 msRcvBuffer 제거한다.
    Do
        iPos = InStr(1, RcvBuffer, Chr(2))
        
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
        If iPos = 0 Then
            Exit Do
        End If
        
        sTmp1 = Left$(RcvBuffer, iPos - 1)
        sTmp2 = Mid$(RcvBuffer, iPos + 3)
        
        RcvBuffer = ""
        RcvBuffer = sTmp1 & sTmp2
    Loop While iPos <> 0
    
    Do
        iPos = InStr(1, RcvBuffer, Chr(3))
        
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
        If iPos = 0 Then
            Exit Do
        End If
        
        sTmp1 = Left$(RcvBuffer, iPos - 5)
        sTmp2 = Mid$(RcvBuffer, iPos + 1)
        
        RcvBuffer = ""
        RcvBuffer = sTmp1 & sTmp2
    Loop While iPos <> 0
    
    '결과구조체 초기화
    Call Init_pResultInfo
    
    
    '작업번호 구하기
    iPos = InStr(RcvBuffer, "CASS/POS")
    If iPos > 0 Then
        sTmp1 = Mid(RcvBuffer, iPos + 9, 7)
            
        If Mid(sTmp1, 5, 1) = "/" Then
            tmpRack = Left(sTmp1, 4)
            tmpPos = Right(sTmp1, 2)
        End If
    End If
    
    iPos = InStr(RcvBuffer, "ID1")
    If iPos > 0 Then
        sTmp2 = Mid(RcvBuffer, iPos + 4, 16)
        ii = InStr(1, sTmp2, vbCr)
        If ii <> 0 Then
            sTmp2 = Mid(sTmp2, 1, ii - 1)
        End If
        tmpBarCd = sTmp2
    End If
    
    iPos = InStr(RcvBuffer, "SEQUENCE")
    If iPos > 0 Then
        tmpSeqNo = Trim(Mid(RcvBuffer, iPos + 8, 7))
    End If
    
       
    '장비에서 검사할 수 있는 모든 항목 저장
    sTotIFCd = "WBC|RBC|HGB|HCT|MCV|MCH|MCHC|RDW|PLT|PCT|MPV|PDW|" _
            & "LY#|MO#|NE#|EO#|BA#|NRBC#|LY%|MO%|NE%|EO%|BA%|NRBC%"
    sIFCd() = Split(sTotIFCd, Chr(124))
    
    '검사명, 검사결과값 얻기
    For ii = 0 To UBound(sIFCd())
        iPos = InStr(RcvBuffer, Trim(sIFCd(ii)))
        
        If iPos > 0 Then
            sTmp = Trim(Mid(RcvBuffer, iPos + 4, 3))
            If sTmp = "Pop" Then
                iPos = 0
            End If
        End If
        
        If iPos > 0 Then
            iPos2 = InStr(iPos, RcvBuffer, Chr(13))
            sTmp = Trim(Mid(RcvBuffer, iPos, iPos2 - iPos))
            
            tmpIFCd = Trim(sIFCd(ii))
            
            sTmp = Trim(Mid(sTmp, Len(tmpIFCd) + 1))
            
            iPos2 = InStr(sTmp, " ")
            If iPos2 > 0 Then
                tmpRst = Trim(Mid(sTmp, 1, iPos2))
                tmpFlag = Trim(Mid(sTmp, iPos2))
            Else
                tmpRst = Trim(sTmp)
                tmpFlag = ""
            End If
            
'            tmpRst = Trim(Mid(sTmp, 5, 6))
'            tmpFlag = Trim(Mid(sTmp, 10))
        
            '--- 결과의 자릿수가 부족해 뒤의 Flag도 표시되는 경우 처리...(2000/11/14 yk)
            iPos = InStr(1, tmpRst, " ")
            If iPos <> 0 Then
                tmpRst = Trim(Mid(tmpRst, 1, iPos - 1))
            End If
            
            'STKS가 업그레이드 된 후 MCHC결과를 잘라내면 SOH가 뒤에 붙는 현상
            If IsNumeric(Right$(tmpRst, 1)) = True Then
            Else
                tmpRst = Left$(tmpRst, Len(tmpRst) - 1)
            End If
                
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .UNIT = .UNIT & Chr(124)
            End With
        End If
    Next ii
    
    '결과 처리
    With pResultInfo
        If .RSTCNT > 0 Then
            .ID = tmpBarCd
            .SEQNO = tmpSeqNo
            .RACK = tmpRack
            .POS = tmpPos
            
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub
'
'   LH-750 Format
'
Private Sub DataEdit_LH750()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim sTmp$, sTmp1$, sTmp2$, sTotIFCd$
    Dim sIFCd() As String
    Dim iPos%, iPos2%, ii%
    
        
    ''Data를 Edit하기 편리하도록
    ''<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>에서
    ''[DATA Block]부분만 제외하고 msRcvBuffer 제거한다.
    Do
        iPos = InStr(1, RcvBuffer, Chr(2))
        
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
        If iPos = 0 Then
            Exit Do
        End If
        
        sTmp1 = Left$(RcvBuffer, iPos - 1)
        sTmp2 = Mid$(RcvBuffer, iPos + 3)
        
        RcvBuffer = ""
        RcvBuffer = sTmp1 & sTmp2
    Loop While iPos <> 0
    
    Do
        iPos = InStr(1, RcvBuffer, Chr(3))
        
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
        If iPos = 0 Then
            Exit Do
        End If
        
        sTmp1 = Left$(RcvBuffer, iPos - 5)
        sTmp2 = Mid$(RcvBuffer, iPos + 1)
        
        RcvBuffer = ""
        RcvBuffer = sTmp1 & sTmp2
    Loop While iPos <> 0
    
    '결과구조체 초기화
    Call Init_pResultInfo
    
    
    '작업번호 구하기
    iPos = InStr(RcvBuffer, "ID1")
    If iPos > 0 Then
        sTmp2 = Mid(RcvBuffer, iPos + 4, 16)
        ii = InStr(1, sTmp2, vbCr)
        If ii <> 0 Then
            sTmp2 = Mid(sTmp2, 1, ii - 1)
        End If
        tmpBarCd = sTmp2
    End If
    
    iPos = InStr(RcvBuffer, "CASSPOS")
    If iPos > 0 Then
        sTmp1 = Mid(RcvBuffer, iPos + 9, 6)
            
        tmpRack = Left(sTmp1, 4)
        tmpPos = Right(sTmp1, 2)
    End If
    
'    iPos = InStr(RcvBuffer, "SEQUENCE")
'    If iPos > 0 Then
'        tmpSeqNo = Trim(Mid(RcvBuffer, iPos + 8, 7))
'    End If
    
       
    '장비에서 검사할 수 있는 모든 항목 저장
    sTotIFCd = "WBC|RBC|HGB|HCT|MCV|MCH|MCHC|RDW|PLT|PCT|MPV|PDW|" _
            & "LY#|MO#|NE#|EO#|BA#|NRBC#|LY%|MO%|NE%|EO%|BA%|NRBC%|" _
            & "RET%|RET#|MRV|MSCV|IRF|HLR%|HLR#"
    sIFCd() = Split(sTotIFCd, Chr(124))
    
    '검사명, 검사결과값 얻기
    For ii = 0 To UBound(sIFCd())
        iPos = InStr(RcvBuffer, Trim(sIFCd(ii)))
        
        If iPos > 0 Then
            sTmp = Trim(Mid(RcvBuffer, iPos + 4, 3))
            If sTmp = "Pop" Then
                iPos = 0
            ElseIf sTmp = "IS" Then
                iPos = InStr(iPos + 4, RcvBuffer, Trim(sIFCd(ii)))
            End If
        End If
        
        If iPos > 0 Then
            iPos2 = InStr(iPos, RcvBuffer, Chr(13))
            sTmp = Trim(Mid(RcvBuffer, iPos, iPos2 - iPos))
            
            tmpIFCd = Trim(sIFCd(ii))
            
            sTmp = Trim(Mid(sTmp, Len(tmpIFCd) + 1))
            
            iPos2 = InStr(sTmp, " ")
            If iPos2 > 0 Then
                tmpRst = Trim(Mid(sTmp, 1, iPos2))
                tmpFlag = Trim(Mid(sTmp, iPos2))
            Else
                tmpRst = Trim(sTmp)
                tmpFlag = ""
            End If
            
'            tmpRst = Trim(Mid(sTmp, 5, 6))
'            tmpFlag = Trim(Mid(sTmp, 10))
        
            '--- 결과의 자릿수가 부족해 뒤의 Flag도 표시되는 경우 처리...(2000/11/14 yk)
            iPos = InStr(1, tmpRst, " ")
            If iPos <> 0 Then
                tmpRst = Trim(Mid(tmpRst, 1, iPos - 1))
            End If
            
            'STKS가 업그레이드 된 후 MCHC결과를 잘라내면 SOH가 뒤에 붙는 현상
            If IsNumeric(Right$(tmpRst, 1)) = True Then
            Else
                tmpRst = Left$(tmpRst, Len(tmpRst) - 1)
            End If
                
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .UNIT = .UNIT & Chr(124)
            End With
        End If
    Next ii
    
    '결과 처리
    With pResultInfo
        If .RSTCNT > 0 Then
            .ID = tmpBarCd
            .SEQNO = tmpSeqNo
            .RACK = tmpRack
            .POS = tmpPos
            
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
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
        Case "LH750"
            Call PhaseCfg_Protocol_LH750
        
        Case "STKS"
            Call PhaseCfg_Protocol_STKS
            
        Case "COULTERT540"
            Call PhaseCfg_Protocol_COULTERT540  '<2008/??/??    mc
            
        Case "ACT5DIFF_ASTM"
            Call PhaseCfg_Protocol_ACT5diff_ASTM  '<2008/05/22    mc
        
        Case "ACTDIFF_ASTM"
            Call PhaseCfg_Protocol_ACTdiff_ASTM  '<2008/07/01    sm
            
        Case "COULTERJT"
            Call PhaseCfg_Protocol_COULTERJT  '<2008/??/??    mc
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_STKS()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
       
        Select Case m_iPhase
            Case 1  'SYN, Blockcount 대기(datablock 이전의 대기상태)
                Select Case Asc(wkDat)
                    Case 22 'SYN에 해당
                        msComm.Output = Chr(22)     'SYN
                        RcvBuffer = ""
                        RcvBuffer = RcvBuffer & wkDat   'wkBuf
                        m_iPhase = 1
                        
                    Case Else   'blockcount-> 2 chars에 해당
                        msComm.Output = Chr(6)   'ACK
                        m_iPhase = 2
                        
                End Select
                
            Case 2  'datablock 수신 상태(one datablock의 끝인 ETX 이전까지)
                Select Case Asc(wkDat)
                    Case 3  'ETX
                        msComm.Output = Chr(6)   'ACK
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 3
                    
                    Case 22 'SYN
                        msComm.Output = Chr(22)  'SYN
                        RcvBuffer = ""
                        m_iPhase = 1
                    
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                        
                End Select
         
            Case 3  '전송이 끝인지 or 다른 datablock 전송의 시작인지 판단하여 상태 변환
                Select Case Asc(wkDat)
                    Case 22 'SYN, 즉 전송의 끝
                        msComm.Output = Chr(6)   'ACK
                        RcvBuffer = RcvBuffer & wkDat
                        
                        '--- Data 편집
                        Call DataEdit_STKS
                        
                        RcvBuffer = ""
                        m_iPhase = 1
                        
                    Case 2  'STX, 즉 다른 datablock 전송 시작
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                        
                End Select
                
        End Select
    Next ix1
    
End Sub
'
'  STKS 결과값 편집
'
Private Sub DataEdit_STKS()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim sTmp$, sTmp1$, sTmp2$, sTotIFCd$
    Dim sIFCd() As String
    Dim iPos%, iPos2%, ii%
            
'    '--- [DATA Block]부분만 제외하고 RcvBuffer에서 제거
'    RcvBuffer = p_Filter_DataBlock(RcvBuffer)
    
    ''Data를 Edit하기 편리하도록
    ''<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>에서
    ''[DATA Block]부분만 제외하고 msRcvBuffer 제거한다.
    Do
        iPos = InStr(1, RcvBuffer, Chr(2))
        
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
        If iPos = 0 Then
            Exit Do
        End If
        
        sTmp1 = Left$(RcvBuffer, iPos - 1)
        sTmp2 = Mid$(RcvBuffer, iPos + 3)
        
        RcvBuffer = ""
        RcvBuffer = sTmp1 & sTmp2
    Loop While iPos <> 0
    
    Do
        iPos = InStr(1, RcvBuffer, Chr(3))
        
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
        If iPos = 0 Then
            Exit Do
        End If
        
        sTmp1 = Left$(RcvBuffer, iPos - 5)
        sTmp2 = Mid$(RcvBuffer, iPos + 1)
        
        RcvBuffer = ""
        RcvBuffer = sTmp1 & sTmp2
    Loop While iPos <> 0
    
    '결과구조체 초기화
    Call Init_pResultInfo
    
    '작업번호/BARCODE 구하기
    iPos = InStr(RcvBuffer, "ID1")
    If iPos > 0 Then
        sTmp2 = Mid(RcvBuffer, iPos + 4, 11)
        ii = InStr(1, sTmp2, vbCr)
        If ii <> 0 Then
            sTmp2 = Mid(sTmp2, 1, ii - 1)
        End If
        tmpBarCd = sTmp2
    End If
    
    iPos = InStr(RcvBuffer, "CASSPOS")
    If iPos > 0 Then
        sTmp1 = Mid(RcvBuffer, iPos + 9, 6)
            
        tmpRack = Left(sTmp1, 4)
        tmpPos = Right(sTmp1, 2)
    End If
    
    '장비에서 검사할 수 있는 모든 항목 저장
    sTotIFCd = "WBC|RBC|HGB|HCT|MCV|MCH|MCHC|RDW|PLT|PCT|MPV|PDW|" _
            & "LY#|MO#|NE#|EO#|BA#|NRBC#|LY%|MO%|NE%|EO%|BA%|NRBC%|" _
            & "RET%|RET#|MRV|MSCV|IRF|HLR%|HLR#"
    sIFCd() = Split(sTotIFCd, Chr(124))
    
    '검사명, 검사결과값 얻기
    For ii = 0 To UBound(sIFCd())
        iPos = InStr(RcvBuffer, Trim(sIFCd(ii)))
        
        If iPos > 0 Then
            sTmp = Trim(Mid(RcvBuffer, iPos + 4, 3))
            If sTmp = "Pop" Then
                iPos = 0
            ElseIf sTmp = "IS" Then
                iPos = InStr(iPos + 4, RcvBuffer, Trim(sIFCd(ii)))
            End If
        End If
        
        If iPos > 0 Then
            iPos2 = InStr(iPos, RcvBuffer, Chr(13))
            sTmp = Trim(Mid(RcvBuffer, iPos, iPos2 - iPos))
            
            tmpIFCd = Trim(sIFCd(ii))
            
            sTmp = Trim(Mid(sTmp, Len(tmpIFCd) + 1))
            
            iPos2 = InStr(sTmp, " ")
            If iPos2 > 0 Then
                tmpRst = Trim(Mid(sTmp, 1, iPos2))
                tmpFlag = Trim(Mid(sTmp, iPos2))
            Else
                tmpRst = Trim(sTmp)
                tmpFlag = ""
            End If
            
'            tmpRst = Trim(Mid(sTmp, 5, 6))
'            tmpFlag = Trim(Mid(sTmp, 10))
        
            '--- 결과의 자릿수가 부족해 뒤의 Flag도 표시되는 경우 처리...(2000/11/14 yk)
            iPos = InStr(1, tmpRst, " ")
            If iPos <> 0 Then
                tmpRst = Trim(Mid(tmpRst, 1, iPos - 1))
            End If
            
            'STKS가 업그레이드 된 후 MCHC결과를 잘라내면 SOH가 뒤에 붙는 현상
            If IsNumeric(Right$(tmpRst, 1)) = True Then
            Else
                tmpRst = Left$(tmpRst, Len(tmpRst) - 1)
            End If
                
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .UNIT = .UNIT & Chr(124)
            End With
        End If
    Next ii
    
    '결과 처리
    With pResultInfo
        If .RSTCNT > 0 Then
            .ID = tmpBarCd
            .SEQNO = tmpSeqNo
            .RACK = tmpRack
            .POS = tmpPos
            
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_COULTERT540()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim sTmp$, sTmp1$, sTmp2$, sTotIFCd$
    Dim sIFCd() As String
    Dim iPos%, iPos2%, ii%
    
    Dim aData() As String
    Dim aField() As String
    Dim iCnt As Integer
    
''    -----------------
''    Date , 10 / 0 / 0
''    test , 6
''     WBC , 7.5
''     RBC , 4.58
''     HGB , 13.03
''     HCT , 42.99
''     MCV , 93.84
''     MCH , 28.41
''     MCHC , 30.27
''     PLT , 202.6
''     LY % ,  37.47
''     LY # ,   2.80
''    -----------------
''    ED
    
    aData = Split(RcvBuffer, vbCrLf)
    sTotIFCd = "WBC|RBC|HGB|HCT|MCV|MCH|MCHC|PLT|LY %|LY #"
    
    For iCnt = 0 To UBound(aData) - 1
        If aData(iCnt) <> "" Then
            If InStr(aData(iCnt), ",") > 0 Then
                aField = Split(aData(iCnt), ",")
                                
                If InStr(sTotIFCd, Trim(aField(0))) > 0 Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & Trim(aField(0)) & Chr(124)
                        .RST1 = .RST1 & Trim(aField(1)) & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                    End With
                End If
            End If
            
        End If
     Next
    
    '결과 처리
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
    
    '결과구조체 초기화
    Call Init_pResultInfo
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_COULTERJT()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim sTmp$, sTmp1$, sTmp2$, sTotIFCd$
    Dim sIFCd() As String
    Dim iPos%, iPos2%, ii%
    
    Dim aData() As String
    Dim aField() As String
    Dim iCnt As Integer
    
''    ---------------
''    Date , 8 / 6 / 30
''    test , 001
''      $
''    ID ,
''     WBC , 0#
''     RBC , 0#
''     HGB , 0#
''     HCT  , -----
''     MCV  , -----
''     MCH  , .....
''     MCHC , -----
''     RDW  , -----
''     PLT , 0#
''     PCT , 0#
''     MPV , 1.7
''     PDW , 10#
''     LY   , .....
''     MO   , .....
''     GR   , .....
''     LY # , .....
''     MO # , .....
''     GR # , .....
''    --------------
    
    aData = Split(RcvBuffer, vbCrLf)
    sTotIFCd = "WBC|RBC|HGB|HCT|MCV|MCH|MCHC|RDW|PLT|PCT|MPV|PDW|LY|MO|GR|LY #|MO #|GR #|"
    
    For iCnt = 0 To UBound(aData) - 1
        If aData(iCnt) <> "" Then
            If InStr(aData(iCnt), ",") > 0 Then
                aField = Split(aData(iCnt), ",")
                
                If UCase(Trim(aField(0))) = "TEST" Then
                    pResultInfo.SEQNO = Trim(aField(1))
                End If
                                
                If InStr(sTotIFCd, Trim(aField(0))) > 0 Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & Trim(aField(0)) & Chr(124)
                        .RST1 = .RST1 & Trim(aField(1)) & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                    End With
                End If
            End If
            
        End If
     Next
    
    '결과 처리
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
    
    '결과구조체 초기화
    Call Init_pResultInfo
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_COULTERT540()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
       
        Select Case Asc(wkDat)
            Case 2  'STX
                RcvBuffer = ""
                'msComm.Output = Chr(6)
            
            Case 3  'ETX
                Call DataEdit_COULTERT540
                RcvBuffer = ""
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_COULTERJT()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
       
        Select Case Asc(wkDat)
            Case 2  'STX
                RcvBuffer = ""
                'msComm.Output = Chr(6)
            
            Case 3  'ETX
                Call DataEdit_COULTERJT
                RcvBuffer = ""
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub

Private Sub DataEdit_ACT5diff_ASTM()
    On Error GoTo ErrHandler
    
    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpKind     As String
    Dim tmpSampType As String
    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpIFCd$, tmpIFCd2$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarmCd$, tmpInstID$
    Dim tmpRstDT$, tmpCmt$

    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
        Case "P"        'Patient Record
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

            Call Init_pResultInfo

        Case "O"
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpField() = Split(RcvBuffer, "|")
            ii = InStr(1, tmpField(2), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpField(2), "^")
                tmpBarCd = Trim(tmpData(0))
                tmpRack = Trim(tmpData(1))
                tmpPos = Trim(tmpData(2))
            Else
                tmpBarCd = Trim(tmpField(2))
            End If
            
            tmpRstDT = Trim(tmpField(7))
            tmpKind = Trim(tmpField(11))
            tmpCmt = Trim(tmpField(25))

            pSampleInfo.ID = tmpBarCd
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            pSampleInfo.KIND = tmpKind
            pSampleInfo.RSTDT = tmpRstDT        '결과일시
            pSampleInfo.CMT1 = tmpCmt           'F: Final, C: corrected report(rerun)

        Case "R"        'Result Record
            '--- 결과데이타 편집
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '5:Reference Ranges
            '6:Result Abnormal Flags
            '8:Result Status
            ' W: suspicion
            ' N: rejected result
            ' F: final result
            ' X: All hematology parameters except BAS# and BAS%
            ' S: BAS# and BAS%
            ' C: Platelet Concentrate Mode
            
            tmpField() = Split(RcvBuffer, "|")
            
            ii = InStr(1, tmpField(2), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpField(2), "^")
                tmpIFCd = Trim(tmpData(3))
                
                If UBound(Split(tmpField(2), "^")) > 3 Then
                    tmpIFCd2 = Trim(tmpData(4))
                Else
                    tmpIFCd2 = ""
                End If
            End If

            tmpRst = Trim(tmpField(3))
            tmpUnit = Trim(tmpField(4))
            tmpFlag = Trim(tmpField(8))
            tmpRstDT = Trim(tmpField(12))

            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS
                .KIND = pSampleInfo.KIND
                .OTHER = pSampleInfo.CMT1
                
                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .INSTID = .INSTID & tmpInstID & Chr(124)
                .RSTDT = .RSTDT & tmpRstDT & Chr(124)
            End With

        Case "C"        'Comment Record
            'Data Alarm 편집
            tmpData() = Split(RcvBuffer, Chr(124))

            tmpAlarmCd = Trim(tmpData(3))
            pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)

        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

            Call Init_pResultInfo

    End Select
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub


Private Sub DataEdit_ACTdiff_ASTM()
    On Error GoTo ErrHandler
    
    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpKind     As String
    Dim tmpSampType As String
    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpIFCd$, tmpIFCd2$, tmpRst$, tmpUnit$, tmpFlag$, tmpAlarmCd$, tmpInstID$
    Dim tmpRstDT$, tmpCmt$

    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
        Case "P"        'Patient Record
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

            Call Init_pResultInfo

        Case "O"
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpField() = Split(RcvBuffer, "|")
            
            tmpSeqNo = Trim(tmpField(2))
          
            pSampleInfo.ID = ""
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = ""
            pSampleInfo.POS = ""
            pSampleInfo.KIND = ""
            pSampleInfo.RSTDT = ""
            pSampleInfo.CMT1 = ""

        Case "R"        'Result Record
            '--- 결과데이타 편집
            
            tmpField() = Split(RcvBuffer, "|")
            
            ii = InStr(1, tmpField(2), "!")
            If ii <> 0 Then
                tmpData() = Split(tmpField(2), "!")
                tmpIFCd = Trim(tmpData(3))
            End If

            tmpRst = Trim(tmpField(3))
            
            ii = InStr(tmpRst, "!")
            If ii > 0 Then
                tmpFlag = Mid(tmpRst, ii + 1)
                tmpRst = Mid(tmpRst, 1, ii - 1)
            End If
            
            tmpUnit = Trim(tmpField(4))
            tmpRstDT = Trim(tmpField(12))

            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS
                .KIND = pSampleInfo.KIND
                .OTHER = pSampleInfo.CMT1
                
                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .INSTID = .INSTID & tmpInstID & Chr(124)
                .RSTDT = .RSTDT & tmpRstDT & Chr(124)
            End With

        Case "C"        'Comment Record
            'Data Alarm 편집
            tmpData() = Split(RcvBuffer, Chr(124))

            tmpAlarmCd = Trim(tmpData(3))
            pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)

        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

            Call Init_pResultInfo

    End Select
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_ACT5diff_ASTM()
    On Error GoTo ErrRtn
    
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
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

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
                            Call DataEdit_ACT5diff_ASTM
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEdit_ACT5diff_ASTM
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If sState = "Q" Then
                            msComm.Output = Chr(5)
                            
                            If m_sTestMode = "77" Then
                                RaiseEvent PrintSendLog(Chr(5))
                            End If
                        
                            m_iSendPhase = 1
                        End If
                        
                        m_iPhase = 3

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        msComm.Output = Chr(6)   'Send ACK
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 21     'NAK
                        Call DataEdit_ACT5diff_ASTM

                        m_iSendPhase = 1
                        m_iFrameN = 1

                        msComm.Output = Chr(5)   'Send ENQ
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(5))
                        End If

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
                        Call SendOrder_ACT5diff_ASTM

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                        m_iPhase = 2

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        msComm.Output = Chr(5)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(5))
                        End If
                        
                        m_iPhase = 3

                    Case 4      'EOT
                        m_iPhase = 1

                End Select
        End Select
    Next ix1

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
    
End Sub

Private Sub PhaseCfg_Protocol_ACTdiff_ASTM()
    On Error GoTo ErrRtn
    
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
                            Call DataEdit_ACTdiff_ASTM
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEdit_ACTdiff_ASTM
                            RcvBuffer = ""
                        End If

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        msComm.Output = Chr(6)   'Send ACK

                    Case 21     'NAK
                        Call DataEdit_ACTdiff_ASTM

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
        End Select
    Next ix1

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
    
End Sub

Private Sub SendOrder_ACT5diff_ASTM()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sStat   As String
    
    Select Case m_iSendPhase
        Case 0
            m_iSendPhase = 1
            msComm.Output = Chr(5)
            Exit Sub
        
        Case 1
            'Header Record
            'sSendBuff = m_iFrameN & "H|\^&|||BCI|||||||P|D1394-97|20080522154814" & vbCr
            sSendBuff = m_iFrameN & "H|\^&||||||||||||" & vbCr
            
            Call Get_OrderString
            
            'Patient Record
            sSendBuff = sSendBuff & "P|1||" & pSampleInfo.ID & "||^|||U|||||Physician||||||||||||Location|||||||||" & vbCr
            
            'Order Record
            'sSendBuff = sSendBuff & "O|1|" & pSampleInfo.ID & "|" & "^^^" & pSampleInfo.IFCD(1) & "|||" & Format(Now, "yyyyMMddHHmmss") & "||||||||||||||||||O|||||"
            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.ID & "||" & "^^^" & pSampleInfo.IFCD(1) & "|||" & Format(Now, "yyyyMMddHHmmss") & "||||||||||||||||||O|||||" & vbCr
                    
            'Terminator Record
            sSendBuff = sSendBuff & "L|1|N"

            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 241 Then
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
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If
                        
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1
            
            RaiseEvent SendOrderOK("", "")

            sState = "": sReqStatusCd = ""

            'BarCode Mode가 아닌 경우 다음 오더 조회
            RaiseEvent RequestCurOrder
    
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

Public Function GetByOne(ByVal tStr As String, sOriginal As String) As String
    Dim POS%
    
    POS = InStr(tStr, Chr$(124))
    
    If POS = 0 Then
    Else
        GetByOne = Trim$(Mid$(tStr, 1, POS - 1))
        sOriginal = Trim$(Mid$(sOriginal, POS + 1, Len(sOriginal) - POS))
    End If
End Function

Public Function GetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim POS%

    POS = InStr(tStr, sUserSymbol)

    If POS = 0 Then
    Else
        GetByOneUserSymbol = Trim$(Mid$(tStr, 1, POS - 1))
        sOriginal = Trim$(Mid$(sOriginal, POS + 1, Len(sOriginal) - POS))
    End If
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
    m_p_sPID = PropBag.ReadProperty("p_sPID", m_def_p_sPID)
    m_p_sData = PropBag.ReadProperty("p_sData", m_def_p_sData)
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
    Call PropBag.WriteProperty("p_sPID", m_p_sPID, m_def_p_sPID)
    Call PropBag.WriteProperty("p_sData", m_p_sData, m_def_p_sData)
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
    
    m_iOrderFlag = 0
    
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
    m_p_sPID = m_def_p_sPID
    m_p_sData = m_def_p_sData
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
'MemberInfo=13,0,0,0
Public Property Get p_sPID() As String
    p_sPID = m_p_sPID
End Property

Public Property Let p_sPID(ByVal New_p_sPID As String)
    m_p_sPID = New_p_sPID
    PropertyChanged "p_sPID"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sData() As String
    p_sData = m_p_sData
End Property

Public Property Let p_sData(ByVal New_p_sData As String)
    m_p_sData = New_p_sData
    PropertyChanged "p_sData"
End Property

