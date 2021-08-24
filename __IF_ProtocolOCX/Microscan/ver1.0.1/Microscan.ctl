VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl Microscan 
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
Attribute VB_Name = "Microscan"
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
Event AppendData(sID$, sISol$, sSpcCd$, sSpcNm$, sKitCd$, sKitNm$, sOrgCd$, sOrgNm$, sAntiCnt$, sAntiCd$, sAntiNm$, sSRI$, sMIC$, sUrine$, sStatus$, sRstDt$)
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
Private pSampleInfo As SAMPLE_INFO_M
Private pResultInfo As RESULT_INFO_M

Private mManual() As MANUALTBL
Private mPanel() As PANELTBL

'기타
Dim iSpaceCnt   As Integer

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
        Case "MICROSCAN"
            Call PhaseCfg_Protocol_Microscan
                    
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_Microscan()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1 As Integer
    Dim i   As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1  '대기
                Select Case Asc(wkDat)
                    Case 2      'STX
                        m_iPhase = 2
                        
                    Case 5      'ENQ
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case Else
                        m_iPhase = 1
                        
                End Select

            Case 2  '결과 수신
                Select Case Asc(wkDat)
                    Case 2     'STX
                    
                    Case 5      'ENQ
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 3     'ETX
                        Call DataEditResponse_Microscan
                        RcvBuffer = ""
                        
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat

                End Select

            Case 3  '오더 전송
                Select Case Asc(wkDat)
                    Case 3     'ETX
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                    Case 4      'EOT
                        If sState = "Q" Then
                            m_iSendPhase = 1
                            msComm.Output = Chr(5)
                            
                            If m_sTestMode = "77" Then
                                RaiseEvent PrintSendLog(Chr(5))
                            End If
                        Else
                            
                        End If
                        
                    Case 6      'ACK
                        Call SendOrder_Microscan

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iPhase = 3

                End Select

        End Select
    Next ix1
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub


' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_Microscan()
    On Error GoTo ErrRtn

    Dim RecType     As String   'Record Type
    Dim tmpField()  As String
    
    Dim TmpAntiCd   As String
    Dim TmpAntiNm   As String
    Dim tmpSRI      As String
    Dim tmpMIC      As String
    Dim tmpRstFlag  As String
    Dim tmpAntiSeq  As String
    
    RecType = Mid$(RcvBuffer, 1, 1)
    
    If RecType = """" Then
        RecType = Mid(RcvBuffer, 2, 1)
    End If

    Select Case RecType
        Case "H"        'Header Record
            Call Init_pResultInfo
            
        Case "P"        'Patient Record
            Call Init_pResultInfo
            tmpField() = Split(RcvBuffer, ",")
            
        Case "A"
            sState = "Q"
            m_iSendPhase = 1
            m_iPhase = 3
            
        Case "B"
            tmpField() = Split(RcvBuffer, ",")
            
            If Trim(tmpField(2)) <> "" Then
                pResultInfo.ID = Trim(Mid(Trim(tmpField(2)), 2, Len(Trim(tmpField(2))) - 2))
            End If
            
            If Trim(tmpField(3)) <> "" Then
                pResultInfo.PATNO = Trim(Mid(Trim(tmpField(3)), 2, Len(Trim(tmpField(3))) - 2))
            End If
            
            If Trim(tmpField(6)) <> "" Then
                pResultInfo.SPCCD = Trim(Mid(Trim(tmpField(6)), 2, Len(Trim(tmpField(6))) - 2))
            End If
            
            If Trim(tmpField(7)) <> "" Then
                pResultInfo.SPCNM = Trim(Mid(Trim(tmpField(7)), 2, Len(Trim(tmpField(7))) - 2))
            End If
            
            If Trim(tmpField(8)) <> "" Then
                pResultInfo.URINE = Trim(Mid(Trim(tmpField(8)), 1, 1))
            Else
                pResultInfo.URINE = "N"
            End If
            
        Case "F"

        Case "R"
            tmpField() = Split(RcvBuffer, ",")
            
            If Trim(tmpField(2)) <> "" Then
                pResultInfo.ISOL = Trim(Mid(Trim(tmpField(2)), 2, Len(Trim(tmpField(2))) - 2))
            End If
            
            If Trim(tmpField(4)) <> "" Then
                pResultInfo.KITCD = Trim(Mid(Trim(tmpField(4)), 2, Len(Trim(tmpField(4))) - 2))
            End If
            
            If Trim(tmpField(5)) <> "" Then
                pResultInfo.KITNM = Trim(Mid(Trim(tmpField(5)), 2, Len(Trim(tmpField(5))) - 2))
            End If
            
            If Trim(tmpField(6)) <> "" Then
                pResultInfo.RSTDT = Trim(tmpField(6))
            End If
            
            If Trim(tmpField(11)) <> "" Then
                pResultInfo.ORGCD = Trim(Mid(Trim(tmpField(11)), 2, Len(Trim(tmpField(11))) - 2))
            End If
            
            If Trim(tmpField(12)) <> "" Then
                pResultInfo.ORGNM = Trim(Mid(Trim(tmpField(12)), 2, Len(Trim(tmpField(12))) - 2))
            End If
            
            'Isolate Status
            'F:Finalized, P:Preliminary
            If Trim(tmpField(37)) <> "" Then
                pResultInfo.STATUS = Trim(Mid(Trim(tmpField(37)), 1, 1))
            End If
            
        Case "M"        'Result Record
            tmpField() = Split(RcvBuffer, ",")
            
            If Trim(tmpField(25)) <> "" Then
                tmpRstFlag = Mid(Trim(tmpField(25)), 1, 1)
            End If
        
            If pResultInfo.URINE = "Y" Then   'Urine
                If Trim(tmpField(8)) <> "" Then
                    tmpSRI = Mid(Trim(tmpField(8)), 2, Len(Trim(tmpField(8))) - 2)
                Else
                    tmpSRI = ""
                End If
            Else
                If Trim(tmpField(7)) <> "" Then
                    tmpSRI = Mid(Trim(tmpField(7)), 2, Len(Trim(tmpField(7))) - 2)
                Else
                    tmpSRI = ""
                End If
            End If
            
            If tmpRstFlag = "N" And tmpSRI <> "" Then
                If Trim(tmpField(2)) <> "" Then
                    TmpAntiCd = Mid(Trim(tmpField(2)), 2, Len(Trim(tmpField(2))) - 2)
                End If
                
                If Trim(tmpField(3)) <> "" Then
                    TmpAntiNm = Mid(Trim(tmpField(3)), 2, Len(Trim(tmpField(3))) - 2)
                End If
                
                If Trim(tmpField(4)) <> "" Then
                    tmpMIC = Mid(Trim(tmpField(4)), 2, Len(Trim(tmpField(4))) - 2)
                End If
                
                If TmpAntiCd <> "" Then
                    pResultInfo.ANTIRST = pResultInfo.ANTIRST + 1
                    pResultInfo.ANTICD = pResultInfo.ANTICD & TmpAntiCd & "|"
                    pResultInfo.ANTINM = pResultInfo.ANTINM & TmpAntiNm & "|"
                    pResultInfo.SRI = pResultInfo.SRI & tmpSRI & "|"
                    pResultInfo.MIC = pResultInfo.MIC & tmpMIC & "|"
                End If
            End If

            tmpAntiSeq = Replace(Trim(tmpField(1)), """", "")
            
            If IsNumeric(tmpAntiSeq) = False Then   '마지막 항생제 SEQ
                If pResultInfo.ORGCD <> "" Then
                    With pResultInfo
                        RaiseEvent AppendData(.ID, .ISOL, .SPCCD, .SPCNM, .KITCD, .KITNM, .ORGCD, .ORGNM, CStr(.ANTIRST), .ANTICD, .ANTINM, .SRI, .MIC, .URINE, .STATUS, .RSTDT)
                    End With
                End If
                
                Call Init_pResultInfo_ResultInfo
            End If
       
        Case "L"
            If pResultInfo.ORGCD <> "" Then
                With pResultInfo
                    RaiseEvent AppendData(.ID, .ISOL, .SPCCD, .SPCNM, .KITCD, .KITNM, .ORGCD, .ORGNM, CStr(.ANTIRST), .ANTICD, .ANTINM, .SRI, .MIC, .URINE, .STATUS, .RSTDT)
                End With
            End If
            
            Call Init_pResultInfo_ResultInfo
            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
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
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
        
        .ISOL = ""
        .SPCCD = ""
        .SPCNM = ""
        .KITCD = ""
        .KITNM = ""
        .ORGCD = ""
        .ORGNM = ""
        .ANTIRST = 0
        .ANTICD = ""
        .ANTINM = ""
        .SRI = ""
        .MIC = ""
        .URINE = ""
        .STATUS = ""
        
        .TISOL = ""
        .TKITCD = ""
        .TKITNM = ""
        .TORGCD = ""
        .TORGNM = ""
        .TANTIRST = ""
        .TANTICD = ""
        .TANTINM = ""
        .TSRI = ""
        .TMIC = ""
        .TSTATUS = ""
    End With
    
    pSampleInfo.CMT1 = ""
    
End Sub

'
'   결과정보 구조체 초기화
'
Private Sub Init_pResultInfo_ResultInfo()
    
    With pResultInfo
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
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
        
        .ISOL = ""
        .KITCD = ""
        .KITNM = ""
        .ORGCD = ""
        .ORGNM = ""
        .ANTIRST = 0
        .ANTICD = ""
        .ANTINM = ""
        .SRI = ""
        .MIC = ""
        .STATUS = ""
        
        .TISOL = ""
        .TKITCD = ""
        .TKITNM = ""
        .TORGCD = ""
        .TORGNM = ""
        .TANTIRST = ""
        .TANTICD = ""
        .TANTINM = ""
        .TSRI = ""
        .TMIC = ""
        .TSTATUS = ""
    End With
    
    pSampleInfo.CMT1 = ""
    
End Sub

'
'   환자 Order 전송
'
Public Sub SendOrder_Microscan()
    On Error GoTo Err_Rtn

    Dim i           As Integer
    Dim sSendBuff   As String
    Dim ChkSum      As String
    Dim sPatNo      As String
    Dim sPatNm      As String
    Dim sFirstNm    As String
    Dim sLastNm     As String
    Dim sSex        As String
    Dim sDept       As String
    Dim sWard       As String
    Dim sRoom       As String
    Dim sPKind      As String
    Dim sPanel      As String
    Dim sMKind      As String
    Dim sManual     As String
    Dim sJDate      As String
    
    Dim ix1         As Integer
    Dim ix2         As Integer
    Dim ix3         As Integer
    Dim bPnlMach    As Boolean
    
    Select Case m_iSendPhase
        Case 1
            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, "", "")
            Call Init_pResultInfo

            RaiseEvent RequestCurOrder("", "", "", "")
            
            Call Get_OrderString
                        
            If pSampleInfo.ORDCNT > 0 Then
                sDept = pSampleInfo.RACK
                sWard = pSampleInfo.POS
                
                If InStr(pSampleInfo.CMT1, "^") > 0 Then
                    sPatNo = Trim(Split(pSampleInfo.CMT1, "^")(0))
                    sFirstNm = Trim(Split(pSampleInfo.CMT1, "^")(1))
                    sLastNm = Trim(Split(pSampleInfo.CMT1, "^")(2))
                    sSex = Trim(Split(pSampleInfo.CMT1, "^")(3))
                     
                    sSendBuff = "P,""L"",""" & sPatNo & """,""" & sLastNm & """,""" & sFirstNm & """,," & sSex & ",,,,,,," & sWard & ",," & sDept & ",,,,,,,,,"
                Else
                    sSendBuff = "P,""L"",""" & pSampleInfo.CMT1 & """,,""" & "" & """,," & sSex & ",,,,,,," & sWard & ",," & sDept & ",,,,,,,,,"
                End If
        
                m_iSendPhase = 2
            Else
                sSendBuff = "L,""L"",N,"
                m_iSendPhase = 4
            End If
        
        Case 2
            sWard = pSampleInfo.POS
            
            If InStr(pSampleInfo.CMT1, "^") > 0 Then
                sPatNo = Trim(Split(pSampleInfo.CMT1, "^")(0))
                
                sSendBuff = "B,""L"",""" & pSampleInfo.ID & """,""" & sPatNo & """,,,""" & pSampleInfo.SPCCD & """,,," & Format(Now, "yyyyMMdd") & ",,,,,,,,,,""" & sWard & """,,,,"
            Else
                sSendBuff = "B,""L"",""" & pSampleInfo.ID & """,""" & pSampleInfo.CMT1 & """,,,""" & pSampleInfo.SPCCD & """,,," & Format(Now, "yyyyMMdd") & ",,,,,,,,,,""" & sWard & """,,,,"
            End If
            
                      
            m_iSendPhase = 3
            
        Case 3
            sJDate = Format(Now, "yyyyMMdd")
            
            sPanel = Trim(Split(pSampleInfo.CONTAINER, "^")(0))
            sManual = Trim(Split(pSampleInfo.CONTAINER, "^")(1))
            
''            N   NBC44   OXI-          -
''            N   NBC44   OXI+          +
''            P   PBC28   Strep/beta-   -
''            P   PBC28   Strep/beta+   +
''            P   PBC28   Staph         1
            
            '<Neg, Pos Panel 설정
            For ix1 = 1 To UBound(mManual)
                If sPanel = mManual(ix1).Panel And sManual = mManual(ix1).Manual Then
                    bPnlMach = True
                    
                    sPKind = ConvertPanelInfo(1, sPanel)
                    ''sMKind = ConvertManualInfo(1, sManual)
                    
                    Select Case sPKind
                        Case "N"
                            Select Case mManual(ix1).MKind
                                Case "-"
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,N,,,,,,,,,,,,,,,,,,,,,,,,"
                                   
                                Case "+"
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,P,,,,,,,,,,,,,,,,,,,,,,,,"
                                    
                                Case Else
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
                                    
                            End Select
                            
                            Exit For
                            
                        Case "P"
                            Select Case mManual(ix1).MKind
                                Case "-"    'Streptococcaceae -
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,,,,,,,N,,,,,,,,,,,,1,,,,,,"
                                    
                                Case "+"    'Streptococcaceae +
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,,,,,,,Y,,,,,,,,,,,,1,,,,,,"
                                
                                Case "2"    'Staph/Related Genera
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,,,,,,,,,,,,,,,,,,,2,,,,,,"
                                
                                Case Else
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
    
                            End Select
                            
                            Exit For
                            
                        Case "R"
                            Select Case mManual(ix1).MKind
                                Case "1", "2", "3", "4"    'Gran Neg Bacilli, Gran Pos Bacilli, Cocci, Clostridia
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,N,,,,,,,,,,,,,,,,N,,,,,0 ,,N,," & mManual(ix1).MKind & ",,,0,P,,"
                                Case Else
                                    sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
    
                            End Select
                            
                            Exit For
                       
                        Case Else
                            sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
                           
                    End Select
                End If
            Next ix1
            '>
            
            If bPnlMach = False Then
                sSendBuff = "R,""L"",""" & pSampleInfo.SEQNO & """,""" & pSampleInfo.ID & """,""" & sPanel & """,," & sJDate & ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,"
            End If
            
            m_iSendPhase = 1
            
        Case 4
            m_iPhase = 1
            m_iSendPhase = 1
            
            msComm.Output = Chr(4)

            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If
            
            Exit Sub
    End Select

''    ChkSum = ChkSum_ASTM(sSendBuff)
''    sSendBuff = sSendBuff & ChkSum
    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10) & Chr(3)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10) & Chr(3))
    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

Private Function ConvertPanelInfo(ByVal iMode As Integer, ByVal sComp As String) As String
    Dim i       As Integer
    Dim sReturn As String
    
    Select Case iMode
        Case 1  'Panel Code -> Panel Kind
            For i = 1 To UBound(mPanel)
                If mPanel(i).Panel = sComp Then
                    sReturn = mPanel(i).PKind
                    Exit For
                End If
            Next
        
    End Select
    
    ConvertPanelInfo = sReturn
End Function

Private Function ConvertManualInfo(ByVal iMode As Integer, ByVal sComp As String) As String
    Dim i       As Integer
    Dim sReturn As String
    
    Select Case iMode
        Case 1  'Manual Name -> Manual Kind
            For i = 1 To UBound(mManual)
                If mManual(i).Manual = sComp Then
                    sReturn = mManual(i).MKind
                    Exit For
                End If
            Next
        
    End Select
    
    ConvertManualInfo = sReturn
End Function

Public Sub SetPanelInfo(ByVal Panel As Variant)
    Dim i As Integer
    Dim aPanel() As String
    
    aPanel = Panel
    ReDim mPanel(UBound(aPanel))
    
    For i = 1 To UBound(aPanel)
        mPanel(i).PKind = Split(aPanel(i), "^")(0)
        mPanel(i).Panel = Split(aPanel(i), "^")(1)
        mPanel(i).Key = Split(aPanel(i), "^")(2)
    Next
         
End Sub

Public Sub SetManualInfo(ByVal Manual As Variant)
    Dim i As Integer
    Dim aManaul() As String
    
    aManaul = Manual
    ReDim mManual(UBound(aManaul))
    
    For i = 1 To UBound(aManaul)
        mManual(i).PKind = Split(aManaul(i), "^")(0)
        mManual(i).Panel = Split(aManaul(i), "^")(1)
        mManual(i).MKind = Split(aManaul(i), "^")(2)
        mManual(i).Manual = Split(aManaul(i), "^")(3)
        mManual(i).Key = Split(aManaul(i), "^")(4)
    Next
         
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
        .CONTAINER = m_p_sRerunGbn
        .SPCCD = m_p_sSpcCd
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

