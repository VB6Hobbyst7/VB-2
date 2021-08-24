VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl URINE 
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
Attribute VB_Name = "URINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_pType = 0
Const m_def_NoRstDiscard = True
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
Dim m_pType As Variant
Dim m_NoRstDiscard As Boolean
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$)
Event DispMsg(sMsg$)
Event RequestNextOrder()
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


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

'For MiditronM
Dim bRstFlag    As Boolean
Dim msSampleNo  As String       '단방향에서 장비번호 체크 사용
Dim msPreSampleNo   As String   '              "

'For Urisys2400
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim RstEnd  As String

Private Sub PhaseCfg_Protocol_UriScan()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2         ' STX
                RcvBuffer = ""
                
            Case 3         ' ETX
                '--- 결과편집/등록
                Call DataEdit_UriScan
                                            
                Do
                    DoEvents
                Loop Until msComm.OutBufferCount = 0
                
            Case Else      '
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1

End Sub
Private Sub PhaseCfg_Protocol_Urometer()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 13       ' CR
                '--- 결과편집/등록
                Call DataEdit_Urometer
                
                RcvBuffer = ""
                                            
                Do
                    DoEvents
                Loop Until msComm.OutBufferCount = 0
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1

End Sub

Private Sub PhaseCfg_Protocol_Urometer120()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 10       ' LF
            Case 13       ' CR
                '--- 결과편집/등록
                Call DataEdit_Urometer120
                
                RcvBuffer = ""
                                            
                Do
                    DoEvents
                Loop Until msComm.OutBufferCount = 0
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1

End Sub
Private Sub DataEdit_Urometer()
    On Error GoTo ErrRtn

    Dim ii      As Integer
    Dim tmpID   As String
    Dim tmpSeqNo    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer
    
    Dim aData() As String

    Dim RecType$   'Record Type

    RecType = Mid$(RcvBuffer, 1, 3)

    Select Case RecType
        Case "ID "
            '결과정보 구조체 초기화
            Call Init_pResultInfo
            
            ''tmpSeqNo = Trim(Mid(RcvBuffer, 8, 4))
            
            aData = Split(RcvBuffer, ":")
            RcvBuffer = Replace(aData(1), Space(1), "")
            
            tmpSeqNo = Trim(Mid(RcvBuffer, 1, 4))
            tmpID = Trim(Mid(RcvBuffer, 6))
            tmpID = Replace(tmpID, vbCr, "")
            tmpID = Replace(tmpID, vbLf, "")
                        
            pResultInfo.ID = tmpID
            pResultInfo.SEQNO = tmpSeqNo

        Case "BLD", "BIL", "URO", "KET", "PRO", "NIT", "GLU", "LEU"

            sRst1 = Trim(Mid(RcvBuffer, 7, 5))
            If UCase(sRst1) = "NEG" Then sRst1 = "-"

            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & RecType & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With

            If RecType = "LEU" Then
                '--- MICRO 자동입력
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 3
                    .IFCD = .IFCD & "WBC|RBC|EPI|"
                    .RST1 = .RST1 & "0-1|0-1|0-1|"
                    .RST2 = .RST2 & "|||"
                    .UNIT = .UNIT & "|||"
                    .FLAG = .FLAG & "|||"
                End With

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    ''.SEQNO = tmpSeqNo

                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "", "")
                    End If
                End With

            End If

        Case "p.H", "S.G"

            sRst1 = Trim(Mid(RcvBuffer, 12, 7))

            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & RecType & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With

        Case Else
    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_Urometer120()
    On Error GoTo ErrRtn

    Dim ii      As Integer
    Dim tmpID   As String
    Dim tmpSeqNo    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer
    
    Dim aData() As String

    Dim RecType$   'Record Type
    
''~MAR/29/2011 11:50:57
''Name:             Sex:
''Ward:             Age:
''SD (10)      (1672)
''BLD       +    10RBC/ul
''BIL -neg
''URO    norm    0.1mg/dl
''KET -neg
''PRO -neg
''NIT     pos    0.1mg/dl
''GLU -neg
''pH 6#
''S.G             <1.005
''LEU     +/-    10WBC/ul
''(SN=12071334)
''ID(             )
''OP(00000000) LOT(000000) ~


    RecType = Trim(Mid$(RcvBuffer, 1, 3))

    Select Case RecType
        Case "ID("
            '결과정보 구조체 초기화
            Call Init_pResultInfo
            
            ''tmpSeqNo = Trim(Mid(RcvBuffer, 8, 4))
            
            aData = Split(RcvBuffer, "(")
            RcvBuffer = Replace(aData(1), Space(1), "")
            
            tmpID = Trim(Mid(aData(1), 1, InStr(aData(1), ")") - 1))
                                   
            pResultInfo.ID = tmpID

        Case "BLD", "BIL", "URO", "KET", "PRO", "NIT", "GLU", "LEU"

            sRst1 = Trim(Mid(RcvBuffer, 7, 5))
            If UCase(sRst1) = "NEG" Then sRst1 = "-"

            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & RecType & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With

            If RecType = "LEU" Then
                '--- MICRO 자동입력
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 3
                    .IFCD = .IFCD & "WBC|RBC|EPI|"
                    .RST1 = .RST1 & "0-1|0-1|0-1|"
                    .RST2 = .RST2 & "|||"
                    .UNIT = .UNIT & "|||"
                    .FLAG = .FLAG & "|||"
                End With

                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    ''.SEQNO = tmpSeqNo

                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "", "")
                    End If
                End With

            End If
            
        Case "pH", "S.G"
            sRst1 = Trim(Mid(RcvBuffer, 16, 7))

            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & RecType & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
        Case Else
    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_US3100R()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd$, tmpRack$, tmpPos$
    Dim tmpSeqNo$, tmpKind$, tmpDate$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    Dim aRow()  As String
    Dim aData() As String
    Dim ii%, iChk%
    
    iChk = 0
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    If InStr(RcvBuffer, ",") = 0 Then Exit Sub
    
    ' 0              1    2  3         4     5         6          7
    '402695080     ,0001,08,N00000001,0000, Strip  9L,2004/03/18,15:15:49,
    '8              9              10             11             12             13             14             15             16             17             18             19
    '0 normal      ,0  -          ,0  -          ,0  -          ,0  -          ,0  -          ,0 6.5         ,0  -          ,0  -          ,1             ,0 1.022       ,0 0000L YELLOW  01 -
    
    aRow() = Split(RcvBuffer, ",")
    
    tmpBarCd = Trim(aRow(0))
    tmpRack = Trim(aRow(1))
    tmpPos = Trim(aRow(2))
    
    tmpSeqNo = Trim(aRow(3))
    tmpKind = Trim(aRow(5))
    tmpDate = Trim(aRow(6) & " " & aRow(7))
    
    For ii = 8 To UBound(aRow())
        If Trim(aRow(ii)) = "" Then
            Exit For
        End If

        If ii = 17 Then
            iChk = 1
        End If
            
        If Left(aRow(ii), 1) = "0" Then
            tmpIFCd = Trim(ii - 7 - iChk)
        
            If ii = 19 Then     'Color,Turbi
                '---Color
                tmpIFCd = 11
                tmpRst1 = Trim(Mid(aRow(ii), 7, 10))
                tmpRst2 = ""
                
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst1 & Chr(124)
                    .RST2 = .RST2 & tmpRst2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
                
                '---Turbi
                tmpIFCd = 12
                tmpRst1 = Trim(Mid(aRow(ii), 19, 2))
                tmpRst2 = ""
                
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst1 & Chr(124)
                    .RST2 = .RST2 & tmpRst2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            Else
                tmpRst1 = Trim(Mid(aRow(ii), 3))
            
                Erase aData()
                aData() = Split(tmpRst1, Space(1))
                
                tmpRst1 = Trim(aData(0))
                If UBound(aData()) >= 1 Then
                    tmpRst2 = Trim(aData(UBound(aData())))
                Else
                    tmpRst2 = ""
                End If
            
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst1 & Chr(124)
                    .RST2 = .RST2 & tmpRst2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            End If
        End If
    Next ii

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = tmpSeqNo
        .RACK = tmpRack
        .POS = tmpPos
        .KIND = tmpKind
        
        If m_NoRstDiscard = True Then
            If .RSTCNT > 0 Then
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, "", "")
            End If
        Else
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, "", "")
        End If
    End With
        
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Sub DataEdit_US2100R()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd    As String
    Dim tmpSeqNo$, tmpKind$, tmpDate$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    Dim aRow()  As String
    Dim aData() As String
    Dim ii%
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    If InStr(RcvBuffer, ",") = 0 Then Exit Sub
    
    ' 0       1 2        3        4        5              6              7              8              9              10             11             12             13             14
    'N000002,0,Strip 10,12/20/99,14:41:58,0 normal      ,0  -          ,0  -          ,0  -          ,0  -          ,0  -          ,0 6.5         ,0  -          ,0*2+    75    ,0 1.010       
    
    aRow() = Split(RcvBuffer, ",")
    
    tmpBarCd = Trim(aRow(0))
    tmpSeqNo = Trim(aRow(3))
    tmpKind = Trim(aRow(5))
    tmpDate = Trim(aRow(6) & " " & aRow(7))
    
    For ii = 8 To UBound(aRow())
        If Trim(aRow(ii)) = "" Then
            Exit For
        End If
        
        If Left(aRow(ii), 1) = "0" Then
            tmpIFCd = Trim(ii - 7)
            tmpRst1 = Trim(Mid(aRow(ii), 3))
        
            Erase aData()
            aData() = Split(tmpRst1, Space(1))
            
            tmpRst1 = Trim(aData(0))
            If UBound(aData()) >= 1 Then
                tmpRst2 = Trim(aData(UBound(aData())))
            Else
                tmpRst2 = ""
            End If
                
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst1 & Chr(124)
                .RST2 = .RST2 & tmpRst2 & Chr(124)
                .UNIT = .UNIT & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
        ElseIf Left(aRow(ii), 1) = "1" Then
            '결과 Error인 경우도 등록처리...2006/9/1 yk
            If m_NoRstDiscard = False Then
                tmpIFCd = Trim(ii - 7)
                tmpRst1 = ""
                tmpRst2 = ""
                                    
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst1 & Chr(124)
                    .RST2 = .RST2 & tmpRst2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            End If
        End If
    Next ii

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = tmpSeqNo
        .KIND = tmpKind
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, "", "")
        End If
    End With
        
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_UriScan()
    On Error GoTo ErrRtn
    
    Dim ii      As Integer
    Dim tmpSeqNo    As String
    Dim tmpBarCd    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer

    Dim aRowData()  As String
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    '설정된 장비코드 편집
    m_p_sTIFCd = "BLD|BIL|URO|KET|PRO|NIT|GLU|p.H|S.G|LEU|VTC|"
    ReDim sTestCd(10) As String
    sTestCd() = Split(m_p_sTIFCd, Chr(124))
    
    iPos = InStr(RcvBuffer, "ID_NO:")
    If iPos <> 0 Then
        If InStr(Mid(RcvBuffer, iPos), Chr(13)) > 0 Then
            aRowData() = Split(Mid(RcvBuffer, iPos + 6), Chr(13))
                
            tmpSeqNo = Trim(Mid(aRowData(0), 1, 4))
            tmpBarCd = Trim(Mid(aRowData(0), 6))
        Else
            tmpSeqNo = Trim(Mid(RcvBuffer, iPos + 6, 4))
        End If
    End If
    
    '--- 결과값 편집
    For ii = 1 To 11
        sIFCd = Trim(sTestCd(ii - 1))
                
        iPos = InStr(RcvBuffer, sIFCd)
        
        If iPos = 0 Then
        Else
            sRstData = Mid$(RcvBuffer, iPos, 15)
            'sRst1 = Trim$(Mid$(sRstData, 11, 5))
            sRst1 = Trim$(Mid$(sRstData, 4, 7))
            sRst1 = sRst1 & "^" & Trim$(Mid$(sRstData, 11, 5))
                        
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & sIFCd & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
        End If
    Next ii

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = tmpSeqNo
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "", "")
        End If
    End With
        
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_UrineQuick()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd    As String
    Dim tmpSeqNo$, tmpKind$, tmpDate$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    Dim aRow()  As String
    Dim aData() As String
    Dim ii%
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    If InStr(RcvBuffer, ",") = 0 Then Exit Sub
    
    ' 0       1 2        3        4        5              6              7              8              9              10             11             12             13             14
    'N000002,0,Strip 10,12/20/99,14:41:58,0 normal      ,0  -          ,0  -          ,0  -          ,0  -          ,0  -          ,0 6.5         ,0  -          ,0*2+    75    ,0 1.010       
    
    aRow() = Split(RcvBuffer, ",")
    
    tmpBarCd = Trim(aRow(0))
    tmpSeqNo = Trim(aRow(3))
    tmpKind = Trim(aRow(5))
    tmpDate = Trim(aRow(6) & " " & aRow(7))
    
    For ii = 8 To UBound(aRow())
        If Trim(aRow(ii)) = "" Then
            Exit For
        End If
        
        If Left(aRow(ii), 1) = "0" Then
            tmpIFCd = Trim(ii - 7)
            tmpRst1 = Trim(Mid(aRow(ii), 3))
        
            Erase aData()
            aData() = Split(tmpRst1, Space(1))
            
            tmpRst1 = Trim(aData(0))
            If UBound(aData()) >= 1 Then
                tmpRst2 = Trim(aData(UBound(aData())))
            Else
                tmpRst2 = ""
            End If
                
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst1 & Chr(124)
                .RST2 = .RST2 & tmpRst2 & Chr(124)
                .UNIT = .UNIT & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
        ElseIf Left(aRow(ii), 1) = "1" Then
            '결과 Error인 경우도 등록처리...2006/9/1 yk
            If m_NoRstDiscard = False Then
                tmpIFCd = Trim(ii - 7)
                tmpRst1 = ""
                tmpRst2 = ""
                                    
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst1 & Chr(124)
                    .RST2 = .RST2 & tmpRst2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            End If
        End If
    Next ii

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = tmpSeqNo
        .KIND = tmpKind
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, "", "")
        End If
    End With
        
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
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
        Case "CLINITEK500"
            Call PhaseCfg_Protocol_Clinitek500
        
        Case "CLINITEK500_CCS"
            Call PhaseCfg_Protocol_Clinitek500_CCS
            
        Case "MIDITRONM"
            Call PhaseCfg_Protocol_MiditronM
        
        Case "MIDITRONJR"
            Call PhaseCfg_Protocol_MiditronJr
        
        Case "URISCAN"
            Call PhaseCfg_Protocol_UriScan
            
        Case "UROMETER"
            Call PhaseCfg_Protocol_Urometer
            
        Case "UROMETER120"
            Call PhaseCfg_Protocol_Urometer120
            
        Case "US2100R", "US3100R"
            Call PhaseCfg_Protocol_US2100R
            
        Case "URINEQUICK"
            Call PhaseCfg_Protocol_UrineQuick
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_Clinitek500()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2         ' STX
                RcvBuffer = ""
                
            Case 3         ' ETX
                '--- 결과편집/등록
                Call DataEdit_Clinitek500
                                            
            Case 21        ' NAK
                
            Case Else      '
                RcvBuffer = RcvBuffer + wkDat
                
         End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_Clinitek500_CCS()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5  'ENQ
                        m_iPhase = 2
                        
                        msComm.Output = Chr(6)
                    Case Else
                    
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                        
                    Case 10     'LF
                        Call DataEdit_Clinitek500_CCS
                        RcvBuffer = ""
                        
                        msComm.Output = Chr(6)
                        
                    Case 3      'ETX
                        msComm.Output = Chr(6)
                    
                    Case 4      'EOT
                        RcvBuffer = ""
                        m_iPhase = 1
                    
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                    
                    Case 23     'ETB
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        
                End Select
        End Select
    Next ix1

End Sub

Private Sub DataEdit_MiditronJr()
    On Error GoTo ErrRtn
    
    Dim ii      As Integer
    Dim tmpSeqNo    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer
    Dim iPos1   As Integer
    Dim iPos2   As Integer
    Dim sTmp    As String
    Dim iTmp    As Integer
    
    
    If Left$(RcvBuffer, 1) = "<" Then
        bRstFlag = True
        RcvBuffer = ""
        Exit Sub
    ElseIf Left$(RcvBuffer, 1) = ":" Then
        bRstFlag = False
        Exit Sub
    End If
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    '설정된 장비코드 편집
    m_p_sTIFCd = "SG|PH|LEU|NIT|PRO|GLU|KET|UBG|BIL|ERY|"
    ReDim sTestCd(m_p_iOrdCnt) As String
    sTestCd() = Split(m_p_sTIFCd, Chr(124))
    
    tmpSeqNo = Trim(Mid(RcvBuffer, 16, 4))
    
    '--- 같은 환자 여러번 넘어올 때의 처리
    '검사시간에 해당
    msSampleNo = Mid(RcvBuffer, 16, 19)
    
    '이전과 검사시간이 같은지 체크
    If msSampleNo = msPreSampleNo Then Exit Sub
    '-------------------------------------
    
    '--- 결과값 편집
    iPos = 1
    For ii = 1 To 10
        sIFCd = Trim(sTestCd(ii - 1))
        If Trim(sIFCd) <> "" Then
            '설정된 장비코드에 해당하는 결과값 조회
            If ii = 10 Then
                iPos1 = InStr(iPos, RcvBuffer, Trim(sTestCd(ii - 1)))
                iPos2 = InStr(iPos1 + 1, RcvBuffer, "NAG")
            Else
                iPos1 = InStr(iPos, RcvBuffer, Trim(sTestCd(ii - 1)))
                iPos2 = InStr(iPos1 + 1, RcvBuffer, Trim(sTestCd(ii)))
            End If
            
            If iPos1 = 0 And iPos2 = 0 Then
                Exit For
            End If
            
            sRstData = Mid(RcvBuffer, iPos1 + Len(sIFCd), iPos2 - iPos1 - Len(sIFCd))
            
            iTmp = InStr(Trim(sRstData), Space(1))
            If iTmp = 0 Then
                sRst1 = Trim(sRstData)
                sRst2 = ""
                sUnit = ""
            Else
                sRst1 = Trim(Mid(Trim(sRstData), 1, iTmp - 1))
                sRst2 = Trim(Mid(Trim(sRstData), iTmp + 1))
                
                iTmp = InStr(Trim(sRst2), Space(1))
                If iTmp = 0 Then
                    iTmp = InStr(Trim(sRst2), "/")
                    If iTmp = 0 Then
                        sUnit = ""
                    Else
                        sUnit = Trim(sRst2)
                        sRst2 = ""
                    End If
                Else
                    '단위가 포함된 경우 편집
                    sUnit = Trim(Mid(sRst2, 1, iTmp - 1))
                    sRst2 = Trim(Mid(sRst2, iTmp + 1))
                End If
            End If
            
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & sIFCd & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
            
            iPos = iPos1
        End If
    Next ii
    
    msPreSampleNo = msSampleNo
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .SEQNO = tmpSeqNo
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "", "")
        End If
    End With
                                            
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Sub DataEdit_MiditronM()
    On Error GoTo ErrRtn
    
    Dim ii      As Integer
    Dim tmpSeqNo    As String
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst1   As String
    Dim sRst2   As String
    Dim sUnit   As String
    Dim iPos    As Integer
    Dim iPos1   As Integer
    Dim iPos2   As Integer
    Dim sTmp    As String
    Dim iTmp    As Integer
    
    
    If Left$(RcvBuffer, 1) = "<" Then
        bRstFlag = True
        RcvBuffer = ""
        Exit Sub
    ElseIf Left$(RcvBuffer, 1) = ":" Then
        bRstFlag = False
        Exit Sub
    End If
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    '설정된 장비코드 편집
    m_p_sTIFCd = "SG|PH|LEU|NIT|PRO|GLU|KET|UBG|BIL|ERY|"
    ReDim sTestCd(m_p_iOrdCnt) As String
    sTestCd() = Split(m_p_sTIFCd, Chr(124))
    
    tmpSeqNo = Trim(Mid(RcvBuffer, 16, 4))
    
    '--- 같은 환자 여러번 넘어올 때의 처리
    '검사시간에 해당
    msSampleNo = Mid(RcvBuffer, 16, 19)
    
    '이전과 검사시간이 같은지 체크
    If msSampleNo = msPreSampleNo Then Exit Sub
    '-------------------------------------
    
    '--- 결과값 편집
    iPos = 1
    For ii = 1 To 10    'm_p_iOrdCnt
        sIFCd = Trim(sTestCd(ii - 1))
        If Trim(sIFCd) <> "" Then
            '설정된 장비코드에 해당하는 결과값 조회
            If ii = 10 Then
                iPos1 = InStr(iPos, RcvBuffer, Trim(sTestCd(ii - 1)))
                iPos2 = InStr(iPos1 + 1, RcvBuffer, "NAG")
            Else
                iPos1 = InStr(iPos, RcvBuffer, Trim(sTestCd(ii - 1)))
                iPos2 = InStr(iPos1 + 1, RcvBuffer, Trim(sTestCd(ii)))
            End If
            
            If iPos1 = 0 And iPos2 = 0 Then
                Exit For
            End If
            
            sRstData = Mid(RcvBuffer, iPos1 + Len(sIFCd), iPos2 - iPos1 - Len(sIFCd))
            
            iTmp = InStr(Trim(sRstData), Space(1))
            If iTmp = 0 Then
                sRst1 = Trim(sRstData)
                sRst2 = ""
                sUnit = ""
            Else
                sRst1 = Trim(Mid(Trim(sRstData), 1, iTmp - 1))
                sRst2 = Trim(Mid(Trim(sRstData), iTmp + 1))
                
                iTmp = InStr(Trim(sRst2), Space(1))
                If iTmp = 0 Then
                    iTmp = InStr(Trim(sRst2), "/")
                    If iTmp = 0 Then
                        sUnit = ""
                    Else
                        sUnit = Trim(sRst2)
                        sRst2 = ""
                    End If
                Else
                    '단위가 포함된 경우 편집
                    sUnit = Trim(Mid(sRst2, 1, iTmp - 1))
                    sRst2 = Trim(Mid(sRst2, iTmp + 1))
                End If
            End If
            
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & sIFCd & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & sRst2 & Chr(124)
                .UNIT = .UNIT & sUnit & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
            
            iPos = iPos1
        End If
    Next ii
    
    msPreSampleNo = msSampleNo
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .SEQNO = tmpSeqNo
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "", "")
        End If
    End With
                                            
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_Clinitek500_Old()
'    On Error GoTo ErrRtn
'
'    Dim iPos    As Integer
'    Dim ii      As Integer
'    Dim tmpSeqNo$
'    Dim sTestCd()   As String
'    Dim sRstData    As String
'    Dim sIFCd   As String
'    Dim sRst    As String
'    Dim iTmp    As Integer
'
'
'    '결과정보 구조체 초기화
'    Call Init_pResultInfo
'
'    '설정된 장비코드 편집
'    ReDim sTestCd(m_p_iOrdCnt) As String
'    sTestCd() = Split(m_p_sTIFCd, Chr(124))
'
'
'    '--- Data 편집
'    'Chr(13) & chr(10)을 절삭
'    RcvBuffer = Mid$(RcvBuffer, 3)
'
'    'SerialNo, 날짜 절삭
'    tmpSeqNo = Trim(Mid(RcvBuffer, 2, 6))
'    iPos = InStr(RcvBuffer, Chr(10))
'    RcvBuffer = Mid$(RcvBuffer, iPos + 1)
'
''    'SampleNo 얻기...
''    iPos = InStr(RcvBuffer, Chr(10))
''    RcvBuffer = Mid$(RcvBuffer, iPos + 1)
'
''    'Color : 절삭...
''    cnt = InStr(RcvBuffer, Chr$(10))
''    RcvBuffer = Mid$(RcvBuffer, cnt + 1)
''
''    'Clarity : 절삭...
''    cnt = InStr(RcvBuffer, Chr$(10))
''    RcvBuffer = Mid$(RcvBuffer, cnt + 1)
'
'    pResultInfo.SEQNO = tmpSeqNo
'
'    Do While True
'        iPos = InStr(RcvBuffer, Chr(10))
'        If iPos = 0 Then
'            Exit Do
'        End If
'
'        sRstData = Left$(RcvBuffer, iPos - 2)
'        RcvBuffer = Mid$(RcvBuffer, iPos + 1)
'
'        '--- 결과값 편집
'        For ii = 1 To m_p_iOrdCnt
'            sIFCd = Trim(sTestCd(ii - 1))
'            If Trim(sIFCd) <> "" Then
'                '설정된 장비코드에 해당하는 결과값 조회
'                iPos = InStr(1, sRstData, Trim(sTestCd(ii - 1)))
'                If iPos <> 0 Then
'                    '결과값
'                    sRst = Trim(Mid(sRstData, Len(sIFCd) + 1))
'                    If Left(sRst, 1) = "*" Then
'                        sRst = Trim(Mid(sRst, 2))
'                    End If
'                    iTmp = InStr(1, sRst, Space(1))
'                    If iTmp > 0 Then
'                        sRst = Trim(Left(sRst, iTmp - 1))
'                    End If
'
'                    'Data 누적
'                    With pResultInfo
'                        .RSTCNT = .RSTCNT + 1
'                        .IFCD = .IFCD & sIFCd & Chr(124)
'                        .RST1 = .RST1 & sRst & Chr(124)
'                        .RST2 = .RST2 & Chr(124)
'                        .UNIT = .UNIT & Chr(124)
'                        .FLAG = .FLAG & Chr(124)
'                    End With
'
'                    Exit For
'                End If
'            End If
'        Next ii
'    Loop
'
'    '결과값 등록/화면 표시 처리...
'    With pResultInfo
'        If .RSTCNT > 0 Then
'            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
'        End If
'    End With
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
'    End If
End Sub

Private Sub DataEdit_Clinitek500()
    On Error GoTo ErrRtn
    
    Dim iPos    As Integer
    Dim ii      As Integer
    Dim tmpSeqNo$
    Dim sTestCd()   As String
    Dim sRstData    As String
    Dim sIFCd   As String
    Dim sRst    As String
    Dim iTmp    As Integer
    
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    '설정된 장비코드 편집
    If pType = "B" Then         '2007/8/8 yk
        m_p_sTIFCd = "SG|pH|LEU|NIT|PRO|GLU|KET|URO|BIL|BLO|"
    Else
        m_p_sTIFCd = "SG|pH|LEU|NIT|PRO|GLU|KET|UBG|BIL|BLD|"
    End If
    ReDim sTestCd(m_p_iOrdCnt) As String
    sTestCd() = Split(m_p_sTIFCd, Chr(124))
    
    
    '--- Data 편집
    'Chr(13) & chr(10)을 절삭
    RcvBuffer = Mid$(RcvBuffer, 3)
    
    'SerialNo, 날짜 절삭
    tmpSeqNo = Trim(Mid(RcvBuffer, 2, 6))
    iPos = InStr(RcvBuffer, Chr(10))
    RcvBuffer = Mid$(RcvBuffer, iPos + 1)
    
'    'SampleNo 얻기...
'    iPos = InStr(RcvBuffer, Chr(10))
'    RcvBuffer = Mid$(RcvBuffer, iPos + 1)

'    'Color : 절삭...
'    cnt = InStr(RcvBuffer, Chr$(10))
'    RcvBuffer = Mid$(RcvBuffer, cnt + 1)
'
'    'Clarity : 절삭...
'    cnt = InStr(RcvBuffer, Chr$(10))
'    RcvBuffer = Mid$(RcvBuffer, cnt + 1)
    
    pResultInfo.SEQNO = tmpSeqNo
    
    Do While True
        iPos = InStr(RcvBuffer, Chr(10))
        If iPos = 0 Then
            Exit Do
        End If
        
        sRstData = Left$(RcvBuffer, iPos - 2)
        RcvBuffer = Mid$(RcvBuffer, iPos + 1)

        '--- 결과값 편집
        For ii = 1 To UBound(sTestCd())     ' m_p_iOrdCnt
            sIFCd = Trim(sTestCd(ii - 1))
            If Trim(sIFCd) <> "" Then
                '설정된 장비코드에 해당하는 결과값 조회
                iPos = InStr(1, sRstData, Trim(sTestCd(ii - 1)))
                If iPos <> 0 Then
                    '결과값
                    sRst = Trim(Mid(sRstData, Len(sIFCd) + 1))
                    If Left(sRst, 1) = "*" Then
                        sRst = Trim(Mid(sRst, 2))
                    End If
                    iTmp = InStr(1, sRst, Space(1))
                    If iTmp > 0 Then
                        sRst = Trim(Left(sRst, iTmp - 1))
                    End If
                    
                    'Data 누적
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & sIFCd & Chr(124)
                        .RST1 = .RST1 & sRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                    
                    Exit For
                End If
            End If
        Next ii
    Loop
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "", "")
        End If
    End With
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEdit_Clinitek500_CCS()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim i       As Integer
    Dim tmpData()   As String
    Dim tmpField()  As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpKind$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$, tmpAlarmCd$
    Dim sRst1Tmp$

    RecType = Mid$(RcvBuffer, 2, 1)

    Select Case RecType
        Case "H"        'Header Record
        
        Case "P"        'Patient Record
            'Seq 번호를 OCX에서 저장하도록 수정, 상호 20100501
            tmpData = Split(RcvBuffer, vbCr)
    
            tmpField() = Split(tmpData(0), "|")
            tmpSeqNo = Trim(tmpField(2))
            
            
            tmpBarCd = Trim(tmpField(3))
            
                        
            'SeqNo 구조체에 저장
            pResultInfo.SEQNO = tmpSeqNo
            pResultInfo.ID = tmpBarCd
            
                    
        Case "R"        'Result Record
'            3R|1|N|GLU|1|NEGATIVE|1|0|A
'              R|2|N|BIL|2|NEGATIVE|1|0|A
'              R|3|N|KET|3|NEGATIVE|1|0|A
'              R|4|N|SG|4|>=1.030|6|0|A

            tmpData = Split(RcvBuffer, vbCr)
            
            For i = 0 To UBound(tmpData) - 2
                tmpField() = Split(tmpData(i), "|")
                tmpIFCd = Trim(tmpField(3))
                tmpRst = Trim(tmpField(5))
                
                sRst1Tmp = tmpRst
                
                If sRst1Tmp = "ERROR" Then
                    tmpRst = ""
                End If
                                
                If InStr(sRst1Tmp, "^") > 0 Then
                    tmpRst = Split(sRst1Tmp, "^")(0)
                    tmpUnit = Split(sRst1Tmp, "^")(1)
                End If
                            
                '결과정보 구조체에 저장
                With pResultInfo
                    '결과값 누적
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & tmpRef & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                    .FLAG = .FLAG & tmpFlag & Chr(124)
                End With
            Next
            
        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD, .KIND, "", "")
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
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
Private Sub PhaseCfg_Protocol_MiditronJr()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX
                RcvBuffer = ""
                
            Case 13     'CR
                '--- 결과편집/등록
                Call DataEdit_MiditronJr
                
                If bRstFlag = True Then
                    msComm.Output = Chr(2) & Chr(62) & Chr(3) & Chr(51) & Chr(63) & Chr(13)
                End If

            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_MiditronM()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX
                RcvBuffer = ""
                
            Case 13     'CR
                '--- 결과편집/등록
                Call DataEdit_MiditronM
                
                If bRstFlag = True Then
                    msComm.Output = Chr(2) & Chr(62) & Chr(3) & Chr(51) & Chr(63) & Chr(13)
                End If

            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
        End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_US2100R()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2          ' STX
                RcvBuffer = ""
                
            Case 3, 10      'ETX,LF
                '--- 결과편집/등록
                If UCase(m_EqName) = "US2100R" Then
                    Call DataEdit_US2100R
                ElseIf UCase(m_EqName) = "US3100R" Then
                    Call DataEdit_US3100R
                End If
                
                msComm.Output = Chr(6)
                
                RcvBuffer = ""
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
         End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_UrineQuick()
    
    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2          ' STX
                RcvBuffer = ""
                
            Case 3, 10      'ETX,LF
                '--- 결과편집/등록
                Call DataEdit_UrineQuick
                
                msComm.Output = Chr(6)
                
                RcvBuffer = ""
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
                
         End Select
    Next ix1
    
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

    m_NoRstDiscard = PropBag.ReadProperty("NoRstDiscard", m_def_NoRstDiscard)
    m_pType = PropBag.ReadProperty("pType", m_def_pType)
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

    Call PropBag.WriteProperty("NoRstDiscard", m_NoRstDiscard, m_def_NoRstDiscard)
    Call PropBag.WriteProperty("pType", m_pType, m_def_pType)
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
    
    '변수 초기화
    bRstFlag = False
    
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
    m_NoRstDiscard = m_def_NoRstDiscard
    m_pType = m_def_pType
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
'MemberInfo=0,0,0,True
Public Property Get NoRstDiscard() As Boolean
    NoRstDiscard = m_NoRstDiscard
End Property

Public Property Let NoRstDiscard(ByVal New_NoRstDiscard As Boolean)
    m_NoRstDiscard = New_NoRstDiscard
    PropertyChanged "NoRstDiscard"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get pType() As Variant
    pType = m_pType
End Property

Public Property Let pType(ByVal New_pType As Variant)
    m_pType = New_pType
    PropertyChanged "pType"
End Property

