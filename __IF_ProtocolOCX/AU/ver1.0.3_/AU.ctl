VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl AU 
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
   End
End
Attribute VB_Name = "AU"
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
Event SendOrderOK(sID$, sRack$, sPos$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)
Event RequestNextOrder()

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

'
'   AU-5400(���ڵ� ���)
'
Private Sub DataEditResponse_AU5400()
    On Error GoTo ErrRtn

    Dim sBC     As String
    Dim sLC     As String
    Dim iETBpos%, ii%, kk%
    Dim sTmpBuf1$, sTmpBuf2$, sTmp$
    Dim sSampNo$, sSampType$
    Dim tmpIFCd$, tmpRst$, tmpFlag$, tmpRstDT$
    Dim iPos%
    Dim sUnitNo$

    'Data�� Edit�ϱ� ���ϵ��� <ETB> ����
    Do
        iETBpos = InStr(1, RcvBuffer, Chr(23))

        If iETBpos = 0 Then Exit Do

        sTmpBuf1 = Left$(RcvBuffer, iETBpos - 1)
        sTmpBuf2 = Mid$(RcvBuffer, iETBpos + 18)
        RcvBuffer = ""
        RcvBuffer = sTmpBuf1 & sTmpBuf2
    Loop While iETBpos <> 0


    sBC = Mid$(RcvBuffer, 1, 1)
    sLC = Mid$(RcvBuffer, 2, 1)

    Select Case sBC
        Case "R"
            If sLC = "B" Or sLC = "E" Then Exit Sub     'B:Start of T.R.I inquiry, E:End of T.R.I inquiry

            With pSampleInfo
                .RACK = Mid(RcvBuffer, 3, 4)
                .POS = Mid(RcvBuffer, 7, 2)
                sSampType = Trim(Mid(RcvBuffer, 9, 1))  'SampleType(' ':Serum,'U':Urine, 'X':other)
                .SEQNO = Mid(RcvBuffer, 10, 4)
'                .ID = Trim(Mid(RcvBuffer, 21, 20))
                .ID = Trim(Mid(RcvBuffer, 14, 21))      '2007/5/30 yk
                iPos = InStr(.ID, Chr(3))   'BARCODE ����
                If iPos <> 0 Then
                    .ID = Mid(.ID, 1, iPos - 1)
                End If

                If UCase(Left(.ID, 3)) = "ERR" Then
                    RaiseEvent DispMsg(.RACK & "/" & .POS & " - BARCODE READ ERROR!!!")
                    Exit Sub
                End If
            End With

            'Order ����...
            Call SendOrder_AU5400

        Case "D"
            If sLC = "B" Or sLC = "E" Then Exit Sub

            '������� �ʱ�ȭ
            Call Init_pResultInfo

            'Sample ���� ����
            If sLC = "Q" Then
                'QC Result
                With pResultInfo
                    'DQ       Q001                      01E1035  13.7r 
                    'DQ       Q001                      01E11
                    .RACK = Trim(Mid$(RcvBuffer, 3, 4))
                    .POS = Trim(Mid$(RcvBuffer, 36, 2))
                    sUnitNo = Trim(Mid$(RcvBuffer, 39, 2))
                    .SEQNO = Mid$(RcvBuffer, 10, 4)         'SampleNo
                    .ID = .SEQNO & "-" & .POS & " Unit" & sUnitNo
                    .KIND = "Q"
                End With
    
                '�������
                For ii = 1 To 100   'm_iTotalItemCnt   '100
                    sTmp = Mid(RcvBuffer, 41 + 10 * (ii - 1), 1)
    
    '                If Asc(sTmp) = 3 Then Exit For
                    If Asc(sTmp) = 3 Or Trim(sTmp) = "" Then Exit For
                    
                    tmpIFCd = Format(Val(Mid(RcvBuffer, 41 + 10 * (ii - 1), 2)), "00")
                    tmpRst = Trim(Mid(RcvBuffer, 41 + 10 * (ii - 1) + 2, 6))
                    tmpFlag = Trim(Mid(RcvBuffer, 41 + 10 * (ii - 1) + 8, 2))
    
                    If Left(tmpRst, 1) = "." Then
                        tmpRst = "0" & tmpRst
                    End If
    
                    tmpRstDT = Format(Now, "YYYYMMDDHHNNSS")
    
                    '����� ����
                    If Trim(tmpIFCd) <> "" Then
                        With pResultInfo
                            .RSTCNT = .RSTCNT + 1
    
                            .IFCD = .IFCD & tmpIFCd & Chr(124)
                            .RST1 = .RST1 & tmpRst & Chr(124)
                            .RST2 = .RST2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & tmpFlag & Chr(124)
                            .RSTDT = .RSTDT & tmpRstDT & Chr(124)
                        End With
                    End If
                Next ii
            Else
                '�Ϲ� Result
                With pResultInfo
                    .RACK = Mid$(RcvBuffer, 3, 4)
                    .POS = Mid$(RcvBuffer, 7, 2)
                    sSampType = Trim(Mid(RcvBuffer, 9, 1))  'SampleType(' ':Serum,'U':Urine, 'X':other)
                    .SEQNO = Mid$(RcvBuffer, 10, 4)
    '                .ID = Trim(Mid$(RcvBuffer, 21, 20))
                    .ID = Trim(Mid$(RcvBuffer, 14, 21))     '2007/5/30 yk
                    'BARCODE ����
                    iPos = InStr(.ID, Chr(3))
                    If iPos <> 0 Then
                        .ID = Mid(.ID, 1, iPos - 1)
                    End If
                    iPos = InStr(.ID, " ")
                    If iPos <> 0 Then
                        .ID = Mid(.ID, 1, iPos - 1)
                    End If
    
                    Select Case Left(.SEQNO, 1)
                        Case "A": .KIND = "C"   'Calibration
                        Case "Q": .KIND = "Q"   'QC Sample
                        Case Else: .KIND = ""
                    End Select
    
                    If Trim(.ID) = "" Then Exit Sub
                End With
    
                '�������
                For ii = 1 To 100   'm_iTotalItemCnt   '100
                    sTmp = Mid(RcvBuffer, 39 + 10 * (ii - 1), 1)
    
    '                If Asc(sTmp) = 3 Then Exit For
                    If Asc(sTmp) = 3 Or Trim(sTmp) = "" Then Exit For
                    
                    tmpIFCd = Format(Val(Mid(RcvBuffer, 39 + 10 * (ii - 1), 2)), "00")
                    tmpRst = Trim(Mid(RcvBuffer, 39 + 10 * (ii - 1) + 2, 6))
                    tmpFlag = Trim(Mid(RcvBuffer, 39 + 10 * (ii - 1) + 8, 2))
    
                    If Left(tmpRst, 1) = "." Then
                        tmpRst = "0" & tmpRst
                    End If
    
                    tmpRstDT = Format(Now, "YYYYMMDDHHNNSS")
    
                    '����� ����
                    If Trim(tmpIFCd) <> "" Then
                        With pResultInfo
                            .RSTCNT = .RSTCNT + 1
    
                            .IFCD = .IFCD & tmpIFCd & Chr(124)
                            .RST1 = .RST1 & tmpRst & Chr(124)
                            .RST2 = .RST2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & tmpFlag & Chr(124)
                            .RSTDT = .RSTDT & tmpRstDT & Chr(124)
                        End With
                    End If
                Next ii
            End If
            
            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .KIND, .RST1, "")
                End If
            End With

        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit ���� - (" & Err.Description & ")")
    End If
End Sub

Private Sub PhaseCfg_Protocol_AU5400()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
                 
        Select Case m_iPhase
            Case 1      '===== STX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        RcvBuffer = ""
                End Select
                
            Case 2      '===== ETX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        
                    Case 3      '----- ETX ����
                        RcvBuffer = RcvBuffer & wkDat
                        
                        Call DataEditResponse_AU5400
                                                
                        m_iPhase = 1
                    
                    Case Else   '----- ���� ����
                        RcvBuffer = RcvBuffer & wkDat
                
                End Select
         End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_AU640()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
                 
        Select Case m_iPhase
            Case 1      '===== STX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        RcvBuffer = ""
                End Select
                
            Case 2      '===== ETX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        
                    Case 3      '----- ETX ����
                        RcvBuffer = RcvBuffer & wkDat
                        
                        Call DataEditResponse_AU640
                        
                        m_iPhase = 1
                    
                    Case Else   '----- ���� ����
                        RcvBuffer = RcvBuffer & wkDat
                
                End Select
         End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_AU680()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
                 
        Select Case m_iPhase
            Case 1      '===== STX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        RcvBuffer = ""
                End Select
                
            Case 2      '===== ETX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        
                    Case 3      '----- ETX ����
                        RcvBuffer = RcvBuffer & wkDat
                        
                        Call DataEditResponse_AU680
                        
                        m_iPhase = 1
                    
                    Case Else   '----- ���� ����
                        RcvBuffer = RcvBuffer & wkDat
                
                End Select
         End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_AU400()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
                 
        Select Case m_iPhase
            Case 1      '===== STX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        RcvBuffer = ""
                End Select
                
            Case 2      '===== ETX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        
                    Case 3      '----- ETX ����
                        RcvBuffer = RcvBuffer & wkDat
                        Call DataEditResponse_AU400
                
                        m_iPhase = 1
                    
                    Case Else   '----- ���� ����
                        RcvBuffer = RcvBuffer & wkDat
                
                End Select
         End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_AU600()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
                 
        Select Case m_iPhase
            Case 1      '===== STX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        RcvBuffer = ""
                End Select
                
            Case 2      '===== ETX ���
                Select Case Asc(wkDat)
                    Case 2      '----- STX ����
                        m_iPhase = 2
                        
                    Case 3      '----- ETX ����
                        RcvBuffer = RcvBuffer & wkDat
                        Call DataEditResponse_AU600
                
                        m_iPhase = 1
                    
                    Case Else   '----- ���� ����
                        RcvBuffer = RcvBuffer & wkDat
                
                End Select
         End Select
    Next ix1
    
End Sub
'
'   AU-640(���ڵ� ���)
'
Private Sub DataEditResponse_AU640()
    On Error GoTo ErrRtn

    Dim sBC     As String
    Dim sLC     As String
    Dim iETBpos%, ii%, kk%
    Dim sTmpBuf1$, sTmpBuf2$, sTmp$
    Dim sSampNo As String
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim iPos%


    'Data�� Edit�ϱ� ���ϵ��� <ETB> ����
    Do
        iETBpos = InStr(1, RcvBuffer, Chr(23))

        If iETBpos = 0 Then Exit Do

        sTmpBuf1 = Left$(RcvBuffer, iETBpos - 1)
        sTmpBuf2 = Mid$(RcvBuffer, iETBpos + 18)
        RcvBuffer = ""
        RcvBuffer = sTmpBuf1 & sTmpBuf2
    Loop While iETBpos <> 0


    sBC = Mid$(RcvBuffer, 1, 1)
    sLC = Mid$(RcvBuffer, 2, 1)

    Select Case sBC
        Case "R"
            If sLC = "B" Or sLC = "E" Then Exit Sub

            With pSampleInfo
                .RACK = Mid(RcvBuffer, 3, 4)
                .POS = Mid(RcvBuffer, 7, 2)
                .SEQNO = Mid(RcvBuffer, 10, 4)
                .ID = Trim(Mid(RcvBuffer, 21, 20))
                iPos = InStr(.ID, Chr(3))   'BARCODE ����
                If iPos <> 0 Then
                    .ID = Mid(.ID, 1, iPos - 1)
                End If

                If UCase(Left(.ID, 3)) = "ERR" Then
                    RaiseEvent DispMsg(.RACK & "/" & .POS & " - BARCODE READ ERROR!!!")
                    Exit Sub
                End If
            End With

            'Order ����...
            Call SendOrder_AU640

        Case "D"
            If sLC = "B" Or sLC = "E" Then Exit Sub

            '������� �ʱ�ȭ
            Call Init_pResultInfo

            'Sample ���� ����
            With pResultInfo
                .RACK = Mid$(RcvBuffer, 3, 4)
                .POS = Mid$(RcvBuffer, 7, 2)
                .SEQNO = Mid$(RcvBuffer, 10, 4)
                .ID = Trim(Mid$(RcvBuffer, 21, 20))
                'BARCODE ����
                iPos = InStr(.ID, Chr(3))
                If iPos <> 0 Then
                    .ID = Mid(.ID, 1, iPos - 1)
                End If
                iPos = InStr(.ID, " ")
                If iPos <> 0 Then
                    .ID = Mid(.ID, 1, iPos - 1)
                End If

                If Trim(.ID) = "" Then Exit Sub
            End With

            If m_iTotalItemCnt = 0 Then m_iTotalItemCnt = 100
            
            '�������
            For ii = 1 To m_iTotalItemCnt   '100
                sTmp = Mid(RcvBuffer, 39 + 10 * (ii - 1), 1)

                If Asc(sTmp) = 3 Then Exit For

                tmpIFCd = Format(Val(Mid(RcvBuffer, 39 + 10 * (ii - 1), 2)), "00")
                tmpRst = Trim(Mid(RcvBuffer, 39 + 10 * (ii - 1) + 2, 6))
                tmpFlag = Trim(Mid(RcvBuffer, 39 + 10 * (ii - 1) + 8, 2))

                If Left(tmpRst, 1) = "." Then
                    tmpRst = "0" & tmpRst
                End If

                '����� ����
                If Trim(tmpIFCd) <> "" Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1

                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & tmpFlag & Chr(124)
                    End With
                End If
            Next ii

            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "")
                End If
            End With

        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit ���� - (" & Err.Description & ")")
    End If
End Sub
'
'   AU-680(���ڵ� ���)
'
Private Sub DataEditResponse_AU680()
    On Error GoTo ErrRtn

    Dim sBC     As String
    Dim sLC     As String
    Dim iETBpos%, ii%, kk%
    Dim sTmpBuf1$, sTmpBuf2$, sTmp$
    Dim sSampNo As String
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim iPos%


    'Data�� Edit�ϱ� ���ϵ��� <ETB> ����
    Do
        iETBpos = InStr(1, RcvBuffer, Chr(23))

        If iETBpos = 0 Then Exit Do

        sTmpBuf1 = Left$(RcvBuffer, iETBpos - 1)
        sTmpBuf2 = Mid$(RcvBuffer, iETBpos + 18)
        RcvBuffer = ""
        RcvBuffer = sTmpBuf1 & sTmpBuf2
    Loop While iETBpos <> 0


    sBC = Mid$(RcvBuffer, 1, 1)
    sLC = Mid$(RcvBuffer, 2, 1)

    Select Case sBC
        Case "R"
            If sLC = "B" Or sLC = "E" Then Exit Sub

            With pSampleInfo
                .RACK = Mid(RcvBuffer, 3, 4)
                .POS = Mid(RcvBuffer, 7, 2)
                .SEQNO = Mid(RcvBuffer, 10, 4)
                .ID = Trim(Mid(RcvBuffer, 21, 20))
                iPos = InStr(.ID, Chr(3))   'BARCODE ����
                If iPos <> 0 Then
                    .ID = Mid(.ID, 1, iPos - 1)
                End If

                If UCase(Left(.ID, 3)) = "ERR" Then
                    RaiseEvent DispMsg(.RACK & "/" & .POS & " - BARCODE READ ERROR!!!")
                    Exit Sub
                End If
            End With

            'Order ����...
            Call SendOrder_AU680

        Case "D", "d"
            If sLC = "B" Or sLC = "E" Then Exit Sub

            If sLC = "Q" Then
                pResultInfo.KIND = "Q"  'QC Data
            End If
            
            '������� �ʱ�ȭ
            Call Init_pResultInfo

            'Sample ���� ����
            With pResultInfo
                .RACK = Mid$(RcvBuffer, 3, 4)
                .POS = Mid$(RcvBuffer, 7, 2)
                .SEQNO = Mid$(RcvBuffer, 10, 4)
                .ID = Trim(Mid$(RcvBuffer, 14, 24))
'                .ID = Trim(Mid$(RcvBuffer, 21, 20))'����
                'BARCODE ����
                iPos = InStr(.ID, Chr(3))
                If iPos <> 0 Then
                    .ID = Mid(.ID, 1, iPos - 1)
                End If
                iPos = InStr(.ID, " ")
                If iPos <> 0 Then
                    .ID = Mid(.ID, 1, iPos - 1)
                End If

                If Trim(.ID) = "" Then Exit Sub
            End With

            If m_iTotalItemCnt = 0 Then m_iTotalItemCnt = 120
            
            '�������
            For ii = 1 To m_iTotalItemCnt   '120
                sTmp = Mid(RcvBuffer, 39 + 11 * (ii - 1), 1)

                If Asc(sTmp) = 3 Then Exit For

                tmpIFCd = Format(Val(Mid(RcvBuffer, 39 + 11 * (ii - 1), 3)), "000")
                tmpRst = Trim(Mid(RcvBuffer, 39 + 11 * (ii - 1) + 3, 6))
                tmpFlag = Trim(Mid(RcvBuffer, 39 + 11 * (ii - 1) + 9, 2))

                If Left(tmpRst, 1) = "." Then
                    tmpRst = "0" & tmpRst
                End If

                '����� ����
                If Trim(tmpIFCd) <> "" Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1

                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & tmpFlag & Chr(124)
                    End With
                End If
            Next ii

            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .KIND, "", "")
                End If
            End With

        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit ���� - (" & Err.Description & ")")
    End If
End Sub
'
'   AU-400(���ڵ� ���)
'
Private Sub DataEditResponse_AU400()
    On Error GoTo ErrRtn
    
    Dim sBC     As String
    Dim sLC     As String
    Dim iETBpos%, ii%, kk%
    Dim sTmpBuf1$, sTmpBuf2$, sTmp$
    Dim sSampNo As String
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    
    'Data�� Edit�ϱ� ���ϵ��� <ETB> ����
    Do
        iETBpos = InStr(1, RcvBuffer, Chr(23))
        
        If iETBpos = 0 Then Exit Do
        
        sTmpBuf1 = Left$(RcvBuffer, iETBpos - 1)
        sTmpBuf2 = Mid$(RcvBuffer, iETBpos + 18)
        RcvBuffer = ""
        RcvBuffer = sTmpBuf1 & sTmpBuf2
    Loop While iETBpos <> 0
    
    
    sBC = Mid$(RcvBuffer, 1, 1)
    sLC = Mid$(RcvBuffer, 2, 1)
    
    Select Case sBC
        Case "R"
            If sLC = "B" Or sLC = "E" Then Exit Sub
            
            With pSampleInfo
                .RACK = Mid(RcvBuffer, 3, 4)
                .POS = Mid(RcvBuffer, 7, 2)
                .SEQNO = Mid(RcvBuffer, 10, 4)
                .ID = Trim(Mid(RcvBuffer, 14, 20))
                
                If UCase(Left(.ID, 3)) = "ERR" Then
                    RaiseEvent DispMsg(.RACK & "/" & .POS & " - BARCODE READ ERROR!!!")
                    Exit Sub
                End If
            End With
            
            'Order ����...
            Call SendOrder_AU400
            
        Case "D"
            If sLC = "B" Or sLC = "E" Then Exit Sub
            
            '������� �ʱ�ȭ
            Call Init_pResultInfo
            
            'Sample ���� ����
            With pResultInfo
                .RACK = Mid$(RcvBuffer, 3, 4)
                .POS = Mid$(RcvBuffer, 7, 2)
                .SEQNO = Mid$(RcvBuffer, 10, 4)
                .ID = Mid$(RcvBuffer, 14, 20)
                .ID = Trim(.ID)
                
                If Trim(.ID) = "" Then Exit Sub
            End With
            
            '�������
            For ii = 1 To m_iTotalItemCnt   '100
                sTmp = Mid(RcvBuffer, 39 + 10 * (ii - 1), 1)
                
                If Asc(sTmp) = 3 Then Exit For
                
                tmpIFCd = Format(Val(Mid(RcvBuffer, 39 + 10 * (ii - 1), 2)), "00")
                tmpRst = Trim(Mid(RcvBuffer, 39 + 10 * (ii - 1) + 2, 6))
                tmpFlag = Trim(Mid(RcvBuffer, 39 + 10 * (ii - 1) + 8, 2))
                
                If Left(tmpRst, 1) = "." Then
                    tmpRst = "0" & tmpRst
                End If
                
                '����� ����
                If Trim(tmpIFCd) <> "" Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & tmpFlag & Chr(124)
                    End With
                End If
            Next ii

            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "")
                End If
            End With
            
        Case Else
        
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit ���� - (" & Err.Description & ")")
    End If
End Sub

Private Sub Edit_Data_au5400()
'    On Error GoTo ErrHandler
'
'    '<---- COBAS ��񿡼� �ַ� ��� S --->
'    Dim sBC          As String
'    Dim sLC          As String
'    Dim iBCpos       As Integer
'    Dim iLCpos       As Integer
'
'    Dim iErrCode     As Integer
'    Dim sGeneralErrCode    As String
''<---- COBAS ��񿡼� �ַ� ��� E --->
'
'    Dim sJDate     As String
'    Dim sJGbn      As String
'    Dim sJNo      As String
'
'    Dim sIFRstCd    As String   '�������̽��� �˻��׸��ڵ�
'
'    Dim sBarCd      As String
'    Dim sSampNo    As String
'    Dim sRack      As String
'    Dim sPos       As String
'
'    Dim sRst     As String
'    Dim sRst2    As String
'
'    Dim sTotRst     As String
'    Dim sTotRst2    As String
'    Dim sTotRstCd   As String
'
'    Dim i           As Integer
'    Dim iTestStart  As Integer
'    Dim iPos        As Integer
'    Dim iRstCnt     As Integer
'
'    Dim sTmp        As String
'    Dim sBuf        As String
'
'    sBC = Mid(msRcvBuffer, 1, 1)
'    sLC = Mid(msRcvBuffer, 2, 1)
'
'    Select Case sBC
'        Case "R"
'            If sLC = "B" Or sLC = "E" Then
'                Exit Sub
'            End If
'
'            With gOrderTable
'                'AU5400���� Order Request�� �ѱ� Rack
'                .sRack = Mid(msRcvBuffer, 3, 4)
'                'AU5400���� Order Request�� �ѱ� Pos
'                .sPos = Mid(msRcvBuffer, 7, 2)
'                'AU5400���� Order Request�� �ѱ� ���ù�ȣ(����Ϸù�ȣ)
'                .sSampNo = Mid(msRcvBuffer, 9, 5)
'                'AU5400���� Order Request�� �ѱ� ���ڵ��ȣ
'                .sSampID = Trim(Mid(msRcvBuffer, 14, Val(gOrdCfg.sFSize(3))))
'
'                If UCase(Left(.sSampID, 3)) = "ERR" Then
'                    ViewMsgLog .sRack & " " & .sPos & " - BARCODE READING ����"
'
'                    Exit Sub
'                End If
'            End With
'
'            msSndState = "S"
'
'            'Order Request ��û ���� ��
'            Call Order_Input
'
'            Exit Sub
'
'        Case "D"
'            If sLC = "B" Or sLC = "E" Then
'                Exit Sub
'            End If
'
'            sRack = Mid(msRcvBuffer, 3, 4)
'            sPos = Mid(msRcvBuffer, 7, 2)
'            sSampNo = Trim(Mid(msRcvBuffer, 9, 5))
'            sBarCd = Trim(Mid(msRcvBuffer, 14, Val(gRstcfg.sFSize(3))))
'
'            If UCase(Left(sSampNo, 1)) = "Q" Or UCase(Left(sSampNo, 1)) = "C" Then
'                lblResult = "QC or CAL !!"
'
'                Exit Sub
'            End If
'
'            iTestStart = 14 + Val(gRstcfg.sFSize(3)) + 5
'
'            '--- �������
'            For i = 1 To 100
'                sTmp = Mid(msRcvBuffer, iTestStart + 10 * (i - 1), 1)
'
'                If Asc(sTmp) = 3 Then Exit For
'
'                sIFRstCd = CStr(Val(Mid(msRcvBuffer, iTestStart + 10 * (i - 1), 2)))
'
'                sRst = Trim(Mid(msRcvBuffer, iTestStart + 10 * (i - 1) + 2, 6))
'
'
'                If Left(sRst, 1) = "." Then
'                    sRst = "0" & sRst
'                End If
'
'                If sIFRstCd <> "" Then
'                    sTotRst = sTotRst & sRst & Chr(124)
'                    sTotRst2 = sTotRst2 & Chr(124)
'                    sTotRstCd = sTotRstCd & sIFRstCd & Chr(124)
'                    iRstCnt = iRstCnt + 1
'                End If
'            Next
'
'            If Len(sBarCd) = Val(gRstcfg.sFSize(3)) Then
'                sJNo = sBarCd
'            Else
'                sJNo = sSampNo
'            End If
'
'            If sJNo <> "" Then
'                Call DisplayResultOkBySex(3, Format(dtpLabDate.Value, "YYYYMMDD"), "", _
'                                            "", "", sJNo, sRack, sPos, "", "", "", "", "", "", _
'                                            iRstCnt, sTotRstCd, sTotRst, sTotRst2, "", "")
'            End If
'
'        Case Else
'    End Select
'
'    Exit Sub
'
'ErrHandler:
'    ViewMsg "Edit_Data ���� �߻�" & "(" & CStr(Err.Number) & " : " & Err.Description & ")"
End Sub

'
'   ���ڵ� ��� ���ϴ� AU-600
'
Private Sub DataEditResponse_AU600()
    On Error GoTo ErrRtn
    
    Dim sBC     As String
    Dim sLC     As String
    Dim iETBpos%, ii%, kk%
    Dim sTmpBuf1$, sTmpBuf2$, sTmp$
    Dim sSampNo As String
    Dim tmpIFCd$, tmpRst$
    
    
    'Data�� Edit�ϱ� ���ϵ��� <ETB> ����
    Do
        iETBpos = InStr(1, RcvBuffer, Chr(23))
        
        If iETBpos = 0 Then
            Exit Do
        End If
        
        sTmpBuf1 = Left$(RcvBuffer, iETBpos - 1)
        sTmpBuf2 = Mid$(RcvBuffer, iETBpos + 21)
        RcvBuffer = ""
        RcvBuffer = sTmpBuf1 & sTmpBuf2
    Loop While iETBpos <> 0
    
    
    sBC = Mid$(RcvBuffer, 1, 1)
    sLC = Mid$(RcvBuffer, 2, 1)
    
    Select Case sBC
        Case "R"
            If sLC = "B" Or sLC = "E" Then
                Exit Sub
            End If
            
            sSampNo = Mid$(RcvBuffer, 10, 4)
            
            'Order ����...
            Call SendOrder_AU600(sSampNo)
            
        Case "D"
            If sLC = "B" Or sLC = "E" Then
                Exit Sub
            End If
            
            '������� �ʱ�ȭ
            Call Init_pResultInfo
            
            'Sample ���� ����
            With pResultInfo
                .RACK = Mid$(RcvBuffer, 3, 4)
                .POS = Mid$(RcvBuffer, 7, 2)
                .SEQNO = Mid$(RcvBuffer, 10, 4)
                .ID = Mid$(RcvBuffer, 14, 20)    '16)
                .ID = Trim(.ID)
                
                If Trim(.ID) = "" Then Exit Sub
            End With
            
            '�������
            For ii = 1 To m_iTotalItemCnt   '100
                sTmp = Mid(RcvBuffer, 39 + 10 * (ii - 1), 1)
                
                If Asc(sTmp) = 3 Then Exit For
                
                tmpIFCd = Format(Val(Mid(RcvBuffer, 39 + 10 * (ii - 1), 2)), "00")
                tmpRst = Trim(Mid(RcvBuffer, 39 + 10 * (ii - 1) + 2, 6))
                
                '����� ����
                If Trim(tmpIFCd) <> "" Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & Chr(124)
                    End With
                End If
'                For kk = 1 To m_iTotalItemCnt
'                    If Trim(tmpIFCd) = gIFItem(k).s09 Then
'                        sTotTestCd = sTotTestCd & gIFItem(k).s02 & gIFItem(k).s03 & gIFItem(k).s04 & gIFItem(k).s05 & gIFItem(k).s06 & Chr(124)
'                        sTotRst = sTotRst & sRst & Chr(124)
'                        iRstCnt = iRstCnt + 1
'                        Exit For
'                    End If
'                Next k
            Next ii

            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "")
                End If
            End With
            
        Case Else
        
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit ���� - (" & Err.Description & ")")
    End If
End Sub

Private Sub SendOrder_AU5400()
    On Error GoTo ErrRtn

    'ȯ���� Order ����
    Dim SendBuf As String
    Dim sTestCd As String
    Dim ii      As Integer
    Dim iCnt    As Integer
    Dim tmpData()   As String

    '���� ������ ���� ��ȸ
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)

    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
        End With
        Exit Sub
    End If

    m_iPhase = 1

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
        .ORDCNT = iCnt      '���� �˻� ������ �׸� ����
    End With

    'Send Message ����
    SendBuf = Chr$(2)
    SendBuf = SendBuf & "S" & Space(1) & pSampleInfo.RACK & pSampleInfo.POS & Space$(1)
    SendBuf = SendBuf & pSampleInfo.SEQNO
    SendBuf = SendBuf & Space(20 - Len(pSampleInfo.ID)) & pSampleInfo.ID & Space$(4) & "E"

    'Order ���ڿ��� ������
    sTestCd = ""

    Dim tmpOrd(1 To 99) As String
    For ii = 1 To pSampleInfo.ORDCNT
        tmpOrd(Val(pSampleInfo.IFCD(ii))) = Format(pSampleInfo.IFCD(ii), "00")
    Next ii
    For ii = 1 To 99
        If Trim(tmpOrd(ii)) <> "" Then
            sTestCd = sTestCd & Trim(tmpOrd(ii))
        End If
    Next ii
''    For ii = 1 To pSampleInfo.ORDCNT
''        sTestCd = sTestCd & Format(pSampleInfo.IFCD(ii), "00")
''    Next ii

    SendBuf = SendBuf & sTestCd & Chr$(3)


    Call Sleep(500)

    msComm.Output = SendBuf

    'Order ���� �Ϸ�
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

    'Log �ۼ�
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuf)
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder �����߻� - " & Err.Description)
    End If
End Sub
Private Sub SendOrder_AU640()
    On Error GoTo ErrRtn

    'ȯ���� Order ����
    Dim SendBuf As String
    Dim sTestCd As String
    Dim ii      As Integer
    Dim iCnt    As Integer
    Dim tmpData()   As String

    '���� ������ ���� ��ȸ
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)

    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
        End With
        Exit Sub
    End If

    m_iPhase = 1

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
        .ORDCNT = iCnt      '���� �˻� ������ �׸� ����
    End With

    'Send Message ����
    SendBuf = Chr$(2)
    SendBuf = SendBuf & "S" & Space(1) & pSampleInfo.RACK & pSampleInfo.POS & Space$(1)
    SendBuf = SendBuf & pSampleInfo.SEQNO
    SendBuf = SendBuf & Space(20 - Len(pSampleInfo.ID)) & pSampleInfo.ID & Space$(4) & "E"

    'Order ���ڿ��� ������
    sTestCd = ""

    Dim tmpOrd(1 To 99) As String
    For ii = 1 To pSampleInfo.ORDCNT
        tmpOrd(Val(pSampleInfo.IFCD(ii))) = Format(pSampleInfo.IFCD(ii), "00")
    Next ii
    For ii = 1 To 99
        If Trim(tmpOrd(ii)) <> "" Then
            sTestCd = sTestCd & Trim(tmpOrd(ii))
        End If
    Next ii
''    For ii = 1 To pSampleInfo.ORDCNT
''        sTestCd = sTestCd & Format(pSampleInfo.IFCD(ii), "00")
''    Next ii

    SendBuf = SendBuf & sTestCd & Chr$(3)


    Call Sleep(500)

    msComm.Output = SendBuf

    'Order ���� �Ϸ�
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

    'Log �ۼ�
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuf)
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder �����߻� - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_AU680()
    On Error GoTo ErrRtn

    'ȯ���� Order ����
    Dim SendBuf As String
    Dim sTestCd As String
    Dim ii      As Integer
    Dim iCnt    As Integer
    Dim tmpData()   As String

    '���� ������ ���� ��ȸ
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)

    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
        End With
        Exit Sub
    End If

    m_iPhase = 1

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
        .ORDCNT = iCnt      '���� �˻� ������ �׸� ����
    End With

    'Send Message ����
    SendBuf = Chr$(2)
    SendBuf = SendBuf & "S" & Space(1) & pSampleInfo.RACK & pSampleInfo.POS & Space$(1)
    SendBuf = SendBuf & pSampleInfo.SEQNO
    SendBuf = SendBuf & Space(20 - Len(pSampleInfo.ID)) & pSampleInfo.ID & Space$(4) & "E"

    'Order ���ڿ��� ������
    sTestCd = ""

    Dim tmpOrd(1 To 120) As String '��ȣ ����
    For ii = 1 To pSampleInfo.ORDCNT
        tmpOrd(Val(pSampleInfo.IFCD(ii))) = Format(pSampleInfo.IFCD(ii), "000") & "0"
    Next ii
    For ii = 1 To 120 '��ȣ ����
        If Trim(tmpOrd(ii)) <> "" Then
            sTestCd = sTestCd & Trim(tmpOrd(ii))
        End If
    Next ii
''    For ii = 1 To pSampleInfo.ORDCNT
''        sTestCd = sTestCd & Format(pSampleInfo.IFCD(ii), "00")
''    Next ii

    SendBuf = SendBuf & sTestCd & Chr$(3)


    Call Sleep(500)

    msComm.Output = SendBuf

    'Order ���� �Ϸ�
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

    'Log �ۼ�
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuf)
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder �����߻� - " & Err.Description)
    End If
End Sub


Private Sub SendOrder_AU400()
    On Error GoTo ErrRtn

    'ȯ���� Order ����
    Dim SendBuf As String
    Dim sTestCd As String
    Dim ii      As Integer
    Dim iCnt    As Integer
    Dim tmpData()   As String

    '���� ������ ���� ��ȸ
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)

    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
        End With
        Exit Sub
    End If

    m_iPhase = 1

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
        .ORDCNT = iCnt      '���� �˻� ������ �׸� ����
    End With

    'Send Message ����
    SendBuf = Chr$(2)
    SendBuf = SendBuf & "S" & Space(1) & pSampleInfo.RACK & pSampleInfo.POS & Space$(1)
    SendBuf = SendBuf & pSampleInfo.SEQNO
    SendBuf = SendBuf & Space(20 - Len(pSampleInfo.ID)) & pSampleInfo.ID & Space$(4) & "E"

    'Order ���ڿ��� ������
    sTestCd = ""
    For ii = 1 To pSampleInfo.ORDCNT
        sTestCd = sTestCd & Format(pSampleInfo.IFCD(ii), "00")
    Next ii

    SendBuf = SendBuf & sTestCd & Chr$(3)


    Call Sleep(100)

    msComm.Output = SendBuf

    'Order ���� �Ϸ�
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

    'Log �ۼ�
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuf)
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder �����߻� - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_AU600(ByVal sSampNo As String)
    On Error GoTo ErrRtn
    
    'ȯ���� Order ����
    Dim SendBuf As String
    Dim sTestCd As String
    Dim ii      As Integer
    Dim iCnt    As Integer
    Dim tmpData()   As String
    
    '���� ������ ���� ��ȸ
    RaiseEvent RequestCurOrder("", "", "", "")
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        Exit Sub
    End If
    
    m_iPhase = 1
    
    ReDim tmpData(m_p_iOrdCnt) As String
    tmpData() = Split(m_p_sTIFCd, Chr(124))
    
    With pSampleInfo
        .ID = m_p_sID
        .SEQNO = sSampNo
        .RACK = Space(4)
        .POS = Space(2)
        .ORDCNT = m_p_iOrdCnt
    
        ReDim .IFCD(.ORDCNT)
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '���� �˻� ������ �׸� ����
    End With
    
    'Send Message ����
    SendBuf = Chr$(2)
    SendBuf = SendBuf & "S " & pSampleInfo.RACK & pSampleInfo.POS & Space$(1)
    SendBuf = SendBuf & pSampleInfo.SEQNO
    SendBuf = SendBuf & Space(4) & pSampleInfo.ID & Space$(4) & "E"
    
    'Order ���ڿ��� ������
    sTestCd = ""
    For ii = 1 To pSampleInfo.ORDCNT
        sTestCd = sTestCd & pSampleInfo.IFCD(ii)
    Next ii

    SendBuf = SendBuf & sTestCd & Chr$(3)
    
    
    Call Sleep(100)
    
    msComm.Output = SendBuf
    
    'Order ���� �Ϸ�
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
    'Log �ۼ�
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuf)
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder �����߻� - " & Err.Description)
    End If
End Sub


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
        Case "AU600"
            Call PhaseCfg_Protocol_AU600
            
        Case "AU400"
            Call PhaseCfg_Protocol_AU400
            
        Case "AU640"
            Call PhaseCfg_Protocol_AU640
        
        Case "AU5400", "AU2700"
            Call PhaseCfg_Protocol_AU5400
            
        Case "AU680"
            Call PhaseCfg_Protocol_AU680 '��ȣ �߰�
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
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
        .ORDCNT = iCnt      '���� �˻� ������ �׸� ����
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
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .KIND = ""
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
    
End Sub
'����ҿ��� �Ӽ����� �ε��մϴ�.
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
    m_iOrderFlag = PropBag.ReadProperty("iOrderFlag", m_def_iOrderFlag)
    m_iTotalItemCnt = PropBag.ReadProperty("iTotalItemCnt", m_def_iTotalItemCnt)
'    m_iLenID = PropBag.ReadProperty("iLenID", m_def_iLenID)
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
'    Call PropBag.WriteProperty("iLenID", m_iLenID, m_def_iLenID)
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
'    m_iLenID = m_def_iLenID
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

