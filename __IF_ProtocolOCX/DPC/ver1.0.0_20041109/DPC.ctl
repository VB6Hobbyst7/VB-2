VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl DPC 
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
Attribute VB_Name = "DPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�⺻ �Ӽ� ��:
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
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sRack$, sPos$)
Event DispMsg(sMsg$)
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


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
Dim iSpaceCnt   As Integer

'For DPC2000
Dim sHeaderInfo()   As String



Private Sub PhaseCfg_Protocol_DPC2000()
            
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
        Select Case m_iPhase
            Case 1            'ENQ ���
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2      '<LF> ���
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                        
                    Case 10     '<LF>
                        Call DataEditResponse_DPC2000       'Data ����
                        
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                                                
                    Case 4      'EOT
                        If sState = "Q" Then
                            msComm.Output = Chr(5)
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3
                        
                    Case 5      'ENQ
                        msComm.Output = Chr(6)   'Send ACK
                        
                    Case 21     'NAK
                        msComm.Output = Chr(5)   'Send ENQ
                        m_iPhase = 1
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                End Select
            
            Case 3      'ACK ���
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        If sState = "Q" Then
                            Call SendOrder_DPC2000
                        End If
                    
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    
                    Case 21     'NAK
                        msComm.Output = Chr(5)
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
Private Sub DataEditResponse_Immulite2000_�����������()
'
'    Dim RecType     As String   'Record Type
'    Dim Tmp     As String
'    Dim i       As Integer
'    Dim ChkR    As Integer
'    Dim RP_Chk  As Integer
'
'    Dim TempTBL(4)  As String
'    Dim tmpRack     As String
'    Dim tmpPos      As String
'    Dim tmpData()   As String
'
'    Dim sRst        As String
'    Dim sRst2       As String
'    Dim sIFRstCd    As String   '�������̽��� �˻��׸��ڵ�
'
'    Dim iCRow   As Integer
'
'
'    RecType = Mid$(RcvBuffer, 2, 1)
'
'    Select Case RecType
'        Case "H"        'Header Record
'            Erase sHeaderInfo()
'            sHeaderInfo() = Split(RcvBuffer, "|")   'HEADER ���� ����
'        Case "M"
'        Case "P"        'Patient Record
'        Case "Q"        'Order Request Record
'            TempTBL(0) = ""
'            For i = 7 To 21
'                Tmp = Mid$(RcvBuffer, i, 1)
'                If Tmp = "|" Or Tmp = "^" Then
'                    Exit For
'                Else
'                    TempTBL(0) = TempTBL(0) & Tmp
'                End If
'            Next i
'            If Trim(Trim(TempTBL(0))) <> "" Then    'BarCode ID�� �� �Ѿ�Դ��� �˻�
'                State = "Q"
'                sSampID = Trim(TempTBL(0))
'            Else
'                State = ""
'                sSampID = ""
'            End If
'
'        Case "O"
'            TempTBL(1) = ""
'            For i = 6 To 21
'                Tmp = Mid$(RcvBuffer, i, 1)
'                If Tmp = "|" Or Tmp = "^" Then
'                    Exit For
'                Else
'                    TempTBL(1) = TempTBL(1) & Tmp
'                End If
'            Next i
'
'            tmpRack = "": tmpPos = ""   '0
''            tmpData() = Split(RcvBuffer, "|")
''            i = InStr(1, tmpData(4), "^")
''            If i <> 0 Then
''                tmpRack = Mid(tmpData(4), i + 1, 1)
''                tmpData(4) = Mid(tmpData(4), i + 4)     '3)     2001/2/20
''            End If
''            i = InStr(1, tmpData(4), "^")
''            If i <> 0 Then
''                tmpPos = Mid(tmpData(4), 1, i - 1)
''            End If
'
'            sSampID = Trim(TempTBL(1))      '���ڵ�
''            gOrderTable.sRack = Format(tmpRack, "0000")
''            gOrderTable.sPos = Format(tmpPos, "0")
'
'        Case "R"        'Result Record
'            For i = 2 To 3
'                TempTBL(i) = ""
'            Next i
'            ChkR = False
'
'            For i = 9 To Len(RcvBuffer)
'                Tmp = Mid$(RcvBuffer, i, 1)
'                If ChkR = False Then    'yk 8/22
'                    If Tmp = "|" Then
''                    If Tmp = "^" Then
'                        ChkR = True
''                        i = i + 3   '3  9/19
'                    Else
'                        TempTBL(2) = TempTBL(2) & Tmp
'                    End If
'                Else
'                    If Tmp = "|" Then
'                        Exit For
'                    Else
'                        TempTBL(3) = TempTBL(3) & Tmp
'                    End If
'                End If
'            Next i
'
'            sIFRstCd = TempTBL(2)
'            sRst = TempTBL(3)
'            '��� ����(�ڸ��� ��)
'            sRst = Edit_Result(sIFRstCd, sRst)
'
''            'SCL�� ��� ���ϴ� HCG ������ ��� �������� �ʾ� �������� ó�� �߰�...2001/5/7 YK
''            sRst = Edit_ResultUnit(sIFRstCd, sRst)
'
'
'            '--- Local DB�� ��� ���
'            Call Append_LocalDB(sSampID, sIFRstCd, sRst)
'
'            '--- spdList�� ��� ǥ��/spdBarCode���� �ش� ID ����
'            Call Disp_spdList(sSampID, sIFRstCd, sRst, iCRow)
'    End Select
    
End Sub

' *=====================================================*
' *               Data���� & ����ó��                   *
' *=====================================================*
Private Sub DataEditResponse_DPC2000()
    On Error GoTo ErrRtn
    
    Dim RecType     As String       'Record Type
    Dim sResType    As String
    Dim sResState   As String
    Dim sResData    As String
    
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$
    
    Dim tmpField()  As String
    Dim tmpData()   As String
   
        
    RecType = Mid$(RcvBuffer, 2, 1)
    
    Select Case RecType
        Case "H"        'Header Record
            Erase sHeaderInfo()
            sHeaderInfo() = Split(RcvBuffer, "|")   'HEADER ���� ����
            
        Case "P"        'Patient Record
            Call Init_pResultInfo
            
        Case "Q"        'Order Request Record
            tmpField() = Split(RcvBuffer, "|")
            tmpData() = Split(tmpField(2), "^")
            tmpBarCd = Trim(tmpData(1))
                        
            If tmpBarCd = "" Then
                sState = ""
                pSampleInfo.ID = ""
                Exit Sub
            Else
                sState = "Q"
                pSampleInfo.ID = tmpBarCd
            End If

        Case "O"        'Order Record
            tmpField() = Split(RcvBuffer, "|")
            tmpBarCd = Trim(tmpField(2))
            tmpSeqNo = ""
            tmpRack = ""
            tmpPos = ""
            
            pSampleInfo.ID = tmpBarCd
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            
        Case "R"        'Result Record
            '--- �������Ÿ ����
            'tmpData(2): TESTCD
            '    "  (3): RESULT
            '    "  (4): UNIT
            '    "  (5): ����ġ ����
            '    "  (6): Result Abnormal Flags
            tmpField() = Split(RcvBuffer, "|")
            
            sResData = Trim(tmpField(3))
            sResState = Trim(tmpField(8))
            tmpUnit = Trim(tmpField(4))
            tmpFlag = Trim(tmpField(6))
            
            'TestCd/Result Type Edit
            tmpData() = Split(Trim(tmpField(2)), "^")
            tmpIFCd = Trim(tmpData(3))
            
            tmpRst = sResData
            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
        
            '������� ����ü�� ����
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS

                '����� ����
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
            End With
        
            '����� ���/ȭ�� ǥ�� ó��...(DPC2000�� ��� 'R'���� �������ؾ� ��)
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With
            
        Case "C"        'Comment Record
        
        Case "L"        'Msg Terminater Record
            Call Init_pResultInfo
                        
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_DPC2000()
    On Error GoTo ErrRtn
    
    Dim sTmp    As String
    Dim ChkS    As String
    Dim TestDat As String
    Dim i       As Integer
        
    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If
    
    Select Case m_iSendPhase
        Case 1      'Header Record
            sTmp = m_iFrameN & "H|" & sHeaderInfo(1) & "||||" & sHeaderInfo(5) & "||" & sHeaderInfo(7) _
                    & "|8N1|||P|1|1919990520173537" & Chr(13) & Chr(3)
            m_iSendPhase = 2
            
        Case 2      'Patient Record
            sTmp = m_iFrameN & "P|1|" & Trim(pSampleInfo.ID) & "|||||||||||||||||||||||||||||||" & Chr(13) & Chr(3)
            m_iSendPhase = 3
            
        Case 3      'Order Record
            TestDat = ""
            '----- �˻��׸� ��ȸ/����
            RaiseEvent RequestCurOrder(pSampleInfo.ID, "", "")
            
            Call Get_OrderString
            
            For i = 1 To pSampleInfo.ORDCNT
                TestDat = TestDat & "^^^" & pSampleInfo.IFCD(i) & "\"
            Next i
            If pSampleInfo.ORDCNT > 0 Then
                TestDat = Left(TestDat, Len(TestDat) - 1)       '"\" Cutting
            End If
            
            'BarCode �����
            sTmp = m_iFrameN & "O|1|" & Trim(pSampleInfo.ID) & "||" & TestDat & "|R||||||N||||||||||||||Q" & Chr(13) & Chr(3)
            m_iSendPhase = 4
            
        Case 4      'Terminator Record
            sTmp = m_iFrameN & "L|1|F" & Chr(13) & Chr(3)
            m_iSendPhase = 5
            
        Case 5      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1: m_iPhase = 1: m_iSendPhase = 1
            sState = ""
            
            'Barcode Mode�� ��� ���ۿϷ� �̺�Ʈ �߻�
            RaiseEvent SendOrderOK(pSampleInfo.ID, "", "")
                
            Exit Sub
            
    End Select
    
    'CheckSum ���
    ChkS = ChkSum_ASTM(sTmp)
    
    msComm.Output = Chr(2) & sTmp & ChkS & Chr(13) & Chr(10)
    
    m_iFrameN = m_iFrameN + 1

    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sTmp & ChkS & Chr(13) & Chr(10))
    End If
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder ���� - " & Err.Description)
    End If
End Sub


Private Sub SendOrder_Immulite2000_�����������()
'
'    Dim Tmp   As String
'    Dim ChkS  As String
'    Dim TestDat As String
'    Dim i   As Integer
'    Dim sTmpData()  As String
'
'
'    If FrameN > 7 Then
'        FrameN = 0
'    End If
'
'    Select Case Snd_Phase
'        Case 1      'H
''            Tmp = FrameN & "H|\^&||||111 Canfield Ave^Randolph^NJ^07869||(201)927-2828|8N1|||P|1|1919990520173537" & Chr(13) & Chr(3)
'            Tmp = FrameN & "H|" & sHeaderInfo(1) & "||||" & sHeaderInfo(5) & "||" & sHeaderInfo(7) & "|8N1|||P|1|1919990520173537" & Chr(13) & Chr(3)
'            Snd_Phase = 2
'
'        Case 2      'P                                             '13
'            Tmp = FrameN & "P|1|" & Trim(sSampID) & "|||||||||||||||||||||||||||||||" & Chr(13) & Chr(3)
'            Snd_Phase = 3
'
'        Case 3      'O
'            TestDat = ""
'            '----- �˻��׸� ��ȸ
'            Call Get_OrderString
'            If gOrderTable.iOrdCnt = 0 Then
''                pnlComm.Caption = "�������̽� ���� �׸��� �������� �ʽ��ϴ�!!"
'                '2001-05-21 KHS modified : Order��� O,L ��Ŷ �־����
'                'Exit Sub
'            End If
'
'            For i = 1 To gOrderTable.iOrdCnt
'                'HCG(Triple Mark?)�ϰ�� 20������ ���ش�.
'                If TestDat = "" Then
''                    If Trim$(ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))) = "HCG" Then
''                        TestDat = "^^^" & Trim$(ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))) & "^20"    'gOrderTable.sIFSeq(i)) '& "^0"
''                    Else
'                        TestDat = "^^^" & Trim$(gOrderTable.sIFCD(i))    'gOrderTable.sIFSeq(i)) '& "^0"
''                    End If
'                Else
''                    If Trim$(ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))) = "HCG" Then
''                        TestDat = TestDat & "\^^^" & Trim$(ConvertIFItemInfo(6, gOrderTable.sIFSeq(i))) & "^20"     'gOrderTable.sIFSeq(i)) '& "^0"
''                    Else
'                        TestDat = TestDat & "\^^^" & Trim$(gOrderTable.sIFCD(i))     'gOrderTable.sIFSeq(i)) '& "^0"
''                    End If
'                End If
'            Next i
'            '-------------------
'
'            Tmp = FrameN & "O|1|" & Trim(sSampID) & "||" & TestDat & "|R||||||N||||||||||||||Q" & Chr(13) & Chr(3)
'            Snd_Phase = 4
'
'        Case 4      'T
'            Tmp = FrameN & "L|1|F" & Chr(13) & Chr(3)
'            Snd_Phase = 5
'
'        Case 5      'EOT
'            Comm1.Output = Chr(4)   'EOT
'            FrameN = 1
'            Phase = 1
'            Snd_Phase = 1
'            State = ""
'            Exit Sub
'
'    End Select
'
'    ChkS = Chk_Sum(Tmp)
''    Print #2, Tmp
'    Comm1.Output = Chr(2) & Tmp & ChkS & Chr(13) & Chr(10)
'    FrameN = FrameN + 1
    
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
        Case "DPC2000"
            '���ڵ� ���
            Call PhaseCfg_Protocol_DPC2000
            
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
        MsgBox "PortOpen Error!!! " & Err.Description, vbCritical
        RaiseEvent DispMsg(Err.Description)
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
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr ���� - " & Err.Description)
    End If
End Function

