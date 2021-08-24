VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl BIOLIS 
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
Attribute VB_Name = "BIOLIS"
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
'Event RequestCurOrder(sID$, sRack$, sPos$)
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
Dim iSpaceCnt   As Integer

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
        Case "BIOLIS"
            If m_bUseBarcode = True Then
                '���ڵ� ���
'                Call PhaseCfg_Protocol_e411_BarcodeMode
            Else
                '���ڵ� ��� ����
                Call PhaseCfg_Protocol_BIOLIS
            End If
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub


'
'   e411 ���ڵ� ��� ���ϴ� ����
'
Private Sub PhaseCfg_Protocol_BIOLIS()
    
    Dim wkDat   As String
    Dim ix1     As Integer
       
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
       
        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                        
                    Case 10     '<LF>
                        Call DataEditResponse_BIOLIS
                        
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                        
                    Case 4      'EOT
                        m_iPhase = 1
                                                
                    Case 5      'ENQ
                        msComm.Output = Chr(6)   'Send ACK
                        
                    Case 21     'NAK
                        msComm.Output = Chr(5)   'Send ENQ
                        m_iPhase = 1
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                End Select
            
            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        Call SendOrder_BIOLIS     'Order ����
                        
                    Case 5      'ENQ
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                                                
                    Case 21     'NAK
                        m_iSendPhase = m_iSendPhase - 1
                        m_iFrameN = m_iFrameN - 1
                        m_iPhase = 3
                        
                        Call SendOrder_BIOLIS     'Order ����
                        
                    Case 4      'EOT
                        m_iPhase = 1
                        
                End Select
        End Select
    Next ix1
    
End Sub

' *=====================================================*
' *               Data���� & ����ó��                   *
' *=====================================================*
Private Sub DataEditResponse_BIOLIS()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim i       As Integer
    Dim aData()   As String
    Dim aField()  As String
    Dim ii      As Integer
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpKind$, tmpRstType$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$, tmpAlarmCd$

    RecType = Mid$(RcvBuffer, 2, 1)

    Select Case RecType
        Case "H"        'Header Record
            sState = ""
            
        Case "M"
        Case "P"        'Patient Record
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
            'Q|1|ALL||ALL||||||||O<CR>
            aData() = Split(RcvBuffer, "|")
            
            tmpBarCd = Trim(aData(2))
            If tmpBarCd <> "" Then      '2004/4/2 yk
                sState = "Q"
                pSampleInfo.ID = UCase(tmpBarCd)
            Else
                sState = ""
                pSampleInfo.ID = ""
            End If
            
        Case "O"
            'O|1|12345|^2^12|^^^1^GOT^0|R||||||||||Serum||||||||||F<CR>
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            aField() = Split(RcvBuffer, Chr(124))
            
            tmpBarCd = Trim(aField(2))
            
            aData() = Split(aField(3), "^")
            tmpSeqNo = Trim(aData(0))
            tmpRack = Trim(aData(1))
            tmpPos = Trim(aData(2))
            
            tmpKind = Trim(aField(11))

            pSampleInfo.ID = UCase(tmpBarCd)
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            pSampleInfo.KIND = tmpKind
            
        Case "R"        'Result Record
            '#Example of transmission when measurement succeeded
            'R|1|^^^1^GOT^0|54.5143|IU/L|8 TO 38|H||F||||20010618145805<CR>
            
            '#Example of transmission when measurement failed
            'R|1|^^^1^GOT^0||IU/L|8 TO 38|N||X||||20010618145805<CR>
            
            '--- �������Ÿ ����
            aData() = Split(RcvBuffer, "|")

            aField() = Split(aData(2), "^")
            tmpIFCd = Trim(aField(3))       'Test Item No.

            tmpRst = Trim(aData(3))
            '--- ������� "^" �� ��� ����
            ii = InStr(1, tmpRst, "^")
            If ii <> 0 Then tmpRst = Mid(tmpRst, ii + 1)

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
            
            tmpUnit = Trim(aData(4))
            tmpFlag = Trim(aData(6))
            If tmpFlag = "N" Then tmpFlag = ""
            
            tmpRstType = Trim(aData(8))
            If tmpRstType = "X" Then
                tmpRst = ""
            End If

            '������� ����ü�� ����
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS
                .KIND = pSampleInfo.KIND
                
                '����� ����
                .RSTCNT = 1
                .IFCD = tmpIFCd & Chr(124)
                .RST1 = tmpRst & Chr(124)
                .RST2 = Chr(124)
                .UNIT = tmpUnit & Chr(124)
                .FLAG = tmpFlag & Chr(124)
                .RSTDT = Chr(124)           '����Ͻ�
            
                '����� ���/ȭ�� ǥ�� ó��...
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD, .KIND, "", "")
            End With
            
            Call Init_pResultInfo
            
'            With pResultInfo
'                .ID = pSampleInfo.ID
'                .SEQNO = pSampleInfo.SEQNO
'                .RACK = pSampleInfo.RACK
'                .POS = pSampleInfo.POS
'                .KIND = pSampleInfo.KIND
'
'                '����� ����
'                .RSTCNT = .RSTCNT + 1
'                .IFCD = .IFCD & tmpIFCd & Chr(124)
'                .RST1 = .RST1 & tmpRst & Chr(124)
'                .RST2 = .RST2 & Chr(124)
'                .UNIT = .UNIT & tmpUnit & Chr(124)
'                .FLAG = .FLAG & tmpFlag & Chr(124)
'                .RSTDT = .RSTDT & Chr(124)          '����Ͻ�
'            End With

        Case "C"        'Comment Record
'            'Data Alarm ����
'            tmpField() = Split(RcvBuffer, Chr(124))
'
'            tmpData() = Split(tmpField(3), "^")
'            tmpAlarmCd = Trim(tmpData(0))
'            If tmpAlarmCd = "0" Then        '2005/4/20 yk
'                tmpAlarmCd = ""
'            End If
''            pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)
'            pResultInfo.ALARMCD = Replace(pResultInfo.ALARMCD, "####", tmpAlarmCd)      '2006/8/17 yk
            
        Case "L"
            '����� ���/ȭ�� ǥ�� ó��...
'            With pResultInfo
'                If .RSTCNT > 0 Then
'                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD, .KIND, "", "")
'                End If
'            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
'
'   e411 ���ڵ� ���
'
Private Sub PhaseCfg_Protocol_e411_BarcodeMode()

'    Dim wkDat   As String
'    Dim ix1     As Integer
'
'    For ix1 = 1 To Len(wkBuf)
'        wkDat = Mid$(wkBuf, ix1, 1)
'
'        Select Case m_iPhase
'            Case 1
'                Select Case Asc(wkDat)
'                    Case 5      'ENQ
'                        msComm.Output = Chr(6)
'                        m_iPhase = 2
'                    Case Else
'                        m_iPhase = 1
'                End Select
'
'            Case 2
'                Select Case Asc(wkDat)
'                    Case 2      'STX
'                        RcvBuffer = ""
'
'                    Case 10     '<LF>
'                        Call DataEditResponse_e411
'
'                        m_iPhase = 2
'                        msComm.Output = Chr(6)
'
'                    Case 4      'EOT
'                        If sState = "Q" Then
'                            msComm.Output = Chr(5)
'                            m_iSendPhase = 1
'                        End If
'                        m_iPhase = 3
'
'                    Case 5      'ENQ
'                        msComm.Output = Chr(6)   'Send ACK
'
'                    Case 21     'NAK
'                        msComm.Output = Chr(5)   'Send ENQ
'                        m_iPhase = 1
'
'                    Case Else
'                        RcvBuffer = RcvBuffer & wkDat
'                        m_iPhase = 2
'                End Select
'
'            Case 3
'                Select Case Asc(wkDat)
'                    Case 6      'ACK
'                        If sState = "Q" Then
'                            Call SendOrder_e411     'Order ����
'                        End If
'
'                    Case 5      'ENQ
'                        msComm.Output = Chr(6)
'                        m_iPhase = 2
'
'                    Case 21     'NAK
'                        msComm.Output = Chr(5)
'                        m_iPhase = 3
'
'                    Case 4      'EOT
'                        m_iPhase = 1
'                End Select
'
'        End Select
'    Next ix1
'
End Sub

Private Sub SendOrder_BIOLIS()
    On Error GoTo ErrRtn

    Dim sTmp    As String
    Dim ChkS    As String
    Dim TestDat As String
    Dim i       As Integer
    Dim sTmpData()  As String
    Dim sActionCd   As String
    Dim sReportType As String

    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If

    Select Case m_iSendPhase
        Case 0
            m_iSendPhase = 1
            msComm.Output = Chr(5)
            Exit Sub

        Case 1      'H
            sTmp = m_iFrameN & "H|\^&|||HOST^P_1|||||Prestige24i^SYSTEM1||P|1|" & Format(Now, "YYYYMMDDHHNNSS") & vbCr & Chr(3)
            m_iSendPhase = 2

        Case 2      'P
            sTmp = m_iFrameN & "P|1" & vbCr & Chr(3)
            m_iSendPhase = 3

        Case 3      'O
            TestDat = ""
            '----- �˻��׸� ��ȸ/����
            Call Get_OrderString
            
            'Order Record
            'O|1|123456|^1^20|^^^1^GOT^0��^^^11^LDH^0��^^^42^Ca^0|R||||||||||Serum||||||||||O<CR>
            
            sActionCd = "N"
            sReportType = "O"

            For i = 1 To pSampleInfo.ORDCNT
                If Trim$(pSampleInfo.IFCD(i)) = "" Then
                Else
                    TestDat = TestDat & "^^^" & Trim$(pSampleInfo.IFCD(i)) & "^^0\"  'item���� ���۾���
                End If
            Next i
            If pSampleInfo.ORDCNT > 0 Then
                TestDat = Left(TestDat, Len(TestDat) - 1)       '"\" Cutting
            End If
            '-------------------

            sTmp = m_iFrameN & "O|1|" & Trim(pSampleInfo.ID) & "|" _
                    & "^" & Trim(Val(pSampleInfo.RACK)) & "^" & Trim(Val(pSampleInfo.POS)) & "|" _
                    & TestDat & "|R||||||" _
                    & sActionCd & "||||Serum||||||||||" & sReportType & vbCr & Chr(3)
            
            m_iSendPhase = 4

        Case 4      'T
            sTmp = m_iFrameN & "L|1|N" & vbCr & Chr(3)
            m_iSendPhase = 5

        Case 5      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1: m_iPhase = 1: m_iSendPhase = 1
            sState = ""

            If m_bUseBarcode = True Then
                'Barcode Mode�� ��� ���ۿϷ� �̺�Ʈ �߻�
                RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
            Else
                'BarCode Mode�� �ƴ� ��� ���� ���� ��ȸ
                RaiseEvent RequestNextOrder
            End If

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
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
        .KIND = ""
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
