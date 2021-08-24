VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl SYSMEX 
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
Attribute VB_Name = "SYSMEX"
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event SendOrderOK(sID$)
Event DispMsg(sMsg$)


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
        Case "SE9000"
            Call PhaseCfg_Protocol_SE9000
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub
'
'   SE-9000 ���ڵ� ���(���μ������������)
'
Private Sub DataEdit_SE9000()
    On Error GoTo ErrRtn

    Dim sBC     As String
    Dim sLC     As String

    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack As String
    Dim tmpPos  As String
    Dim ii      As Integer
    Dim tmpRst()    As String       '��� �ӽ� ����

    Dim sChk$, sChk2$


    sBC = Mid(RcvBuffer, 1, 2)
    sLC = Mid(RcvBuffer, 3, 1)

    Select Case sBC
        Case "R1"
            pSampleInfo.ID = ""

'            sChk = Mid$(RcvBuffer, 3, 1)
'            sChk2 = Mid$(RcvBuffer, 5, 1)
            sChk = Mid$(RcvBuffer, 5, 1)
            sChk2 = Mid$(RcvBuffer, 7, 1)

            If sChk <> "+" And sChk <> "-" And sChk2 <> "Q" Then
                Exit Sub
            End If

            tmpBarCd = Mid(RcvBuffer, 3, 13)
            tmpBarCd = Mid$(tmpBarCd, 3)
            pSampleInfo.ID = Trim(tmpBarCd)

            Call SendOrder_SE9000
            m_iPhase = 2

            Exit Sub

        Case "D1"
            Select Case sLC
                Case "U"
                    '������� �ʱ�ȭ
                    Call Init_pResultInfo

                    tmpRack = Mid(RcvBuffer, 10, 4)
                    tmpPos = Mid(RcvBuffer, 14, 2)
                    tmpBarCd = Trim(Mid(RcvBuffer, 22, 13))

                    'For ������� SE-9000
'                    sChk = Mid$(RcvBuffer, 22, 1)
'                    sChk2 = Mid$(RcvBuffer, 24, 1)
                    sChk = Mid$(RcvBuffer, 24, 1)
                    sChk2 = Mid$(RcvBuffer, 26, 1)

                    If sChk <> "+" And sChk <> "-" And sChk2 <> "Q" Then
                    Else
'                        tmpBarCd = Mid$(tmpBarCd, 3, 7) & "0" & Mid(tmpBarCd, 10)
                        tmpBarCd = Mid$(tmpBarCd, 3, 7) & Mid(tmpBarCd, 10)
                    End If
                    '--------------------

                    ReDim tmpRst(25) As String

                    'WBC
                    tmpRst(1) = Mid(RcvBuffer, 63, 5)
                    If tmpRst(1) = Space(5) Then
                        tmpRst(1) = "N"
                    Else
                        tmpRst(1) = Format(Val(Format(tmpRst(1), "@@@.@@")), "0.00")
                    End If

                    'RBC
                    tmpRst(2) = Mid(RcvBuffer, 69, 4)

                    If tmpRst(2) = Space(4) Then
                        tmpRst(2) = "N"
                    Else
                        tmpRst(2) = Format(Val(Format$(tmpRst(2), "@@.@@")), "0.00")
                    End If

                    'HGB, HCT, MCV, MCH, MCHC
                    For ii = 3 To 7
                        tmpRst(ii) = Mid(RcvBuffer, 74 + (ii - 3) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    'PLT
                    tmpRst(8) = Mid(RcvBuffer, 99, 4)

                    If tmpRst(8) = Space(4) Then
                        tmpRst(8) = "N"
                    Else
                        tmpRst(8) = Trim(Val(Format(tmpRst(8), "@@@@")))
                    End If

                    'LYMPH%, MONO%, NEUT%, EO%, BASO%
                    For ii = 9 To 13
                        tmpRst(ii) = Mid(RcvBuffer, 104 + (ii - 9) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    'LYMPH#, MONO#, NEUT#, EO#, BASO#
                    For ii = 14 To 18
                        tmpRst(ii) = Mid(RcvBuffer, 129 + (ii - 14) * 6, 5)

                        If tmpRst(ii) = Space(5) Then
                            tmpRst(ii) = "N"
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@@")), "0.00")
                        End If
                    Next ii

                    'RDW-CV(%), RDW-SD(fL), PDW(fL), MPV(fL), P-LCR
                    For ii = 19 To 23
                        tmpRst(ii) = Mid(RcvBuffer, 159 + (ii - 19) * 5, 4)

                        If tmpRst(ii) = Space(4) Then
                            tmpRst(ii) = ""
                        Else
                            tmpRst(ii) = Format(Val(Format(tmpRst(ii), "@@@.@")), "0.0")
                        End If
                    Next ii

                    '--- �Ʒ� �׸���� KX-21���� �˻����� ����(SE-9000���� �˻�)
                    'RET%
                    tmpRst(24) = Mid$(RcvBuffer, 189, 4)

                    If tmpRst(24) = "    " Then
                        tmpRst(24) = "N"
                    Else
                        tmpRst(24) = Format$(Val(Format$(tmpRst(24), "@@.@@")), "0.00")
                    End If
                    'RET#
                    tmpRst(25) = Mid$(RcvBuffer, 194, 4)

                    If tmpRst(25) = "    " Then
                        tmpRst(25) = "N"
                    Else
                        tmpRst(25) = Trim(Val("0." & tmpRst(25)))
                    End If
                    '�̻� ������ �Ÿ���
                    For ii = 24 To 25
                        If Val(tmpRst(ii)) = "0" Then
                            tmpRst(ii) = "-"
                        End If
                    Next ii
                    '--- �������...SE-9000������ �˻�...

                    With pResultInfo
                        .ID = tmpBarCd
                        .RACK = tmpRack
                        .POS = tmpPos

                        For ii = 1 To 25
                            .RSTCNT = .RSTCNT + 1

                            .IFCD = .IFCD & Trim(ii) & Chr(124)
                            .RST1 = .RST1 & tmpRst(ii) & Chr(124)
                            .RST2 = .RST2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & Chr(124)
                        Next ii
                    End With

                    '��� ó��
                    With pResultInfo
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                    End With

                    'Query ���� ����� ���� ���� ��츦 ����
                    If m_iOrderFlag = 1 Then
                        Call SendOrder_SE9000
                        m_iPhase = 2
                    Else
                        m_iPhase = 1
                    End If

                Case "C"

                Case Else

            End Select

        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit ���� �߻� - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_SE9000()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                    
                    Case 3      'ETX
                        Call DataEdit_SE9000
                        
                        msComm.Output = Chr(6)       'ACK

                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                End Select
                
            Case 2
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        RaiseEvent SendOrderOK(pSampleInfo.ID)
                        
                        'Order�� ������ �ٽ� �ʱ� ����
                        m_iPhase = 1
                        m_iOrderFlag = 0
                        
                    Case 21
                        Call SendOrder_SE9000
                    
                    Case Else
                        m_iPhase = 1
                        m_iOrderFlag = 0
                End Select
        End Select
    Next ix1
    
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

Private Sub SendOrder_SE9000()
    On Error GoTo ErrRtn
    
    Dim SendBuf$, sTestCd$, sBuf$
    Dim iCbcDef As Integer
    
    iCbcDef = 0

    '-----DIFF CHECK
    If Left$(pSampleInfo.ID, 1) = "+" Then
        iCbcDef = 2     ' CBC & DIFF
    Else
        iCbcDef = 1     ' CBC only
    End If
    
    Select Case iCbcDef
        Case 1      '===== CBC Only
'           sTestCd = "1111111100000000001111111111111"
'                      1234567890123456789012345678901 <31> ' for number test
            sTestCd = "1111111100000000001111100000000"   'CBC ONLY
        Case 2      '===== CBC + DIFF
            sTestCd = "1111111111111111111111111111111"
'                      1234567890123456789012345678901 <31> ' for number test
        Case Else   '===== CBC + DIFF
'           sTestCd = "1111111111111111111111111111111"
'                      1234567890123456789012345678901 <31> ' for number test
            sTestCd = "1111111100000000001111100000000"   'CBC ONLY
    End Select
    
    SendBuf = "S"
    SendBuf = SendBuf & "1"
    SendBuf = SendBuf & Format(Now, "YYYYMMDD")
    SendBuf = SendBuf & Right(String(13, "0") & pSampleInfo.ID, 13)
    SendBuf = SendBuf & Space$(4)
    SendBuf = SendBuf & Space$(2)
    SendBuf = SendBuf & "1"
    SendBuf = SendBuf & Right(String(13, "0") & pSampleInfo.ID, 13)
    SendBuf = SendBuf & Space$(25)
    SendBuf = SendBuf & "1"
    SendBuf = SendBuf & Space$(8)
    SendBuf = SendBuf & Space$(15)
    SendBuf = SendBuf & Space$(8)
    SendBuf = SendBuf & Space$(20)
    SendBuf = SendBuf & Space$(20)
    SendBuf = SendBuf & sTestCd
    
    msComm.Output = Chr(2) & SendBuf & Chr(3)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & SendBuf & Chr(3))
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder �����߻� - " & Err.Description)
    End If
End Sub

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
    
    m_iOrderFlag = 0
    
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

