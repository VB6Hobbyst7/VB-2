VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl CHORUS 
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   LockControls    =   -1  'True
   ScaleHeight     =   3075
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
      InputMode       =   1
   End
End
Attribute VB_Name = "CHORUS"
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
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


'===== User Define
'�������̽����� ���
Dim RcvBuffer   As String
Dim wkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String
Dim miETB As Integer

'����ü ����
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'��Ÿ
Dim iSpaceCnt   As Integer

Dim sSndPacket As String
Dim maSndPacket() As String
Dim miPacketCnt As Integer
Dim msSndState As String

Dim bData As Boolean
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
        Case "CHORUS"
                Call PhaseCfg_Protocol_CHORUS
                
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_CHORUS_temp()
'
'    Dim wkDat   As String
'    Dim ix1     As Integer
'    Dim iLen    As Integer
'
''    If Asc(Left(wkBuf, 1)) = 2 Then
''        If Len(wkBuf) < 2 Then Exit Sub
''
''        iLen = Asc(Mid$(wkBuf, 2, 1))
''
''        RcvBuffer = Mid(wkBuf, 3, iLen)
''
''        Text2 = Text2 & RcvBuffer & vbCrLf
''
''        If iLen = 50 Then
''            Call DataEditResponse_CHORUS
''            RcvBuffer = ""
''        End If
''
''        msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
''        If sTestMode = "77" Then
''            RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
''        End If
''    End If
'
'            Else
'                Exit For
'            End If
'
'            If iLen > 0 Then
'                RcvBuffer = Mid(wkBuf, ix1 + 1, iLen)
'                ix1 = ix1 + iLen
'
'                Text2 = Text2 & RcvBuffer & vbCrLf
'
'                msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
'                If sTestMode = "77" Then
'                    RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
'                End If
'            End If
'
'            If iLen = 50 Then
'                Call DataEditResponse_CHORUS
'                RcvBuffer = ""
'            End If
'
'            If ix1 > Len(wkBuf) Then Exit For
'        End If
'
'        RcvBuffer = RcvBuffer & wkDat
'
'        'ENQ����
'        If InStr(RcvBuffer, "CD") > 0 Then
'            RcvBuffer = ""
'            'ACK����
'            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
'            If sTestMode = "77" Then
'                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
'            End If
'
'        'EOT����
'        ElseIf Len(RcvBuffer) = 4 And (Mid(RcvBuffer, 2, 2) = "?" Or Mid(RcvBuffer, 2, 2) = "") Then
'            RcvBuffer = ""
'            'ACK����
'            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
'            If sTestMode = "77" Then
'                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
'            End If
'
'            bData = False
'        '���DATA
'        ElseIf Len(RcvBuffer) = 51 Then
'            Call DataEditResponse_CHORUS
'            RcvBuffer = ""
'            'ACK����
'            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
'            If sTestMode = "77" Then
'                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
'            End If
'
'            bData = True
'        '����߿� chksum + STX �ϰ�� (�Ŵ����� �ȳѾ��)
'        ElseIf bData = True And wkDat = "" Then
'            RcvBuffer = ""
'        End If
'    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_CHORUS()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
'        If wkDat = "?" Then wkDat = " 0"
        
        RcvBuffer = RcvBuffer & wkDat
        
'        If Asc(wkDat) = 2 And Len(RcvBuffer) > 4 Then
        If Len(RcvBuffer) = 5 And (InStr(RcvBuffer, "") > 0 Or InStr(RcvBuffer, "") > 0) Then
            RcvBuffer = ""
            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
            bData = False
            
        ElseIf Len(RcvBuffer) = 54 And Left(RcvBuffer, 2) = "2" Then
            Call DataEditResponse_CHORUS
            RcvBuffer = ""
            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
            
            bData = True
        ElseIf bData = True And wkDat = "" Then
            RcvBuffer = ""
            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
            bData = False
        End If
    Next ix1
            
'            Text2 = Text2 & " :" & RcvBuffer & vbCrLf
'
'            RcvBuffer = ""
'            'ACK����
'            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
''            If sTestMode = "77" Then
''                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
''            End If
'
''        'EOT����
''        ElseIf (Len(RcvBuffer) = 4 Or Len(RcvBuffer) = 5) And (Mid(RcvBuffer, 2, 2) = "?" Or Mid(RcvBuffer, 2, 2) = "") Then
''            Text2 = Text2 & RcvBuffer & vbCrLf
''
''            RcvBuffer = ""
''            'ACK����
''            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
''            If sTestMode = "77" Then
''                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
''            End If
''
''            bData = False
'        '���DATA
'        ElseIf Len(RcvBuffer) >= 51 Then
'            Text2 = Text2 & RcvBuffer & vbCrLf
'
'            Call DataEditResponse_CHORUS
'            RcvBuffer = ""
'            'ACK����
'            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
''            If sTestMode = "77" Then
''                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
''            End If
'
'            bData = True
'        '����߿� chksum + STX �ϰ�� (�Ŵ����� �ȳѾ��)
'        ElseIf bData = True And wkDat = "" Then
''            Text2 = Text2 & RcvBuffer & vbCrLf
'
'            RcvBuffer = ""
'        End If
'    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_CHORUS_back()
'
'    Dim wkDat   As String
'    Dim ix1     As Integer
'
'    For ix1 = 1 To Len(wkBuf)
'        wkDat = Mid$(wkBuf, ix1, 1)
'
'        RcvBuffer = RcvBuffer & wkDat
'
'        'ENQ����
'        If InStr(RcvBuffer, "CD") > 0 Then
'            RcvBuffer = ""
'            'ACK����
'            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
'            If sTestMode = "77" Then
'                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
'            End If
'
'        'EOT����
'        ElseIf Len(RcvBuffer) = 4 And (Mid(RcvBuffer, 2, 2) = "?" Or Mid(RcvBuffer, 2, 2) = "") Then
'            RcvBuffer = ""
'            'ACK����
'            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
'            If sTestMode = "77" Then
'                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
'            End If
'
'            bData = False
'        '���DATA
'        ElseIf Len(RcvBuffer) = 51 Then
'            Call DataEditResponse_CHORUS
'            RcvBuffer = ""
'            'ACK����
'            msComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)
'            If sTestMode = "77" Then
'                RaiseEvent PrintSendLog(Chr(2) & Chr(1) & Chr(4) & Chr(5))
'            End If
'
'            bData = True
'        '����߿� chksum + STX �ϰ�� (�Ŵ����� �ȳѾ��)
'        ElseIf bData = True And wkDat = "" Then
'            RcvBuffer = ""
'        End If
'    Next ix1
    
End Sub
' *=====================================================*
' *               Data���� & ����ó��                   *
' *=====================================================*
Private Sub DataEditResponse_CHORUS()
    On Error GoTo ErrRtn

    Dim sRecType As String   'Record Type
    Dim i        As Integer
    Dim iLoop    As Integer
    Dim tmpData()   As String
    Dim tmpInfo()   As String
    Dim tmpPacket()   As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpRst2$, tmpUnit$, tmpRef$, tmpFlag$
    
    RcvBuffer = Replace(RcvBuffer, Chr(0), Chr(32))
    
    Call Init_pResultInfo
            
    pSampleInfo.SEQNO = ""
    pSampleInfo.ID = Trim(Mid(RcvBuffer, 3, 19))
    pSampleInfo.RACK = ""
    pSampleInfo.POS = ""
    
    tmpIFCd = Trim(Mid(RcvBuffer, 22, 7))
    tmpRst = Trim(Mid(RcvBuffer, 30, 12))
    tmpRst2 = Trim(Mid(RcvBuffer, 29, 1))
    tmpUnit = Trim(Mid(RcvBuffer, 42, 10))
    tmpFlag = ""

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
        .RST2 = .RST2 & tmpRst2 & Chr(124)
        .UNIT = .UNIT & tmpUnit & Chr(124)
        .FLAG = .FLAG & tmpFlag & Chr(124)
    End With
            
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With

    Call Init_pResultInfo

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
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

Private Sub Command1_Click()
RcvBuffer = ""
End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            Dim Buffer As Variant
            Buffer = msComm.Input
            
'            wkBuf = msComm.Input
            wkBuf = StrConv(Buffer, vbUnicode)
            
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
    
    bData = False
    
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

