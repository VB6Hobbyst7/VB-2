VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl CX 
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
Attribute VB_Name = "CX"
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
Event RequestCurOrder(sID$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
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

'����ü ����
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'��Ÿ
Dim iSpaceCnt   As Integer

'For CX
Private pCXInfo As CXINFO




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
        Case "CX9"
            Call PhaseCfg_Protocol_CX9
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_CX9()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
             
        Select Case m_iPhase
            Case 1      '===== EOT ���
                Select Case Asc(wkDat)
                    Case 4      '----- EOT ����
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
                
            Case 2      '===== SOH ���
                Select Case Asc(wkDat)
                    Case 1      '----- SOH ����
                        msComm.Output = Chr(6)  'ACK �۽�
                        piAckEtx = 2
                        m_iPhase = 3
                        RcvBuffer = ""
                End Select
                
            Case 3      '===== LF ���
                Select Case Asc(wkDat)
                    Case 10     '----- LF ����
                        Select Case piAckEtx
                            Case 1
                                msComm.Output = Chr(6)  'ACK �۽�
                                piAckEtx = 2
                            Case 2
                                msComm.Output = Chr(3)  'ETX �۽�
                                piAckEtx = 1
                        End Select
                        m_iPhase = 4
                        
                    Case Else   '----- ���� ����
                        RcvBuffer = RcvBuffer & wkDat
                End Select
            
            Case 4      '===== EOT ���
                Select Case Asc(wkDat)
                    Case 4      '----- EOT ����
                        m_iPhase = 1
                        
                        ' Interface���� ���� ����Ÿ ����
                        Call DataEditResponse_CX9
                        
                        If pbContension = True Then
                            Sleep (500)
                            msComm.Output = Chr(4) & Chr(1) 'EOT+SOH �۽�
                            pbContension = False
                            m_iPhase = 5
                        End If
                        
                    Case Else   '----- ���� ����
                        RcvBuffer = RcvBuffer + wkDat
                        m_iPhase = 3
                End Select
            
            Case 5      '===== ACK ���
                Select Case Asc(wkDat)
                    Case 6      '----- ACK ����
                        Call SendOrder_CX9
                        m_iPhase = 6
                    Case 4      '----- EOT ����
                        pbContension = True
                        m_iPhase = 2
                    Case Else
                End Select
            
            Case 6      '===== ETX ���
                Select Case Asc(wkDat)
                    Case 3      '----- ETX ���� (ORDER�־��� ��츸 ����)
                        msComm.Output = Chr(4)      'EOT
                        m_iPhase = 1
                        
                    Case 21     '----- NAK ����
                        msComm.Output = psNakBuf    'NAK message ������
                        
                End Select
        
        End Select
    Next ix1
    
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
            Erase .IFCD     '2003/4/16
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
            
            If m_sTestMode = "77" Then
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
    
    '���� �ʱ�ȭ(CX �迭)
    piAckEtx = 1: pbContension = False
    
    
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

Private Sub DataEditResponse_CX9()
    On Error GoTo ErrRtn
    
    Dim iPos1%, iPos2%, ix1%
    Dim sSF     As String
    Dim sFC     As String
    Dim sRC     As String
    Dim sHQ     As String
    
    Dim tmpField()  As String
    Dim tmpData()   As String
    
    Dim tmpBarCd$, tmpRack$, tmpPos$, tmpSeqNo$, tmpSpcCd$, tmpTestType$
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    
    iPos1 = InStr(RcvBuffer, "[")
    iPos2 = InStr(RcvBuffer, "]")
    
    Do While (iPos1 > 0)
        sSF = Mid$(RcvBuffer, iPos1 + 4, 3)
        sFC = Mid$(RcvBuffer, iPos1 + 8, 2)
        sRC = Trim(Val(Trim(Mid$(RcvBuffer, iPos1 + 11, 2))))
        sHQ = Mid$(RcvBuffer, iPos1 + 11, iPos2 - iPos1 - 11)   'HOST QUERY
        
        Select Case sSF
            ' ===== Order Arr.
            Case "701"
                Select Case sFC
                    Case "02"
                        Select Case sRC
                            Case "0"
                                With pSampleInfo
                                    If .ORDCNT > 0 Then
                                        RaiseEvent SendOrderOK(.ID, .SEQNO, .RACK, .POS)
                                    Else
                                        '��ȸ�� ������ ���� ��� ȯ������ ����ü �ʱ�ȭ
                                        Call Init_pResultInfo
                                
                                        RaiseEvent SendOrderOK("", "", "", "")
                                    End If
                                End With

                                Call SendNextOrder
                                
                            Case "1"
                                RaiseEvent DispMsg("[701,02,01] SYNTAX ERROR")
                                
                            Case "2"
                                RaiseEvent DispMsg("[701,02,02] BUSY")
                                
                            Case "3"
                                RaiseEvent DispMsg("[701,02,03] INVALID CHEMISTRY REQUESTED")
                                
                            Case "4"
                                RaiseEvent DispMsg("[701,02,04] INVALID ORDAC REQUESTED")
                                
                            Case "5"
                                RaiseEvent DispMsg("[701,02,05] INVALID CHEMISTRY COMBINATION PROGRAMMED")
                                
                            Case "6"
                                RaiseEvent DispMsg("[701,02,06] CONTROL NOT CONFIGURED")
                                
                            Case "7"
                                RaiseEvent DispMsg("[701,02,07] CALIBRATOR SECTOR ONLY")
                                
                            Case "8"
                                RaiseEvent DispMsg("[701,02,08] MODE MISMATCH")
                                
                            Case "9"
                                RaiseEvent DispMsg("[701,02,09] CX7 ERROR")
                                
                            Case "10"
                                RaiseEvent DispMsg("[701,02,10] COMPLETED SAMPLE")
                                
                            Case "11"
                                RaiseEvent DispMsg("[701,02,11] Incompatible Fluid Types")
                                
                            Case "12"
                                RaiseEvent DispMsg("[701,02,12] Incompatible Test Types")
                                
                            Case "13"
                                RaiseEvent DispMsg("[701,02,13] Incompatible Patient Name")
                        End Select
                        
                    Case "04"   'Clear Sector/Sample IDs�� �����Ƿ� ���ǹ� (������ ���Ͽ�!)
                        Select Case sRC
                            Case "0"
                                m_iPhase = 5
                                msComm.Output = Chr(4) & Chr(1) 'EOT+SOH �۽�
                                
                            Case "1"
                                RaiseEvent DispMsg("[701,04,01] BAD MESSAGE")
                                
                            Case "2"
                                RaiseEvent DispMsg("[701,04,02] BUSY")
                                
                            Case "3"
                                RaiseEvent DispMsg("[701,04,03] CX7 ERROR")
                                
                            Case "4"
                                RaiseEvent DispMsg("[701,04,04] NOT EXISTENT ERROR")
                        End Select
                        
                    Case "06"   'HOST QUERY
                        With pCXInfo
                            .CURINDEX = 0
                            Erase .BARCODE()
                            
                            tmpField() = Split(sHQ, ",")
                        
                            For ix1 = 0 To UBound(tmpField())
                                If ix1 > 6 Then Exit For
                                
                                .BARCODE(ix1 + 1) = Trim(tmpField(ix1))
                            Next ix1
                        End With
                        
                        Call SendNextOrder
                        
                End Select
                
            ' ===== Result Arr.
            Case "702"
                Select Case sFC
                    Case "01"
                        Call Init_pResultInfo
                        
                        tmpField() = Split(RcvBuffer, ",")
                        tmpSeqNo = Trim(tmpField(5))
                        tmpRack = Trim(tmpField(7))
                        tmpPos = Trim(tmpField(8))
                        tmpTestType = Trim(tmpField(9))
                        tmpSpcCd = Trim(tmpField(11))
                        tmpBarCd = Trim(tmpField(12))
                        
                        With pResultInfo
                            .ID = tmpBarCd
                            .SEQNO = tmpSeqNo
                            .RACK = tmpRack
                            .POS = tmpPos
                            .KIND = tmpTestType
                        End With
                        
                    Case "03"
                        tmpField() = Split(RcvBuffer, ",")
                        
                        tmpIFCd = Trim(tmpField(10))
                        tmpFlag = Trim(tmpField(22))
                        If tmpFlag = "NA" Then
                            tmpFlag = ""
                        End If
                        tmpRst = Trim(tmpField(15))     '25))
                        
                        If (IsNumeric(tmpRst) = False) Then
                            tmpRst = ""
                        End If
                        If Left(tmpRst, 1) = "." Then
                            tmpRst = "0" & tmpRst
                        End If
                        
                        '������� ����ü�� ����
                        With pResultInfo
                            '����� ����
                            .RSTCNT = .RSTCNT + 1
                            .IFCD = .IFCD & tmpIFCd & Chr(124)
                            .RST1 = .RST1 & tmpRst & Chr(124)
                            .RST2 = .RST2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & tmpFlag & Chr(124)
                        End With
            
                    Case "05"
                        '����� ���/ȭ�� ǥ�� ó��...
                        With pResultInfo
                            If .RSTCNT > 0 Then
                                'SEQNO ��� TESTTYPE �Ѱ���(�����Ƿ����)
                                RaiseEvent AppendData(.ID, .KIND, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
'                                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                            End If
                        End With
            
                        Call Init_pResultInfo
                    
                End Select
        End Select
        
        iPos1 = InStr(2, RcvBuffer, "[")
        If iPos1 <> 0 Then
            RcvBuffer = Mid(RcvBuffer, iPos1)
            iPos1 = 1
        End If
    Loop
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit �����߻� - " & Err.Description)
    End If
End Sub

Private Sub SendNextOrder()
    On Error GoTo ErrRtn
    
    Dim sTmp$
    
    With pCXInfo
        .CURINDEX = .CURINDEX + 1
        If .CURINDEX > 7 Then
            .CURINDEX = 1
        End If
        pSampleInfo.ID = .BARCODE(.CURINDEX)
    End With
   
    '----- �˻��׸� ��ȸ
    RaiseEvent RequestCurOrder(pSampleInfo.ID)

    Call Get_OrderString
    
    If pSampleInfo.ORDCNT > 0 Then
        m_iPhase = 5
        msComm.Output = Chr(4) & Chr(1)
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order_Next �����߻�" & "(" & Err.Description & ")")
    End If
End Sub
Private Sub SendOrder_CX9()
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim sTestCd     As String
    Dim iCnt%, iChk%, ix1
    Dim sChksums As String   '���Ǿ��� checksum
    
    '----- Order �����
    sTestCd = ""
    For iCnt = 1 To pSampleInfo.ORDCNT
        sTestCd = sTestCd & "," & Trim$(pSampleInfo.IFCD(iCnt)) & " ,0"
    Next iCnt
   
    '----- ������ Order Format �����
    sSendBuff = "[ 0,701,01"
    sSendBuff = sSendBuff & ",00"   'RACK
    sSendBuff = sSendBuff & ",00"   'POS
    sSendBuff = sSendBuff & ",0,RO"
    sSendBuff = sSendBuff & ",SE"   'Serum:SE, Urine:UR
    sSendBuff = sSendBuff & "," & Left(Trim(Trim(pSampleInfo.ID)) & Space(11), 11)  '���ڵ��ȣ(left justified)
    sSendBuff = sSendBuff & "," & String(20, " ")
    sSendBuff = sSendBuff & "," & String(25, " ")
    sSendBuff = sSendBuff & "," & String(25, " ")
    sSendBuff = sSendBuff & "," & String(18, " ")
    sSendBuff = sSendBuff & "," & String(15, " ")
    sSendBuff = sSendBuff & "," & String(1, " ")
    sSendBuff = sSendBuff & "," & String(12, " ")
    sSendBuff = sSendBuff & "," & String(18, " ")
    sSendBuff = sSendBuff & "," & String(6, " ")
    sSendBuff = sSendBuff & "," & String(4, " ")
    sSendBuff = sSendBuff & "," & String(20, " ")
    sSendBuff = sSendBuff & ",000"
    sSendBuff = sSendBuff & ",5"
    sSendBuff = sSendBuff & "," & String(6, " ")
    sSendBuff = sSendBuff & ",M"
    sSendBuff = sSendBuff & "," & String(25, " ")
    sSendBuff = sSendBuff & "," & String(7, " ")
    sSendBuff = sSendBuff & "," & String(4, " ")
    sSendBuff = sSendBuff & "," & String(4, " ")
    sSendBuff = sSendBuff & "," & String(6, " ")
    
    '----- ���ڵ��ȣ�� �ο��� Order ����
    sSendBuff = sSendBuff & "," & Format(Trim(pSampleInfo.ORDCNT), "000")
    
    '----- Order ���̱�
    sSendBuff = sSendBuff & sTestCd & "]"
    
    '----- checksum ���
    iChk = 0
    For ix1 = 1 To Len(sSendBuff)
        iChk = iChk + Asc(Mid$(sSendBuff, ix1, 1))
    Next ix1
    iChk = iChk Mod 256
    iChk = 256 - iChk
    
    sChksums = Right$("0" & Hex$(iChk), 2)
    
    sSendBuff = sSendBuff & sChksums & Chr(13) & Chr(10)
    psNakBuf = sSendBuff
    
    msComm.Output = sSendBuff
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendBuff)
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order ���۽� �����߻� - " & Err.Description)
    End If
End Sub


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

