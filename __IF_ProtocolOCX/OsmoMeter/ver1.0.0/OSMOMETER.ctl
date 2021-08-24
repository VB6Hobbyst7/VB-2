VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl OSMOMETER 
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
Attribute VB_Name = "OSMOMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�⺻ �Ӽ� ��:
'Const m_def_iSMPLen = 0
'Const m_def_iBCLen = 0
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
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,CommPort
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "��� ��Ʈ ��ȣ�� ��ȯ�ϰų� �����մϴ�."
    CommPort = MSComm.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    MSComm.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Private Function fConvertErrMsg(ByVal sCd As String) As String
   
    Select Case UCase(m_EqName)
        Case "OCSENSOR"
            Select Case sCd
                Case "10": fConvertErrMsg = "��ü���ڵ� �ǵ� ����"
                Case "20": fConvertErrMsg = "��ü���ڵ� �ߺ�����(���ϳ� üũ)"
                Case "01": fConvertErrMsg = "��ü�� ����/���� ����"
                Case "02": fConvertErrMsg = "�þ���� ����"
                Case "03": fConvertErrMsg = "RBC"
                Case "04": fConvertErrMsg = "PRC"
                Case "05": fConvertErrMsg = "OR"
                Case "06": fConvertErrMsg = "UR"
                Case "07": fConvertErrMsg = "���� ���ֿ���"
                Case "08": fConvertErrMsg = "�þ� ���ֿ���"
                Case "09": fConvertErrMsg = "�ͼ� ����"
                Case "0A": fConvertErrMsg = "�þ� ��ũ ����"
                Case "0B": fConvertErrMsg = "�˷��� ����"
                Case "90": fConvertErrMsg = "�̰��� ����"
                Case Else: fConvertErrMsg = ""
            End Select
            
        Case "OCSENSOR2"
            Select Case sCd
                Case "R":  fConvertErrMsg = "�þ���� ����"
                Case "S":  fConvertErrMsg = "���þ���"
                Case "E":  fConvertErrMsg = "��ü���ڵ� �ǵ� ����"
                Case "D":  fConvertErrMsg = "Sample ����NG"
                Case "B":  fConvertErrMsg = "�þ� Blank NG"
                Case "T":  fConvertErrMsg = "���� Sensor Error"
                Case "#":  fConvertErrMsg = "Sample �ٷ�����"
                Case "V":  fConvertErrMsg = "Over Range"
                Case "U":  fConvertErrMsg = "Buffer ����"
                Case Else: fConvertErrMsg = ""
            End Select
            
        Case "OCSENSORMICRO"
            Select Case sCd
                Case "0A": fConvertErrMsg = "Cell blank error"
                Case "01": fConvertErrMsg = "No sample solution"
                Case "02": fConvertErrMsg = "No reagent solution"
                Case "04": fConvertErrMsg = "PRC(Pro-zone check)"
                Case "05": fConvertErrMsg = "Over range"
                Case "10": fConvertErrMsg = "Barcode error"
                Case "90": fConvertErrMsg = "No sample"
                Case "1A": fConvertErrMsg = "Barcode error + Cell blank error"
                Case "11": fConvertErrMsg = "Barcode error + No sample solution"
                Case "12": fConvertErrMsg = "Barcode error + No reagent solution"
                Case "14": fConvertErrMsg = "Barcode error + PRC"
                Case "15": fConvertErrMsg = "Barcode error + Over range"
                Case Else: fConvertErrMsg = ""
            End Select
    End Select
        
End Function

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
        Case "OCSENSOR", "OCSENSOR2", "OCSENSORMICRO"
            Call PhaseCfg_Protocol_OCSENSOR
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_OCSENSOR()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
       
        Select Case Asc(wkDat)
            Case 5      'ENQ
                MSComm.Output = Chr(6)
                
            Case 2      'STX
                RcvBuffer = ""
                
            Case 3      'ETX
                Select Case UCase(m_EqName)
                    Case "OCSENSOR"
                        Call DataEdit_OCSENSOR
                    Case "OCSENSOR2"
                        Call DataEdit_OCSENSOR2
                    Case "OCSENSORMICRO"
                        Call DataEdit_OCSENSORMicro
                End Select
                
                RcvBuffer = ""
                MSComm.Output = Chr(6)
                
            Case 4      'EOT
                RcvBuffer = ""
            
            Case Else
                RcvBuffer = RcvBuffer & wkDat

        End Select
    Next ix1
    
End Sub
Private Sub DataEdit_OCSENSOR()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd$, tmpSeq$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    Dim sDataFlag$
    Dim tmpErrCd$, sErrMsg$
    Dim tmpDate$, tmpTime$
    
'    If Len(RcvBuffer) < 69 Then
'        Exit Sub
'    End If
    
    sDataFlag = Left(RcvBuffer, 1)      'Q:QC, �� ��:�Ϲ� ����
       
    tmpDate = Mid(RcvBuffer, 5, 10)     '��������
    tmpTime = Mid(RcvBuffer, 15, 5)     '�����ð�
    
    tmpRack = Trim(Mid(RcvBuffer, 20, 3))
    tmpPos = Trim(Mid(RcvBuffer, 23, 2))
    tmpSeq = Trim(Mid(RcvBuffer, 26, 4))
    tmpBarCd = Trim(Mid(RcvBuffer, 31, 15))
    
    '--- ������� ����
    tmpIFCd = "TEST"
    tmpRst1 = Trim(Mid(RcvBuffer, 56, 9))   '����ġ
    tmpRst2 = Trim(Mid(RcvBuffer, 65, 2))   '�������
    
    tmpErrCd = Trim(Mid(RcvBuffer, 68, 2))
    sErrMsg = fConvertErrMsg(tmpErrCd)
    '��� ���� �߻��� �����޽��� ȭ�� ǥ��
    If Trim(sErrMsg) <> "" Then
        RaiseEvent DispMsg(sErrMsg)
    End If
    
    
    '������� ����ü�� ����
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = tmpSeq
        .RACK = tmpRack
        .POS = tmpPos
        .RSTCNT = 1
        .IFCD = tmpIFCd & Chr(124)
        .RST1 = tmpRst1 & Chr(124)
        .RST2 = tmpRst2 & Chr(124)
        .UNIT = Chr(124)
        .FLAG = tmpErrCd & Chr(124)
    End With
    
    '����� ���/ȭ�� ǥ�� ó��...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With

    Call Init_pResultInfo

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit �����߻� - " & Err.Description)
    End If
End Sub
Private Sub DataEdit_OCSENSORMicro()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd$, tmpSeq$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    Dim sDataFlag$
    Dim tmpErrCd$, sErrMsg$
    Dim tmpDate$, tmpTime$
    
'    If Len(RcvBuffer) < 69 Then
'        Exit Sub
'    End If
       
    tmpDate = Mid(RcvBuffer, 5, 10)     '��������
    tmpTime = Mid(RcvBuffer, 15, 5)     '�����ð�
    
    tmpSeq = Trim(Mid(RcvBuffer, 26, 4))
    tmpBarCd = Trim(Mid(RcvBuffer, 31, 15))
    
    '--- ������� ����
    tmpIFCd = "OB"
    tmpRst1 = Trim(Mid(RcvBuffer, 56, 9))   '����ġ
    tmpRst2 = Trim(Mid(RcvBuffer, 65, 2))   '�������
    
    tmpErrCd = Trim(Mid(RcvBuffer, 68, 2))
    sErrMsg = fConvertErrMsg(tmpErrCd)
    
    '��� ���� �߻��� �����޽��� ȭ�� ǥ��
    If Trim(sErrMsg) <> "" Then
        RaiseEvent DispMsg(sErrMsg)
    End If
    
    
    '������� ����ü�� ����
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = tmpSeq
        .RACK = ""
        .POS = ""
        .RSTCNT = 1
        .IFCD = tmpIFCd & Chr(124)
        .RST1 = tmpRst1 & Chr(124)
        .RST2 = tmpRst2 & Chr(124)
        .UNIT = Chr(124)
        .FLAG = tmpErrCd & Chr(124)
    End With
    
    '����� ���/ȭ�� ǥ�� ó��...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With

    Call Init_pResultInfo

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit �����߻� - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_OCSENSOR2()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd$, tmpSeq$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$
    Dim sDataFlag$
    Dim tmpErrCd$, sErrMsg$
    Dim tmpDate$, tmpTime$
    
'    If Len(RcvBuffer) < 69 Then
'        Exit Sub
'    End If
    
    sDataFlag = Left(RcvBuffer, 1)      'Q:QC, �� ��:�Ϲ� ����
    
    tmpDate = Mid(RcvBuffer, 2, 8)      '��������
    
    tmpRack = ""
    tmpPos = ""
    tmpSeq = Trim(Mid(RcvBuffer, 10, 4))
    '-- ��¥����(11, 1)
    '-- space(12, 1)
    
    If bUseBarcode = True Then      '2006/4/24 yk
        tmpBarCd = Trim(Mid(RcvBuffer, 15, 16))
    End If
    
    '--- ������� ����
    tmpIFCd = "OB"
    tmpRst1 = Trim(Mid(RcvBuffer, 30, 4))   '����ġ
    tmpRst2 = Trim(Mid(RcvBuffer, 34, 1))   '�������
    
    tmpErrCd = Trim(Mid(RcvBuffer, 35, 1))
    sErrMsg = fConvertErrMsg(tmpErrCd)
    
    '��� ���� �߻��� �����޽��� ȭ�� ǥ��
    If Trim(sErrMsg) <> "" Then
        RaiseEvent DispMsg(sErrMsg)
    End If
    
    
    '������� ����ü�� ����
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = tmpSeq
        .RACK = tmpRack
        .POS = tmpPos
        .RSTCNT = 1
        .IFCD = tmpIFCd & Chr(124)
        .RST1 = tmpRst1 & Chr(124)
        .RST2 = tmpRst2 & Chr(124)
        .UNIT = Chr(124)
        .FLAG = tmpErrCd & Chr(124)
    End With
    
    '����� ���/ȭ�� ǥ�� ó��...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With

    Call Init_pResultInfo

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit �����߻� - " & Err.Description)
    End If
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
    End With
    
End Sub
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,RTSEnable
Public Property Get RTSEnable() As Boolean
Attribute RTSEnable.VB_Description = "���� ��û ���� ���������� ���θ� �����մϴ�."
    RTSEnable = MSComm.RTSEnable
End Property

Public Property Let RTSEnable(ByVal New_RTSEnable As Boolean)
    MSComm.RTSEnable() = New_RTSEnable
    PropertyChanged "RTSEnable"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,RThreshold
Public Property Get RThreshold() As Integer
Attribute RThreshold.VB_Description = "������ ������ ���� ��ȯ�ϰų� �����մϴ�."
    RThreshold = MSComm.RThreshold
End Property

Public Property Let RThreshold(ByVal New_RThreshold As Integer)
    MSComm.RThreshold() = New_RThreshold
    PropertyChanged "RThreshold"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,Settings
Public Property Get Settings() As String
Attribute Settings.VB_Description = "���� �ӵ�, �и�Ƽ, ������ ��Ʈ, �ߴ� ��Ʈ �Ű� ������ ��ȯ�ϰų� �����մϴ�."
    Settings = MSComm.Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    MSComm.Settings() = New_Settings
    PropertyChanged "Settings"
End Property

Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case MSComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            wkBuf = MSComm.Input
            
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

    MSComm.CommPort = PropBag.ReadProperty("CommPort", 1)
    MSComm.RTSEnable = PropBag.ReadProperty("RTSEnable", False)
    MSComm.RThreshold = PropBag.ReadProperty("RThreshold", 0)
    MSComm.Settings = PropBag.ReadProperty("Settings", "9600,n,8,1")
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
'    m_iSMPLen = PropBag.ReadProperty("iSMPLen", m_def_iSMPLen)
'    m_iBCLen = PropBag.ReadProperty("iBCLen", m_def_iBCLen)
End Sub

'�Ӽ����� ����ҿ� ����մϴ�.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CommPort", MSComm.CommPort, 1)
    Call PropBag.WriteProperty("RTSEnable", MSComm.RTSEnable, False)
    Call PropBag.WriteProperty("RThreshold", MSComm.RThreshold, 0)
    Call PropBag.WriteProperty("Settings", MSComm.Settings, "9600,n,8,1")
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
'    Call PropBag.WriteProperty("iSMPLen", m_iSMPLen, m_def_iSMPLen)
'    Call PropBag.WriteProperty("iBCLen", m_iBCLen, m_def_iBCLen)
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
        MSComm.PortOpen = True
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
'    m_iSMPLen = m_def_iSMPLen
'    m_iBCLen = m_def_iBCLen
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
    MSComm.Output = Chr(iChr)
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr ���� - " & Err.Description)
    End If
End Function
'
'
''���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
''MemberInfo=7,0,0,0
'Public Property Get iBCLen() As Integer
'    iBCLen = m_iBCLen
'End Property
'
'Public Property Let iBCLen(ByVal New_iBCLen As Integer)
'    m_iBCLen = New_iBCLen
'    PropertyChanged "iBCLen"
'End Property
'
