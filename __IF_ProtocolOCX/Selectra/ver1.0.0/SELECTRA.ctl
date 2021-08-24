VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl SELECTRA 
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
Attribute VB_Name = "SELECTRA"
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
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
'Event SendOrderOK(sID$, sRack$, sPos$)
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

Private Sub PhaseCfg_Protocol_SELECTRA()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'stx
                RcvBuffer = ""
                
            Case 3      'etx
                If UCase(m_EqName) = "SELECTRAXL" Then
                    Call DataEditResponse_SELECTRAXL
                    
                ElseIf UCase(m_EqName) = "SELECTRA_B" Then
                    Call DataEditResponse_SELECTRA_B
                    
                ElseIf UCase(m_EqName) = "SELECTRA_SPLIT" Then
                    Call DataEditResponse_SELECTRA_SPLIT
                
                Else
                    Call DataEditResponse_SELECTRA
                End If
                RcvBuffer = ""
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
        
End Sub
Private Sub DataEditResponse_SELECTRA()
    On Error GoTo ErrRtn
    
    Dim iPos        As Integer
    Dim ii          As Integer
    Dim sResType    As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$
    Dim tmpCnt$, tmpRCnt%
    Dim sRetVal     As String
    
    Dim tmpField()  As String
    Dim tmpData()   As String
    
    ' Get Test Name
    iPos = InStr(RcvBuffer, "{")
    
    ' Get Result
    iPos = iPos + 1
    Select Case Mid$(RcvBuffer, iPos, 1)
        Case "R", "r"
            iPos = iPos + 7
            sResType = Mid(RcvBuffer, iPos, 1)
            If Trim(sResType) <> "N" Then
                Exit Sub
            End If
            
            iPos = iPos + 2     '10
            tmpBarCd = Trim$(Mid$(RcvBuffer, iPos, 12))     'Sample No
           
            ' Get Test Count
            iPos = iPos + 48
            tmpCnt = Trim$(Mid$(RcvBuffer, iPos, 2))
            
            If Not IsNumeric(tmpCnt) Then
                RaiseEvent DispMsg("�̻� ��� ����...")
                Exit Sub
            End If
            tmpRCnt = Val(tmpCnt)
            
            Call Init_pResultInfo
            
            '������� ����ü�� ����
            With pResultInfo
                .ID = tmpBarCd
            End With
            
            'Get TestNm & Result
            iPos = iPos + 3
            For ii = 1 To tmpRCnt
                tmpIFCd = Trim(Mid$(RcvBuffer, (ii - 1) * 40 + iPos, 4))
                tmpRst = Trim(Mid$(RcvBuffer, (ii - 1) * 40 + iPos + 5, 7))
                tmpUnit = Trim(Mid(RcvBuffer, (ii - 1) * 40 + iPos + 33, 6))
                
                '����� ����
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            Next ii
    
            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

        Case "q"
            tmpField() = Split(Mid$(RcvBuffer, iPos), ";")
            
            sRetVal = Trim(tmpField(2))
            tmpBarCd = Trim(tmpField(3))
            
            If sRetVal = "0" Then   '���������� ��������
                RaiseEvent SendOrderOK(tmpBarCd, "", "", "")
                RaiseEvent RequestNextOrder
            End If
            
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEditResponse_SELECTRA_B()
    On Error GoTo ErrRtn
    
    Dim iPos        As Integer
    Dim ii          As Integer
    Dim sResType    As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$
    Dim tmpCnt$, tmpRCnt%, iFieldPos%
    Dim sRetVal     As String
    
    Dim tmpField()  As String
    Dim tmpData()   As String
    
    ' Get Test Name
    iPos = InStr(RcvBuffer, "{")
    
    ' Get Result
    iPos = iPos + 1
    Select Case Mid$(RcvBuffer, iPos, 1)
        Case "R", "r"
            iPos = iPos + 7
            sResType = Mid(RcvBuffer, iPos, 1)
            If Trim(sResType) <> "N" Then
                Exit Sub
            End If
            
            iPos = iPos + 2     '10
            tmpBarCd = Trim$(Mid$(RcvBuffer, iPos, 12))     'Sample No
           
            ' Get Test Count
            iPos = iPos + 48
            tmpCnt = Trim$(Mid$(RcvBuffer, iPos, 2))
            
            If Not IsNumeric(tmpCnt) Then
                RaiseEvent DispMsg("�̻� ��� ����...")
                Exit Sub
            End If
            tmpRCnt = Val(tmpCnt)
            
            Call Init_pResultInfo
            
            '������� ����ü�� ����
            With pResultInfo
                .ID = tmpBarCd
            End With
            
            'Get TestNm & Result
            iPos = iPos + 3
            
            tmpField() = Split(Mid(RcvBuffer, iPos), ";")
            iFieldPos = 0
            
            For ii = 1 To tmpRCnt
                tmpIFCd = Trim(tmpField(iFieldPos))
                iFieldPos = iFieldPos + 1
                
                tmpRst = Trim(tmpField(iFieldPos))
                iFieldPos = iFieldPos + 2
                
                tmpUnit = Trim(tmpField(iFieldPos))
                iFieldPos = iFieldPos + 1
                
                '����� ����
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            Next ii
    
            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

        Case "q"
            tmpField() = Split(Mid$(RcvBuffer, iPos), ";")
            
            sRetVal = Trim(tmpField(2))
            tmpBarCd = Trim(tmpField(3))
            
            If sRetVal = "0" Then   '���������� ��������
                RaiseEvent SendOrderOK(tmpBarCd, "", "", "")
                RaiseEvent RequestNextOrder
            End If
            
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEditResponse_SELECTRAXL()
    On Error GoTo ErrRtn
    
    Dim iPos        As Integer
    Dim ii          As Integer
    Dim sResType    As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$
    Dim tmpCnt$, tmpRCnt%
    Dim sRetVal     As String
    
    Dim sType       As String
    Dim aData()     As String
    
    '���� ������
    '{R;      ;N;1001        ;                    ;           ;M;;20-MAY-2002; 8:10; 1;L;ALP ;;           ;------;                       ;   ;U/l   ;}
    
    If Left(RcvBuffer, 1) <> "{" Then Exit Sub
    
    sType = Mid(RcvBuffer, 2, 1)
    
    Select Case sType
        Case "R", "r"       'result
            aData() = Split(RcvBuffer, ";")
            
            sResType = Trim(aData(2))
            If Trim(sResType) <> "N" Then Exit Sub
            
            tmpBarCd = Trim(aData(3))
           
            ' Get Test Count
            tmpCnt = Trim(aData(10))
            If Not IsNumeric(tmpCnt) Then
                RaiseEvent DispMsg("�̻� ��� ����...")
                Exit Sub
            End If
            tmpRCnt = Val(tmpCnt)
            
            Call Init_pResultInfo
            
            '������� ����ü�� ����
            With pResultInfo
                .ID = tmpBarCd
            End With
            
            'Get TestNm & Result
            For ii = 11 To 11 + (tmpRCnt * 8) Step 8
                If ii >= (11 + (tmpRCnt * 8)) Then Exit For
            
                tmpIFCd = Trim(aData(ii + 1))
                tmpRst = Trim(aData(ii + 4))
                tmpFlag = Trim(aData(ii + 5))
                tmpUnit = Trim(aData(ii + 7))
                
                '����� ����
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .FLAG = .FLAG & tmpFlag & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                End With
            Next ii
    
            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

        Case "q"
'            aData() = Split(Mid$(RcvBuffer, iPos), ";")
'
'            sRetVal = Trim(aData(2))
'            tmpBarCd = Trim(aData(3))
'
'            If sRetVal = "0" Then   '���������� ��������
'                RaiseEvent SendOrderOK(tmpBarCd, "", "", "")
'                RaiseEvent RequestNextOrder
'            End If
            
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEditResponse_SELECTRA_SPLIT()
    On Error GoTo ErrRtn
    
    Dim iPos        As Integer
    Dim ii          As Integer
    Dim sResType    As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$
    Dim tmpCnt$, tmpRCnt%
    Dim sRetVal     As String
    
    Dim sType       As String
    Dim aData()     As String
    
    '���� ������
    '{R;    ;N;0811190105  ;                    ;          ;M; 8;GOT ;37     ;                N     ;U/l   ;GPT ;19     ;                      ;U/l   ;GGTS;71     ;                N     ;U/l   ;GLUC;145.7  ;                N     ;mg/dl ;CHOL;185    ;                      ;mg/dl ;TRIG;97     ;                      ;mg/dl ;HDL ;43.9   ;                      ;mg/dl ;LDL ;128.8  ;                      ;mg/dl ;}
    
    If Left(RcvBuffer, 1) <> "{" Then Exit Sub
    
    sType = Mid(RcvBuffer, 2, 1)
    
    Select Case sType
        Case "R", "r"       'result
            aData() = Split(RcvBuffer, ";")
            
            sResType = Trim(aData(2))
            If Trim(sResType) <> "N" Then Exit Sub
            
            tmpBarCd = Trim(aData(3))
           
            ' Get Test Count
            tmpCnt = Trim(aData(7))
            If Not IsNumeric(tmpCnt) Then
                RaiseEvent DispMsg("�̻� ��� ����...")
                Exit Sub
            End If
            tmpRCnt = Val(tmpCnt)
            
            Call Init_pResultInfo
            
            '������� ����ü�� ����
            With pResultInfo
                .ID = tmpBarCd
            End With
            
            'Get TestNm & Result
            For ii = 7 To 6 + (tmpRCnt * 4)
                ''If ii >= (11 + (tmpRCnt * 8)) Then Exit For
            
                tmpIFCd = Trim(aData(ii + 1))
                tmpRst = Trim(aData(ii + 2))
                tmpFlag = Trim(aData(ii + 3))
                tmpUnit = Trim(aData(ii + 4))
                
                '����� ����
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .FLAG = .FLAG & tmpFlag & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                End With
                
                ii = ii + 3
            Next ii
    
            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

        Case "q"
'            aData() = Split(Mid$(RcvBuffer, iPos), ";")
'
'            sRetVal = Trim(aData(2))
'            tmpBarCd = Trim(aData(3))
'
'            If sRetVal = "0" Then   '���������� ��������
'                RaiseEvent SendOrderOK(tmpBarCd, "", "", "")
'                RaiseEvent RequestNextOrder
'            End If
            
    End Select
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
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
        Case "SELECTRA", "SELECTRAXL", "SELECTRA_B", "SELECTRA_SPLIT"
            '���ڵ� ���
            Call PhaseCfg_Protocol_SELECTRA
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub
Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=14
Public Function SendOrder_Selectra() As Variant
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim ii      As Integer
    Dim sOrdStr As String

    
    '--- �˻��׸� ����
    Call Get_OrderString
    
    
    sSendBuff = ""
    sOrdStr = String(26, "0")
    
    '�˻��׸� Order�ڵ� �߰�
    For ii = 1 To pSampleInfo.ORDCNT
        If Trim$(pSampleInfo.IFCD(ii)) <> "" Then
            Mid$(sOrdStr, Val(pSampleInfo.IFCD(ii)), 1) = "1"
        End If
    Next ii
    
    sSendBuff = "{Q;" & Left(Trim(pSampleInfo.ID) & Space(12), 12) _
                & ";N;" & Space(20) & ";" & Space(11) & ";M;" & sOrdStr & ";}"
                
    msComm.Output = Chr(2) & sSendBuff & Chr(3)
    
    If m_sTestMode = 77 Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(3))
    End If
            
    '���۵� ������ �ִ� ��� ȭ��ǥ��
    If pSampleInfo.ORDCNT > 0 Then
'        RaiseEvent SendOrderOK(pSampleInfo.ID, "", "", "")
    Else
        '��ȸ�� ������ ���� ��� ȯ������ ����ü �ʱ�ȭ
        Call Init_pResultInfo

'        RaiseEvent SendOrderOK("", "", "", "")
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder ���� - " & Err.Description)
    End If
End Function

