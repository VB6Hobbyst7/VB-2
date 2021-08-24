VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl LH750 
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
Attribute VB_Name = "LH750"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�⺻ �Ӽ� ��:
Const m_def_p_sPID = "0"
Const m_def_p_sData = "0"
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
Dim m_p_sPID As String
Dim m_p_sData As String
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
Event SendOrderOK(sID$, sRetCd$)
Event RequestCurOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
'Event SendOrderOK(sID$)
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

Private Sub PhaseCfg_Protocol_LH750()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1      ''SYN, Blockcount ���(datablock ������ ������)
                Select Case Asc(wkDat)
                    Case 22     'SYN�� �ش�
                        msComm.Output = Chr(22)     'SYN
                        RcvBuffer = RcvBuffer & wkDat   'wkBuf
                        m_iPhase = 1

                    Case Else   'blockcount-> 2 chars�� �ش�
                        msComm.Output = Chr(6)      'ACK
                        m_iPhase = 2
                End Select

            Case 2      ''datablock ���� ����(one datablock�� ���� ETX ��������)
                Select Case Asc(wkDat)
                    Case 3      'ETX
                        msComm.Output = Chr(6)      'ACK
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 3

                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                End Select

            Case 3      ''������ ������ or �ٸ� datablock ������ �������� �Ǵ��Ͽ� ���� ��ȯ
                Select Case Asc(wkDat)
                    Case 22     'SYN, �� ������ ��
                        msComm.Output = Chr(6)      'ACK
                        RcvBuffer = RcvBuffer & wkDat

                        Call DataEdit_LH750

                        RcvBuffer = ""
                        m_iPhase = 1

                    Case 2  'STX, �� �ٸ� datablock ���� ����
                        'ix1 = ix1 + 3   'manual dataformat ���� p.11
                        ''�ϴ��� �� ���۹ް� edit_data���� �ɷ����� ������ �ٲ�. 1998-05-21 ������
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2

                End Select

            '--- ORDER ���� ����
            Case 4
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        Call SendOrder_LH750
                        m_iPhase = 5

                    Case 22     'SYN
                        msComm.Output = Chr(22)
                        m_iPhase = 1

                End Select

            Case 5
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        Call SendOrder_LH750
                        m_iPhase = 6

                    Case Else   'NAK -> RECEIVER ABORT
                        m_iPhase = 1

                End Select

            Case 6
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        msComm.Output = Chr(5)      'ENQ
                        m_iPhase = 7

                    Case Else
                        m_iPhase = 1

                End Select

            Case 7
                Select Case Asc(wkDat)
                    Case 6      'ACK

                    Case 16     'DLE
                        m_iPhase = 8

                End Select

            Case 8      'RETURN CODE ���
                Select Case Asc(wkDat)
                    Case 65, 66, 67, 68, 69, 70     'A, B, C, D, E, F
                        m_iPhase = 1
                        m_iSendPhase = 1
                        RaiseEvent SendOrderOK(pSampleInfo.ID, wkDat)

                    Case Else
                        m_iPhase = 1
                        m_iSendPhase = 1
                        RaiseEvent SendOrderOK("", wkDat)

                End Select

        End Select
    Next ix1

End Sub
Private Sub SendOrder_LH750()
    On Error GoTo ErrRtn
    
    Dim sSend   As String * 256
    Dim sSendStr    As String
    Dim sChkSum As String
    
    Select Case m_iSendPhase
        Case 1
            RaiseEvent RequestCurOrder
            
            Call Get_OrderString
            
            If pSampleInfo.ORDCNT = 0 Then
                Exit Sub
            End If
                
            msComm.Output = "01"
            m_iSendPhase = m_iSendPhase + 1
            
            Exit Sub

        Case 2
            sSend = pSampleInfo.KIND
                
            sChkSum = ChkSum_LH750(sSend)
            
            sSendStr = Chr(2) & Format(1, "00") & sSend & sChkSum & Chr(3)
            
            msComm.Output = sSendStr
            
            If sTestMode = "77" Then
                RaiseEvent PrintSendLog(sSendStr)
            End If

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder ���� - " & Err.Description)
    End If
End Sub

Private Function ChkSum_LH750(ByVal sPara As String) As String

   Dim crcmsb%, xt  As Single, tmpx%, tmpx2 As Integer
   Dim crclsb%, cr$, x%, test$, i%
   
   crcmsb = 255
   crclsb = 255
   
   For i = 1 To Len(sPara)
      x = Asc(Mid$(sPara, i, 1)) Xor crcmsb
      tmpx = Int(x / 16)
      x = x Xor tmpx
      xt = x
      tmpx = (xt * 16) Mod 256
      tmpx2 = Int(x / 8)
      crcmsb = (crclsb Xor tmpx2 Xor tmpx) Mod 256
      tmpx = (xt * 32) Mod 256
      crclsb = (x Xor tmpx) Mod 256
   Next i

   crclsb = crclsb Xor 255
   crcmsb = crcmsb Xor 255
   ChkSum_LH750 = Right("0" & Hex$(crcmsb), 2) & Right$("0" & Hex$(crclsb), 2)
    
End Function
'
'   LH-750 Format
'
Private Sub DataEdit_LH750()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    
    Dim tmpIFCd$, tmpRst$, tmpFlag$
    Dim sTmp$, sTmp1$, sTmp2$, sTotIFCd$
    Dim sIFCd() As String
    Dim iPos%, iPos2%, ii%
    
        
    ''Data�� Edit�ϱ� ���ϵ���
    ''<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>����
    ''[DATA Block]�κи� �����ϰ� msRcvBuffer �����Ѵ�.
    Do
        iPos = InStr(1, RcvBuffer, Chr(2))
        
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
        If iPos = 0 Then
            Exit Do
        End If
        
        sTmp1 = Left$(RcvBuffer, iPos - 1)
        sTmp2 = Mid$(RcvBuffer, iPos + 3)
        
        RcvBuffer = ""
        RcvBuffer = sTmp1 & sTmp2
    Loop While iPos <> 0
    
    Do
        iPos = InStr(1, RcvBuffer, Chr(3))
        
        '<STX>[MS Char][NS Char][DATA Block][MS Char][NS Char][MS Char][NS Char]<ETX>
        If iPos = 0 Then
            Exit Do
        End If
        
        sTmp1 = Left$(RcvBuffer, iPos - 5)
        sTmp2 = Mid$(RcvBuffer, iPos + 1)
        
        RcvBuffer = ""
        RcvBuffer = sTmp1 & sTmp2
    Loop While iPos <> 0
    
    '�������ü �ʱ�ȭ
    Call Init_pResultInfo
    
    
    '�۾���ȣ ���ϱ�
    iPos = InStr(RcvBuffer, "ID1")
    If iPos > 0 Then
        sTmp2 = Mid(RcvBuffer, iPos + 4, 16)
        ii = InStr(1, sTmp2, vbCr)
        If ii <> 0 Then
            sTmp2 = Mid(sTmp2, 1, ii - 1)
        End If
        tmpBarCd = sTmp2
    End If
    
    iPos = InStr(RcvBuffer, "CASSPOS")
    If iPos > 0 Then
        sTmp1 = Mid(RcvBuffer, iPos + 9, 6)
            
        tmpRack = Left(sTmp1, 4)
        tmpPos = Right(sTmp1, 2)
    End If
    
'    iPos = InStr(RcvBuffer, "SEQUENCE")
'    If iPos > 0 Then
'        tmpSeqNo = Trim(Mid(RcvBuffer, iPos + 8, 7))
'    End If
    
       
    '��񿡼� �˻��� �� �ִ� ��� �׸� ����
    sTotIFCd = "WBC|RBC|HGB|HCT|MCV|MCH|MCHC|RDW|PLT|PCT|MPV|PDW|" _
            & "LY#|MO#|NE#|EO#|BA#|NRBC#|LY%|MO%|NE%|EO%|BA%|NRBC%|" _
            & "RET%|RET#|MRV|MSCV|IRF|HLR%|HLR#"
    sIFCd() = Split(sTotIFCd, Chr(124))
    
    '�˻��, �˻����� ���
    For ii = 0 To UBound(sIFCd())
        iPos = InStr(RcvBuffer, Trim(sIFCd(ii)))
        
        If iPos > 0 Then
            sTmp = Trim(Mid(RcvBuffer, iPos + 4, 3))
            If sTmp = "Pop" Then
                iPos = 0
            ElseIf sTmp = "IS" Then
                iPos = InStr(iPos + 4, RcvBuffer, Trim(sIFCd(ii)))
            End If
        End If
        
        If iPos > 0 Then
            iPos2 = InStr(iPos, RcvBuffer, Chr(13))
            sTmp = Trim(Mid(RcvBuffer, iPos, iPos2 - iPos))
            
            tmpIFCd = Trim(sIFCd(ii))
            
            sTmp = Trim(Mid(sTmp, Len(tmpIFCd) + 1))
            
            iPos2 = InStr(sTmp, " ")
            If iPos2 > 0 Then
                tmpRst = Trim(Mid(sTmp, 1, iPos2))
                tmpFlag = Trim(Mid(sTmp, iPos2))
            Else
                tmpRst = Trim(sTmp)
                tmpFlag = ""
            End If
            
'            tmpRst = Trim(Mid(sTmp, 5, 6))
'            tmpFlag = Trim(Mid(sTmp, 10))
        
            '--- ����� �ڸ����� ������ ���� Flag�� ǥ�õǴ� ��� ó��...(2000/11/14 yk)
            iPos = InStr(1, tmpRst, " ")
            If iPos <> 0 Then
                tmpRst = Trim(Mid(tmpRst, 1, iPos - 1))
            End If
            
            'STKS�� ���׷��̵� �� �� MCHC����� �߶󳻸� SOH�� �ڿ� �ٴ� ����
            If IsNumeric(Right$(tmpRst, 1)) = True Then
            Else
                tmpRst = Left$(tmpRst, Len(tmpRst) - 1)
            End If
                
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .UNIT = .UNIT & Chr(124)
            End With
        End If
    Next ii
    
    '��� ó��
    With pResultInfo
        If .RSTCNT > 0 Then
            .ID = tmpBarCd
            .SEQNO = tmpSeqNo
            .RACK = tmpRack
            .POS = tmpPos
            
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit ���� �߻� - " & Err.Description)
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
        Case "LH750"
            Call PhaseCfg_Protocol_LH750
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub

Private Sub Get_OrderString()

    Dim sOrder  As String
        
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        Exit Sub
    End If
    
    With pSampleInfo
        .ID = m_p_sID
        .ORDCNT = m_p_iOrdCnt
        
        sOrder = sOrder & Chr(1) & Right("00" & Hex(3), 2) & Chr(13) & Chr(10)
        
        sOrder = sOrder & "WLAD" & Chr(13) & Chr(10)
        sOrder = sOrder & "ID " & Trim(m_p_sSeq) & vbCr & vbLf
        sOrder = sOrder & "TS " & Trim(m_p_sTIFCd) & "," & Trim(m_p_sID) & "," & Chr(13) & Chr(10)
        
        .KIND = sOrder
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
    m_p_sPID = PropBag.ReadProperty("p_sPID", m_def_p_sPID)
    m_p_sData = PropBag.ReadProperty("p_sData", m_def_p_sData)
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
    Call PropBag.WriteProperty("p_sPID", m_p_sPID, m_def_p_sPID)
    Call PropBag.WriteProperty("p_sData", m_p_sData, m_def_p_sData)
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
    m_p_sPID = m_def_p_sPID
    m_p_sData = m_def_p_sData
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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sPID() As String
    p_sPID = m_p_sPID
End Property

Public Property Let p_sPID(ByVal New_p_sPID As String)
    m_p_sPID = New_p_sPID
    PropertyChanged "p_sPID"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sData() As String
    p_sData = m_p_sData
End Property

Public Property Let p_sData(ByVal New_p_sData As String)
    m_p_sData = New_p_sData
    PropertyChanged "p_sData"
End Property

