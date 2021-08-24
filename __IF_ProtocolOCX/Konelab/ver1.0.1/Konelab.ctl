VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl Konelab 
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
Attribute VB_Name = "Konelab"
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
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
'Event SendOrderOK(sID$, sRack$, sPos$)
Event SendOrderOK(sID$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event DispMsg(sMsg$)
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


'===== User Define
'�������̽����� ���
Dim RcvBuffer       As String
Dim wkBuf           As String
Dim sState          As String
Dim sReqStatusCd    As String
Dim sRcvState       As String
Dim sSndstate       As String

'For ASTM
Dim sSndH As String
Dim sSndP As String
Dim sSndO As String
Dim sSndL As String

'for ����׸�
Dim sTcho   As String
Dim sTpro   As String
Dim sAlb    As String
Dim sTg     As String
Dim sBun    As String
Dim sCrea   As String
Dim sTbil   As String
Dim sDbil   As String
Dim sHdlc   As String
Dim sFe     As String
Dim sUibc   As String

'����ü ����
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'��Ÿ
Dim iSpaceCnt   As Integer

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
        Case "KONELAB"
            Call PhaseCfg_Protocol_Konelab
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_Konelab()
    Dim wkdat As String
    Dim ix1 As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkdat = Mid$(wkBuf, ix1, 1)
             
        Select Case m_iPhase
            'ENQ ��� ����
            Case 1
                Select Case Asc(wkdat)
                    'ENQ
                    Case 5
                        sRcvState = ""
                        sSndstate = ""
                        RcvBuffer = ""
                        
                        'ACK ����
                        msComm.Output = Chr(6)
                        If m_sTestMode = 77 Then
                           RaiseEvent PrintSendLog("<ACK>")
                        End If
                        
                        m_iPhase = 2
                    Case Else
                        sRcvState = ""
                        sSndstate = ""
                        m_iPhase = 1
                End Select
            
            'Packet ����, Packet �м�(Edit_Data)
            Case 2
                Select Case Asc(wkdat)
                    'STX
                    Case 2
                    'EOT
                    Case 4
                        Call Edit_Data
                    'ENQ
                    Case 5
                        'ACK ����
                        msComm.Output = Chr(6)
                        If m_sTestMode = 77 Then
                           RaiseEvent PrintSendLog("<ACK>")
                        End If
                    'LF
                    Case 10
                        RcvBuffer = RcvBuffer & wkdat
                        'ACK ����
                        msComm.Output = Chr(6)
                        If m_sTestMode = 77 Then
                            RaiseEvent PrintSendLog("<ACK>")
                        End If
                    'NAK
                    Case 21
                        'ENQ ����
                        msComm.Output = Chr(5)
                        If m_sTestMode = 77 Then
                            RaiseEvent PrintSendLog("<ENQ_NAK_P2>")
                        End If
                    Case Is < 0
                    
                    Case Else
                        RcvBuffer = RcvBuffer & wkdat
                End Select
                
            'SendOrder���� ENQ���� ACK ������
            Case 3
                Select Case Asc(wkdat)
                    'ACK
                    Case 6
                        If sSndstate = "E" Then
                            'Packet ����
                            msComm.Output = sSndH
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(sSndH)
                            End If
                            
                            sSndstate = "H"
                            m_iPhase = 3
                        ElseIf sSndstate = "H" Then
                            'Packet ����
                            msComm.Output = sSndP
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(sSndP)
                            End If
                            
                            sSndstate = "P"
                            m_iPhase = 3
                        ElseIf sSndstate = "P" Then
                            'Packet ����
                            msComm.Output = sSndO
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(sSndO)
                            End If
                            
                            sSndstate = "O"
                            m_iPhase = 3
                        ElseIf sSndstate = "O" Then
                            'Packet ����
                            msComm.Output = sSndL
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(sSndL)
                            End If
                            
                            sSndstate = "L"
                            m_iPhase = 3
                        ElseIf sSndstate = "L" Then
                            'EOT ����
                            msComm.Output = Chr(4)
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog("<EOT>")
                                RaiseEvent SendOrderOK(pSampleInfo.ID)
                            End If
                            
'                            Call SendOrder_Konelab
'                            Call DisplayOrderOK("AFTER_ORDER")
                            
                            sSndstate = ""
                            sSndH = "": sSndP = "": sSndO = "": sSndL = ""
                            m_iPhase = 1
                            
'                            cmdSendOrd_Click
                        End If
                    'NAK
                    Case 21
                        If sSndstate = "E" Then
                            msComm.Output = Chr(5)
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog("<ENQ_NAK_P3>")
                            End If

                            sSndstate = "E"
                            m_iPhase = 3
                        ElseIf sSndstate = "H" Then
                            msComm.Output = sSndH
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(sSndH)
                            End If

                            sSndstate = "H"
                            m_iPhase = 3
                        ElseIf sSndstate = "P" Then
                            msComm.Output = sSndP
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(sSndP)
                            End If

                            sSndstate = "P"
                            m_iPhase = 3
                        ElseIf sSndstate = "O" Then
                            msComm.Output = sSndO
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(sSndO)
                            End If

                            sSndstate = "O"
                            m_iPhase = 3
                        ElseIf sSndstate = "L" Then
                            msComm.Output = sSndL
                            If m_sTestMode = 77 Then
                                RaiseEvent PrintSendLog(sSndL)
                            End If

                            sSndstate = "L"
                            m_iPhase = 3
                        End If
                    'ENQ
                    Case 5
                        'ACK ����
                        msComm.Output = Chr(6)
                        If m_sTestMode = 77 Then
                            RaiseEvent PrintSendLog("<ACK>")
                        End If

                        RcvBuffer = ""
                        m_iPhase = 2
                End Select
        End Select
    Next
End Sub

' *=====================================================*
' *               Data���� & ����ó��                   *
' *=====================================================*
Private Sub Edit_Data()
    On Error GoTo ErrHandler
    

    Dim iErrCode     As Integer
    Dim sGeneralErrCode    As String


    Dim sJDate      As String
    Dim sJGbn       As String
    Dim sJNo        As String
    Dim sIFSpcCd    As String   '�������̽��� ��ü�ڵ�
    Dim sIFRstCd    As String   '�������̽��� �˻��׸��ڵ�
    Dim sRxData     As String
    Dim sSex        As String
    Dim sSampNo     As String
   
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpRst2$, tmpUnit$, tmpRef$, tmpFlag$

    Dim i           As Integer
    
    Dim sTmp As String
    Dim iPos As Integer
    Dim iETBpos As Integer
    Dim sRecType As String
    Dim sBuf      As String
    
    '### Rack Or Tray ��İ� Conflict ����
'    Call ProtectConflict("Y")
    
    sRxData = ""
    sRxData = RcvBuffer

   'sRecType �ʱ�ȭ
    sRecType = "S"
    
    Do While sRecType <> ""
        sTmp = GetByOneUserSymbol(sRxData, sRxData, vbLf)
        
        sRecType = Mid(sTmp, 2, 1)
        
        If sRecType = "" Then
           Exit Do
        End If
        
        If sRecType = "H" Then
        ElseIf sRecType = "Q" Then
            sBuf = Split(sTmp, Chr(124))(2)
            pSampleInfo.ID = Split(sBuf, "^")(1)
            pSampleInfo.RACK = Split(sBuf, "^")(2)
            pSampleInfo.POS = Split(sBuf, "^")(3)
            sSndstate = "Q"
'            sRcvState = "R"
            
        ElseIf sRecType = "P" Then
        ElseIf sRecType = "O" Then
        
            Call Init_pResultInfo
            
            sJDate = ""
            tmpBarCd = ""
            tmpRack = ""
            tmpPos = ""
            
            '����׸� �ʱ�ȭ
            sTcho = ""
            sTpro = ""
            sAlb = ""
            sTg = ""
            sBun = ""
            sCrea = ""
            sTbil = ""
            sDbil = ""
            sHdlc = ""
            sFe = ""
            sUibc = ""
            
            '3O|1|1|^1^1|^^^TotT3^1|||||||||||Serum||||||||||F
            'O1
            Call GetByOne(sTmp, sTmp)

            'O2
            Call GetByOne(sTmp, sTmp)

            '03
            sBuf = GetByOne(sTmp, sTmp)
            iPos = InStr(1, sBuf, "-")
            
            If InStr(1, sBuf, "^") = 0 And InStr(1, sBuf, "-") = 0 Then
                tmpBarCd = Trim(sBuf)
            Else
                If iPos = 0 Then
                    tmpBarCd = GetByOneUserSymbol(sBuf, sBuf, "^")
                Else
                    sJDate = GetByOneUserSymbol(sBuf, sBuf, "-")
                    tmpBarCd = CStr(CInt(Val(sBuf)))
                End If
            End If
        
            
            pResultInfo.ID = tmpBarCd
            
            'O4
            sBuf = GetByOne(sTmp, sTmp)
            Call GetByOneUserSymbol(sBuf, sBuf, "^")
            tmpRack = GetByOneUserSymbol(sBuf, sBuf, "^")
            tmpPos = sBuf
            
        ElseIf sRecType = "R" Then
            sRcvState = "R"
         
            'R1
            Call GetByOne(sTmp, sTmp)

            'R2
            Call GetByOne(sTmp, sTmp)

            'R3
            sBuf = GetByOne(sTmp, sTmp)
            sBuf = Mid(sBuf, 4)
            tmpIFCd = GetByOneUserSymbol(sBuf, sBuf, "^")
            
            'R4
            sBuf = GetByOne(sTmp, sTmp)
            tmpRst = sBuf
            
            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
            
            'for ����׸�
            If tmpIFCd = "TCHO" Then
                sTcho = tmpRst
            ElseIf tmpIFCd = "TP" Then
                sTpro = tmpRst
            ElseIf tmpIFCd = "ALB" Then
                sAlb = tmpRst
            ElseIf tmpIFCd = "TG" Then
                sTg = tmpRst
            ElseIf tmpIFCd = "BUN" Then
                sBun = tmpRst
            ElseIf tmpIFCd = "CREA" Then
                sCrea = tmpRst
            ElseIf tmpIFCd = "TBIL" Then
                sTbil = tmpRst
            ElseIf tmpIFCd = "DBIL" Then
                sDbil = tmpRst
            ElseIf tmpIFCd = "HDL" Then
                sHdlc = tmpRst
            ElseIf tmpIFCd = "FE" Then
                sFe = tmpRst
            ElseIf tmpIFCd = "UIBC" Then
                sUibc = tmpRst
            End If
            
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
            
'            With pResultInfo
'            If .RSTCNT > 0 Then
'                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
'            End If
'            End With
                            
        ElseIf sRecType = "L" Then
            
            If sSndstate = "Q" Then
                Call SendOrder_Konelab
            End If
            
            If sRcvState = "R" Then
                With pResultInfo
                    If .RSTCNT > 0 Then
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                    End If
                End With
            End If
                        
        Else
        End If
    Loop
        
    If sRcvState = "R" Then
        If (sSndstate = "E") Or (sSndstate = "H") Or (sSndstate = "P") Or (sSndstate = "O") Or (sSndstate = "L") Then
            
            'ENQ ����
            msComm.Output = Chr(5)
            
            If m_sTestMode = 77 Then
               RaiseEvent PrintSendLog("<ENQ>")
            End If
            
            sSndstate = "E"
            m_iPhase = 3
        Else
            m_iPhase = 1
        End If
    End If
    
    sRcvState = ""
    
    Exit Sub
    
ErrHandler:
    sRcvState = ""
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Sub SendOrder_Konelab()
    On Error GoTo ErrHandler
    
'ȯ���� Order ����
    Dim SendBuff As String
    Dim i%, j%, k%, iOrdCnt%
    Dim vIFCnt, vTmp
    Dim sTmp$, sTIFOrdCd$, sOrdList$, sIFSeq$, sBuf$, sTIFSeq$
    Dim objOrd As Object
    
    SendBuff = ""
    
    RaiseEvent RequestCurOrder(pSampleInfo.ID, "", pSampleInfo.RACK, pSampleInfo.POS)
            
    sBuf = ""

    For i = 1 To m_p_iOrdCnt
        sTmp = GetByOneUserSymbol(m_p_sTIFCd, m_p_sTIFCd, Chr(124))
    
        If sTmp = "" Then
        Else
            'sBuf = sBuf & "^^^" & sTmp & "^0.0" & "\"
            If Trim(sTmp) = "HbA1c" Then
                sBuf = sBuf & "^^^" & sTmp & "\" & "^^^" & "Hb" & "\" & "^^^" & "HbA" & "\"
            Else
                sBuf = sBuf & "^^^" & sTmp & "^0.0" & "\"
'                sBuf = sBuf & "^^^" & sTmp & "\"
            End If
        End If
    Next
    
    '�� ���� "\" ����
    sBuf = Mid(sBuf, 1, Len(sBuf) - 1)
    
    '1H|\^&|||Host LIS|||||ACCESS||P|1' + CR + ETX
    'sSndH = "1H|\^&|||Host LIS|||||ACCESS||P|1" & vbCr & Chr(3)
    '1H|\^&|||20^1^5.0.1|||||||P
    sSndH = "1H|\^&|||Host|||||||P" & vbCr & Chr(3)
    sSndH = Chr(2) & sSndH & ASTM_CheckSum(sSndH) & vbCr & vbLf
    
'''    gOrderTable.sSampID = gOrderTable.sJDate & "-" & Format(gOrderTable.sJNo, "0000")
    
    '2P|1|' + PATIENT ID + CR + ETX
    'sSndP = "2P|1|" & gOrderTable.sSampID & vbCr & Chr(3)
    '???
    'sSndP = "2P|1|" & pSampleInfo.ID & vbCr & Chr(3)
    sSndP = "2P|1|" & vbCr & Chr(3)
    sSndP = Chr(2) & sSndP & ASTM_CheckSum(sSndP) & vbCr & vbLf
    
    '3O|1|1234567890|^1^1|^^^TSH^0\^^^FT4^0|R||||||N||||Serum' + CR + ETX
    'sSndO = "3O|1|" & gOrderTable.sSampID & "|^" & CStr(Val(gOrderTable.sRack)) & "^" & CStr(Val(gOrderTable.sPos)) & "|" & sBuf & "|R||||||N||||Serum" & vbCr & Chr(3)
    '???
    'sSndO = "3O|1|" & gOrderTable.sSampID & "^0.0^" & CStr(Val(gOrderTable.sRack)) & "^" & CStr(Val(gOrderTable.sPos)) & "||" & sBuf & "|R||||||N||||1||||||||||O" & vbCr & Chr(3)
    'sSndO = "3O|1|" & pSampleInfo.ID & "^^" & CStr(Val(pSampleInfo.RACK)) & "^" & CStr(Val(pSampleInfo.POS)) & "||" & sBuf & "|R||||||N||||1||||||||||O" & vbCr & Chr(3)
'    sSndO = "3O|1|" & pSampleInfo.ID & "^^" & "^" & "||" & sBuf & "|R||||||N||||1||||||||||O" & vbCr & Chr(3)
    sSndO = "3O|1|" & pSampleInfo.ID & "||" & sBuf & "|R||||||X||||1||||||||||O" & vbCr & Chr(3)
    sSndO = Chr(2) & sSndO & ASTM_CheckSum(sSndO) & vbCr & vbLf
    
    sSndL = "4L|1" & vbCr & Chr(3)
    sSndL = Chr(2) & sSndL & ASTM_CheckSum(sSndL) & vbCr & vbLf
        
    '<ENQ> ����
    msComm.Output = Chr(5)
    
    If m_sTestMode = 77 Then
       RaiseEvent PrintSendLog("<ENQ>")
    End If


    '<ENQ>�� ���� ����
    sSndstate = "E"
    m_iPhase = 3
    
    Exit Sub
    
ErrHandler:
    sSndstate = ""
    m_iPhase = 1
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder Error - " & Err.Description)
    End If
End Sub

Private Function ASTM_CheckSum(ByVal sBuf$) As String
    Dim iCnt%
    Dim iSum%
    
    For iCnt = 1 To Len(sBuf)
        iSum = iSum + Val(Asc(Mid(sBuf, iCnt, 1)))
    Next
    
    iSum = iSum Mod 256
    
    ASTM_CheckSum = Right("0" & CStr(Hex(iSum)), 2)
End Function

Public Function GetByOne(ByVal tStr As String, sOriginal As String) As String
    Dim POS%
    
    POS = InStr(tStr, Chr$(124))
    
    If POS = 0 Then
    Else
        GetByOne = Trim$(Mid$(tStr, 1, POS - 1))
        sOriginal = Trim$(Mid$(sOriginal, POS + 1, Len(sOriginal) - POS))
    End If
End Function

Public Function GetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim POS%

    POS = InStr(tStr, sUserSymbol)

    If POS = 0 Then
    Else
        GetByOneUserSymbol = Trim$(Mid$(tStr, 1, POS - 1))
        sOriginal = Trim$(Mid$(sOriginal, POS + 1, Len(sOriginal) - POS))
    End If
End Function

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

