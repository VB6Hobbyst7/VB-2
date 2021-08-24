VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl BD 
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   1815
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
Attribute VB_Name = "BD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'�⺻ �Ӽ� ��:
Const m_def_p_sPatInfo = 0
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
Private m_p_sPatInfo As Variant
Private m_EqName As String
Private m_bUseBarcode As Boolean
Private m_iPhase As Integer
Private m_iSendPhase As Integer
Private m_sTestMode As String
Private m_iFrameN As Integer
Private m_p_sID As String
Private m_p_sSeq As String
Private m_p_sRack As String
Private m_p_sPos As String
Private m_p_iOrdCnt As Integer
Private m_p_sTIFCd As String
Private m_PortOpen As Boolean
Private m_OpenPW As String
Private m_EditPW As String
'�̺�Ʈ ����:
Event AppendData(sID$, sSeq$, sRack$, sPos$, sOrgrst$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTRst3$, sUnit$, sTFlag$, sQCGbn$)
Event SendOrderOK(sID$, sSeqno$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sSeqno$)
Event DispMsg(sMsg$)

'===== User Define
'�������̽����� ���
Private f_strRcvBuffer  As String
Private f_strWkBuf      As String
Private f_strState      As String
Private f_blnSend       As Boolean
Private f_bEndChk       As Boolean
Private f_bSTXChk       As Boolean

'����ü ����
Private f_typSampleInfo As SAMPLE_INFO
Private f_typResultInfo As RESULT_INFO

'��Ÿ
Private f_intSpaceCnt   As Integer
Private f_strOrganism   As String
Private f_strOrgaRslt   As String
Private f_subTestCode   As String
Private f_subStartDte   As String

Private Sub SendOrder_BD()

    On Error GoTo ErrRtn
    
    Dim sTmp    As String
    Dim ChkS    As String
    Dim strDta1()   As String
    
    Dim i       As Integer
    
    If m_iFrameN >= 7 Then
        m_iFrameN = 1
    End If
    
    Do While True
        Select Case m_iSendPhase
            Case 1      'Header Record
                sTmp = m_iFrameN & "H|\^&|||Becton Dickinson||||||||V1.0|" & Format$(Now, "YYYYMMDD") & vbCr
                '----- �˻��׸� ��ȸ/����
                RaiseEvent RequestCurOrder(f_typSampleInfo.SEQNO)
    
                Call Get_OrderString
    
                '��� ������ ���� ���
'                If f_typSampleInfo.ORDCNT > 0 Then
                    m_iSendPhase = 2
'                Else
'                    m_iSendPhase = 4
'                End If
                
            Case 2      'Patient Record
                '-- ��Ʈ��ȣ~����~��~����~�˻��ڵ�
                If InStr(f_typSampleInfo.OTHER, "~") > 0 Then
                    strDta1 = Split(f_typSampleInfo.OTHER, "~")
                Else
                    ReDim strDta1(0 To 5) As String
                    strDta1(0) = f_typSampleInfo.OTHER  '-- ��Ϲ�ȣ
                    strDta1(1) = "" '-- ����
                    strDta1(2) = "" '-- ��
                    strDta1(3) = "" '-- ����
                    strDta1(4) = "" '-- �˻��ڵ�
                End If
                
                sTmp = sTmp & "P|1||" & strDta1(0) & "||^^^^|||" & strDta1(1) & "||" & f_typSampleInfo.ID & "^^^^||" & strDta1(4) & "||" & f_typSampleInfo.ID & "|^^^^|||||^^^^||||||" & strDta1(3) & "|" & strDta1(2) & "|||||||" & vbCr
                m_iSendPhase = 3
                
            Case 3      'Order Record
                
                'BarCode �����
                sTmp = sTmp & "O|1|" & f_typSampleInfo.SEQNO & "^^^||^^^^||||||^^||^^||^||^^|^^^^|||||^|||||" & vbCr
                m_iSendPhase = 5
                
            Case 4
'                msComm.Output = Chr(4) & Chr(5)
                sTmp = sTmp & "Q|1|||||||||||A" & vbCr
                
                m_iSendPhase = 5
                
            Case 5      'Terminator Record
                sTmp = sTmp & "L|1|N" & vbCr & Chr(3)
                
                MSComm.Output = Chr(5) & Chr(2) & sTmp & ChkSum_ASTM(sTmp) & vbCrLf & Chr(4)
                
'                Call Sleep(500)
                
                m_iSendPhase = 6
                m_iPhase = 3
                
                If sTestMode = "77" Then
                    RaiseEvent PrintSendLog(Chr(5) & Chr(2) & sTmp & ChkSum_ASTM(sTmp) & vbCrLf & Chr(4))
                End If
                
'                m_iFrameN = 1:  m_iSendPhase = 7
'                m_iPhase = 3
'
'                Exit Sub
                
            Case 6      'EOT
                m_iFrameN = 1:  m_iPhase = 2:   m_iSendPhase = 7
                f_strState = "Q"
                
                '-- 05/11/14 YEJ
                'Barcode Mode�� ��� ���ۿϷ� �̺�Ʈ �߻�
'                RaiseEvent SendOrderOK(f_typSampleInfo.ID, f_typSampleInfo.SEQNO)
'                m_iPhase = 3
                
                Exit Sub
        End Select
        
        If m_iSendPhase = 7 Then Exit Do
    Loop
    
    m_iFrameN = m_iFrameN + 1

'    If sTestMode = "77" Then
'        RaiseEvent PrintSendLog(Chr(2) & sTmp & ChkSum_ASTM(sTmp) & vbCrLf & Chr(4))
'    End If
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder ���� - " & Err.Description)
    End If
End Sub

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
        Case "BD_BP"
            Call PhaseCfg_Protocol_BD
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_BD()
            
    Dim wkdat   As String
    Dim ix1     As Long
    
    For ix1 = 1 To Len(f_strWkBuf)
        wkdat = Mid$(f_strWkBuf, ix1, 1)
                 
        Select Case m_iPhase
            Case 1            'ENQ ���
                Select Case Asc(wkdat)
                    Case 5
                        f_bEndChk = True: f_bSTXChk = False
                        MSComm.Output = Chr(6)
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2      '<LF> ���
                Select Case Asc(wkdat)
                    Case 2  '-STX
                        If f_bEndChk = True Then
                            f_strRcvBuffer = ""
                        Else
                            f_bSTXChk = True
                        End If
                        f_bEndChk = True
                        
                    Case 3  'ETX
                        MSComm.Output = Chr(6)
                        If f_bEndChk = True Then
                            Call DataEditResponse_BD
                            f_strRcvBuffer = ""
                        End If
'                        msComm.Output = Chr(6)
                    
'''                    Case 3  'ETX
'''
'''                        Call DataEditResponse_BD
'''
'''                        f_strRcvBuffer = ""
'''                        m_iPhase = 2
'''                        If f_blnSend And f_typSampleInfo.ID <> "" Then
'''                            msComm.Output = Chr(6)
'''                        Else
'''                            msComm.Output = Chr(21)
'''                        End If
                        
                    Case 5  '-ENQ
                        f_bEndChk = True: f_bSTXChk = False
                        MSComm.Output = Chr(6)
                        
                    Case 6      'ACK
                        If f_strState = "Q" Then
'                            Call SendOrder_BD          '-- 05/11/14 YEJ Add
                            RaiseEvent SendOrderOK(f_typSampleInfo.ID, f_typSampleInfo.SEQNO)   '-- 05/11/14 YEJ Add
                            f_strState = ""
                        End If
                    
                    Case 21 '-NAK
                        
                    Case 23 '-ETB
                        f_bEndChk = False
                        MSComm.Output = Chr(6)
                        
                    Case Else
                        If f_bEndChk = True Then
                            If f_bSTXChk = True Then
                                f_bSTXChk = False
                            Else
                                f_strRcvBuffer = f_strRcvBuffer & wkdat
                            End If
                        End If
                        
                End Select
            
            Case 3      'ACK ���
                Select Case Asc(wkdat)
                    Case 6      'ACK
                        If f_strState = "Q" Then
'                            Call SendOrder_BD          '-- 05/11/14 YEJ Add
                            RaiseEvent SendOrderOK(f_typSampleInfo.ID, f_typSampleInfo.SEQNO)   '-- 05/11/14 YEJ Add
                            f_strState = ""
                        End If
                        m_iPhase = 1    '-- 05/11/14 YEJ
                    Case 5      'ENQ
                        f_bEndChk = True: f_bSTXChk = False
                        MSComm.Output = Chr(6)
                        m_iPhase = 2
                        
                    Case 21     'NAK
                        MSComm.Output = Chr(5)
                        m_iPhase = 3
                        
                    Case 3      'EOT
                        m_iPhase = 1
                End Select
                
        End Select
    Next ix1

End Sub


' *=====================================================*
' *               Data���� & ����ó��                   *
' *=====================================================*
Private Sub DataEditResponse_BD()

    On Error GoTo ErrHandler
    
    Dim strRecord() As String
    Dim intIdx  As Integer
    
    Dim strPacket$, strRecType$
    Dim strSampID$, strRst1$, strRst2$, strRst3$, strDate1$, strDate2$
    Dim strData1$(), strData2$(), strData3$
    Dim strIFCd As String, strOrganism  As String
    Dim strRstgbn   As String

    If f_strRcvBuffer = "" Then Exit Sub
     
    strRecord = Split(f_strRcvBuffer, vbCr)
    
    For intIdx = 0 To UBound(strRecord) '- 1
        
        strRecord(intIdx) = Replace(strRecord(intIdx), vbLf, "")
        
        If Mid(strRecord(intIdx), 2, 1) = "H" Or Mid(strRecord(intIdx), 3, 1) = "H" Then
            Call Init_f_typResultInfo
            f_strOrganism = ""
            f_strOrgaRslt = ""
            f_subTestCode = ""
            f_subStartDte = ""
            
            f_blnSend = False
        Else
            Select Case Mid(strRecord(intIdx), 1, 1)
                Case "H"        'Header Record
                    Call Init_f_typResultInfo
                    f_strOrganism = ""
                    f_strOrgaRslt = ""
                    f_subTestCode = ""
                    f_subStartDte = ""
                    f_blnSend = False
                Case "M"
                Case "P"        'Patient Record
                    strData1() = Split(strRecord(intIdx), "|")
                
                    If UBound(strData1) >= 12 Then
                        f_typSampleInfo.ID = Trim(strData1(10))     '-- barcode
                        f_typResultInfo.ID = Trim(strData1(10))
                        f_subTestCode = Trim(strData1(12))          '-- �˻��ڵ�
                    Else
                        f_typSampleInfo.ID = ""
                        f_typResultInfo.ID = ""
                        f_subTestCode = ""
                    End If
                Case "Q"        'Order Request Record
                    strData1() = Split(strRecord(intIdx), "|")
                    strData2() = Split(strData1(2), "^")
                    
                    If InStr(strData1(2), "^") > 0 Then
                        strSampID = strData2(1)
                    Else
                        strSampID = ""
                    End If
                    
                    If strSampID = "" Then
                        f_strState = ""
                        f_typSampleInfo.SEQNO = ""
                        Exit Sub
                    Else
                        f_strState = "Q"
                        f_typSampleInfo.SEQNO = strSampID
                    End If
                    
                Case "O"
                    f_strState = ""
                    strData1() = Split(strRecord(intIdx), "|")
                    
                    If InStr(strData1(2), "^") > 0 Then
                        strData2() = Split(strData1(2), "^")
                        f_typSampleInfo.SEQNO = Trim(strData2(0))
                        f_typResultInfo.SEQNO = Trim(strData2(0))
                        f_strOrganism = Trim(strData2(2))
                    Else
                        f_typSampleInfo.SEQNO = strData1(2)
                        f_typResultInfo.SEQNO = strData1(2)
                        f_strOrganism = ""
                    End If
                    
                    strRst1 = "": strRst2 = "": strRst3 = ""
                    f_subStartDte = Trim(strData1(7)):    strDate2 = Trim(strData1(14))
                    
                Case "R"
                    strData1() = Split(strRecord(intIdx), "|")
                    strData2() = Split(strData1(2), "^")
                    strRstgbn = Trim(strData2(3))
                    Select Case strRstgbn
                        Case "GND"          '-- �յ������
                                            strData2 = Split(strData1(3), "^")
                                            
                                            strRst1 = Trim(strData2(0))
                                            strRst2 = ""
                                            strRst3 = ""
                                            
                                            If strData1(11) = "" Then
                                                strRst2 = Format$(Now, "yyyy-MM-dd HH:mm:ss")
                                            Else
                                                strRst2 = Format$(strData1(11), "0000-00-00 00:00:00")
                                            End If
                                            
                                            strRst3 = Format$(strData1(12), "0000-00-00 00:00:00")
                                            
                                            strData2 = Split(strData1(13), "^")
                                            
                                            If UBound(strData2) >= 4 Then
                                            Else
                                                ReDim strData2(0 To 4) As String
                                                
                                                strData2(3) = "":   strData2(4) = ""
                                            End If
                                            
                                            '������� ����ü�� ����
                                            With f_typResultInfo
                                                .ID = f_typSampleInfo.ID
                                                .SEQNO = f_typSampleInfo.SEQNO
                                                .RACK = strData2(3) & "/" & strData2(4)
'                                                .POS = ""
                                                .QCGBN = f_typSampleInfo.QCGBN
                                                '����� ����
                                                .RSTCNT = .RSTCNT + 1
                                                .IFCD = .IFCD & "BECTECT" & Chr(124)
                                                .RST1 = .RST1 & strRst1 & Chr(124)
                                                .RST2 = .RST2 & strRst2 & Chr(124)
                                                .RST3 = .RST3 & strRst3 & Chr(124)
                                                .UNIT = .UNIT & "" & Chr(124)
                                                .FLAG = .FLAG & "" & Chr(124)
                                            End With
                                                
                                            strRst1 = "":   strRst2 = "":   strRst3 = ""
                                            strDate1 = "":  strDate2 = ""
                                            
                                            
                        Case "GEN_MGIT"     '-- BACTEC MGIT 960 �յ���
'                                            strData2 = Split(strData1(3), "^")
'
'                                            strRst1 = Trim(strData2(0))
'                                            strRst2 = ""
'                                            strRst3 = ""
'
'                                            strRst2 = Trim(strData1(11))
'                                            If Trim(strRst2) <> "" Then
'                                                strRst2 = Mid(strRst2, 3, 6) & "-" & Right(strRst2, 6)
'                                            End If
'
'                                            strRst3 = Trim(strData1(12))
'                                            If Trim(strRst3) <> "" Then
'                                                strRst2 = Mid(strRst3, 3, 6) & "-" & Right(strRst3, 6)
'                                            End If
'
'                                            '������� ����ü�� ����
'                                            With f_typResultInfo
'                                                .ID = f_typSampleInfo.ID
'                                                .SEQNO = f_typSampleInfo.SEQNO
'                                                .RACK = ""
''                                                .POS = ""
'                                                .QCGBN = f_typSampleInfo.QCGBN
'                                                '����� ����
'                                                .RSTCNT = .RSTCNT + 1
'                                                .IFCD = .IFCD & "BECTECT" & Chr(124)
'                                                .RST1 = .RST1 & strRst1 & Chr(124)
'                                                .RST2 = .RST2 & strRst2 & Chr(124)
'                                                .RST3 = .RST3 & strRst3 & Chr(124)
'                                                .UNIT = .UNIT & "" & Chr(124)
'                                                .FLAG = .FLAG & "" & Chr(124)
'                                            End With
'
'                                            strRst1 = "":   strRst2 = "":   strRst3 = ""
'                                            strDate1 = "":  strDate2 = ""
                        
                        Case "GEN_PROBETEC" '-- DBProbeTec ET
                        
                        Case "AST"          '-- AST ���
                                            If UBound(strData1) >= 3 Then
                                            strIFCd = "PHOENIX"
                                            
                                            strData2 = Split(strData1(2), "^")
                                            strRst1 = Trim(strData2(5))
                                            
                                            strData2 = Split(strData1(3), "^")
                                            
                                            If UBound(strData2) >= 3 Then
                                                strRst2 = Trim(strData2(1))
                                                strRst3 = IIf(Trim$(strData2(4)) = "", Trim(strData2(3)), Trim$(strData2(4)))
                                                If strRst3 = "" Then strRst3 = strData2(2)
                                                
                                                '������� ����ü�� ����
                                                With f_typResultInfo
                                                    .ID = f_typSampleInfo.ID
                                                    .SEQNO = f_typSampleInfo.SEQNO
                                                    .RACK = ""
    '                                                .POS =""
                                                    .QCGBN = f_typSampleInfo.QCGBN
                                                    '����� ����
                                                    .IFCD = .IFCD & "PHOENIX" & Chr(124)
                                                    .RSTCNT = .RSTCNT + 1
                                                    .RST1 = .RST1 & strRst1 & Chr(124)
                                                    .RST2 = .RST2 & strRst2 & Chr(124)
                                                    .RST3 = .RST3 & strRst3 & Chr(124)
                                                    .UNIT = .UNIT & "" & Chr(124)
                                                    .FLAG = .FLAG & "" & Chr(124)
                                                End With
                                            End If
                                            
                                            strRst1 = "":   strRst2 = "":   strRst3 = ""
                                            strDate1 = "":  strDate2 = ""
                                            End If
                                            
                        Case "AST_MIC"     '-- AST ���
'                                            If UBound(strData1) >= 14 Then
'                                                f_strOrgaRslt = strData1(14)
'                                                strData2 = Split(f_strOrgaRslt, "^")
'                                                f_strOrgaRslt = strData2(1)
'                                            End If

                                            strData2 = Split(strData1(2), "^")
                                            strRst1 = Trim(strData2(5))

                                            strData2 = Split(strData1(3), "^")

                                            If UBound(strData2) >= 3 Then
                                                strRst2 = Trim(strData2(1))
                                                strRst3 = Trim(strData2(3))
                                                If strRst3 = "" Then strRst3 = Trim(strData2(2))
                                                
                                                '������� ����ü�� ����
                                                With f_typResultInfo
                                                    .ID = f_typSampleInfo.ID
                                                    .SEQNO = f_typSampleInfo.SEQNO
    '                                                .RACK = ""
    '                                                .POS =""
                                                    .QCGBN = f_typSampleInfo.QCGBN
                                                    '����� ����
                                                    .RSTCNT = .RSTCNT + 1
                                                    .IFCD = .IFCD & "PHOENIX" & Chr(124)
                                                    .RST1 = .RST1 & strRst1 & Chr(124)
                                                    .RST2 = .RST2 & strRst2 & Chr(124)
                                                    .RST3 = .RST3 & strRst3 & Chr(124)
                                                    .UNIT = .UNIT & "" & Chr(124)
                                                    .FLAG = .FLAG & "" & Chr(124)
                                                End With
                                            End If

                                            strRst1 = "":   strRst2 = "":   strRst3 = ""
                                            strDate1 = "":  strDate2 = ""
                        
                        Case "AST_DIA"
                        Case "ID"           '-- ���̸�
                                            strData2 = Split(strData1(3), "^")
                                            f_strOrganism = Trim(strData2(1))
                                            
                        Case "OTHERE"       '-- ��Ÿ
                    End Select
                Case "C"
                    strData1 = Split(strRecord(intIdx), "|")
                    f_strOrgaRslt = Trim(strData1(3))
                        
                Case "L"
                    strData1() = Split(strRecord(intIdx), "|")
                    
                    f_blnSend = True
                    If f_strState = "Q" And strData1(2) <> "A" Then
                        m_iSendPhase = 1
                        Call SendOrder_BD
                    ElseIf f_strState <> "Q" Then
                        '����� ���/ȭ�� ǥ�� ó��...
                        With f_typResultInfo
                            If .RSTCNT > 0 Or f_strOrganism <> "" Then
                                If .RSTCNT = 0 And f_strOrganism <> "" Then
                                    .IFCD = .IFCD & "PHOENIX" & Chr(124)
                                    .RST1 = .RST1 & "" & Chr(124)
                                    .RST2 = .RST2 & "" & Chr(124)
                                    .RST3 = .RST3 & "" & Chr(124)
                                    .UNIT = .UNIT & "" & Chr(124)
                                    .FLAG = .FLAG & "" & Chr(124)
                                    .RSTCNT = .RSTCNT + 1
                                End If
                                
                                RaiseEvent AppendData(.ID, .SEQNO, .RACK, f_subTestCode, f_strOrganism & "^" & f_strOrgaRslt, 1, .IFCD, .RST1, .RST2, .RST3, .UNIT, .FLAG, .QCGBN)
                            End If
                        End With
            
                        Call Init_f_typResultInfo
                        f_strOrganism = ""
                        f_strOrgaRslt = ""
                        f_subTestCode = ""
                        f_subStartDte = ""
                    End If
                    
            End Select
        End If
        
    Next
    
    Exit Sub
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With f_typSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
            Erase .IFCD
        End With
        
        Exit Sub
    End If
    
    With f_typSampleInfo
        .ID = m_p_sID
        .SEQNO = m_p_sSeq
        .RACK = m_p_sRack
        .POS = m_p_sPos
        .ORDCNT = 1      '���� �˻� ������ �׸� ����
        ReDim Preserve .IFCD(1 To 1) As String
        .IFCD(1) = ""
        .OTHER = m_p_sPatInfo
    End With
        
End Sub

'
'   ������� ����ü �ʱ�ȭ
'
Private Sub Init_f_typResultInfo()
    
    With f_typResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .QCGBN = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .RST3 = ""
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

    f_strWkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case MSComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            f_strWkBuf = MSComm.Input
            
            If sTestMode = "77" Then
                RaiseEvent PrintRcvLog(f_strWkBuf)
            End If
                                
            If f_intSpaceCnt = 30 Then
                f_intSpaceCnt = 0
            End If
            f_intSpaceCnt = f_intSpaceCnt + 2
            
            RaiseEvent DispMsg(Space(f_intSpaceCnt) & "���� Interface �۾� ��...")
            
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
    m_p_sPatInfo = PropBag.ReadProperty("p_sPatInfo", m_def_p_sPatInfo)
    
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
    Call PropBag.WriteProperty("p_sPatInfo", m_p_sPatInfo, m_def_p_sPatInfo)
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
    m_p_sPatInfo = m_def_p_sPatInfo

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

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=14,0,0,0
Public Property Get p_sPatInfo() As Variant
    p_sPatInfo = m_p_sPatInfo
End Property

Public Property Let p_sPatInfo(ByVal New_p_sPatInfo As Variant)
    m_p_sPatInfo = New_p_sPatInfo
    PropertyChanged "p_sPatInfo"
End Property

