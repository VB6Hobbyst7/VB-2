VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.UserControl ROTORGENE 
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   LockControls    =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   7890
   Begin VB.TextBox txtTestNm 
      Height          =   285
      Left            =   30
      TabIndex        =   3
      Text            =   "�˻��׸��"
      Top             =   45
      Width           =   1155
   End
   Begin FPSpread.vaSpread spdExcRst 
      Height          =   2895
      Left            =   30
      TabIndex        =   2
      Top             =   360
      Width           =   7800
      _Version        =   196608
      _ExtentX        =   13758
      _ExtentY        =   5106
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   20
      SpreadDesigner  =   "ROTORGENE.ctx":0000
      UserResize      =   2
   End
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
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   0
      Top             =   600
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
Attribute VB_Name = "ROTORGENE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�⺻ �Ӽ� ��:
Const m_def_p_sCmt1 = ""
Const m_def_p_sSpcCd = 0
Const m_def_p_sTSVol = "0"
Const m_def_p_sRerunGbn = "0"
Const m_def_p_bSIndex = 0
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
Dim m_p_sCmt1 As String
Dim m_p_sSpcCd As Variant
Dim m_p_sTSVol As String
Dim m_p_sRerunGbn As String
Dim m_p_bSIndex As Boolean
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
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTInstID$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sRack$, sPos$, sKind$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
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
Dim iSpaceCnt   As Integer

'For E-170/Hitachi7600
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim sNextSend   As String

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

Private Function ConvertDataAlarmCode(ByVal sEqNm As String, ByVal sCode As String) As String
    
    Dim sTmp    As String
    
    ConvertDataAlarmCode = "": sTmp = ""
    
    Select Case UCase(sEqNm)
        Case "HITACHI7600"
            Select Case Trim(sCode)
                Case "0": sTmp = ""
                Case "1": sTmp = "ADC?"
                Case "2": sTmp = "Cell?"
                Case "3": sTmp = "Sampl"
                Case "4": sTmp = "Reagn"
                Case "5": sTmp = "ABS?"
                Case "6": sTmp = "Prozon"
                Case "7": sTmp = "Limt0"
                Case "8": sTmp = "Limt1"
                Case "9": sTmp = "Limt2"
                Case "10": sTmp = "Lin."
                Case "11": sTmp = "Lin8."
                Case "12": sTmp = "S1Abs?"
                Case "13": sTmp = "Dup"
                Case "14": sTmp = "Std?"
                Case "15": sTmp = "Sens"
                Case "16": sTmp = "Calib"
                Case "17": sTmp = "SDI"
                Case "18": sTmp = "Noise"
                Case "19": sTmp = "Level"
                Case "20": sTmp = "Slope?"
                Case "21": sTmp = "Margin"
                Case "22": sTmp = "I.Std"
                Case "23": sTmp = "R.Over"
                Case "24": sTmp = "Cmp.T"
                Case "25": sTmp = "Cmp.TI"
                Case "26": sTmp = "LIMTH"
                Case "27": sTmp = "LIMTL"
                Case "28": sTmp = "Random"
                Case "29": sTmp = "Systm1"
                Case "30": sTmp = "Systm2"
                Case "31": sTmp = "Systm3"
                Case "32": sTmp = "Systm4"
                Case "33": sTmp = "Systm5"
                Case "34": sTmp = "Systm6"
                Case "35": sTmp = "QCErr1"
                Case "36": sTmp = "QCErr2"
                Case "37": sTmp = "Calc?"
                Case "38": sTmp = "Over"
                Case "39": sTmp = "???"
                Case "42": sTmp = "Edited"
                Case "44": sTmp = "ReptH"
                Case "45": sTmp = "ReptL"
                Case "51": sTmp = "Resp1"
                Case "52": sTmp = "Resp2"
                Case "53": sTmp = "Condi"
            End Select
        
        Case Else
        
    End Select
    
    ConvertDataAlarmCode = Trim(sTmp)
    
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
        Case "ROTORGENE"
            'Call PhaseCfg_Protocol_RotorGene
        
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub

' *=====================================================*
' *               Data���� & ����ó��                   *
' *=====================================================*
Public Sub DataEditResponse_RotorGene()
    On Error GoTo ErrRtn
    
    Dim iRow       As Integer
    Dim sTestNm    As String
    Dim sRst       As String
    Dim vBarCd, vCtval1, vCtval2, vRst
        
    With spdExcRst
        For iRow = 1 To .MaxRows
            Call Init_pResultInfo
                  
            RaiseEvent DispMsg(Space(iSpaceCnt) & "���� Interface �۾� ��...")
            
            Call .GetText(1, iRow, vBarCd)
            Call .GetText(2, iRow, vCtval1)
            Call .GetText(3, iRow, vRst)
            sRst = Trim(CStr(vRst))
            
            '<Ct�� ó��..
            If CStr(vCtval1) = "" Then
                Call .GetText(5, iRow, vCtval2)
                
                If Val(vCtval2) <= 25 Then
                    sRst = "Target Not Detected"
                ElseIf Val(vCtval2) > 25 Then
                    sRst = "Test Invalidated"
                End If
            End If
            '>
                         
            With pResultInfo
                .ID = Trim(CStr(vBarCd))
                .RST1 = sRst & Chr(124)
                .RST2 = "" & Chr(124)
                .FLAG = "" & Chr(124)
                .IFCD = Trim(txtTestNm.Text) & Chr(124)
                .RSTCNT = "1"   '�Ѱ��� File�� �� �׸��̱� ����...
                
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .INSTID, .ALARMCD, .KIND, .RSTDT, .OTHER)
                
            End With
            
            Call Init_pResultInfo
        Next
    End With
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Public Sub DisplaySpdExcRst1(ByVal sFileNm As String)
    On Error GoTo ErrRtn
    
    Dim blnFlag   As Boolean
    Dim bSpdSet   As Boolean
    Dim intIdx    As Integer
    Dim iRow      As Integer
    Dim iPos      As Integer
    
    Dim sTestNm   As String
    Dim sBarCd    As String
    Dim sCtval    As String
    Dim sRst      As String
    Dim sSeq      As String
    Dim sType     As String
    Dim sSplit()  As String
    
    Dim objExcell   As Object
    
    Set objExcell = CreateObject("Excel.Application")
    
    objExcell.Workbooks.Open sFileNm
    
    objExcell.Visible = False
    
    blnFlag = False
    bSpdSet = False
    
    spdExcRst.MaxRows = 0
    
    intIdx = 0
    iRow = 0

    '<Test Name Edit
    sSplit = Split(sFileNm, "\")
    sTestNm = sSplit(UBound(sSplit))
    
    If InStr(sTestNm, "_") > 0 Then
        iPos = InStr(sTestNm, "_")
        sTestNm = Mid(sTestNm, 1, iPos - 1)
    Else
        iPos = InStr(sTestNm, ".")
        sTestNm = Mid(sTestNm, 1, iPos - 1)
        sTestNm = Left(Trim(sTestNm), 3)
    End If
    '>
    
    txtTestNm.Text = sTestNm
    
    Do
        intIdx = intIdx + 1
        
        sSeq = objExcell.Worksheets(1).Range("A" & CStr(intIdx)).Value      'Seq
        sBarCd = objExcell.Worksheets(1).Range("B" & CStr(intIdx)).Value    'BarCode
        sType = objExcell.Worksheets(1).Range("C" & CStr(intIdx)).Value     'TYpe
        sCtval = objExcell.Worksheets(1).Range("D" & CStr(intIdx)).Value    'Ct value
        sRst = objExcell.Worksheets(1).Range("F" & CStr(intIdx)).Value      'Result
        
        If IsNumeric(sSeq) And Trim(sType) <> "Standard" Then
            bSpdSet = True
            With spdExcRst
                .MaxRows = .MaxRows + 1
                iRow = iRow + 1
                Call .SetText(1, iRow, sBarCd)
                Call .SetText(2, iRow, sCtval)
                Call .SetText(3, iRow, sRst)
            End With
        ElseIf sSeq = "" And bSpdSet = True Then
            Exit Do
        End If
    Loop
        
    bSpdSet = False
    
    objExcell.Workbooks.Close
    Set objExcell = Nothing
       
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If

End Sub

Public Sub DisplaySpdExcRst2(ByVal sFileNm As String)
    On Error GoTo ErrRtn
    
    Dim blnFlag   As Boolean
    Dim bSpdSet   As Boolean
    Dim intIdx    As Integer
    Dim iRow      As Integer
    Dim iPos      As Integer
    
    Dim sTestNm   As String
    Dim sBarCd    As String
    Dim sCtval    As String
    Dim sRst      As String
    Dim sSeq      As String
    Dim sType     As String
    
    Dim objExcell   As Object
    
    Set objExcell = CreateObject("Excel.Application")
    
    objExcell.Workbooks.Open sFileNm
    
    objExcell.Visible = False
    
    blnFlag = False
    bSpdSet = False
   
    intIdx = 0
    iRow = 0
    
    Do
        intIdx = intIdx + 1
        
        sSeq = objExcell.Worksheets(1).Range("A" & CStr(intIdx)).Value      'Seq
        sBarCd = objExcell.Worksheets(1).Range("B" & CStr(intIdx)).Value    'BarCode
        sType = objExcell.Worksheets(1).Range("C" & CStr(intIdx)).Value     'TYpe
        sCtval = objExcell.Worksheets(1).Range("D" & CStr(intIdx)).Value    'Ct value
        sRst = objExcell.Worksheets(1).Range("F" & CStr(intIdx)).Value      'Result
        
        If IsNumeric(sSeq) And Trim(sType) <> "Standard" Then
            bSpdSet = True
            With spdExcRst
                iRow = iRow + 1
                Call .SetText(4, iRow, sBarCd)
                Call .SetText(5, iRow, sCtval)
                Call .SetText(6, iRow, sRst)
            End With
        ElseIf sSeq = "" And bSpdSet = True Then
            Exit Do
        End If
    Loop
        
    bSpdSet = False
    
    objExcell.Workbooks.Close
    Set objExcell = Nothing
       
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If

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
        .QCGBN = ""
        .KIND = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .INSTID = ""
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
End Sub
'
'   ȯ�� Order ����
'
Private Sub SendOrder_DPE_Batch()
'    On Error GoTo Err_Rtn
'
'    Dim sSendBuff   As String
'    Dim iCnt    As Integer
'    Dim ChkSum  As String
'    Dim sStat   As String
'
'    Select Case m_iSendPhase
'        Case 0
'            m_iSendPhase = 1
'            msComm.Output = Chr(5)
'            Exit Sub
'
'        Case 1
'            'Header Record
''            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^REPLY|P|1" & vbCr
'            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||H7600^1|TSDWN^BATCH|P|1" & vbCr
'
'            'Patient Record
'            sSendBuff = sSendBuff & "P|1" & vbCr
'
''            '----- �˻��׸� ��ȸ
''            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.KIND)
'
'            Call Get_OrderString
'
'            'Order Record
''            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) _
''                    & "^" & pSampleInfo.SPCCD & "^" _
''                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|" & pSampleInfo.KIND & "|"
'            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Left(Trim(pSampleInfo.ID) & Space(13), 13) _
'                    & "^1^" _
'                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|R1|"
'
'            '�˻��׸� Order�ڵ� �߰�
'            For iCnt = 1 To pSampleInfo.ORDCNT
'                '���� ����
'                Select Case Trim$(pSampleInfo.IFCD(iCnt))
'                    Case "989", "990", "991"
'                        'ISE �׸��� �ִ� ��� ��ü �˻�(ISE �˻�� 2���� ���ո� ���� ����)
'                        If Val(InStr(1, sSendBuff, "989" & "/")) = 0 Then
'                            sSendBuff = sSendBuff & "^^^989/\"
'                        End If
'                        If Val(InStr(1, sSendBuff, "990" & "/")) = 0 Then
'                            sSendBuff = sSendBuff & "^^^990/\"
'                        End If
'                        If Val(InStr(1, sSendBuff, "991" & "/")) = 0 Then
'                            sSendBuff = sSendBuff & "^^^991/\"
'                        End If
'
'                    Case Else
'                        '�Ϲ��׸�
'                        sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "/\"
'                End Select
'            Next iCnt
'
'            If pSampleInfo.SINDEX = True Then           'Serum Index
'                sSendBuff = sSendBuff & "^^^992/\^^^993/\^^^994/\"
'            End If
'
'            If pSampleInfo.ORDCNT > 0 And Trim(sReqStatusCd) <> "A" Then
'                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
'            End If
'
'            'STAT RACK�� ���� ó���߰�
''            If Left(pSampleInfo.RACK, 1) = "4" Then
''                sStat = "S"
''            Else
'                sStat = "R"
''            End If
'
'            If Trim(pSampleInfo.CMT1) = "" Then
'                sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
'                        & "^^^^||||||O" & vbCr
'            Else
'                'Comment ���� ������ ��� Comment1 ���� Ư������ ����(2005/8/1 yk)
'                sSendBuff = sSendBuff & "|" & sStat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
'                        & Trim(pSampleInfo.CMT1) & "^^^^||||||O" & vbCr
'            End If
'
'            'Terminator Record
'            sSendBuff = sSendBuff & "L|1|N"
'
'
'            '--- Text�� ������ 240byte�� �Ѿ ��� ó�� �߰�...
'            If Len(sSendBuff) >= 241 Then
'                sNextSend = Mid(sSendBuff, 241)
'                sSendBuff = Left(sSendBuff, 240)
'                sSendBuff = sSendBuff & Chr(23)
'
'                m_iFrameN = m_iFrameN + 1
'                m_iSendPhase = 2
'            Else
'                sSendBuff = sSendBuff & Chr(13) & Chr(3)
'                GoTo Send_Terminate
'            End If
'
'        Case 2
'            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
'            sNextSend = ""
'
'Send_Terminate:
'            m_iSendPhase = 3
'
'        Case 3      'EOT
'            msComm.Output = Chr(4)   'EOT
'            m_iFrameN = 1
'            m_iPhase = 3
'            m_iSendPhase = 1
'
'            sState = "": sReqStatusCd = ""
'
'            'BarCode Mode�� �ƴ� ��� ���� ���� ��ȸ
'            RaiseEvent RequestNextOrder
'
'            Exit Sub
'    End Select
'
'    ChkSum = ChkSum_ASTM(sSendBuff)
'    sSendBuff = sSendBuff & ChkSum
'    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)
'
'    If m_sTestMode = "77" Then
'        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
'    End If
'
''    '���۵� ������ �ִ� ��� ȭ��ǥ��
''    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
''        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
''            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
''        End If
''    Else
''        '��ȸ�� ������ ���� ��� ȯ������ ����ü �ʱ�ȭ
''        Call Init_pResultInfo
''
''        RaiseEvent SendOrderOK("", "", "", "")
''    End If
'
'Err_Rtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("Order ���۽� �����߻� - " & Err.Description)
'    End If
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
            .SINDEX = False
            .CMT1 = ""
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
        .SINDEX = m_p_bSIndex
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
        
        .CMT1 = m_p_sCmt1
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
    m_p_bSIndex = PropBag.ReadProperty("p_bSIndex", m_def_p_bSIndex)
    m_p_sRerunGbn = PropBag.ReadProperty("p_sRerunGbn", m_def_p_sRerunGbn)
    m_p_sTSVol = PropBag.ReadProperty("p_sTSVol", m_def_p_sTSVol)
    m_p_sSpcCd = PropBag.ReadProperty("p_sSpcCd", m_def_p_sSpcCd)
    m_p_sCmt1 = PropBag.ReadProperty("p_sCmt1", m_def_p_sCmt1)
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
    Call PropBag.WriteProperty("p_bSIndex", m_p_bSIndex, m_def_p_bSIndex)
    Call PropBag.WriteProperty("p_sRerunGbn", m_p_sRerunGbn, m_def_p_sRerunGbn)
    Call PropBag.WriteProperty("p_sTSVol", m_p_sTSVol, m_def_p_sTSVol)
    Call PropBag.WriteProperty("p_sSpcCd", m_p_sSpcCd, m_def_p_sSpcCd)
    Call PropBag.WriteProperty("p_sCmt1", m_p_sCmt1, m_def_p_sCmt1)
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
    
    '���� �ʱ�ȭ
    bEndChk = True: bSTXChk = False
    
    
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
'    m_iStartSampleNo = m_def_iStartSampleNo
    m_p_bSIndex = m_def_p_bSIndex
    m_p_sRerunGbn = m_def_p_sRerunGbn
    m_p_sTSVol = m_def_p_sTSVol
    m_p_sSpcCd = m_def_p_sSpcCd
    m_p_sCmt1 = m_def_p_sCmt1
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
'MemberInfo=0,0,0,0
Public Property Get p_bSIndex() As Boolean
    p_bSIndex = m_p_bSIndex
End Property

Public Property Let p_bSIndex(ByVal New_p_bSIndex As Boolean)
    m_p_bSIndex = New_p_bSIndex
    PropertyChanged "p_bSIndex"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sRerunGbn() As String
    p_sRerunGbn = m_p_sRerunGbn
End Property

Public Property Let p_sRerunGbn(ByVal New_p_sRerunGbn As String)
    m_p_sRerunGbn = New_p_sRerunGbn
    PropertyChanged "p_sRerunGbn"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sTSVol() As String
    p_sTSVol = m_p_sTSVol
End Property

Public Property Let p_sTSVol(ByVal New_p_sTSVol As String)
    m_p_sTSVol = New_p_sTSVol
    PropertyChanged "p_sTSVol"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=14,0,0,0
Public Property Get p_sSpcCd() As Variant
    p_sSpcCd = m_p_sSpcCd
End Property

Public Property Let p_sSpcCd(ByVal New_p_sSpcCd As Variant)
    m_p_sSpcCd = New_p_sSpcCd
    PropertyChanged "p_sSpcCd"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,
Public Property Get p_sCmt1() As String
    p_sCmt1 = m_p_sCmt1
End Property

Public Property Let p_sCmt1(ByVal New_p_sCmt1 As String)
    m_p_sCmt1 = New_p_sCmt1
    PropertyChanged "p_sCmt1"
End Property

