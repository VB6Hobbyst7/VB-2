VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmIIS204 
   BackColor       =   &H00DBE6E6&
   Caption         =   "�˻��� ��ȸ"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "frmIIS204.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&P)"
      Height          =   495
      Left            =   11490
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ������(&C)"
      Height          =   495
      Left            =   12705
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   13913
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   8567
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1335
      Left            =   68
      TabIndex        =   7
      Top             =   -15
      Width           =   15105
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DBE6E6&
         Height          =   330
         Left            =   3675
         Picture         =   "frmIIS204.frx":0CCA
         Style           =   1  '�׷���
         TabIndex        =   2
         Top             =   825
         Width           =   405
      End
      Begin VB.TextBox txtEqpCd 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   1515
         MaxLength       =   8
         TabIndex        =   1
         Top             =   840
         Width           =   2160
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ȸ(&Q)"
         Height          =   495
         Left            =   7260
         Style           =   1  '�׷���
         TabIndex        =   3
         Top             =   735
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFromDt 
         Height          =   330
         Left            =   1515
         TabIndex        =   0
         Top             =   270
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25427969
         CurrentDate     =   38330
      End
      Begin MedControls1.LisLabel lblEqpNm 
         Height          =   345
         Left            =   4110
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   825
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   609
         BackColor       =   16252919
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   330
         Left            =   3285
         TabIndex        =   13
         Top             =   270
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25427969
         CurrentDate     =   38330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "~"
         Height          =   180
         Left            =   3015
         TabIndex        =   14
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� �˻����"
         Height          =   180
         Left            =   330
         TabIndex        =   10
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "�� ��������"
         Height          =   180
         Left            =   330
         TabIndex        =   9
         Top             =   345
         Width           =   960
      End
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   405
      Left            =   75
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1365
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   714
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� �˻��� ����Ʈ"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   6690
      Left            =   75
      TabIndex        =   12
      Top             =   1800
      Width           =   15090
      _Version        =   393216
      _ExtentX        =   26617
      _ExtentY        =   11800
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      MaxRows         =   22
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIIS204.frx":1B0C
      TextTip         =   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3765
      TabIndex        =   16
      Top             =   1500
      Width           =   540
   End
   Begin VB.Label lblCnt 
      BackStyle       =   0  '����
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4425
      TabIndex        =   15
      Top             =   1500
      Width           =   450
   End
End
Attribute VB_Name = "frmIIS204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIIS204.frm (�츮LIS�� �����Ҷ� ���)
'   �ۼ���  :
'   ��  ��  : �˻��� ��ȸ��
'   �ۼ���  : 2004-12-17
'   ��  ��  :
'       1. 1.1.2:  (2004-12-17)
'       2. 1.1.3:  (2004-12-28)
'          - ��ȸ�� Spread�� ����, ��ȸ����ǥ��
'          - ��½� ����, ��ȸ���� ���
'       3. 1.2.3:  (2005-06-14)
'   ��  ��  :
'       1. �̻��� ���� ���? �ϴ��� �̻����� �����ϰ� ��������!
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady�� Column Enum
Private Enum TReadyEnum
    ccNo = 1
    ccPtId = 2
    ccName = 3
    ccAccNo = 4
    ccBarNo = 5
    ccSexAge = 6
    ccStatFg = 7
    ccWardId = 8
    ccDept = 9
    ccSpcNm = 10
    ccTestNms = 11
    ccRcvNm = 12
    ccRcvDt = 13
    ccRmk = 14
End Enum

Private mEqpChoice        As clsIISEqpChoice    '������ ���� Ŭ����
Private WithEvents mCode  As clsIISCodeList     '�ڵ帮��Ʈ Ŭ����
Attribute mCode.VB_VarHelpID = -1

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = "�˻�����ȸ"
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    Set mEqpChoice = New clsIISEqpChoice
    
    Call CtlClear
    Call ShowBasicEqp
End Sub

Private Sub Form_Deactivate()
    Me.WindowState = vbMinimized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mEqpChoice = Nothing
    Set frmIIS204 = Nothing
End Sub

Private Sub cmdQuery_Click()
    Dim Rs          As ADODB.Recordset
    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
    Dim strEqpCd    As String           '����ڵ�
    Dim strFromDt   As String           'From Date
    Dim strToDt     As String           'To Date
    Dim strKey      As String           'Spread�� Ű(SpcYy+SpcNo)
    Dim strSpcYy    As String           '���ڵ��ȣ(����)
    Dim strSpcNo    As String           '���ڵ��ȣ(����)
    Dim strTemp     As String
    
    strEqpCd = Trim$(txtEqpCd.Text)
    strFromDt = Format$(dtpFromDt.Value, "YYYYMMDD")
    strToDt = Format$(dtpToDt.Value, "YYYYMMDD")
    If strEqpCd = "" Then
        MsgBox "��� �����ϼ���.", vbInformation, "����"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Call mTblClear(tblReady)
    
On Error GoTo Errors
    Set objAccInfo = New clsIISAccInfo
    Set Rs = objAccInfo.GetTargetSpcs(strEqpCd, strFromDt, strToDt)
    If Not (Rs.BOF Or Rs.EOF) Then
        With tblReady
            Do Until Rs.EOF
                strSpcYy = Rs.Fields("SPCYY").Value
                strSpcNo = Rs.Fields("SPCNO").Value
                strKey = strSpcYy & strSpcNo
                If strTemp <> strKey Then
                    '## �ٸ� ���ڵ��ȣ �϶��� ������� ǥ��
                    If .MaxRows <= .DataRowCnt Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    Else
                        .Row = .DataRowCnt + 1
                    End If
                    
                    .Col = TReadyEnum.ccNo:      .Value = .Row
                    .Col = TReadyEnum.ccPtId:    .Value = Rs.Fields("PTID").Value & ""
                    .Col = TReadyEnum.ccName:    .Value = Rs.Fields("NAME").Value & ""
                    .Col = TReadyEnum.ccAccNo:   .Value = Rs.Fields("WORKAREA").Value & "-" & _
                                                          Mid$(Rs.Fields("ACCDT").Value, 3) & "-" & _
                                                          Rs.Fields("ACCSEQ").Value
                    .Col = TReadyEnum.ccBarNo:   .Value = strSpcYy & "-" & strSpcNo
                    .Col = TReadyEnum.ccSexAge:  .Value = Rs.Fields("SEX").Value & "" & "/" & _
                                                          mGetAge(Mid$(Rs.Fields("SSN").Value & "", 1, 6))
                    .Col = TReadyEnum.ccStatFg:  .Value = IIf(Rs.Fields("STATFG").Value & "" = "1", "Y", "")
                    .Col = TReadyEnum.ccWardId:  .Value = Rs.Fields("WARDID").Value & ""
                    .Col = TReadyEnum.ccDept:    .Value = Rs.Fields("DEPTCD").Value & ""
                    .Col = TReadyEnum.ccSpcNm:   .Value = Rs.Fields("SPCNM").Value & ""
                    .Col = TReadyEnum.ccTestNms: .Value = Rs.Fields("TESTNM").Value & ""
                    .Col = TReadyEnum.ccRcvNm:   .Value = Rs.Fields("RCVNM").Value & ""
                    .Col = TReadyEnum.ccRcvDt:   .Value = Format$(Rs.Fields("RCVDT").Value & "", "####-##-##") & " " & _
                                                          Mid$(Rs.Fields("RCVTM").Value & "", 1, 2) & ":" & _
                                                          Mid$(Rs.Fields("RCVTM").Value & "", 3, 2)
                                                          
                    '## 1.2.3:  (2005-06-14)
                    '   - ó�渮��ũ�� ��ȸ�ϵ��� ����
                    .Col = TReadyEnum.ccRmk:     .Value = Rs.Fields("MESG").Value & ""
                    strTemp = strKey
                Else
                    '## ���� ���ڵ��ȣ �϶��� �˻�� ǥ��
                    .Col = TReadyEnum.ccTestNms
                    .Value = .Value & "," & Rs.Fields("TESTNM").Value & ""
                End If
                Rs.MoveNext
            Loop
            
            lblCnt.Caption = CStr(.DataRowCnt)
        End With
    End If

    Rs.Close
    Set Rs = Nothing
    Set objAccInfo = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
Errors:
    Set Rs = Nothing
    Set objAccInfo = Nothing
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "����"
End Sub

Private Sub cmdSearch_Click()
    Set mCode = New clsIISCodeList
    
    With mCode
        .Caption = "�˻���� ����Ʈ"
        .HeaderCd = "����ڵ�"
        .HeaderCdNm = "����"
        .CodeListByRs mEqpChoice.GetUsingEqp
    End With
    Set mCode = Nothing
    
    SendKeys "{TAB}"
End Sub

Private Sub cmdPrint_Click()
    Dim objPrint    As clsIISPrint  '��� Ŭ����
    Dim strHeader1  As String       '������1
    Dim strHeader2  As String       '������2
    Dim strHeader3  As String       '������3
    Dim strBody     As String       '��¹ٵ�
    Dim i           As Long
    
    If tblReady.DataRowCnt < 1 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    strHeader1 = "���˻��󸮽�Ʈ��"
    
    strHeader2 = "�� �˻���� : " & lblEqpNm.Caption & " (" & txtEqpCd.Text & ")" & DIV & "5" & DIV & "1"
    strHeader2 = strHeader2 & vbTab & "�� ����Ͻ� : " & Format$(Now, "YYYY-MM-DD HH:MM") & _
                 Space(30) & "�� ��ȸ�Ǽ� : " & lblCnt.Caption & _
                 DIV & "5" & DIV & "1"
    
    strHeader3 = "����" & DIV & "5" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "������ȣ" & DIV & "15" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "ȯ��ID" & DIV & "40" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "ȯ�ڸ�" & DIV & "57" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "S/A" & DIV & "75" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "����" & DIV & "85" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "����" & DIV & "95" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "�����" & DIV & "105" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "��ü��" & DIV & "120" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "�˻��׸�" & DIV & "135" & DIV & "1"
    
    With tblReady
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = TReadyEnum.ccNo:      strBody = strBody & .Value & DIV & "5" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccAccNo:   strBody = strBody & vbTab & .Value & DIV & "15" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccPtId:    strBody = strBody & vbTab & .Value & DIV & "40" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccName:    strBody = strBody & vbTab & .Value & DIV & "57" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccSexAge:  strBody = strBody & vbTab & .Value & DIV & "75" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccStatFg:  strBody = strBody & vbTab & .Value & DIV & "85" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccWardId:  strBody = strBody & vbTab & .Value & DIV & "95" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccDept:    strBody = strBody & vbTab & .Value & DIV & "105" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccSpcNm:   strBody = strBody & vbTab & .Value & DIV & "120" & DIV & "0" & DIV & "0"
            .Col = TReadyEnum.ccTestNms: strBody = strBody & vbTab & .Value & DIV & "135" & DIV & "1" & DIV & "1" & vbTab
        Next i
        strBody = Mid$(strBody, 1, Len(strBody) - 1)
    End With
    
    Set objPrint = New clsIISPrint
    
    With objPrint
        .PrinterHeader1 = strHeader1
        .PrinterHeader2 = strHeader2
        .PrinterHeader3 = strHeader3
        .PrinterBody = strBody
        .CallPrint
    End With
    Set objPrint = Nothing
    
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub tblReady_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim strTestNms As String    '�˻��׸��
    Dim strRemarks As String    'ó�渮��ũ
    
    If Row < 1 Then Exit Sub
    With tblReady
        .Row = Row: .Col = TReadyEnum.ccTestNms
        If .Value = "" Then Exit Sub
        
        strTestNms = vbCrLf & "   ## �˻��׸� ##" & vbCrLf
        strTestNms = strTestNms & Space(3) & .Value & vbCrLf
        
        '## 1.2.3:  (2005-06-14)
        '   - ó�渮��ũ�� �����ϸ� ������ ǥ���ϵ��� ����
        .Row = Row: .Col = TReadyEnum.ccRmk
        strRemarks = .Value
        
        If strRemarks <> "" Then
            strTestNms = strTestNms & vbCrLf & "   ## ó�渮��ũ ##" & vbCrLf
            strTestNms = strTestNms & Space(3) & strRemarks & vbCrLf
        End If
        
        MultiLine = 1
        TipWidth = 6000
        TipText = strTestNms
        Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �˻����1���� ������ ���ǥ��
'-----------------------------------------------------------------------------'
Private Sub ShowBasicEqp()
    With mEqpChoice
        If .GetEqp Then
            If .EqpCd1 = "" Then GoTo EndLine
            txtEqpCd.Text = .EqpCd1
            lblEqpNm.Caption = .EqpNm1
            
            '## ��Ŀ���� "��ȸ"��ư����
            SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}"
        End If
    End With
    Exit Sub
    
EndLine:
    '## ��Ŀ���� ����� ��ư����
    SendKeys "{TAB}": SendKeys "{TAB}"
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    dtpFromDt.Value = Format(Now - 1, "YYYY-MM-DD")
    dtpToDt.Value = Now
    txtEqpCd.Text = ""
    lblEqpNm.Caption = "":  lblCnt.Caption = ""
    Call mTblClear(tblReady)
End Sub

'-----------------------------------------------------------------------------'
'   ��� : CodeList���� �̺�Ʈ ó��1
'-----------------------------------------------------------------------------'
Private Sub mCode_SelectedItem(ByRef pSelItem As String)
    Dim strEqpCd As String      '����ڵ�
    
    strEqpCd = mGetP(pSelItem, 1, DIV)
    If strEqpCd = Trim(txtEqpCd.Text) Then
        MsgBox "�ش� ����ڵ�� �̹� ���õǾ� �ֽ��ϴ�.", vbInformation, "����"
        pSelItem = ""
    Else
        txtEqpCd.Text = strEqpCd
        lblEqpNm.Caption = mGetP(pSelItem, 2, DIV)
    End If
End Sub
