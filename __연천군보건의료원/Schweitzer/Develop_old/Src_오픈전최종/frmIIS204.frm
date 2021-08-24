VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmIIS204 
   BackColor       =   &H00DBE6E6&
   Caption         =   "검사대상 조회"
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
   WindowState     =   2  '최대화
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DBE6E6&
      Caption         =   "출 력(&P)"
      Height          =   495
      Left            =   11490
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   12705
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   495
      Left            =   13913
      Style           =   1  '그래픽
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
         Style           =   1  '그래픽
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
         Caption         =   "조 회(&Q)"
         Height          =   495
         Left            =   7260
         Style           =   1  '그래픽
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
         Caption         =   "▶ 검사장비"
         Height          =   180
         Left            =   330
         TabIndex        =   10
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 접수일자"
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
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "■ 검사대상 리스트"
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
         Name            =   "굴림"
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
      BackStyle       =   0  '투명
      Caption         =   "개수 :"
      BeginProperty Font 
         Name            =   "굴림"
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
      BackStyle       =   0  '투명
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "굴림"
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
'   파일명  : frmIIS204.frm (우리LIS랑 조인할때 사용)
'   작성자  :
'   내  용  : 검사대상 조회폼
'   작성일  : 2004-12-17
'   버  전  :
'       1. 1.1.2:  (2004-12-17)
'       2. 1.1.3:  (2004-12-28)
'          - 조회시 Spread에 순번, 조회개수표시
'          - 출력시 순번, 조회개수 출력
'       3. 1.2.3:  (2005-06-14)
'   메  모  :
'       1. 미생물 장비는 어떻게? 일단은 미생물은 제외하고 개발했음!
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady의 Column Enum
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

Private mEqpChoice        As clsIISEqpChoice    '사용장비 선택 클래스
Private WithEvents mCode  As clsIISCodeList     '코드리스트 클래스
Attribute mCode.VB_VarHelpID = -1

Private Sub Form_Activate()
    mdiIISMain.lblMenuNm = "검사대상조회"
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
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim strEqpCd    As String           '장비코드
    Dim strFromDt   As String           'From Date
    Dim strToDt     As String           'To Date
    Dim strKey      As String           'Spread의 키(SpcYy+SpcNo)
    Dim strSpcYy    As String           '바코드번호(연도)
    Dim strSpcNo    As String           '바코드번호(순번)
    Dim strTemp     As String
    
    strEqpCd = Trim$(txtEqpCd.Text)
    strFromDt = Format$(dtpFromDt.Value, "YYYYMMDD")
    strToDt = Format$(dtpToDt.Value, "YYYYMMDD")
    If strEqpCd = "" Then
        MsgBox "장비를 선택하세요.", vbInformation, "정보"
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
                    '## 다른 바코드번호 일때는 모든정보 표시
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
                    '   - 처방리마크를 조회하도록 수정
                    .Col = TReadyEnum.ccRmk:     .Value = Rs.Fields("MESG").Value & ""
                    strTemp = strKey
                Else
                    '## 같은 바코드번호 일때는 검사명만 표시
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
    MsgBox Err.Description, vbCritical, "오류"
End Sub

Private Sub cmdSearch_Click()
    Set mCode = New clsIISCodeList
    
    With mCode
        .Caption = "검사장비 리스트"
        .HeaderCd = "장비코드"
        .HeaderCdNm = "장비명"
        .CodeListByRs mEqpChoice.GetUsingEqp
    End With
    Set mCode = Nothing
    
    SendKeys "{TAB}"
End Sub

Private Sub cmdPrint_Click()
    Dim objPrint    As clsIISPrint  '출력 클래스
    Dim strHeader1  As String       '출력헤더1
    Dim strHeader2  As String       '출력헤더2
    Dim strHeader3  As String       '출력헤더3
    Dim strBody     As String       '출력바디
    Dim i           As Long
    
    If tblReady.DataRowCnt < 1 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    strHeader1 = "『검사대상리스트』"
    
    strHeader2 = "※ 검사장비 : " & lblEqpNm.Caption & " (" & txtEqpCd.Text & ")" & DIV & "5" & DIV & "1"
    strHeader2 = strHeader2 & vbTab & "※ 출력일시 : " & Format$(Now, "YYYY-MM-DD HH:MM") & _
                 Space(30) & "※ 조회건수 : " & lblCnt.Caption & _
                 DIV & "5" & DIV & "1"
    
    strHeader3 = "순번" & DIV & "5" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "접수번호" & DIV & "15" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "환자ID" & DIV & "40" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "환자명" & DIV & "57" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "S/A" & DIV & "75" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "응급" & DIV & "85" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "병동" & DIV & "95" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "진료과" & DIV & "105" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "검체명" & DIV & "120" & DIV & "0"
    strHeader3 = strHeader3 & vbTab & "검사항목" & DIV & "135" & DIV & "1"
    
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
    Dim strTestNms As String    '검사항목명
    Dim strRemarks As String    '처방리마크
    
    If Row < 1 Then Exit Sub
    With tblReady
        .Row = Row: .Col = TReadyEnum.ccTestNms
        If .Value = "" Then Exit Sub
        
        strTestNms = vbCrLf & "   ## 검사항목 ##" & vbCrLf
        strTestNms = strTestNms & Space(3) & .Value & vbCrLf
        
        '## 1.2.3:  (2005-06-14)
        '   - 처방리마크가 존재하면 툴팁에 표시하도록 수정
        .Row = Row: .Col = TReadyEnum.ccRmk
        strRemarks = .Value
        
        If strRemarks <> "" Then
            strTestNms = strTestNms & vbCrLf & "   ## 처방리마크 ##" & vbCrLf
            strTestNms = strTestNms & Space(3) & strRemarks & vbCrLf
        End If
        
        MultiLine = 1
        TipWidth = 6000
        TipText = strTestNms
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검사장비1에서 설정된 장비표시
'-----------------------------------------------------------------------------'
Private Sub ShowBasicEqp()
    With mEqpChoice
        If .GetEqp Then
            If .EqpCd1 = "" Then GoTo EndLine
            txtEqpCd.Text = .EqpCd1
            lblEqpNm.Caption = .EqpNm1
            
            '## 포커스를 "조회"버튼으로
            SendKeys "{TAB}": SendKeys "{TAB}": SendKeys "{TAB}"
        End If
    End With
    Exit Sub
    
EndLine:
    '## 포커스를 장비선택 버튼으로
    SendKeys "{TAB}": SendKeys "{TAB}"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    dtpFromDt.Value = Format(Now - 1, "YYYY-MM-DD")
    dtpToDt.Value = Now
    txtEqpCd.Text = ""
    lblEqpNm.Caption = "":  lblCnt.Caption = ""
    Call mTblClear(tblReady)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼의 이벤트 처리1
'-----------------------------------------------------------------------------'
Private Sub mCode_SelectedItem(ByRef pSelItem As String)
    Dim strEqpCd As String      '장비코드
    
    strEqpCd = mGetP(pSelItem, 1, DIV)
    If strEqpCd = Trim(txtEqpCd.Text) Then
        MsgBox "해당 장비코드는 이미 선택되어 있습니다.", vbInformation, "정보"
        pSelItem = ""
    Else
        txtEqpCd.Text = strEqpCd
        lblEqpNm.Caption = mGetP(pSelItem, 2, DIV)
    End If
End Sub
