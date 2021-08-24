VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frm157BarReprint 
   BackColor       =   &H00DBE6E6&
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   10905
   Begin VB.CheckBox chkSelAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "전체선택(&A)"
      ForeColor       =   &H00553755&
      Height          =   255
      Left            =   300
      TabIndex        =   13
      Top             =   1050
      Width           =   1350
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1050
      Left            =   8760
      TabIndex        =   8
      Top             =   -45
      Width           =   2070
      Begin VB.TextBox txtLabelCnt 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   915
         TabIndex        =   9
         Text            =   "1"
         Top             =   390
         Width           =   570
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1485
         TabIndex        =   10
         Top             =   390
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         BuddyControl    =   "txtLabelCnt"
         BuddyDispid     =   196611
         OrigLeft        =   3840
         OrigTop         =   330
         OrigRight       =   4080
         OrigBottom      =   645
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   2
         Left            =   75
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   390
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "출력장수"
         Appearance      =   0
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장"
         Height          =   180
         Left            =   1770
         TabIndex        =   11
         Tag             =   "151"
         Top             =   450
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00EAE7E3&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8190
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00EAE7E3&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00EAE7E3&
      Caption         =   "재출력(&S)"
      Height          =   510
      Left            =   6870
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8505
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1050
      Left            =   210
      TabIndex        =   3
      Top             =   -45
      Width           =   1815
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "환자 ID "
         Height          =   225
         Index           =   0
         Left            =   285
         TabIndex        =   5
         Top             =   270
         Width           =   945
      End
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "바코드"
         Height          =   225
         Index           =   1
         Left            =   285
         TabIndex        =   4
         Top             =   615
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   7080
      Left            =   210
      TabIndex        =   12
      Tag             =   "10114"
      Top             =   1350
      Width           =   10605
      _Version        =   196608
      _ExtentX        =   18706
      _ExtentY        =   12488
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   14737632
      MaxCols         =   29
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis157.frx":0000
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin VB.Frame fraSearchKey 
      BackColor       =   &H00DBE6E6&
      Height          =   1050
      Index           =   0
      Left            =   2025
      TabIndex        =   15
      Top             =   -45
      Width           =   6735
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1095
         TabIndex        =   19
         Top             =   210
         Width           =   1935
      End
      Begin VB.ComboBox cboOrdDate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Lis157.frx":43DF
         Left            =   4080
         List            =   "Lis157.frx":43E1
         Style           =   2  '드롭다운 목록
         TabIndex        =   18
         Top             =   600
         Width           =   1860
      End
      Begin VB.OptionButton optDuration 
         BackColor       =   &H00DBE6E6&
         Caption         =   "최근 1개월"
         Height          =   285
         Index           =   0
         Left            =   3300
         TabIndex        =   17
         Top             =   240
         Width           =   1170
      End
      Begin VB.OptionButton optDuration 
         BackColor       =   &H00DBE6E6&
         Caption         =   "기간제한없슴"
         Height          =   285
         Index           =   1
         Left            =   4575
         TabIndex        =   16
         Top             =   240
         Width           =   1410
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   1095
         TabIndex        =   20
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         BackColor       =   15597309
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "환자 ID"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   6
         Left            =   90
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "성    명"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   1
         Left            =   3105
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "처 방 일"
         Appearance      =   0
      End
      Begin VB.Label lblOrdDtCnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6015
         TabIndex        =   21
         Top             =   705
         Width           =   90
      End
   End
   Begin VB.Frame fraSearchKey 
      BackColor       =   &H00DBE6E6&
      Height          =   1050
      Index           =   1
      Left            =   2025
      TabIndex        =   2
      Top             =   -45
      Width           =   6735
      Begin VB.TextBox txtSpcNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1725
         TabIndex        =   31
         Top             =   225
         Width           =   2025
      End
      Begin VB.TextBox txtWorkArea 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3795
         MaxLength       =   2
         TabIndex        =   30
         Top             =   225
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAccDt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4620
         MaxLength       =   8
         TabIndex        =   29
         Top             =   225
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtAccNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5895
         TabIndex        =   28
         Top             =   225
         Visible         =   0   'False
         Width           =   705
      End
      Begin MedControls1.LisLabel lblPtNm1 
         Height          =   330
         Left            =   1725
         TabIndex        =   6
         Top             =   600
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   582
         BackColor       =   15597309
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   ""
         Appearance      =   0
         LeftGab         =   100
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   345
         Index           =   3
         Left            =   645
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "검체번호"
         Appearance      =   0
      End
      Begin MedControls1.LisLabel LisLabel4 
         Height          =   330
         Index           =   4
         Left            =   645
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         BackColor       =   10392451
         ForeColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Caption         =   "성    명"
         Appearance      =   0
      End
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   5715
         X2              =   5865
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   4425
         X2              =   4575
         Y1              =   390
         Y2              =   390
      End
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "☞ 처방내역을 검색중입니다..."
      ForeColor       =   &H00553755&
      Height          =   270
      Left            =   2100
      TabIndex        =   14
      Top             =   1080
      Width           =   8085
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   210
      Shape           =   4  '둥근 사각형
      Top             =   1020
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CCFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      Height          =   300
      Left            =   2040
      Shape           =   4  '둥근 사각형
      Top             =   1020
      Width           =   8775
   End
End
Attribute VB_Name = "frm157BarReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objPatient  As clsPatient
Private objSql      As clsLISSqlStatement

Private tmpAccDt    As String
Private OrdFg       As Boolean
Private ClearFg     As Boolean
Private SelFg       As Boolean
Private blnInitFg   As Boolean
Private PtFg        As Boolean
Private SelAllFg    As Boolean

Public Event FormClose()

Private Sub cboOrdDate_Click()

    If txtPtId.Text = "" Then
        txtPtId.SetFocus
        Exit Sub
    End If
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    If Screen.ActiveControl.Name = optSearchKey(0).Name Then Exit Sub
      
    MouseRunning
    lblMessage.Caption = lblPtNm.Caption & " 님의 처방내역을 조회중입니다.."
    Call DisplayOrder
    lblMessage.Caption = ""
    MouseDefault
    cmdReprint.Enabled = True
    If OrdFg Then
        tblOrdSheet.SetFocus
    Else
        cmdReprint.Enabled = False
        txtPtId.SetFocus
        Call txtPtId_GotFocus
    End If

End Sub

Private Sub chkSelAll_Click()
   
    Dim i As Integer
    
    SelFg = True
        With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            .Value = chkSelAll.Value
        Next
    End With
    SelFg = False
 
End Sub

Private Sub cmdClear_Click()
    Call ClearRtn
    If optSearchKey(0).Value Then
        txtPtId.Text = ""
        txtPtId.SetFocus
    Else
        txtWorkArea.Text = ""
        txtAccDt.Text = ""
        txtAccNo.Text = ""
'        txtWorkArea.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set objPatient = Nothing
    Set objSql = Nothing
    RaiseEvent FormClose
End Sub

Private Sub cmdReprint_Click()
    Dim MyBar               As clsBarcode
    Dim tmpLabNo            As Variant
    Dim TestNames           As String
    Dim BarBuffer(1 To 15)  As String
    Dim AccFg               As Boolean
    Dim i                   As Long
    
    TestNames = ""
    
    MouseRunning
    lblMessage.Caption = " Barcode Label을 출력중입니다."
    
    Set MyBar = New clsBarcode
'    Set MyBar.MyDB = dbconn
    Set MyBar.TableInfo = New clsTables
    Set MyBar.FieldInfo = New clsFields
    
    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            If .Value = 1 Then
                Call .GetText(7, i + 1, tmpLabNo)
                .Col = 7
                If .Value <> tmpLabNo Then
                    Erase BarBuffer
                    .Col = 20:  BarBuffer(1) = LABName
                                
                    .Col = 18:  TestNames = TestNames & .Value & ","
                    .Col = 13:  BarBuffer(2) = .Value           'WorkArea
                    .Col = 16:  BarBuffer(3) = Mid(.Value, 3)   'AccDt
                    .Col = 14:  BarBuffer(4) = .Value           'AccSeq
                    .Col = 19:  BarBuffer(5) = .Value           'SpcNo
                                BarBuffer(6) = objPatient.ptid             '환자ID
                                BarBuffer(7) = Trim(objPatient.PtNm)  '환자명
                    .Col = 12:  BarBuffer(8) = .Value           '검체명
                    .Col = 15:  BarBuffer(9) = .Value           '보관코드
                    .Col = 17:  BarBuffer(10) = .Value           'StatFg
                    .Col = 27:
                            If .Value = "" Then                 '진료과코드
                                  .Col = 22: BarBuffer(11) = .Value
                            Else
                                BarBuffer(11) = .Value        '병동ID
'                                .Col = 21
'                                If Trim(.Value) <> "" Then
'                                    BarBuffer(11) = BarBuffer(11) & "/" & .Value
'                                End If
                            End If
                            
                    '-- 변경 희망채혈일시 -> 실제채혈일시 By M.G.Choi 2006.01.19
                    '----------------------------------------------------------------------------
                    .Col = 28: BarBuffer(12) = .Value 'Mid(strColDt, 5, 2) & "/" & Mid(strColDt, 7, 2)
                    .Col = 29: BarBuffer(13) = Format(.Value, "0#:0#") 'strColTm
                    '----------------------------------------------------------------------------
                    
                    '** 원본 --------------------------------------------------------------------
'                    .Col = 8:  BarBuffer(12) = Mid(.Value, 5, 2) & "/" & Mid(.Value, 7, 2)      '처방일
'                    .Col = 24: BarBuffer(13) = .Value           '희망채혈일시
                    '----------------------------------------------------------------------------
                     BarBuffer(14) = TestNames                  '검사명
                     BarBuffer(15) = txtLabelCnt.Text           '라벨출력장수
                    .Col = 23: AccFg = IIf(.Value >= enStsCd.StsCd_LIS_Accession, True, False)  'Status
                    Call MyBar.Label_PrintOut(BarBuffer(1), BarBuffer(2), BarBuffer(3), BarBuffer(4), BarBuffer(5), BarBuffer(6), _
                                              BarBuffer(7), BarBuffer(8), BarBuffer(9), BarBuffer(10), BarBuffer(11), _
                                              BarBuffer(12), BarBuffer(13), BarBuffer(14), BarBuffer(15), AccFg)
'                    Call medSleep(1000)
                    TestNames = ""
                Else
                    .Col = 18
                    TestNames = TestNames & .Value & ","
                End If
            End If
        Next
    End With
    Set MyBar = Nothing
   
    Call cmdClear_Click
    MouseDefault
    lblMessage.Caption = ""
            
End Sub

Private Sub Form_Activate()
    If blnInitFg Then Exit Sub
    optSearchKey(0).Value = True
    Call optSearchKey_Click(0)
    blnInitFg = True
End Sub

Private Sub Form_Load()
    PtFg = False
    SelFg = False
    cboOrdDate.Clear
    blnInitFg = False
    lblOrdDtCnt.Caption = ""
    ClearFg = True
    Set objPatient = New clsPatient
    Set objSql = New clsLISSqlStatement

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
    Set objPatient = Nothing
    Set objSql = Nothing
End Sub

Private Sub optSearchKey_Click(Index As Integer)
   
    If Not Me.Visible Then Exit Sub
    
    If Index = 0 Then
        optSearchKey(0).ForeColor = vbBlue
        optSearchKey(1).ForeColor = vbBlack
        fraSearchKey(0).Visible = True
        fraSearchKey(1).Visible = False
        txtSpcNo.Text = ""
        txtWorkArea.Text = ""
        txtAccDt.Text = ""
        txtAccNo.Text = ""
        txtPtId.SetFocus
    Else
        optSearchKey(0).ForeColor = vbBlack
        optSearchKey(1).ForeColor = vbBlue
        fraSearchKey(0).Visible = False
        fraSearchKey(1).Visible = True
        txtPtId.Text = ""
        txtSpcNo.Text = "": txtSpcNo.SetFocus
'        txtWorkArea.SetFocus
    End If
   
    Call ClearRtn
   
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
   
    Dim i As Long
    Dim SvLabNo As String
    Dim SvButtonVal As Integer
    
    If Col <> 1 Then Exit Sub
    If SelFg Then Exit Sub
   
    With tblOrdSheet
        .Row = Row
        .Col = 1:  SvButtonVal = .Value
        .Col = 7:  SvLabNo = Trim(.Value)
        For i = 1 To .DataRowCnt
            If i <> Row Then
                .Row = i
                .Col = 7
                If Trim(.Value) = SvLabNo Then
                    .Col = 1
                    If .Value <> SvButtonVal Then .Value = SvButtonVal
                End If
            End If
        Next
    End With
   
End Sub

Private Sub cboOrdDate_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        tblOrdSheet.SetFocus
    End If

End Sub


'% 환자ID가 변경되면 화면Clear
Private Sub txtPtId_Change()
    If Not ClearFg Then Call ClearRtn
End Sub

'% 환자 ID
Private Sub txtPtId_GotFocus()
    With txtPtId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'% 환자정보 검색
Private Sub txtPtId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cboOrdDate.SetFocus
    End If
End Sub

Private Sub txtPtId_LostFocus()
      
    If txtPtId.Text = "" Then Exit Sub
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If Screen.ActiveControl.Name = optSearchKey(0).Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    
    If IsNumeric(txtPtId.Text) Then txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    
    Set objPatient = Nothing
    Set objPatient = New clsPatient
    
    With objPatient
'        Call .ClearData   '클래스 내 변수 초기화
        If .GETPatient(txtPtId.Text) Then
'        If .PtntQuery(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm         '성명
            PtFg = True
            ClearFg = False
            If Not LoadOrderDate Then
                MsgBox objPatient.PtNm & " 님의 처방내역이 없습니다"
                txtPtId.Text = ""
                txtPtId.SetFocus
                Call txtPtId_GotFocus
                Exit Sub
            End If
        Else
            MsgBox "등록되지 않은 환자ID입니다.. 다시 입력하세요.."
            txtPtId.Text = ""
            ClearFg = True
            PtFg = False
            txtPtId.SetFocus
            Call txtPtId_GotFocus
            Exit Sub
        End If
    End With
    cboOrdDate.SetFocus

End Sub


Private Sub txtSpcNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtSpcNo.Text = "" Then Exit Sub
        Call GetSpcDataQuery
    End If
End Sub

Private Sub GetSpcDataQuery()
    Dim strSpcYY As String
    Dim strSpcNO As String
    Dim strWA    As String
    Dim strAccDt As String
    Dim strAccNo As String
    Dim ii       As Integer
    
    strSpcYY = Mid(txtSpcNo.Text, 1, 2)
    strSpcNO = Mid(txtSpcNo.Text, 3)
    txtWorkArea.Text = "": txtAccDt.Text = "": txtAccNo.Text = "": tmpAccDt = ""
    Call objSql.GetLabNo(strSpcYY, strSpcNO, strWA, strAccDt, strAccNo)
    If strWA = "" Then
        MsgBox "해당검체에 대한 정보가 없거나 임상병리 처방이 아닙니다." & _
               "확인후 출력하십시요.", vbInformation + vbOKOnly, "Info"
        txtSpcNo.Text = ""
        If txtSpcNo.Enabled Then txtSpcNo.SetFocus
        Exit Sub
    End If
    
    txtWorkArea.Text = strWA
    txtAccDt.Text = strAccDt: tmpAccDt = strAccDt
    txtAccNo.Text = strAccNo
    
    Call MouseRunning
    lblMessage.Caption = "접수번호 " & txtWorkArea.Text & "-" & txtAccDt.Text & "-" & txtAccNo.Text & " 를 조회중입니다.."
    Call DisplayOrder
    lblMessage.Caption = ""
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1: .Value = 1
        Next
    End With
    cmdReprint.Enabled = True
    Call cmdReprint_Click
    txtSpcNo.Text = "": txtSpcNo.SetFocus
    cmdReprint.Enabled = False
    Call MouseDefault

End Sub


'% 검색한 처방을 테이블에 디스플레이 한다.
Private Sub DisplayOrder()
   
    Dim i           As Integer
    Dim SqlStmt     As String
    Dim Rs          As Recordset
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
   
    DoEvents
   
    ' 처방내역 검색
    tmpDate = Format(GetSystemDate, CS_DateDbFormat)
    tmpTime = Format(GetSystemDate, CS_TimeDbFormat)
    
    If optSearchKey(0) Then
        SqlStmt = objSql.SqlWardBarReprint(1, txtPtId.Text, Format(cboOrdDate.Text, CS_DateDbFormat)) ', strOrdDiv)
    Else
        SqlStmt = objSql.SqlBarReprint(2, txtWorkArea.Text, tmpAccDt, txtAccNo.Text)
    End If
   
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        If optSearchKey(0) Then
            MsgBox objPatient.PtNm & " 님의 처방내역이 없습니다"
        Else
            MsgBox "해당 접수내역이 없습니다"
        End If
        txtPtId.Text = ""
        Call ClearRtn
        GoTo Nodata
    End If
    
    With tblOrdSheet
      
        .ReDraw = False
        .MaxRows = 0
        If Rs.RecordCount < 20 Then
            .MaxRows = 20
            .Row = Rs.RecordCount + 1
            .Row2 = 20
            .Col = 1: .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .Protect = True
            .BlockMode = False
        Else
            .MaxRows = Rs.RecordCount + 1  '데이타 건수
        End If
        .RowHeight(-1) = 13
        
        'Locking Cells
        .Row = -1
        .Col = 2: .Col2 = .MaxCols
        .BlockMode = True
        .Lock = True
        .Protect = True
        .BlockMode = False
'이건 왜 있는건지.. 나중에 실행해보구 없어도 되는거면 과감히 삭제..
        objPatient.ptid = Trim("" & Rs.Fields("PtId").Value)
        objPatient.PtNm = GetPtNm(objPatient.ptid)    ' Trim("" & rs.Fields("PtNm").Value)
        lblPtNm1.Caption = objPatient.PtNm
'        objPatient.WardID = Trim("" & Rs.Fields("HosilId").Value)
        Call ICSPatientMark(Trim("" & Rs.Fields("PtId").Value), enICSNum.LIS_ALL)
        For i = 1 To Rs.RecordCount
            lblMessage.Caption = lblMessage.Caption & "."
            DoEvents
         
            .Row = i
            .Col = 1
            If optSearchKey(0).Value Then
               .Value = 0
            Else
               .Value = 1
            End If

            .Row = i
            If SvOrdDt <> Trim("" & Rs.Fields("OrdDt").Value) Then
                .Col = 2: .Value = Format("" & Rs.Fields("OrdDt").Value, CS_DateMask)    '처방일
                .Col = 3: .Value = Trim("" & Rs.Fields("OrdNo").Value)                   '처방번호
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)                   '검체
                SvOrdDt = Trim("" & Rs.Fields("OrdDt").Value)
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)                            '처방번호
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)                            '검체
            End If
            If SvOrdNo <> Trim("" & Rs.Fields("OrdNo").Value) Then
                .Col = 3: .Value = Trim("" & Rs.Fields("OrdNo").Value)                   '처방번호
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)                   '검체
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)                            '처방번호
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)                            '검체
            End If
            If SvSpcNm <> Trim("" & Rs.Fields("SpcNm").Value) Then
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)                   '검체
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)
            End If
         
            .Col = 4: .Value = Trim("" & Rs.Fields("TestNm").Value)                      '처방명
                         .ForeColor = DCM_LightBlue                                         '약간 파란색
            .Col = 6: .Value = Choose(Val("" & Rs.Fields("StatFg").Value) + 1, "", "Y")  '응급여부
                         .ForeColor = DCM_Red                                               '빨간색
            .Col = 7: .Value = Trim("" & Rs.Fields("LabNo").Value)       'LabNo
            .Col = 8: .Value = Trim("" & Rs.Fields("OrdDt").Value)       '처방일
            .Col = 9: .Value = Trim("" & Rs.Fields("OrdNo").Value)       '처방번호
            .Col = 10: .Value = Trim("" & Rs.Fields("OrdSeq").Value)     '처방Seq
            .Col = 11: .Value = Trim("" & Rs.Fields("OrdCd").Value)      '검사코드
            .Col = 12: .Value = Trim("" & Rs.Fields("SpcNm").Value)      '검체명
            .Col = 13: .Value = Trim("" & Rs.Fields("WorkArea").Value)   'WorkArea
            .Col = 14: .Value = Trim("" & Rs.Fields("AccSeq").Value)     'AccSeq
            .Col = 15: .Value = Trim("" & Rs.Fields("StoreCd").Value)    '보관코드
            .Col = 16: .Value = Trim("" & Rs.Fields("AccDt").Value)      'AccDt  채혈일
            .Col = 17: .Value = Trim("" & Rs.Fields("StatFg").Value)     '응급여부
            .Col = 18: .Value = Trim("" & Rs.Fields("AbbrNm5").Value)    '약어명
            .Col = 19: .Value = Trim("" & Rs.Fields("SpcYy").Value) & _
                                Format(Val(Rs.Fields("SpcNo").Value), CS_BarFormat)  '검체번호
            .Col = 20: .Value = Trim("" & Rs.Fields("BuildNm").Value)    '건물명
            .Col = 21: .Value = Trim("" & Rs.Fields("HosilId").Value)    '호실코드
            .Col = 22: .Value = Trim("" & Rs.Fields("DeptCd").Value)     '진료과코드
            .Col = 23: .Value = Trim("" & Rs.Fields("StsCd").Value)      'status
            .Col = 24: .Value = Mid(Trim("" & Rs.Fields("ReqTm").Value), 1, 2) & ":" & _
                                Mid(Trim("" & Rs.Fields("ReqTm").Value), 3, 2)       '희망채혈일시
            .Col = 27: .Value = Trim("" & Rs.Fields("WardId").Value)     '병동코드
            
            '-- 변경 희망채혈일시 -> 실제채혈일시 By M.G.Choi 2006.01.19
            .Col = 28: .Value = Trim("" & Rs.Fields("coldt").Value)      '채혈일자
            .Col = 29: .Value = Mid(Trim("" & Rs.Fields("coltm").Value), 1, 4)    '채혈시간
            
            Rs.MoveNext
        Next
        .ReDraw = True
      
    End With
    cmdReprint.Enabled = True
    OrdFg = True
    ClearFg = False
   
Nodata:
    Set Rs = Nothing
   
End Sub

Private Sub ClearRtn()
    With tblOrdSheet
        .Row = -1
        .Col = -1
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
    End With
    txtLabelCnt.Text = "1"
    cboOrdDate.Clear
    lblPtNm.Caption = ""
    lblPtNm1.Caption = ""
    lblOrdDtCnt.Caption = ""
    
    cmdReprint.Enabled = False
    OrdFg = False
    Set objPatient = Nothing
    Set objPatient = New clsPatient
'    Set objPatient.objDb = dbconn
    
    SelFg = False
    ClearFg = True
    lblMessage.Caption = ""
    Call ICSPatientMark
   
End Sub

Public Function LoadOrderDate() As Boolean
    
    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim strOrdDiv As String
    
    objSql.OrderDate = Format(GetSystemDate, "yyyymmdd")
    Set Rs = New Recordset
    Rs.Open objSql.SqlGetOrdDateForBarprint(txtPtId.Text, LIS_ORDDIV, optDuration(0).Value), DBConn
    
    If Rs.EOF Then
        LoadOrderDate = False
    Else
        LoadOrderDate = True
        cboOrdDate.Clear
        While (Not Rs.EOF)
            cboOrdDate.AddItem Format(Rs.Fields("orddt").Value, CS_DateMask)
            Rs.MoveNext
        Wend
        If cboOrdDate.ListCount > 1 Then
            lblOrdDtCnt.Caption = CStr(cboOrdDate.ListCount)
        Else
            lblOrdDtCnt.Caption = ""
        End If
        cboOrdDate.ListIndex = 0
    End If
    Set Rs = Nothing
End Function



