VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Begin VB.Form frm160WardBarReprint 
   BackColor       =   &H00DBE6E6&
   Caption         =   "병동/외래 Barcode Label 재출력"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10905
   ControlBox      =   0   'False
   Icon            =   "Lis160.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   10905
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Index           =   2
      Left            =   8760
      TabIndex        =   20
      Top             =   45
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "출력장수"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Index           =   1
      Left            =   2295
      TabIndex        =   19
      Top             =   45
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "검색조건"
      LeftGab         =   100
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   300
      Index           =   0
      Left            =   75
      TabIndex        =   18
      Top             =   45
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   529
      BackColor       =   8388608
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "처방구분"
      LeftGab         =   100
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1200
      Left            =   75
      TabIndex        =   11
      Top             =   270
      Width           =   2220
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "해부병리"
         Height          =   255
         Index           =   3
         Left            =   900
         TabIndex        =   24
         Tag             =   "A"
         Top             =   765
         Width           =   1170
      End
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "전체"
         ForeColor       =   &H004A4189&
         Height          =   720
         Index           =   0
         Left            =   135
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   240
         Width           =   630
      End
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "혈액은행"
         Height          =   255
         Index           =   2
         Left            =   900
         TabIndex        =   21
         Tag             =   "B"
         Top             =   510
         Width           =   1170
      End
      Begin VB.OptionButton optSearchKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "임상병리"
         Height          =   270
         Index           =   1
         Left            =   900
         TabIndex        =   12
         Tag             =   "L"
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdReprint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "재출력(&S)"
      Height          =   510
      Left            =   6855
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "종료(&X)"
      Height          =   510
      Left            =   9510
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   8505
      Width           =   1320
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "화면지움(&C)"
      Height          =   510
      Left            =   8175
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8505
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBE6E6&
      Height          =   1200
      Left            =   8760
      TabIndex        =   8
      Top             =   270
      Width           =   2085
      Begin VB.TextBox txtLabelCnt 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   285
         TabIndex        =   2
         Text            =   "1"
         Top             =   450
         Width           =   690
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   975
         TabIndex        =   9
         Top             =   435
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtLabelCnt"
         BuddyDispid     =   196615
         OrigLeft        =   3840
         OrigTop         =   330
         OrigRight       =   4080
         OrigBottom      =   645
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "장"
         Height          =   180
         Left            =   1275
         TabIndex        =   10
         Tag             =   "151"
         Top             =   510
         Width           =   195
      End
   End
   Begin VB.CheckBox chkSelAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "전체선택(&A)"
      ForeColor       =   &H00553755&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1500
      Width           =   1350
   End
   Begin FPSpread.vaSpread tblOrdSheet 
      Height          =   6570
      Left            =   75
      TabIndex        =   3
      Tag             =   "10114"
      Top             =   1800
      Width           =   10755
      _Version        =   196608
      _ExtentX        =   18971
      _ExtentY        =   11589
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
      MaxCols         =   28
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "Lis160.frx":08CA
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleCols     =   5
      VisibleRows     =   500
   End
   Begin VB.Frame fraSearchKey 
      BackColor       =   &H00DBE6E6&
      Height          =   1200
      Index           =   0
      Left            =   2295
      TabIndex        =   13
      Top             =   270
      Width           =   6450
      Begin VB.OptionButton optDuration 
         BackColor       =   &H00DBE6E6&
         Caption         =   "기간제한없슴"
         Height          =   285
         Index           =   1
         Left            =   4575
         TabIndex        =   23
         Top             =   315
         Width           =   1410
      End
      Begin VB.OptionButton optDuration 
         BackColor       =   &H00DBE6E6&
         Caption         =   "최근 1개월"
         Height          =   285
         Index           =   0
         Left            =   3300
         TabIndex        =   22
         Top             =   315
         Width           =   1170
      End
      Begin VB.ComboBox cboOrdDate 
         Height          =   300
         ItemData        =   "Lis160.frx":4C60
         Left            =   4080
         List            =   "Lis160.frx":4C6A
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   705
         Width           =   1860
      End
      Begin VB.TextBox txtPtId 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1095
         TabIndex        =   0
         Top             =   270
         Width           =   1935
      End
      Begin MedControls1.LisLabel lblPtNm 
         Height          =   330
         Left            =   1095
         TabIndex        =   14
         Top             =   660
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
         Left            =   75
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   270
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
         Left            =   75
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   660
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   675
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
         TabIndex        =   15
         Top             =   780
         Width           =   90
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   1470
      Width           =   2205
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "☞ 처방내역을 검색중입니다..."
      ForeColor       =   &H00553755&
      Height          =   270
      Left            =   2415
      TabIndex        =   17
      Top             =   1530
      Width           =   8085
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CCFFFF&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      Height          =   300
      Left            =   2295
      Shape           =   4  '둥근 사각형
      Top             =   1470
      Width           =   8550
   End
End
Attribute VB_Name = "frm160WardBarReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsFirst As Boolean

Private MyPatient   As clsPatient
Private MySql       As clsLISSqlStatement

Private tmpAccDt    As String
Private OrdFg       As Boolean
Private ClearFg     As Boolean
Private SelFg       As Boolean
Private MsgFg       As Boolean

Private PtFg        As Boolean
Private SelAllFg    As Boolean

Public Event LastFormUnload()
Public Event ThisFormUnload()

Private Sub cboOrdDate_Click()

    If txtPtId.Text = "" Then
       txtPtId.SetFocus
       Exit Sub
    End If
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
      
    MouseRunning
    
    lblMessage.Caption = "▶  " & lblPtNm.Caption & " 님의 처방내역을 조회중입니다.."
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
    txtPtId.Text = ""
    txtPtId.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set MyPatient = Nothing
    Set MySql = Nothing
    
    RaiseEvent ThisFormUnload
End Sub
Private Function BarPrint_Check() As Boolean
    Dim ii As Integer
    
    With tblOrdSheet
        For ii = 1 To .DataRowCnt
            .Row = ii: .Col = 1
            If .Value = 1 Then
                BarPrint_Check = True
                Exit For
            End If
        Next
    End With
End Function
Private Sub cmdReprint_Click()
    Dim MyBar              As clsBarcode
    Dim tmpLabNo           As Variant
    Dim TestNames          As String
    Dim BarBuffer(0 To 15) As String
    Dim strStatFg          As String
    Dim strSpcNO           As String
    Dim AccFg              As Boolean
    Dim FzFg               As Boolean
    Dim i                  As Long
    
    If BarPrint_Check = False Then
        MsgBox "출력대상을 선택한후 재출력 버튼을 클릭하세요.", vbInformation + vbOKOnly, "출력대상선택"
        Exit Sub
    End If
    
    Set MyBar = New clsBarcode
'    Set MyBar.MyDB = dbconn
    Set MyBar.TableInfo = New clsTables
    Set MyBar.FieldInfo = New clsFields
    TestNames = ""
    
    Screen.MousePointer = vbArrowHourglass
    lblMessage.Caption = " Barcode Label을 출력중입니다."
    With tblOrdSheet
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = 1
            If .Value = 1 Then
                Call .GetText(7, i + 1, tmpLabNo)
                .Col = 7
                If .Value <> tmpLabNo Then
                    Erase BarBuffer
                    .Col = 18:  TestNames = TestNames & .Value & ","
                    .Col = 25:  BarBuffer(0) = .Value                   '처방구분
                    .Col = 20:
                                Select Case BarBuffer(0)
                                    Case LIS_ORDDIV: BarBuffer(1) = LABName
                                    Case BBS_ORDDIV: BarBuffer(1) = BBSName
                                    Case APS_ORDDIV: BarBuffer(1) = APSName
                                End Select
                    .Col = 13:
                                Select Case BarBuffer(0)
                                    Case LIS_ORDDIV: BarBuffer(2) = .Value                   'WorkArea
                                    Case BBS_ORDDIV: BarBuffer(2) = BBSBarNm
                                    Case APS_ORDDIV: BarBuffer(2) = APSBarNm
                                End Select
                    .Col = 16:  BarBuffer(3) = Mid(.Value, 3)           'AccDt
                    .Col = 14:  Select Case BarBuffer(0)
                                Case BBS_ORDDIV:
                                    .Col = 7
                                    BarBuffer(4) = Format(.Value, String(11, "@"))
                                Case Else:
                                    .Col = 14
                                    BarBuffer(4) = IIf(.Value = "0", "", Format(.Value, String(4, "@")))    'SpcNo
                                End Select
                    .Col = 19:  BarBuffer(5) = .Value                   'SpcNo
                                BarBuffer(6) = MyPatient.ptid           '환자ID
                                BarBuffer(7) = Trim(MyPatient.PtNm)
                    .Col = 12:  BarBuffer(8) = .Value                   '검체명
                    .Col = 15:  BarBuffer(9) = .Value                   '보관코드
                    .Col = 17:
                                If BarBuffer(5) = strSpcNO Then
                                    BarBuffer(10) = IIf(strStatFg = "1", strStatFg, .Value) 'StatFg 구분
                                Else
                                    BarBuffer(10) = .Value
                                End If
                    .Col = 27:
                                If .Value = "" Then
                                    .Col = 22: BarBuffer(11) = .Value   '진료과코드
                                    If BarBuffer(11) <> "" Then
'                                        .Col = 21
'                                        If .Value <> "" Then
'                                            BarBuffer(11) = BarBuffer(11) & "/" & .Value
'                                        End If
                                    Else
                                        .Col = 21
                                        If .Value <> "" Then
                                            BarBuffer(11) = .Value
                                        End If
                                    End If
                                Else
                                    BarBuffer(11) = .Value              '병동ID
'                                    .Col = 21
'                                    If .Value <> "" Then
'                                        BarBuffer(11) = BarBuffer(11) & "/" & .Value
'                                    End If
                                End If
                    .Col = 8:   BarBuffer(12) = Mid(.Value, 5, 2) & "/" & _
                                                Mid(.Value, 7, 2)       '처방일
                    .Col = 24:  BarBuffer(13) = .Value                  '희망채혈일시
                                BarBuffer(14) = TestNames               '검사명
                                BarBuffer(15) = txtLabelCnt.Text        '라벨출력장수
                    .Col = 23:
                                AccFg = IIf(Val(.Value) >= 2, True, False)
                    .Col = 26:
                                FzFg = IIf(.Value = "1", True, False)
                    MyBar.WardTmp = ""
                    Call MyBar.Label_PrintOut( _
                            BarBuffer(1), BarBuffer(2), BarBuffer(3), _
                            BarBuffer(4), BarBuffer(5), BarBuffer(6), BarBuffer(7), _
                            BarBuffer(8), BarBuffer(9), BarBuffer(10), BarBuffer(11), _
                            BarBuffer(12), BarBuffer(13), BarBuffer(14), BarBuffer(15), _
                            AccFg, FzFg)
                    
'                    Call medSleep(2000)
                    TestNames = ""
                Else
                    .Col = 17
                        If .Value = "1" Then
                            .Col = 19: strSpcNO = .Value
                            strStatFg = "1"
                        End If
                    .Col = 18
                    TestNames = TestNames & .Value & ","
                End If
            End If
        Next
    End With
   
    Call ClearRtn
    
    Screen.MousePointer = vbDefault
    lblMessage.Caption = ""
    txtPtId.Text = ""
    txtPtId.SetFocus
    Set MyBar = Nothing
End Sub

Private Sub Form_Activate()
    If Not IsFirst Then Exit Sub
    IsFirst = False
    PtFg = False
    SelFg = False
    cboOrdDate.Clear
    optSearchKey(0).Value = True
    optDuration(0).Value = True
    lblOrdDtCnt.Caption = ""
    ClearFg = True
End Sub

Private Sub Form_Load()
    IsFirst = True
    Set MyPatient = New clsPatient
    Set MySql = New clsLISSqlStatement
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ICSPatientMark
End Sub

Private Sub optSearchKey_Click(Index As Integer)
   
On Error GoTo Err_Trap
    
    optSearchKey(0).ForeColor = vbBlue
    optSearchKey(1).ForeColor = vbBlack
    fraSearchKey(0).Visible = True
    fraSearchKey(1).Visible = False

Err_Trap:
    Call ClearRtn
    If txtPtId.Text = "" Then txtPtId.SetFocus
   
End Sub

Private Sub tblOrdSheet_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
   
    Dim i           As Long
    Dim SvLabNo     As String
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
        Call ICSPatientMark(txtPtId.Text, enICSNum.LIS_ALL)
        cboOrdDate.SetFocus
    End If
End Sub

Private Sub txtPtId_LostFocus()
      
    If txtPtId.Text = "" Then Exit Sub
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    If Screen.ActiveControl.Name = optSearchKey(0).Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdExit.Name Then Exit Sub
    If Screen.ActiveControl.Name = cmdClear.Name Then Exit Sub
    
    If IsNumeric(txtPtId.Text) Then
        txtPtId.Text = Format(txtPtId.Text, P_PatientIdFormat)
    End If
    
    Set MyPatient = Nothing
    Set MyPatient = New clsPatient
    
    With MyPatient
'        Call .ClearData   '클래스 내 변수 초기화
'        If .PtntQuery(txtPtId.Text) Then
        If .GETPatient(txtPtId.Text) Then
            lblPtNm.Caption = .PtNm         '성명
            PtFg = True
            ClearFg = False
            If Not LoadOrderDate Then
                MsgFg = True
                MsgBox MyPatient.PtNm & " 님의 처방내역이 없습니다"
                txtPtId.Text = ""
                txtPtId.SetFocus
                MsgFg = False
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
End Sub


'% 검색한 처방을 테이블에 디스플레이 한다.
Private Sub DisplayOrder()
    Dim Rs As Recordset
    Dim i           As Integer
    Dim SqlStmt     As String
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim strAccdt    As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim strOrdDiv   As String
    
    DoEvents
     
    ' 처방내역 검색
    tmpDate = Format(GetSystemDate, CS_DateDbFormat)
    tmpTime = Format(GetSystemDate, CS_TimeDbFormat)
     
    strOrdDiv = GetOrdDiv
         

    SqlStmt = MySql.SqlWardBarReprint(1, txtPtId.Text, Format(cboOrdDate.Text, CS_DateDbFormat), strOrdDiv)
   
    Set Rs = New Recordset
    Rs.Open SqlStmt, DBConn
    
    If Rs.EOF Then
        MsgBox MyPatient.PtNm & " 님의 처방내역이 없습니다", vbInformation, "간호사채혈"
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
            
'        MyPatient.ptid = Trim("" & rs.Fields("PtId").Value)
'        MyPatient.PtNm = Trim("" & rs.Fields("PtNm").Value)
'        MyPatient.ptWardID = Trim("" & rs.Fields("HosilId").Value)
      
        For i = 1 To Rs.RecordCount
            lblMessage.Caption = lblMessage.Caption & "."
            DoEvents
            .Row = i
            .Col = 1: .Value = 0
            .Row = i    '**ButtonClicked 이벤트가 발생하여 Row값이 바뀌므로 다시 한번 셋팅.
            '-- 변경 채혈일
            If strAccdt <> Trim("" & Trim(Rs.Fields("AccDt").Value)) Then
                .Col = 2: .Value = Format("" & Rs.Fields("OrdDt").Value, CS_DateMask)   '처방일
                .Col = 3: .Value = Trim("" & Rs.Fields("OrdNo").Value)   '처방번호
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)   '검체
                strAccdt = Trim("" & Rs.Fields("AccDt").Value)
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)            '처방번호
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)            '검체
            End If
            
            If SvOrdNo <> Trim("" & Rs.Fields("OrdNo").Value) Then
                .Col = 3: .Value = Trim("" & Rs.Fields("OrdNo").Value)   '처방번호
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)   '검체
                SvOrdNo = Trim("" & Rs.Fields("OrdNo").Value)            '처방번호
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)            '검체
            End If
            If SvSpcNm <> Trim("" & Rs.Fields("SpcNm").Value) Then
                .Col = 5: .Value = Trim("" & Rs.Fields("SpcNm").Value)   '검체
                SvSpcNm = Trim("" & Rs.Fields("SpcNm").Value)
            End If
         
            .Col = 4: .Value = Trim("" & Rs.Fields("TestNm").Value)      '처방명
                         '.ForeColor = &HDF6A3E        '약간 파란색
                        Select Case Rs.Fields("orddiv")
                            Case APS_ORDDIV: .ForeColor = &H5E3F00     '&HDF6A3E     '약간 파란색
                            Case BBS_ORDDIV: .ForeColor = &H496835     '&H6C6181     '&H81815A     '약간녹색
                            Case LIS_ORDDIV: .ForeColor = &H553755
                        End Select
            .Col = 6: .Value = Choose(Val("" & Rs.Fields("StatFg").Value) + 1, "", "Y")     '응급여부
                         .ForeColor = &HFF&       '빨간색
            Select Case Trim("" & Rs.Fields("OrdDiv").Value)
                Case APS_ORDDIV, LIS_ORDDIV
                    .Col = 7: .Value = Trim("" & Rs.Fields("LabNo").Value)       'LabNo
                Case BBS_ORDDIV
                    .Col = 7: .Value = Trim("" & Rs.Fields("SpcYy").Value) & "-" & _
                                             Rs.Fields("SpcNo").Value
            End Select
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
                                Format(Val(Rs.Fields("SpcNo").Value), CS_BarFormat)     '검체번호
            .Col = 20: .Value = IIf(P_ApplyBuildingInfo, Trim("" & Rs.Fields("BuildNm").Value), "") '건물명
            .Col = 21: .Value = Trim("" & Rs.Fields("HosilId").Value)    '호실코드
            .Col = 22: .Value = Trim("" & Rs.Fields("DeptCd").Value)     '진료과코드
            .Col = 23: .Value = Trim("" & Rs.Fields("StsCd").Value)      'status
            .Col = 24: .Value = Mid(Trim("" & Rs.Fields("ReqTm").Value), 1, 2) & ":" & _
                                Mid(Trim("" & Rs.Fields("ReqTm").Value), 3, 2)  '희망채혈일시
            .Col = 25: .Value = Trim("" & Rs.Fields("OrdDiv").Value)     'OrdDiv
            .Col = 26: .Value = Trim("" & Rs.Fields("FzFg").Value)       '냉동절편구분
            .Col = 27: .Value = Trim("" & Rs.Fields("wardid").Value)     '병동코드

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


Private Function GetOrdDiv() As String
    
    Dim i As Long
    
    GetOrdDiv = ""
    
    For i = 0 To optSearchKey.Count - 1
        If optSearchKey(i).Value Then
            GetOrdDiv = optSearchKey(i).Tag
            Exit For
        End If
    Next
    
End Function


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
    lblOrdDtCnt.Caption = ""
   
    cmdReprint.Enabled = False
    OrdFg = False
    Set MyPatient = Nothing
    Set MyPatient = New clsPatient
'    Set MyPatient.objDb = dbconn
    
    SelFg = False
    ClearFg = True
    lblMessage.Caption = ""
    chkSelAll.Value = 0
End Sub

Private Function LoadOrderDate() As Boolean

    Dim SqlStmt As String
    Dim Rs As Recordset
    Dim strOrdDiv As String
    
    MySql.OrderDate = Format(GetSystemDate, "yyyymmdd")
    
    Set Rs = New Recordset
    Rs.Open MySql.SqlGetOrdDateForBarprint(txtPtId.Text, GetOrdDiv, optDuration(0).Value), DBConn
    
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

Public Sub Call_PtId_KeyPress()

    Call txtPtId_KeyPress(vbKeyReturn)

End Sub

Public Sub Call_ToDate_LostFocus()

    Call cboOrdDate_KeyDown(vbKeyReturn, 0)
   
End Sub


