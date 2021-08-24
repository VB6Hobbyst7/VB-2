VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FGO0301 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "WorkSheet 출력"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin Threed.SSPanel pnlMain 
      Height          =   3375
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   5953
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin Threed.SSPanel pnlButton 
         Height          =   1125
         Left            =   90
         TabIndex        =   9
         Top             =   2160
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   1984
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand cmdExit 
            Height          =   1005
            Left            =   3690
            TabIndex        =   7
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   1773
            _StockProps     =   78
            Caption         =   "종료 ESC"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   10.5
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   3
            Picture         =   "FGO0301.frx":0000
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   1005
            Left            =   2520
            TabIndex        =   6
            ToolTipText     =   "WorkSheet 출력"
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   1773
            _StockProps     =   78
            Caption         =   "인쇄 F5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   3
            Picture         =   "FGO0301.frx":08DA
         End
      End
      Begin Threed.SSPanel pnlOption 
         Height          =   2055
         Left            =   90
         TabIndex        =   10
         Top             =   90
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   3625
         _StockProps     =   15
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtslipcd 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   2
            Top             =   720
            Width           =   525
         End
         Begin VB.TextBox txtLabSeqS 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   3
            Top             =   1170
            Width           =   765
         End
         Begin VB.TextBox txtLabSeqE 
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2310
            MaxLength       =   5
            TabIndex        =   4
            Top             =   1170
            Width           =   765
         End
         Begin Threed.SSOption optprt 
            Height          =   315
            Index           =   0
            Left            =   1260
            TabIndex        =   5
            Top             =   1620
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "작업번호"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSPanel pnlLabDate 
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "접수일자"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin MSComCtl2.DTPicker dtpLabDateS 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            Top             =   270
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   24444931
            CurrentDate     =   36165
         End
         Begin Threed.SSPanel pnlSlipcd 
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   690
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "SLIP"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin Threed.SSCommand cmdsliph 
            Height          =   330
            Left            =   1800
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   720
            Width           =   270
            _Version        =   65536
            _ExtentX        =   476
            _ExtentY        =   582
            _StockProps     =   78
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RoundedCorners  =   0   'False
            Picture         =   "FGO0301.frx":11B4
         End
         Begin MSComCtl2.DTPicker dtpLabDateE 
            Height          =   315
            Left            =   2910
            TabIndex        =   1
            Top             =   270
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   24444931
            CurrentDate     =   36165
         End
         Begin Threed.SSPanel pnlSpcCd 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   1140
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "작업번호"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin Threed.SSPanel pnlPrtGbn 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1590
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "정렬구분"
            ForeColor       =   8454143
            BackColor       =   8388608
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
         End
         Begin Threed.SSOption optprt 
            Height          =   315
            Index           =   1
            Left            =   2340
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1620
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   78
            Caption         =   "등록번호"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblSlipNm 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  '단일 고정
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
            Left            =   2070
            TabIndex        =   17
            Top             =   720
            Width           =   2325
         End
         Begin VB.Line lineLabSeq 
            BorderWidth     =   2
            X1              =   1980
            X2              =   2340
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line lineLabDate 
            BorderWidth     =   2
            Visible         =   0   'False
            X1              =   2700
            X2              =   2910
            Y1              =   420
            Y2              =   420
         End
      End
   End
End
Attribute VB_Name = "FGO0301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DCJ0101         As DCJ0101
Dim CodeHelp_F      As Integer

Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub cmdPrint_Click()

    Dim DCO0101     As DCO0101
    Dim iRCnt       As Integer
    Dim PrtGbn      As String
    
' 기본 변수 체크
    If Trim(txtLabSeqS.Text) = "" And Trim(txtLabSeqE.Text) = "" Then
        txtLabSeqS.Text = "00001"
        txtLabSeqE.Text = "99999"
    End If
    
    If Trim(txtslipcd.Text) = "" Or Trim(lblSlipNm.Caption) = "" Then
        ViewMsg "SLIP을 선택하여 주십시요"
        txtslipcd.SetFocus
    End If
    
    If IsNumeric(txtLabSeqS.Text) = False Or IsNumeric(txtLabSeqE.Text) = False Then
        ViewMsg "입력하신 검체번호가 잘못되었습니다."
        txtLabSeqS.SetFocus
    End If
    
    If optprt(0).Value = True Then
        PrtGbn = "0"            ' 작업번호순 정렬
    Else
        PrtGbn = "1"            ' 등록번호순 정렬
    End If
    
' 화면 마우스 처리
    Screen.MousePointer = 11
    
    Set DCO0101 = New DCO0101
        iRCnt = DCO0101.Print_WorkSheet(Format(dtpLabDateS.Value, "YYYYMMDD"), txtslipcd.Text, txtLabSeqS.Text, txtLabSeqE.Text, PrtGbn)
    Set DCO0101 = Nothing
    
    If iRCnt = 0 Then
        ViewMsg "출력할 작업대장이 없습니다."
    Else
        ViewMsg Trim(Str(iRCnt)) & " 건의 작업대장이 출력되었습니다."
    End If
    Screen.MousePointer = 1
 
End Sub

Private Sub cmdsliph_Click()

    Dim i%
    Dim j%
    Dim CPart As DCB0101
    Dim sTot01$
    Dim sTot02$
    Dim sTot03$
    
    txtslipcd.SetFocus
    
    Set CPart = New DCB0101
    
    CPart.Get_PART
    
    j = CPart.CurItemCnt
    
    Erase gCodeHlpTable '배열 초기화
    
    ReDim gCodeHlpTable(j) As CodeTBL
    
    With CPart
        sTot01 = .TotField01
        sTot02 = .TotField02
        sTot03 = .TotField03
    End With
    
    Set CPart = Nothing
    
    For i = 1 To j
        gCodeHlpTable(i).sSeq = Format$(i, "00000")
        gCodeHlpTable(i).sCode = GetByOne(sTot01, sTot01) & GetByOne(sTot02, sTot02)
        gCodeHlpTable(i).sCodeNm = GetByOne(sTot03, sTot03)
    Next
    
    giCodeHlpCnt = j
    
    hWndCd = txtslipcd.hwnd
    
    FSO0101.Top = 4200
    FSO0101.Left = 5450
    
    
' Code Help Flag
    CodeHelp_F = True
    
    Load FSO0101
    FSO0101.Show vbModal

End Sub

Private Sub dtpLabDateS_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        txtslipcd.SetFocus
        KeyCode = 0
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Call cmdPrint_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyEscape Then
        Call cmdExit_Click
        KeyCode = 0
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ViewMsg ""
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call InitRegCurFrmTitle

End Sub

Private Sub Form_Activate()

    If CodeHelp_F = False Then
        txtslipcd.Text = fCurUserSlipCd
        lblSlipNm.Caption = fCurUserSlipNm
    End If

End Sub

Private Sub Form_Load()

Dim Sys_Date    As String
    
' 변수 초기화
'    Rnow_row = 1
'    Onow_row = 1
    CodeHelp_F = False
    
' 날짜 초기화
    Set DCJ0101 = New DCJ0101
    Sys_Date = DCJ0101.Get_Date("D")
    dtpLabDateS.Value = Sys_Date
    'dtpLabDateE.Value = Sys_Date
    Set DCJ0101 = Nothing

End Sub

Private Sub txtLabSeqE_GotFocus()
    
    txtLabSeqE.Tag = txtLabSeqE.Text
    Call Txt_Highlight(txtLabSeqE)
    
End Sub

Private Sub txtLabSeqE_LostFocus()

    If Trim(txtLabSeqE.Text) = "" Then
        txtLabSeqE.Text = txtLabSeqS.Text
        Exit Sub
    End If
    
    If txtLabSeqE.Text <> txtLabSeqE.Tag Then
        If Len(txtLabSeqE.Text) < txtLabSeqE.MaxLength Then
            txtLabSeqE.Text = Format(txtLabSeqE.Text, "00000")
        End If
    End If
    
    If Val(txtLabSeqS.Text) > Val(txtLabSeqE.Text) Then
        ViewMsg "검체번호 구간이 잘못되었습니다."
        txtLabSeqS.SetFocus
    End If

End Sub

Private Sub txtLabSeqS_GotFocus()

    txtLabSeqS.Tag = txtLabSeqS.Text
    Call Txt_Highlight(txtLabSeqS)
    
End Sub

Private Sub txtLabSeqS_LostFocus()

    If txtLabSeqS.Text <> txtLabSeqS.Tag Then
        If Len(txtLabSeqS.Text) < txtLabSeqS.MaxLength Then
            txtLabSeqS.Text = Format(txtLabSeqS.Text, "00000")
        End If
    End If

End Sub

Private Sub txtslipcd_Change()

    If Len(txtslipcd.Text) = txtslipcd.MaxLength Then
        
        Set DCJ0101 = New DCJ0101
        
        lblSlipNm.Caption = DCJ0101.Get_SlipNm(txtslipcd.Text)
    
        If lblSlipNm.Caption = "" Then
            ViewMsg "존재하지 않는 Slip Code입니다."
        End If
    
        Set DCJ0101 = Nothing
        
        If CodeHelp_F = False Then
            txtLabSeqS.SetFocus
        Else
            SendKeys "{ENTER}"
        End If
    Else
        lblSlipNm.Caption = ""
    End If

End Sub

Private Sub txtslipcd_GotFocus()

    Call Txt_Highlight(txtslipcd)
    
End Sub

Private Sub txtslipcd_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CodeHelp_F = False
    
End Sub
